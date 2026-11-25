<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonthlyClosing_CyclicBundlingExplain2
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.gbo1 = New System.Windows.Forms.GroupBox
        Me.dgv1 = New System.Windows.Forms.DataGridView
        Me.Col1_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1_CodeNo = New System.Windows.Forms.DataGridViewLinkColumn
        Me.Col1_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col1_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        Me.dgv2 = New System.Windows.Forms.DataGridView
        Me.Col2_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2_CodeNo = New System.Windows.Forms.DataGridViewLinkColumn
        Me.Col2_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col2_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.gbo2 = New System.Windows.Forms.GroupBox
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.gbo1.SuspendLayout()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgv2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbo2.SuspendLayout()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.gbo1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnClose)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.gbo2)
        Me.SplitContainer1.Size = New System.Drawing.Size(882, 653)
        Me.SplitContainer1.SplitterDistance = 305
        Me.SplitContainer1.TabIndex = 0
        '
        'gbo1
        '
        Me.gbo1.Controls.Add(Me.dgv1)
        Me.gbo1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbo1.Location = New System.Drawing.Point(0, 0)
        Me.gbo1.Name = "gbo1"
        Me.gbo1.Size = New System.Drawing.Size(882, 305)
        Me.gbo1.TabIndex = 1
        Me.gbo1.TabStop = False
        Me.gbo1.Text = "Detail Item"
        '
        'dgv1
        '
        Me.dgv1.AllowUserToAddRows = False
        Me.dgv1.AllowUserToDeleteRows = False
        Me.dgv1.AllowUserToResizeColumns = False
        Me.dgv1.AllowUserToResizeRows = False
        Me.dgv1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgv1.ColumnHeadersHeight = 25
        Me.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgv1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1_No, Me.Col1_CodeNo, Me.Col1_Desc, Me.Col1_Qty, Me.Col1_Unit, Me.Col1_Amount, Me.Col1_PrimaryCodeNo})
        Me.dgv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgv1.Location = New System.Drawing.Point(3, 18)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.ReadOnly = True
        Me.dgv1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgv1.RowTemplate.Height = 24
        Me.dgv1.Size = New System.Drawing.Size(876, 284)
        Me.dgv1.TabIndex = 0
        '
        'Col1_No
        '
        Me.Col1_No.HeaderText = "No"
        Me.Col1_No.Name = "Col1_No"
        Me.Col1_No.ReadOnly = True
        Me.Col1_No.Width = 51
        '
        'Col1_CodeNo
        '
        Me.Col1_CodeNo.HeaderText = "Code No"
        Me.Col1_CodeNo.Name = "Col1_CodeNo"
        Me.Col1_CodeNo.ReadOnly = True
        Me.Col1_CodeNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Col1_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Col1_CodeNo.Width = 88
        '
        'Col1_Desc
        '
        Me.Col1_Desc.HeaderText = "Description"
        Me.Col1_Desc.Name = "Col1_Desc"
        Me.Col1_Desc.ReadOnly = True
        Me.Col1_Desc.Width = 104
        '
        'Col1_Qty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col1_Qty.DefaultCellStyle = DataGridViewCellStyle1
        Me.Col1_Qty.HeaderText = "Qty"
        Me.Col1_Qty.Name = "Col1_Qty"
        Me.Col1_Qty.ReadOnly = True
        Me.Col1_Qty.Width = 55
        '
        'Col1_Unit
        '
        Me.Col1_Unit.HeaderText = "Unit"
        Me.Col1_Unit.Name = "Col1_Unit"
        Me.Col1_Unit.ReadOnly = True
        Me.Col1_Unit.Width = 58
        '
        'Col1_Amount
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col1_Amount.DefaultCellStyle = DataGridViewCellStyle2
        Me.Col1_Amount.HeaderText = "Amount"
        Me.Col1_Amount.Name = "Col1_Amount"
        Me.Col1_Amount.ReadOnly = True
        Me.Col1_Amount.Width = 81
        '
        'Col1_PrimaryCodeNo
        '
        Me.Col1_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.Col1_PrimaryCodeNo.Name = "Col1_PrimaryCodeNo"
        Me.Col1_PrimaryCodeNo.ReadOnly = True
        Me.Col1_PrimaryCodeNo.Visible = False
        Me.Col1_PrimaryCodeNo.Width = 140
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(515, 90)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(120, 36)
        Me.btnClose.TabIndex = 0
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'dgv2
        '
        Me.dgv2.AllowUserToAddRows = False
        Me.dgv2.AllowUserToDeleteRows = False
        Me.dgv2.AllowUserToResizeColumns = False
        Me.dgv2.AllowUserToResizeRows = False
        Me.dgv2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgv2.ColumnHeadersHeight = 25
        Me.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgv2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col2_No, Me.Col2_CodeNo, Me.Col2_Desc, Me.Col2_Qty, Me.Col2_Unit, Me.Col2_Amount, Me.Col2_PrimaryCodeNo})
        Me.dgv2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgv2.Location = New System.Drawing.Point(3, 18)
        Me.dgv2.Name = "dgv2"
        Me.dgv2.ReadOnly = True
        Me.dgv2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgv2.RowTemplate.Height = 24
        Me.dgv2.Size = New System.Drawing.Size(876, 323)
        Me.dgv2.TabIndex = 1
        '
        'Col2_No
        '
        Me.Col2_No.HeaderText = "No"
        Me.Col2_No.Name = "Col2_No"
        Me.Col2_No.ReadOnly = True
        Me.Col2_No.Width = 51
        '
        'Col2_CodeNo
        '
        Me.Col2_CodeNo.HeaderText = "Code No"
        Me.Col2_CodeNo.Name = "Col2_CodeNo"
        Me.Col2_CodeNo.ReadOnly = True
        Me.Col2_CodeNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Col2_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Col2_CodeNo.Width = 88
        '
        'Col2_Desc
        '
        Me.Col2_Desc.HeaderText = "Description"
        Me.Col2_Desc.Name = "Col2_Desc"
        Me.Col2_Desc.ReadOnly = True
        Me.Col2_Desc.Width = 104
        '
        'Col2_Qty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col2_Qty.DefaultCellStyle = DataGridViewCellStyle3
        Me.Col2_Qty.HeaderText = "Qty"
        Me.Col2_Qty.Name = "Col2_Qty"
        Me.Col2_Qty.ReadOnly = True
        Me.Col2_Qty.Width = 55
        '
        'Col2_Unit
        '
        Me.Col2_Unit.HeaderText = "Unit"
        Me.Col2_Unit.Name = "Col2_Unit"
        Me.Col2_Unit.ReadOnly = True
        Me.Col2_Unit.Width = 58
        '
        'Col2_Amount
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Col2_Amount.DefaultCellStyle = DataGridViewCellStyle4
        Me.Col2_Amount.HeaderText = "Amount"
        Me.Col2_Amount.Name = "Col2_Amount"
        Me.Col2_Amount.ReadOnly = True
        Me.Col2_Amount.Width = 81
        '
        'Col2_PrimaryCodeNo
        '
        Me.Col2_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.Col2_PrimaryCodeNo.Name = "Col2_PrimaryCodeNo"
        Me.Col2_PrimaryCodeNo.ReadOnly = True
        Me.Col2_PrimaryCodeNo.Visible = False
        Me.Col2_PrimaryCodeNo.Width = 140
        '
        'gbo2
        '
        Me.gbo2.Controls.Add(Me.dgv2)
        Me.gbo2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbo2.Location = New System.Drawing.Point(0, 0)
        Me.gbo2.Name = "gbo2"
        Me.gbo2.Size = New System.Drawing.Size(882, 344)
        Me.gbo2.TabIndex = 2
        Me.gbo2.TabStop = False
        Me.gbo2.Text = "GroupBox1"
        '
        'FrmTransaksiMonthlyClosing_CyclicBundlingExplain2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(882, 653)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 700)
        Me.Name = "FrmTransaksiMonthlyClosing_CyclicBundlingExplain2"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Transaction"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.gbo1.ResumeLayout(False)
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgv2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbo2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents gbo1 As System.Windows.Forms.GroupBox
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
    Friend WithEvents dgv2 As System.Windows.Forms.DataGridView
    Friend WithEvents Col1_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1_CodeNo As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents Col1_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col1_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2_CodeNo As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents Col2_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents gbo2 As System.Windows.Forms.GroupBox
End Class
