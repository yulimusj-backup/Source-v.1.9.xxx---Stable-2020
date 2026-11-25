<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterCuti_Balance
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtLeaveDeduction = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtLeaveAddition = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtQtyLeave = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtTotalBalance = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtTotalUsed = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtTotalBeginning = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.dgvExplain = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNoDocument = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTanggalTransaksi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTanggalTransaksi_End = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyUsed = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryTransaksiRujukan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeTransaksiRujukan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        Me.chkDetail = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvExplain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.dgvExplain)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 5)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(928, 430)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.SplitContainer1)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 10)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Size = New System.Drawing.Size(912, 124)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(4, 19)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtLeaveDeduction)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtLeaveAddition)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtQtyLeave)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtTotalBalance)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtTotalUsed)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtTotalBeginning)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Size = New System.Drawing.Size(904, 101)
        Me.SplitContainer1.SplitterDistance = 444
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 0
        '
        'TxtLeaveDeduction
        '
        Me.TxtLeaveDeduction.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLeaveDeduction.Location = New System.Drawing.Point(132, 68)
        Me.TxtLeaveDeduction.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtLeaveDeduction.Name = "TxtLeaveDeduction"
        Me.TxtLeaveDeduction.ReadOnly = True
        Me.TxtLeaveDeduction.Size = New System.Drawing.Size(307, 22)
        Me.TxtLeaveDeduction.TabIndex = 5
        Me.TxtLeaveDeduction.Text = "0"
        Me.TxtLeaveDeduction.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(5, 71)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(115, 17)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Leave Deduction"
        '
        'TxtLeaveAddition
        '
        Me.TxtLeaveAddition.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLeaveAddition.Location = New System.Drawing.Point(132, 36)
        Me.TxtLeaveAddition.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtLeaveAddition.Name = "TxtLeaveAddition"
        Me.TxtLeaveAddition.ReadOnly = True
        Me.TxtLeaveAddition.Size = New System.Drawing.Size(307, 22)
        Me.TxtLeaveAddition.TabIndex = 3
        Me.TxtLeaveAddition.Text = "0"
        Me.TxtLeaveAddition.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 39)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(102, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Leave Addition"
        '
        'TxtQtyLeave
        '
        Me.TxtQtyLeave.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyLeave.Location = New System.Drawing.Point(132, 4)
        Me.TxtQtyLeave.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtQtyLeave.Name = "TxtQtyLeave"
        Me.TxtQtyLeave.ReadOnly = True
        Me.TxtQtyLeave.Size = New System.Drawing.Size(307, 22)
        Me.TxtQtyLeave.TabIndex = 1
        Me.TxtQtyLeave.Text = "0"
        Me.TxtQtyLeave.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 7)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Qty Leave"
        '
        'TxtTotalBalance
        '
        Me.TxtTotalBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalBalance.Location = New System.Drawing.Point(143, 68)
        Me.TxtTotalBalance.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtTotalBalance.Name = "TxtTotalBalance"
        Me.TxtTotalBalance.ReadOnly = True
        Me.TxtTotalBalance.Size = New System.Drawing.Size(307, 22)
        Me.TxtTotalBalance.TabIndex = 7
        Me.TxtTotalBalance.Text = "0"
        Me.TxtTotalBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(16, 71)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(95, 17)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Total Balance"
        '
        'TxtTotalUsed
        '
        Me.TxtTotalUsed.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalUsed.Location = New System.Drawing.Point(143, 39)
        Me.TxtTotalUsed.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtTotalUsed.Name = "TxtTotalUsed"
        Me.TxtTotalUsed.ReadOnly = True
        Me.TxtTotalUsed.Size = New System.Drawing.Size(307, 22)
        Me.TxtTotalUsed.TabIndex = 5
        Me.TxtTotalUsed.Text = "0"
        Me.TxtTotalUsed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(16, 43)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(77, 17)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Total Used"
        '
        'TxtTotalBeginning
        '
        Me.TxtTotalBeginning.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalBeginning.Location = New System.Drawing.Point(143, 7)
        Me.TxtTotalBeginning.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtTotalBeginning.Name = "TxtTotalBeginning"
        Me.TxtTotalBeginning.ReadOnly = True
        Me.TxtTotalBeginning.Size = New System.Drawing.Size(307, 22)
        Me.TxtTotalBeginning.TabIndex = 3
        Me.TxtTotalBeginning.Text = "0"
        Me.TxtTotalBeginning.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 11)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(115, 17)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Total Beginnning"
        '
        'dgvExplain
        '
        Me.dgvExplain.AllowUserToAddRows = False
        Me.dgvExplain.AllowUserToDeleteRows = False
        Me.dgvExplain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvExplain.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvExplain.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColNoDocument, Me.ColTanggalTransaksi, Me.ColTanggalTransaksi_End, Me.ColQtyUsed, Me.ColBalance, Me.ColPrimaryTransaksiRujukan, Me.ColTipeTransaksiRujukan})
        Me.dgvExplain.Location = New System.Drawing.Point(8, 142)
        Me.dgvExplain.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvExplain.Name = "dgvExplain"
        Me.dgvExplain.ReadOnly = True
        Me.dgvExplain.RowTemplate.Height = 24
        Me.dgvExplain.Size = New System.Drawing.Size(912, 281)
        Me.dgvExplain.TabIndex = 1
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColNoDocument
        '
        Me.ColNoDocument.HeaderText = "Document No"
        Me.ColNoDocument.Name = "ColNoDocument"
        Me.ColNoDocument.ReadOnly = True
        Me.ColNoDocument.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColTanggalTransaksi
        '
        Me.ColTanggalTransaksi.HeaderText = "Date"
        Me.ColTanggalTransaksi.Name = "ColTanggalTransaksi"
        Me.ColTanggalTransaksi.ReadOnly = True
        Me.ColTanggalTransaksi.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTanggalTransaksi.Width = 44
        '
        'ColTanggalTransaksi_End
        '
        Me.ColTanggalTransaksi_End.HeaderText = "End Date"
        Me.ColTanggalTransaksi_End.Name = "ColTanggalTransaksi_End"
        Me.ColTanggalTransaksi_End.ReadOnly = True
        Me.ColTanggalTransaksi_End.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTanggalTransaksi_End.Visible = False
        Me.ColTanggalTransaksi_End.Width = 73
        '
        'ColQtyUsed
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyUsed.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQtyUsed.HeaderText = "Qty Used"
        Me.ColQtyUsed.Name = "ColQtyUsed"
        Me.ColQtyUsed.ReadOnly = True
        Me.ColQtyUsed.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyUsed.Width = 73
        '
        'ColBalance
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalance.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColBalance.HeaderText = "Balance"
        Me.ColBalance.Name = "ColBalance"
        Me.ColBalance.ReadOnly = True
        Me.ColBalance.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBalance.Width = 65
        '
        'ColPrimaryTransaksiRujukan
        '
        Me.ColPrimaryTransaksiRujukan.HeaderText = "Primary Transaksi Rujukan"
        Me.ColPrimaryTransaksiRujukan.Name = "ColPrimaryTransaksiRujukan"
        Me.ColPrimaryTransaksiRujukan.ReadOnly = True
        Me.ColPrimaryTransaksiRujukan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryTransaksiRujukan.Visible = False
        Me.ColPrimaryTransaksiRujukan.Width = 184
        '
        'ColTipeTransaksiRujukan
        '
        Me.ColTipeTransaksiRujukan.HeaderText = "Tipe Transaksi Rujukan"
        Me.ColTipeTransaksiRujukan.Name = "ColTipeTransaksiRujukan"
        Me.ColTipeTransaksiRujukan.ReadOnly = True
        Me.ColTipeTransaksiRujukan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTipeTransaksiRujukan.Visible = False
        Me.ColTipeTransaksiRujukan.Width = 164
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(844, 441)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 28)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'chkDetail
        '
        Me.chkDetail.AutoSize = True
        Me.chkDetail.Location = New System.Drawing.Point(16, 441)
        Me.chkDetail.Name = "chkDetail"
        Me.chkDetail.Size = New System.Drawing.Size(104, 21)
        Me.chkDetail.TabIndex = 3
        Me.chkDetail.Text = "Show Detail"
        Me.chkDetail.UseVisualStyleBackColor = True
        '
        'FrmMasterCuti_Balance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(971, 485)
        Me.Controls.Add(Me.chkDetail)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(979, 522)
        Me.Name = "FrmMasterCuti_Balance"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Leave Application Balance"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvExplain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents dgvExplain As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtQtyLeave As System.Windows.Forms.TextBox
    Friend WithEvents TxtLeaveAddition As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtLeaveDeduction As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalBeginning As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalBalance As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalUsed As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents chkDetail As System.Windows.Forms.CheckBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNoDocument As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTanggalTransaksi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTanggalTransaksi_End As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyUsed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryTransaksiRujukan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeTransaksiRujukan As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
