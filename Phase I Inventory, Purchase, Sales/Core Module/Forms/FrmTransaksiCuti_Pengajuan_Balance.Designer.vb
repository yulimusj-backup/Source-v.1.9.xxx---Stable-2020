<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiCuti_Pengajuan_Balance
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNoDocument = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTanggalTransaksi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyUsed = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryTransaksiRujukan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeTransaksiRujukan = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.GroupBox1.Location = New System.Drawing.Point(12, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(696, 349)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.SplitContainer1)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(684, 101)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 16)
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
        Me.SplitContainer1.Size = New System.Drawing.Size(678, 82)
        Me.SplitContainer1.SplitterDistance = 333
        Me.SplitContainer1.TabIndex = 0
        '
        'TxtLeaveDeduction
        '
        Me.TxtLeaveDeduction.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLeaveDeduction.Location = New System.Drawing.Point(99, 55)
        Me.TxtLeaveDeduction.Name = "TxtLeaveDeduction"
        Me.TxtLeaveDeduction.ReadOnly = True
        Me.TxtLeaveDeduction.Size = New System.Drawing.Size(231, 20)
        Me.TxtLeaveDeduction.TabIndex = 5
        Me.TxtLeaveDeduction.Text = "0"
        Me.TxtLeaveDeduction.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 58)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(89, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Leave Deduction"
        '
        'TxtLeaveAddition
        '
        Me.TxtLeaveAddition.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLeaveAddition.Location = New System.Drawing.Point(99, 29)
        Me.TxtLeaveAddition.Name = "TxtLeaveAddition"
        Me.TxtLeaveAddition.ReadOnly = True
        Me.TxtLeaveAddition.Size = New System.Drawing.Size(231, 20)
        Me.TxtLeaveAddition.TabIndex = 3
        Me.TxtLeaveAddition.Text = "0"
        Me.TxtLeaveAddition.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Leave Addition"
        '
        'TxtQtyLeave
        '
        Me.TxtQtyLeave.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyLeave.Location = New System.Drawing.Point(99, 3)
        Me.TxtQtyLeave.Name = "TxtQtyLeave"
        Me.TxtQtyLeave.ReadOnly = True
        Me.TxtQtyLeave.Size = New System.Drawing.Size(231, 20)
        Me.TxtQtyLeave.TabIndex = 1
        Me.TxtQtyLeave.Text = "0"
        Me.TxtQtyLeave.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Qty Leave"
        '
        'TxtTotalBalance
        '
        Me.TxtTotalBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalBalance.Location = New System.Drawing.Point(107, 55)
        Me.TxtTotalBalance.Name = "TxtTotalBalance"
        Me.TxtTotalBalance.ReadOnly = True
        Me.TxtTotalBalance.Size = New System.Drawing.Size(231, 20)
        Me.TxtTotalBalance.TabIndex = 7
        Me.TxtTotalBalance.Text = "0"
        Me.TxtTotalBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 58)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(73, 13)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Total Balance"
        '
        'TxtTotalUsed
        '
        Me.TxtTotalUsed.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalUsed.Location = New System.Drawing.Point(107, 32)
        Me.TxtTotalUsed.Name = "TxtTotalUsed"
        Me.TxtTotalUsed.ReadOnly = True
        Me.TxtTotalUsed.Size = New System.Drawing.Size(231, 20)
        Me.TxtTotalUsed.TabIndex = 5
        Me.TxtTotalUsed.Text = "0"
        Me.TxtTotalUsed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 35)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(59, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Total Used"
        '
        'TxtTotalBeginning
        '
        Me.TxtTotalBeginning.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalBeginning.Location = New System.Drawing.Point(107, 6)
        Me.TxtTotalBeginning.Name = "TxtTotalBeginning"
        Me.TxtTotalBeginning.ReadOnly = True
        Me.TxtTotalBeginning.Size = New System.Drawing.Size(231, 20)
        Me.TxtTotalBeginning.TabIndex = 3
        Me.TxtTotalBeginning.Text = "0"
        Me.TxtTotalBeginning.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(87, 13)
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
        Me.dgvExplain.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColNoDocument, Me.ColTanggalTransaksi, Me.ColQtyUsed, Me.ColBalance, Me.ColPrimaryTransaksiRujukan, Me.ColTipeTransaksiRujukan})
        Me.dgvExplain.Location = New System.Drawing.Point(6, 115)
        Me.dgvExplain.Name = "dgvExplain"
        Me.dgvExplain.ReadOnly = True
        Me.dgvExplain.Size = New System.Drawing.Size(684, 228)
        Me.dgvExplain.TabIndex = 1
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(633, 358)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 46
        '
        'ColNoDocument
        '
        Me.ColNoDocument.HeaderText = "Document No"
        Me.ColNoDocument.Name = "ColNoDocument"
        Me.ColNoDocument.ReadOnly = True
        Me.ColNoDocument.Width = 98
        '
        'ColTanggalTransaksi
        '
        Me.ColTanggalTransaksi.HeaderText = "Date"
        Me.ColTanggalTransaksi.Name = "ColTanggalTransaksi"
        Me.ColTanggalTransaksi.ReadOnly = True
        Me.ColTanggalTransaksi.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColTanggalTransaksi.Width = 55
        '
        'ColQtyUsed
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyUsed.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQtyUsed.HeaderText = "Qty Used"
        Me.ColQtyUsed.Name = "ColQtyUsed"
        Me.ColQtyUsed.ReadOnly = True
        Me.ColQtyUsed.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColQtyUsed.Width = 76
        '
        'ColBalance
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalance.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColBalance.HeaderText = "Balance"
        Me.ColBalance.Name = "ColBalance"
        Me.ColBalance.ReadOnly = True
        Me.ColBalance.Width = 71
        '
        'ColPrimaryTransaksiRujukan
        '
        Me.ColPrimaryTransaksiRujukan.HeaderText = "Primary Transaksi Rujukan"
        Me.ColPrimaryTransaksiRujukan.Name = "ColPrimaryTransaksiRujukan"
        Me.ColPrimaryTransaksiRujukan.ReadOnly = True
        Me.ColPrimaryTransaksiRujukan.Visible = False
        Me.ColPrimaryTransaksiRujukan.Width = 158
        '
        'ColTipeTransaksiRujukan
        '
        Me.ColTipeTransaksiRujukan.HeaderText = "Tipe Transaksi Rujukan"
        Me.ColTipeTransaksiRujukan.Name = "ColTipeTransaksiRujukan"
        Me.ColTipeTransaksiRujukan.ReadOnly = True
        Me.ColTipeTransaksiRujukan.Visible = False
        Me.ColTipeTransaksiRujukan.Width = 145
        '
        'FrmTransaksiCuti_Pengajuan_Balance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(720, 393)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(736, 431)
        Me.Name = "FrmTransaksiCuti_Pengajuan_Balance"
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
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNoDocument As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTanggalTransaksi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyUsed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryTransaksiRujukan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeTransaksiRujukan As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
