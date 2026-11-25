<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiJournal_EndOfMonth_DetailAPAR
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
        Me.dgvDetail = New System.Windows.Forms.DataGridView
        Me.ColNoTransaksi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExchangeRateLama = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExchangeRateBaru = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDifference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button1 = New System.Windows.Forms.Button
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvDetail
        '
        Me.dgvDetail.AllowUserToAddRows = False
        Me.dgvDetail.AllowUserToDeleteRows = False
        Me.dgvDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNoTransaksi, Me.ColBalance, Me.ColExchangeRateLama, Me.ColExchangeRateBaru, Me.ColDifference})
        Me.dgvDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDetail.Location = New System.Drawing.Point(0, 0)
        Me.dgvDetail.Name = "dgvDetail"
        Me.dgvDetail.ReadOnly = True
        Me.dgvDetail.Size = New System.Drawing.Size(716, 487)
        Me.dgvDetail.TabIndex = 0
        '
        'ColNoTransaksi
        '
        Me.ColNoTransaksi.HeaderText = "Document"
        Me.ColNoTransaksi.Name = "ColNoTransaksi"
        Me.ColNoTransaksi.ReadOnly = True
        Me.ColNoTransaksi.Width = 81
        '
        'ColBalance
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalance.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColBalance.HeaderText = "Balance"
        Me.ColBalance.Name = "ColBalance"
        Me.ColBalance.ReadOnly = True
        Me.ColBalance.Width = 71
        '
        'ColExchangeRateLama
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColExchangeRateLama.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColExchangeRateLama.HeaderText = "Previous Exchange Rate"
        Me.ColExchangeRateLama.Name = "ColExchangeRateLama"
        Me.ColExchangeRateLama.ReadOnly = True
        Me.ColExchangeRateLama.Width = 150
        '
        'ColExchangeRateBaru
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColExchangeRateBaru.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColExchangeRateBaru.HeaderText = "Current Exchange Rate"
        Me.ColExchangeRateBaru.Name = "ColExchangeRateBaru"
        Me.ColExchangeRateBaru.ReadOnly = True
        Me.ColExchangeRateBaru.Width = 143
        '
        'ColDifference
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDifference.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColDifference.HeaderText = "Difference"
        Me.ColDifference.Name = "ColDifference"
        Me.ColDifference.ReadOnly = True
        Me.ColDifference.Width = 81
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(436, 90)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiJournal_EndOfMonth_DetailAPAR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(716, 487)
        Me.Controls.Add(Me.dgvDetail)
        Me.Controls.Add(Me.Button1)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(732, 525)
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiJournal_EndOfMonth_DetailAPAR"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail AR Revaluation"
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvDetail As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ColNoTransaksi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExchangeRateLama As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExchangeRateBaru As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDifference As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
