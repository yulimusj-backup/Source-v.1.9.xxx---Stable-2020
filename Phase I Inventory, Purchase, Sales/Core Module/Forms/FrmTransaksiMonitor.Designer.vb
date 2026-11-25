<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonitor
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiMonitor))
        Me.dgvMonitor = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNamaKomp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIpAdd = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeTrans = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.btnRefreshMonitor = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.btnQueue_Clear = New System.Windows.Forms.Button
        Me.btnQueue_Refresh = New System.Windows.Forms.Button
        Me.prgQueue = New System.Windows.Forms.ProgressBar
        Me.dgvQueue = New System.Windows.Forms.DataGridView
        Me.ColQueue_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQueue_UserName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQueue_ComputerID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQueue_TransactionCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQueue_Timestamp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        CType(Me.dgvMonitor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvQueue, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvMonitor
        '
        Me.dgvMonitor.AllowUserToAddRows = False
        Me.dgvMonitor.AllowUserToDeleteRows = False
        Me.dgvMonitor.AllowUserToOrderColumns = True
        Me.dgvMonitor.AllowUserToResizeColumns = False
        Me.dgvMonitor.AllowUserToResizeRows = False
        Me.dgvMonitor.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvMonitor.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvMonitor.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvMonitor.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColUserName, Me.ColNamaKomp, Me.ColIpAdd, Me.ColTipeTrans})
        Me.dgvMonitor.Location = New System.Drawing.Point(6, 6)
        Me.dgvMonitor.Name = "dgvMonitor"
        Me.dgvMonitor.ReadOnly = True
        Me.dgvMonitor.RowTemplate.Height = 24
        Me.dgvMonitor.Size = New System.Drawing.Size(618, 349)
        Me.dgvMonitor.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 51
        '
        'ColUserName
        '
        Me.ColUserName.HeaderText = "User"
        Me.ColUserName.Name = "ColUserName"
        Me.ColUserName.ReadOnly = True
        Me.ColUserName.Width = 63
        '
        'ColNamaKomp
        '
        Me.ColNamaKomp.HeaderText = "Computer Id"
        Me.ColNamaKomp.Name = "ColNamaKomp"
        Me.ColNamaKomp.ReadOnly = True
        Me.ColNamaKomp.Width = 109
        '
        'ColIpAdd
        '
        Me.ColIpAdd.HeaderText = "IP Address"
        Me.ColIpAdd.Name = "ColIpAdd"
        Me.ColIpAdd.ReadOnly = True
        Me.ColIpAdd.Width = 101
        '
        'ColTipeTrans
        '
        Me.ColTipeTrans.HeaderText = "Transaction Type"
        Me.ColTipeTrans.Name = "ColTipeTrans"
        Me.ColTipeTrans.ReadOnly = True
        Me.ColTipeTrans.Visible = False
        Me.ColTipeTrans.Width = 144
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(527, 88)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Button1"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(6, 386)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(537, 23)
        Me.ProgressBar1.TabIndex = 2
        '
        'Timer2
        '
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 50
        '
        'btnRefreshMonitor
        '
        Me.btnRefreshMonitor.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRefreshMonitor.Location = New System.Drawing.Point(549, 386)
        Me.btnRefreshMonitor.Name = "btnRefreshMonitor"
        Me.btnRefreshMonitor.Size = New System.Drawing.Size(75, 23)
        Me.btnRefreshMonitor.TabIndex = 3
        Me.btnRefreshMonitor.Text = "Refresh"
        Me.btnRefreshMonitor.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 358)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Label1"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(642, 444)
        Me.TabControl1.TabIndex = 5
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.btnRefreshMonitor)
        Me.TabPage1.Controls.Add(Me.ProgressBar1)
        Me.TabPage1.Controls.Add(Me.dgvMonitor)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(634, 415)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Online User"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.btnQueue_Clear)
        Me.TabPage2.Controls.Add(Me.btnQueue_Refresh)
        Me.TabPage2.Controls.Add(Me.prgQueue)
        Me.TabPage2.Controls.Add(Me.dgvQueue)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(634, 415)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Transaction Queue"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'btnQueue_Clear
        '
        Me.btnQueue_Clear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnQueue_Clear.Location = New System.Drawing.Point(551, 384)
        Me.btnQueue_Clear.Name = "btnQueue_Clear"
        Me.btnQueue_Clear.Size = New System.Drawing.Size(75, 23)
        Me.btnQueue_Clear.TabIndex = 6
        Me.btnQueue_Clear.Text = "Clear"
        Me.btnQueue_Clear.UseVisualStyleBackColor = True
        '
        'btnQueue_Refresh
        '
        Me.btnQueue_Refresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnQueue_Refresh.Location = New System.Drawing.Point(470, 384)
        Me.btnQueue_Refresh.Name = "btnQueue_Refresh"
        Me.btnQueue_Refresh.Size = New System.Drawing.Size(75, 23)
        Me.btnQueue_Refresh.TabIndex = 5
        Me.btnQueue_Refresh.Text = "Refresh"
        Me.btnQueue_Refresh.UseVisualStyleBackColor = True
        '
        'prgQueue
        '
        Me.prgQueue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgQueue.Location = New System.Drawing.Point(8, 384)
        Me.prgQueue.Name = "prgQueue"
        Me.prgQueue.Size = New System.Drawing.Size(456, 23)
        Me.prgQueue.TabIndex = 4
        '
        'dgvQueue
        '
        Me.dgvQueue.AllowUserToAddRows = False
        Me.dgvQueue.AllowUserToDeleteRows = False
        Me.dgvQueue.AllowUserToOrderColumns = True
        Me.dgvQueue.AllowUserToResizeColumns = False
        Me.dgvQueue.AllowUserToResizeRows = False
        Me.dgvQueue.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvQueue.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvQueue.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvQueue.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColQueue_No, Me.ColQueue_UserName, Me.ColQueue_ComputerID, Me.ColQueue_TransactionCode, Me.ColQueue_Timestamp})
        Me.dgvQueue.Location = New System.Drawing.Point(8, 6)
        Me.dgvQueue.Name = "dgvQueue"
        Me.dgvQueue.ReadOnly = True
        Me.dgvQueue.RowTemplate.Height = 24
        Me.dgvQueue.Size = New System.Drawing.Size(618, 372)
        Me.dgvQueue.TabIndex = 1
        '
        'ColQueue_No
        '
        Me.ColQueue_No.HeaderText = "No"
        Me.ColQueue_No.Name = "ColQueue_No"
        Me.ColQueue_No.ReadOnly = True
        Me.ColQueue_No.Width = 51
        '
        'ColQueue_UserName
        '
        Me.ColQueue_UserName.HeaderText = "User"
        Me.ColQueue_UserName.Name = "ColQueue_UserName"
        Me.ColQueue_UserName.ReadOnly = True
        Me.ColQueue_UserName.Width = 63
        '
        'ColQueue_ComputerID
        '
        Me.ColQueue_ComputerID.HeaderText = "Computer Id"
        Me.ColQueue_ComputerID.Name = "ColQueue_ComputerID"
        Me.ColQueue_ComputerID.ReadOnly = True
        Me.ColQueue_ComputerID.Width = 109
        '
        'ColQueue_TransactionCode
        '
        Me.ColQueue_TransactionCode.HeaderText = "Transaction Code"
        Me.ColQueue_TransactionCode.Name = "ColQueue_TransactionCode"
        Me.ColQueue_TransactionCode.ReadOnly = True
        Me.ColQueue_TransactionCode.Width = 145
        '
        'ColQueue_Timestamp
        '
        Me.ColQueue_Timestamp.HeaderText = "Timestamp"
        Me.ColQueue_Timestamp.Name = "ColQueue_Timestamp"
        Me.ColQueue_Timestamp.ReadOnly = True
        Me.ColQueue_Timestamp.Width = 102
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 50
        '
        'FrmTransaksiMonitor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(642, 444)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnClose)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiMonitor"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Online User"
        CType(Me.dgvMonitor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvQueue, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvMonitor As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents btnRefreshMonitor As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNamaKomp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIpAdd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeTrans As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents btnQueue_Clear As System.Windows.Forms.Button
    Friend WithEvents btnQueue_Refresh As System.Windows.Forms.Button
    Friend WithEvents prgQueue As System.Windows.Forms.ProgressBar
    Friend WithEvents dgvQueue As System.Windows.Forms.DataGridView
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ColQueue_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQueue_UserName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQueue_ComputerID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQueue_TransactionCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQueue_Timestamp As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
