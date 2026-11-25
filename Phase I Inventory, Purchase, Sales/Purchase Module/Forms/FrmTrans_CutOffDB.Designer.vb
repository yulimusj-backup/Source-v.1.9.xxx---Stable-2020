<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_CutOffDB
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
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.dgvProgress = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColModule = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnStartCancel = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.LblCutOffPeriod = New System.Windows.Forms.Label
        Me.dtpCutOffPeriod = New System.Windows.Forms.DateTimePicker
        Me.Bgw1 = New System.ComponentModel.BackgroundWorker
        Me.FrmDia = New System.Windows.Forms.SaveFileDialog
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.TxtDB = New System.Windows.Forms.TextBox
        CType(Me.dgvProgress, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(9, 470)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(2)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(459, 19)
        Me.ProgressBar1.TabIndex = 1
        '
        'dgvProgress
        '
        Me.dgvProgress.AllowUserToAddRows = False
        Me.dgvProgress.AllowUserToDeleteRows = False
        Me.dgvProgress.AllowUserToResizeColumns = False
        Me.dgvProgress.AllowUserToResizeRows = False
        Me.dgvProgress.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvProgress.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvProgress.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvProgress.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColModule, Me.ColDesc, Me.ColStatus})
        Me.dgvProgress.Location = New System.Drawing.Point(9, 34)
        Me.dgvProgress.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvProgress.Name = "dgvProgress"
        Me.dgvProgress.ReadOnly = True
        Me.dgvProgress.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvProgress.RowTemplate.Height = 24
        Me.dgvProgress.Size = New System.Drawing.Size(520, 431)
        Me.dgvProgress.TabIndex = 2
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColModule
        '
        Me.ColModule.HeaderText = "Module"
        Me.ColModule.Name = "ColModule"
        Me.ColModule.ReadOnly = True
        Me.ColModule.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColModule.Width = 48
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Width = 66
        '
        'ColStatus
        '
        Me.ColStatus.HeaderText = "Status"
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.ReadOnly = True
        Me.ColStatus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColStatus.Width = 43
        '
        'btnStartCancel
        '
        Me.btnStartCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStartCancel.Location = New System.Drawing.Point(472, 470)
        Me.btnStartCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnStartCancel.Name = "btnStartCancel"
        Me.btnStartCancel.Size = New System.Drawing.Size(56, 19)
        Me.btnStartCancel.TabIndex = 3
        Me.btnStartCancel.Text = "Start"
        Me.btnStartCancel.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(472, 492)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(56, 19)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 494)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 13)
        Me.Label1.TabIndex = 5
        '
        'LblCutOffPeriod
        '
        Me.LblCutOffPeriod.AutoSize = True
        Me.LblCutOffPeriod.Location = New System.Drawing.Point(9, 11)
        Me.LblCutOffPeriod.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblCutOffPeriod.Name = "LblCutOffPeriod"
        Me.LblCutOffPeriod.Size = New System.Drawing.Size(70, 13)
        Me.LblCutOffPeriod.TabIndex = 6
        Me.LblCutOffPeriod.Tag = ""
        Me.LblCutOffPeriod.Text = "Cut off period"
        '
        'dtpCutOffPeriod
        '
        Me.dtpCutOffPeriod.CustomFormat = "yyyy"
        Me.dtpCutOffPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpCutOffPeriod.Location = New System.Drawing.Point(83, 10)
        Me.dtpCutOffPeriod.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpCutOffPeriod.Name = "dtpCutOffPeriod"
        Me.dtpCutOffPeriod.ShowUpDown = True
        Me.dtpCutOffPeriod.Size = New System.Drawing.Size(80, 20)
        Me.dtpCutOffPeriod.TabIndex = 7
        '
        'Bgw1
        '
        Me.Bgw1.WorkerSupportsCancellation = True
        '
        'FrmDia
        '
        Me.FrmDia.DefaultExt = "sql"
        Me.FrmDia.Filter = "sql files (*.sql)|*.sql"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(274, 11)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(59, 13)
        Me.LinkLabel1.TabIndex = 10
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "LinkLabel1"
        '
        'TxtDB
        '
        Me.TxtDB.Location = New System.Drawing.Point(168, 8)
        Me.TxtDB.Name = "TxtDB"
        Me.TxtDB.Size = New System.Drawing.Size(100, 20)
        Me.TxtDB.TabIndex = 11
        Me.TxtDB.Text = "smess_iii"
        '
        'FrmTrans_CutOffDB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(535, 514)
        Me.Controls.Add(Me.TxtDB)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.dtpCutOffPeriod)
        Me.Controls.Add(Me.LblCutOffPeriod)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnStartCancel)
        Me.Controls.Add(Me.dgvProgress)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(464, 335)
        Me.Name = "FrmTrans_CutOffDB"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Database cut off"
        CType(Me.dgvProgress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents dgvProgress As System.Windows.Forms.DataGridView
    Friend WithEvents btnStartCancel As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblCutOffPeriod As System.Windows.Forms.Label
    Friend WithEvents dtpCutOffPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColModule As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Bgw1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents FrmDia As System.Windows.Forms.SaveFileDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtDB As System.Windows.Forms.TextBox
End Class
