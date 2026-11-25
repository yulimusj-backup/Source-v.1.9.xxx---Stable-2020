<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2_Cek
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
        Me.dgvCheckResult = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCheckResult = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnStart = New System.Windows.Forms.Button
        Me.chk1 = New System.Windows.Forms.CheckBox
        Me.chk2 = New System.Windows.Forms.CheckBox
        Me.chk3 = New System.Windows.Forms.CheckBox
        Me.chk4 = New System.Windows.Forms.CheckBox
        Me.prg1 = New System.Windows.Forms.ProgressBar
        Me.lblProgress = New System.Windows.Forms.Label
        Me.bgw1 = New System.ComponentModel.BackgroundWorker
        CType(Me.dgvCheckResult, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvCheckResult
        '
        Me.dgvCheckResult.AllowUserToAddRows = False
        Me.dgvCheckResult.AllowUserToDeleteRows = False
        Me.dgvCheckResult.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCheckResult.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCheckResult.ColumnHeadersHeight = 25
        Me.dgvCheckResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCheckResult.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColCheckResult})
        Me.dgvCheckResult.Location = New System.Drawing.Point(12, 66)
        Me.dgvCheckResult.Name = "dgvCheckResult"
        Me.dgvCheckResult.ReadOnly = True
        Me.dgvCheckResult.RowTemplate.Height = 24
        Me.dgvCheckResult.Size = New System.Drawing.Size(891, 522)
        Me.dgvCheckResult.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 51
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.Width = 88
        '
        'ColCheckResult
        '
        Me.ColCheckResult.HeaderText = "Check Result"
        Me.ColCheckResult.Name = "ColCheckResult"
        Me.ColCheckResult.ReadOnly = True
        Me.ColCheckResult.Width = 116
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(828, 594)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(484, 10)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 23)
        Me.btnStart.TabIndex = 2
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'chk1
        '
        Me.chk1.AutoSize = True
        Me.chk1.Location = New System.Drawing.Point(12, 39)
        Me.chk1.Name = "chk1"
        Me.chk1.Size = New System.Drawing.Size(133, 21)
        Me.chk1.TabIndex = 3
        Me.chk1.TabStop = False
        Me.chk1.Text = "Rekap VS Detail"
        Me.chk1.UseVisualStyleBackColor = True
        Me.chk1.Visible = False
        '
        'chk2
        '
        Me.chk2.AutoSize = True
        Me.chk2.Checked = True
        Me.chk2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk2.Location = New System.Drawing.Point(12, 12)
        Me.chk2.Name = "chk2"
        Me.chk2.Size = New System.Drawing.Size(125, 21)
        Me.chk2.TabIndex = 4
        Me.chk2.Text = "Check Item Qty"
        Me.chk2.UseVisualStyleBackColor = True
        '
        'chk3
        '
        Me.chk3.AutoSize = True
        Me.chk3.Checked = True
        Me.chk3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk3.Location = New System.Drawing.Point(236, 12)
        Me.chk3.Name = "chk3"
        Me.chk3.Size = New System.Drawing.Size(239, 21)
        Me.chk3.TabIndex = 6
        Me.chk3.Text = "Rekap End VS Detail Warehouse"
        Me.chk3.UseVisualStyleBackColor = True
        '
        'chk4
        '
        Me.chk4.AutoSize = True
        Me.chk4.Checked = True
        Me.chk4.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk4.Location = New System.Drawing.Point(236, 39)
        Me.chk4.Name = "chk4"
        Me.chk4.Size = New System.Drawing.Size(200, 21)
        Me.chk4.TabIndex = 5
        Me.chk4.Text = "Rekap Out VS Detail Using"
        Me.chk4.UseVisualStyleBackColor = True
        '
        'prg1
        '
        Me.prg1.Location = New System.Drawing.Point(565, 10)
        Me.prg1.Name = "prg1"
        Me.prg1.Size = New System.Drawing.Size(338, 23)
        Me.prg1.TabIndex = 0
        Me.prg1.Visible = False
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(565, 44)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(0, 17)
        Me.lblProgress.TabIndex = 8
        Me.lblProgress.Visible = False
        '
        'bgw1
        '
        Me.bgw1.WorkerReportsProgress = True
        '
        'Form2_Cek
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(915, 629)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.prg1)
        Me.Controls.Add(Me.chk3)
        Me.Controls.Add(Me.chk4)
        Me.Controls.Add(Me.chk2)
        Me.Controls.Add(Me.chk1)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.dgvCheckResult)
        Me.Name = "Form2_Cek"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form2_Cek"
        CType(Me.dgvCheckResult, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvCheckResult As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents chk1 As System.Windows.Forms.CheckBox
    Friend WithEvents chk2 As System.Windows.Forms.CheckBox
    Friend WithEvents chk3 As System.Windows.Forms.CheckBox
    Friend WithEvents chk4 As System.Windows.Forms.CheckBox
    Friend WithEvents prg1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents bgw1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCheckResult As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
