<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiCuti_Pengajuan_ExplainDuration
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
        Me.dgvExplain = New System.Windows.Forms.DataGridView
        Me.btnClose = New System.Windows.Forms.Button
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkShift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColPrimaryWorkShift = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvExplain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvExplain
        '
        Me.dgvExplain.AllowUserToAddRows = False
        Me.dgvExplain.AllowUserToDeleteRows = False
        Me.dgvExplain.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvExplain.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDate, Me.ColWorkShift, Me.ColDuration, Me.ColStatus, Me.ColPrimaryWorkShift})
        Me.dgvExplain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvExplain.Location = New System.Drawing.Point(0, 0)
        Me.dgvExplain.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvExplain.Name = "dgvExplain"
        Me.dgvExplain.RowTemplate.Height = 24
        Me.dgvExplain.Size = New System.Drawing.Size(589, 508)
        Me.dgvExplain.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(463, 464)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 28)
        Me.btnClose.TabIndex = 1
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.Width = 63
        '
        'ColWorkShift
        '
        Me.ColWorkShift.HeaderText = "Work Shift"
        Me.ColWorkShift.Name = "ColWorkShift"
        Me.ColWorkShift.Width = 98
        '
        'ColDuration
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle1.Format = "N1"
        DataGridViewCellStyle1.NullValue = "0"
        Me.ColDuration.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColDuration.HeaderText = "Duration (days)"
        Me.ColDuration.Name = "ColDuration"
        Me.ColDuration.Width = 131
        '
        'ColStatus
        '
        Me.ColStatus.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColStatus.HeaderText = "Status"
        Me.ColStatus.Items.AddRange(New Object() {"Leave", "Sick", "Permit", "Absent"})
        Me.ColStatus.MinimumWidth = 100
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.ToolTipText = "Status"
        '
        'ColPrimaryWorkShift
        '
        Me.ColPrimaryWorkShift.HeaderText = "Primary Work Shift"
        Me.ColPrimaryWorkShift.Name = "ColPrimaryWorkShift"
        Me.ColPrimaryWorkShift.Visible = False
        Me.ColPrimaryWorkShift.Width = 150
        '
        'FrmTransaksiCuti_Pengajuan_ExplainDuration
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(589, 508)
        Me.Controls.Add(Me.dgvExplain)
        Me.Controls.Add(Me.btnClose)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(597, 545)
        Me.Name = "FrmTransaksiCuti_Pengajuan_ExplainDuration"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Explain Duration"
        CType(Me.dgvExplain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvExplain As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkShift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColPrimaryWorkShift As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
