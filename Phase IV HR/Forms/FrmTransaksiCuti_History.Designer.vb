<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiCuti_History
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
        Me.dgvHistory = New System.Windows.Forms.DataGridView
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboLeave = New System.Windows.Forms.ComboBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeaveAplicationDoc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLeaveAplication = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvHistory
        '
        Me.dgvHistory.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvHistory.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvHistory.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDate, Me.ColDesc, Me.ColLeaveAplicationDoc, Me.ColDuration, Me.ColBalance, Me.ColPrimaryLeaveAplication})
        Me.dgvHistory.Location = New System.Drawing.Point(12, 70)
        Me.dgvHistory.Name = "dgvHistory"
        Me.dgvHistory.Size = New System.Drawing.Size(491, 407)
        Me.dgvHistory.TabIndex = 0
        '
        'TextBox3
        '
        Me.TextBox3.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox3.Location = New System.Drawing.Point(71, 12)
        Me.TextBox3.MaxLength = 50
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ReadOnly = True
        Me.TextBox3.Size = New System.Drawing.Size(333, 20)
        Me.TextBox3.TabIndex = 39
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(71, 12)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(28, 20)
        Me.TextBox1.TabIndex = 40
        Me.TextBox1.TabStop = False
        Me.TextBox1.Text = "-1"
        Me.TextBox1.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 13)
        Me.Label2.TabIndex = 41
        Me.Label2.Text = "Employee"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 44)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(37, 13)
        Me.Label4.TabIndex = 43
        Me.Label4.Text = "Period"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "yyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(71, 38)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.ShowUpDown = True
        Me.DateTimePicker1.Size = New System.Drawing.Size(84, 20)
        Me.DateTimePicker1.TabIndex = 42
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(161, 41)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 13)
        Me.Label5.TabIndex = 45
        Me.Label5.Text = "Leave Code"
        '
        'cboLeave
        '
        Me.cboLeave.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLeave.FormattingEnabled = True
        Me.cboLeave.Location = New System.Drawing.Point(232, 38)
        Me.cboLeave.Name = "cboLeave"
        Me.cboLeave.Size = New System.Drawing.Size(172, 21)
        Me.cboLeave.TabIndex = 44
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(387, 95)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 46
        Me.btnClose.Text = "Button1"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.Width = 55
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.Width = 85
        '
        'ColLeaveAplicationDoc
        '
        Me.ColLeaveAplicationDoc.HeaderText = "Leave Aplication Document No"
        Me.ColLeaveAplicationDoc.Name = "ColLeaveAplicationDoc"
        Me.ColLeaveAplicationDoc.Width = 180
        '
        'ColDuration
        '
        Me.ColDuration.HeaderText = "Duration"
        Me.ColDuration.Name = "ColDuration"
        Me.ColDuration.Width = 72
        '
        'ColBalance
        '
        Me.ColBalance.HeaderText = "Balance"
        Me.ColBalance.Name = "ColBalance"
        Me.ColBalance.Width = 71
        '
        'ColPrimaryLeaveAplication
        '
        Me.ColPrimaryLeaveAplication.HeaderText = "Primary Leave Aplication"
        Me.ColPrimaryLeaveAplication.Name = "ColPrimaryLeaveAplication"
        Me.ColPrimaryLeaveAplication.Visible = False
        Me.ColPrimaryLeaveAplication.Width = 148
        '
        'FrmTransaksiCuti_History
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(515, 489)
        Me.Controls.Add(Me.dgvHistory)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cboLeave)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox3)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(531, 527)
        Me.Name = "FrmTransaksiCuti_History"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Leave History"
        CType(Me.dgvHistory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvHistory As System.Windows.Forms.DataGridView
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboLeave As System.Windows.Forms.ComboBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeaveAplicationDoc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLeaveAplication As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
