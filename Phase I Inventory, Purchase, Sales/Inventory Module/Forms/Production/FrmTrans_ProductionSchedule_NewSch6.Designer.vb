<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_NewSch6
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.TxtBackColor = New System.Windows.Forms.TextBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.cboProcess = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvWU = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkUnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkUnitDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryWorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvWU, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TxtBackColor
        '
        Me.TxtBackColor.Location = New System.Drawing.Point(111, 236)
        Me.TxtBackColor.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtBackColor.Name = "TxtBackColor"
        Me.TxtBackColor.Size = New System.Drawing.Size(53, 20)
        Me.TxtBackColor.TabIndex = 15
        Me.TxtBackColor.TabStop = False
        Me.TxtBackColor.Visible = False
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(327, 235)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 19)
        Me.btnOK.TabIndex = 11
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(387, 235)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 19)
        Me.btnCancel.TabIndex = 12
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'cboProcess
        '
        Me.cboProcess.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboProcess.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProcess.FormattingEnabled = True
        Me.cboProcess.Location = New System.Drawing.Point(111, 9)
        Me.cboProcess.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboProcess.Name = "cboProcess"
        Me.cboProcess.Size = New System.Drawing.Size(334, 21)
        Me.cboProcess.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 12)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Process"
        '
        'dgvWU
        '
        Me.dgvWU.AllowUserToAddRows = False
        Me.dgvWU.AllowUserToDeleteRows = False
        Me.dgvWU.AllowUserToResizeColumns = False
        Me.dgvWU.AllowUserToResizeRows = False
        Me.dgvWU.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvWU.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvWU.ColumnHeadersHeight = 25
        Me.dgvWU.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvWU.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColWorkUnitCode, Me.ColWorkUnitDesc, Me.ColRate, Me.ColPrimaryWorkUnit})
        Me.dgvWU.Location = New System.Drawing.Point(9, 32)
        Me.dgvWU.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvWU.MultiSelect = False
        Me.dgvWU.Name = "dgvWU"
        Me.dgvWU.RowHeadersVisible = False
        Me.dgvWU.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvWU.RowTemplate.Height = 24
        Me.dgvWU.Size = New System.Drawing.Size(436, 198)
        Me.dgvWU.TabIndex = 16
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 46
        '
        'ColWorkUnitCode
        '
        Me.ColWorkUnitCode.HeaderText = "Work Unit Code"
        Me.ColWorkUnitCode.Name = "ColWorkUnitCode"
        Me.ColWorkUnitCode.Width = 108
        '
        'ColWorkUnitDesc
        '
        Me.ColWorkUnitDesc.HeaderText = "Description"
        Me.ColWorkUnitDesc.Name = "ColWorkUnitDesc"
        Me.ColWorkUnitDesc.Width = 85
        '
        'ColRate
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRate.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColRate.HeaderText = "Plan Rate (Unit/Hour)"
        Me.ColRate.Name = "ColRate"
        Me.ColRate.Width = 135
        '
        'ColPrimaryWorkUnit
        '
        Me.ColPrimaryWorkUnit.HeaderText = "Primary Work Unit"
        Me.ColPrimaryWorkUnit.Name = "ColPrimaryWorkUnit"
        Me.ColPrimaryWorkUnit.Visible = False
        Me.ColPrimaryWorkUnit.Width = 117
        '
        'FrmTrans_ProductionSchedule_NewSch6
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(454, 262)
        Me.Controls.Add(Me.dgvWU)
        Me.Controls.Add(Me.TxtBackColor)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.cboProcess)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(470, 300)
        Me.Name = "FrmTrans_ProductionSchedule_NewSch6"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Set Rate By Process"
        CType(Me.dgvWU, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtBackColor As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents cboProcess As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvWU As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkUnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkUnitDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryWorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
