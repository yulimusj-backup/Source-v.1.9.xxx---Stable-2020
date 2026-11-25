<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmUpdate_RepairAPAR
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
        Me.btnCek = New System.Windows.Forms.Button
        Me.btnRepair = New System.Windows.Forms.Button
        Me.prg1 = New System.Windows.Forms.ProgressBar
        Me.LblProgress = New System.Windows.Forms.Label
        Me.dgvRepair = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRefParty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalKartu = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSelisih = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryRefParty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.rboAP = New System.Windows.Forms.RadioButton
        Me.rboAR = New System.Windows.Forms.RadioButton
        Me.bgw1 = New System.ComponentModel.BackgroundWorker
        Me.Button1 = New System.Windows.Forms.Button
        CType(Me.dgvRepair, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCek
        '
        Me.btnCek.Location = New System.Drawing.Point(65, 12)
        Me.btnCek.Name = "btnCek"
        Me.btnCek.Size = New System.Drawing.Size(75, 23)
        Me.btnCek.TabIndex = 0
        Me.btnCek.Text = "Cek"
        Me.btnCek.UseVisualStyleBackColor = True
        '
        'btnRepair
        '
        Me.btnRepair.Location = New System.Drawing.Point(65, 41)
        Me.btnRepair.Name = "btnRepair"
        Me.btnRepair.Size = New System.Drawing.Size(75, 23)
        Me.btnRepair.TabIndex = 1
        Me.btnRepair.Text = "Repair"
        Me.btnRepair.UseVisualStyleBackColor = True
        '
        'prg1
        '
        Me.prg1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prg1.Location = New System.Drawing.Point(146, 12)
        Me.prg1.Name = "prg1"
        Me.prg1.Size = New System.Drawing.Size(647, 23)
        Me.prg1.TabIndex = 2
        '
        'LblProgress
        '
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(143, 47)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(13, 17)
        Me.LblProgress.TabIndex = 3
        Me.LblProgress.Text = "-"
        '
        'dgvRepair
        '
        Me.dgvRepair.AllowUserToAddRows = False
        Me.dgvRepair.AllowUserToDeleteRows = False
        Me.dgvRepair.AllowUserToResizeColumns = False
        Me.dgvRepair.AllowUserToResizeRows = False
        Me.dgvRepair.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvRepair.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvRepair.ColumnHeadersHeight = 25
        Me.dgvRepair.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvRepair.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColRefParty, Me.ColTotalKartu, Me.ColTotalBalance, Me.ColSelisih, Me.ColPrimaryRefParty})
        Me.dgvRepair.Location = New System.Drawing.Point(12, 101)
        Me.dgvRepair.Name = "dgvRepair"
        Me.dgvRepair.RowHeadersVisible = False
        Me.dgvRepair.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvRepair.RowTemplate.Height = 24
        Me.dgvRepair.Size = New System.Drawing.Size(781, 283)
        Me.dgvRepair.TabIndex = 4
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 51
        '
        'ColRefParty
        '
        Me.ColRefParty.HeaderText = "Supplier"
        Me.ColRefParty.Name = "ColRefParty"
        Me.ColRefParty.Width = 85
        '
        'ColTotalKartu
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalKartu.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColTotalKartu.HeaderText = "Total Kartu"
        Me.ColTotalKartu.Name = "ColTotalKartu"
        Me.ColTotalKartu.Width = 103
        '
        'ColTotalBalance
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalBalance.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColTotalBalance.HeaderText = "Total Balance"
        Me.ColTotalBalance.Name = "ColTotalBalance"
        Me.ColTotalBalance.Width = 120
        '
        'ColSelisih
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSelisih.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColSelisih.HeaderText = "Selisih"
        Me.ColSelisih.Name = "ColSelisih"
        Me.ColSelisih.Width = 74
        '
        'ColPrimaryRefParty
        '
        Me.ColPrimaryRefParty.HeaderText = "Primary Ref Party"
        Me.ColPrimaryRefParty.Name = "ColPrimaryRefParty"
        Me.ColPrimaryRefParty.Visible = False
        Me.ColPrimaryRefParty.Width = 144
        '
        'rboAP
        '
        Me.rboAP.AutoSize = True
        Me.rboAP.Checked = True
        Me.rboAP.Location = New System.Drawing.Point(12, 13)
        Me.rboAP.Name = "rboAP"
        Me.rboAP.Size = New System.Drawing.Size(47, 21)
        Me.rboAP.TabIndex = 5
        Me.rboAP.TabStop = True
        Me.rboAP.Text = "AP"
        Me.rboAP.UseVisualStyleBackColor = True
        '
        'rboAR
        '
        Me.rboAR.AutoSize = True
        Me.rboAR.Location = New System.Drawing.Point(12, 42)
        Me.rboAR.Name = "rboAR"
        Me.rboAR.Size = New System.Drawing.Size(48, 21)
        Me.rboAR.TabIndex = 6
        Me.rboAR.Text = "AR"
        Me.rboAR.UseVisualStyleBackColor = True
        '
        'bgw1
        '
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(65, 70)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(241, 25)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Cek Balance vs Balance Original"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmUpdate_RepairAPAR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(805, 396)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.rboAR)
        Me.Controls.Add(Me.rboAP)
        Me.Controls.Add(Me.dgvRepair)
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.prg1)
        Me.Controls.Add(Me.btnRepair)
        Me.Controls.Add(Me.btnCek)
        Me.Name = "FrmUpdate_RepairAPAR"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Update AP AR"
        CType(Me.dgvRepair, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCek As System.Windows.Forms.Button
    Friend WithEvents btnRepair As System.Windows.Forms.Button
    Friend WithEvents prg1 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents dgvRepair As System.Windows.Forms.DataGridView
    Friend WithEvents rboAP As System.Windows.Forms.RadioButton
    Friend WithEvents rboAR As System.Windows.Forms.RadioButton
    Friend WithEvents bgw1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRefParty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalKartu As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSelisih As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryRefParty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
