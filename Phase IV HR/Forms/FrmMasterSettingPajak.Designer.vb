<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterSettingPajak
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
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvPTKP = New System.Windows.Forms.DataGridView
        Me.ColPTKP_Golongan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPTKP_PTKP_Tahunan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPTKP_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPTKP_Enum = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPTKP_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvRate = New System.Windows.Forms.DataGridView
        Me.ColRate_Maximum = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRate_Tarif = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRate_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRate_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dgvRateFinal = New System.Windows.Forms.DataGridView
        Me.ColRateFinal_Maksimum = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRateFinal_Tarif = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRateFinal_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRateFinal_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvPTKP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvRate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvRateFinal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(16, 15)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(777, 449)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvPTKP)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabPage1.Size = New System.Drawing.Size(769, 420)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "PTKP"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvPTKP
        '
        Me.dgvPTKP.AllowUserToAddRows = False
        Me.dgvPTKP.AllowUserToDeleteRows = False
        Me.dgvPTKP.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvPTKP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColPTKP_Golongan, Me.ColPTKP_PTKP_Tahunan, Me.ColPTKP_Remarks, Me.ColPTKP_Enum, Me.ColPTKP_Primary})
        Me.dgvPTKP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvPTKP.Location = New System.Drawing.Point(4, 4)
        Me.dgvPTKP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvPTKP.Name = "dgvPTKP"
        Me.dgvPTKP.Size = New System.Drawing.Size(761, 412)
        Me.dgvPTKP.TabIndex = 0
        '
        'ColPTKP_Golongan
        '
        Me.ColPTKP_Golongan.HeaderText = "Criteria"
        Me.ColPTKP_Golongan.Name = "ColPTKP_Golongan"
        Me.ColPTKP_Golongan.Width = 78
        '
        'ColPTKP_PTKP_Tahunan
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPTKP_PTKP_Tahunan.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColPTKP_PTKP_Tahunan.HeaderText = "PTKP Tahunan"
        Me.ColPTKP_PTKP_Tahunan.Name = "ColPTKP_PTKP_Tahunan"
        Me.ColPTKP_PTKP_Tahunan.Width = 130
        '
        'ColPTKP_Remarks
        '
        Me.ColPTKP_Remarks.HeaderText = "Remarks"
        Me.ColPTKP_Remarks.MaxInputLength = 150
        Me.ColPTKP_Remarks.Name = "ColPTKP_Remarks"
        Me.ColPTKP_Remarks.Width = 89
        '
        'ColPTKP_Enum
        '
        Me.ColPTKP_Enum.HeaderText = "Enum"
        Me.ColPTKP_Enum.Name = "ColPTKP_Enum"
        Me.ColPTKP_Enum.Visible = False
        Me.ColPTKP_Enum.Width = 59
        '
        'ColPTKP_Primary
        '
        Me.ColPTKP_Primary.HeaderText = "Primary"
        Me.ColPTKP_Primary.Name = "ColPTKP_Primary"
        Me.ColPTKP_Primary.Visible = False
        Me.ColPTKP_Primary.Width = 66
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvRate)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabPage2.Size = New System.Drawing.Size(769, 420)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Rate"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvRate
        '
        Me.dgvRate.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvRate.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColRate_Maximum, Me.ColRate_Tarif, Me.ColRate_Remarks, Me.ColRate_Primary})
        Me.dgvRate.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvRate.Location = New System.Drawing.Point(4, 4)
        Me.dgvRate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvRate.Name = "dgvRate"
        Me.dgvRate.RowTemplate.Height = 24
        Me.dgvRate.Size = New System.Drawing.Size(761, 412)
        Me.dgvRate.TabIndex = 1
        '
        'ColRate_Maximum
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRate_Maximum.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColRate_Maximum.HeaderText = "Maximum Income"
        Me.ColRate_Maximum.Name = "ColRate_Maximum"
        Me.ColRate_Maximum.Width = 140
        '
        'ColRate_Tarif
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRate_Tarif.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColRate_Tarif.HeaderText = "Tarif (%)"
        Me.ColRate_Tarif.Name = "ColRate_Tarif"
        Me.ColRate_Tarif.Width = 88
        '
        'ColRate_Remarks
        '
        Me.ColRate_Remarks.HeaderText = "Remarks"
        Me.ColRate_Remarks.MaxInputLength = 150
        Me.ColRate_Remarks.Name = "ColRate_Remarks"
        Me.ColRate_Remarks.Width = 89
        '
        'ColRate_Primary
        '
        Me.ColRate_Primary.HeaderText = "Primary"
        Me.ColRate_Primary.Name = "ColRate_Primary"
        Me.ColRate_Primary.Visible = False
        Me.ColRate_Primary.Width = 66
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(693, 471)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 28)
        Me.btnClose.TabIndex = 0
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(585, 471)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(100, 28)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dgvRateFinal)
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(769, 420)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Rate PPH Final Pesangon"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'dgvRateFinal
        '
        Me.dgvRateFinal.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvRateFinal.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColRateFinal_Maksimum, Me.ColRateFinal_Tarif, Me.ColRateFinal_Remarks, Me.ColRateFinal_Primary})
        Me.dgvRateFinal.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvRateFinal.Location = New System.Drawing.Point(3, 3)
        Me.dgvRateFinal.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvRateFinal.Name = "dgvRateFinal"
        Me.dgvRateFinal.RowTemplate.Height = 24
        Me.dgvRateFinal.Size = New System.Drawing.Size(763, 414)
        Me.dgvRateFinal.TabIndex = 2
        '
        'ColRateFinal_Maksimum
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRateFinal_Maksimum.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColRateFinal_Maksimum.HeaderText = "Maximum Income"
        Me.ColRateFinal_Maksimum.Name = "ColRateFinal_Maksimum"
        Me.ColRateFinal_Maksimum.Width = 140
        '
        'ColRateFinal_Tarif
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRateFinal_Tarif.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColRateFinal_Tarif.HeaderText = "Tarif (%)"
        Me.ColRateFinal_Tarif.Name = "ColRateFinal_Tarif"
        Me.ColRateFinal_Tarif.Width = 88
        '
        'ColRateFinal_Remarks
        '
        Me.ColRateFinal_Remarks.HeaderText = "Remarks"
        Me.ColRateFinal_Remarks.MaxInputLength = 150
        Me.ColRateFinal_Remarks.Name = "ColRateFinal_Remarks"
        Me.ColRateFinal_Remarks.Width = 89
        '
        'ColRateFinal_Primary
        '
        Me.ColRateFinal_Primary.HeaderText = "Primary"
        Me.ColRateFinal_Primary.Name = "ColRateFinal_Primary"
        Me.ColRateFinal_Primary.Visible = False
        Me.ColRateFinal_Primary.Width = 81
        '
        'FrmMasterSettingPajak
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(809, 514)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.TabControl1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimizeBox = False
        Me.Name = "FrmMasterSettingPajak"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Income Tax Setting"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvPTKP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvRate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvRateFinal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents dgvPTKP As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents dgvRate As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColPTKP_Golongan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPTKP_PTKP_Tahunan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPTKP_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPTKP_Enum As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPTKP_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRate_Maximum As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRate_Tarif As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRate_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRate_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents dgvRateFinal As System.Windows.Forms.DataGridView
    Friend WithEvents ColRateFinal_Maksimum As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRateFinal_Tarif As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRateFinal_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRateFinal_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
End Class
