<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_NewSch2_ColumnSetting
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
        Me.ChkWorkHour = New System.Windows.Forms.CheckBox
        Me.ChkUnit = New System.Windows.Forms.CheckBox
        Me.ChkQty = New System.Windows.Forms.CheckBox
        Me.ChkRate = New System.Windows.Forms.CheckBox
        Me.ChkShowProgress = New System.Windows.Forms.CheckBox
        Me.ChkShowProcess = New System.Windows.Forms.CheckBox
        Me.ChkDayFinish = New System.Windows.Forms.CheckBox
        Me.ChkDayStart = New System.Windows.Forms.CheckBox
        Me.ChkDesc = New System.Windows.Forms.CheckBox
        Me.ChkCodeNo = New System.Windows.Forms.CheckBox
        Me.ChkWONo = New System.Windows.Forms.CheckBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.ChkOutputLocation = New System.Windows.Forms.CheckBox
        Me.ChkInputLocation = New System.Windows.Forms.CheckBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvItemConfig = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryMasterLine = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ChkProduceQty = New System.Windows.Forms.CheckBox
        Me.ChkGoodQty = New System.Windows.Forms.CheckBox
        Me.ChkRejectQty = New System.Windows.Forms.CheckBox
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvItemConfig, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ChkWorkHour
        '
        Me.ChkWorkHour.AutoSize = True
        Me.ChkWorkHour.Location = New System.Drawing.Point(138, 29)
        Me.ChkWorkHour.Name = "ChkWorkHour"
        Me.ChkWorkHour.Size = New System.Drawing.Size(78, 17)
        Me.ChkWorkHour.TabIndex = 7
        Me.ChkWorkHour.Text = "Work Hour"
        Me.ChkWorkHour.UseVisualStyleBackColor = True
        '
        'ChkUnit
        '
        Me.ChkUnit.AutoSize = True
        Me.ChkUnit.Location = New System.Drawing.Point(138, 6)
        Me.ChkUnit.Name = "ChkUnit"
        Me.ChkUnit.Size = New System.Drawing.Size(45, 17)
        Me.ChkUnit.TabIndex = 6
        Me.ChkUnit.Text = "Unit"
        Me.ChkUnit.UseVisualStyleBackColor = True
        '
        'ChkQty
        '
        Me.ChkQty.AutoSize = True
        Me.ChkQty.Location = New System.Drawing.Point(6, 121)
        Me.ChkQty.Name = "ChkQty"
        Me.ChkQty.Size = New System.Drawing.Size(42, 17)
        Me.ChkQty.TabIndex = 5
        Me.ChkQty.Text = "Qty"
        Me.ChkQty.UseVisualStyleBackColor = True
        '
        'ChkRate
        '
        Me.ChkRate.AutoSize = True
        Me.ChkRate.Location = New System.Drawing.Point(6, 98)
        Me.ChkRate.Name = "ChkRate"
        Me.ChkRate.Size = New System.Drawing.Size(49, 17)
        Me.ChkRate.TabIndex = 4
        Me.ChkRate.Text = "Rate"
        Me.ChkRate.UseVisualStyleBackColor = True
        '
        'ChkShowProgress
        '
        Me.ChkShowProgress.AutoSize = True
        Me.ChkShowProgress.Location = New System.Drawing.Point(138, 98)
        Me.ChkShowProgress.Name = "ChkShowProgress"
        Me.ChkShowProgress.Size = New System.Drawing.Size(67, 17)
        Me.ChkShowProgress.TabIndex = 10
        Me.ChkShowProgress.Text = "Progress"
        Me.ChkShowProgress.UseVisualStyleBackColor = True
        '
        'ChkShowProcess
        '
        Me.ChkShowProcess.AutoSize = True
        Me.ChkShowProcess.Location = New System.Drawing.Point(6, 29)
        Me.ChkShowProcess.Name = "ChkShowProcess"
        Me.ChkShowProcess.Size = New System.Drawing.Size(64, 17)
        Me.ChkShowProcess.TabIndex = 1
        Me.ChkShowProcess.Text = "Process"
        Me.ChkShowProcess.UseVisualStyleBackColor = True
        '
        'ChkDayFinish
        '
        Me.ChkDayFinish.AutoSize = True
        Me.ChkDayFinish.Location = New System.Drawing.Point(138, 75)
        Me.ChkDayFinish.Name = "ChkDayFinish"
        Me.ChkDayFinish.Size = New System.Drawing.Size(119, 17)
        Me.ChkDayFinish.TabIndex = 9
        Me.ChkDayFinish.Text = "Day of Week Finish"
        Me.ChkDayFinish.UseVisualStyleBackColor = True
        '
        'ChkDayStart
        '
        Me.ChkDayStart.AutoSize = True
        Me.ChkDayStart.Location = New System.Drawing.Point(138, 52)
        Me.ChkDayStart.Name = "ChkDayStart"
        Me.ChkDayStart.Size = New System.Drawing.Size(114, 17)
        Me.ChkDayStart.TabIndex = 8
        Me.ChkDayStart.Text = "Day of Week Start"
        Me.ChkDayStart.UseVisualStyleBackColor = True
        '
        'ChkDesc
        '
        Me.ChkDesc.AutoSize = True
        Me.ChkDesc.Location = New System.Drawing.Point(6, 75)
        Me.ChkDesc.Name = "ChkDesc"
        Me.ChkDesc.Size = New System.Drawing.Size(79, 17)
        Me.ChkDesc.TabIndex = 3
        Me.ChkDesc.Text = "Description"
        Me.ChkDesc.UseVisualStyleBackColor = True
        '
        'ChkCodeNo
        '
        Me.ChkCodeNo.AutoSize = True
        Me.ChkCodeNo.Location = New System.Drawing.Point(6, 52)
        Me.ChkCodeNo.Name = "ChkCodeNo"
        Me.ChkCodeNo.Size = New System.Drawing.Size(68, 17)
        Me.ChkCodeNo.TabIndex = 2
        Me.ChkCodeNo.Text = "Code No"
        Me.ChkCodeNo.UseVisualStyleBackColor = True
        '
        'ChkWONo
        '
        Me.ChkWONo.AutoSize = True
        Me.ChkWONo.Location = New System.Drawing.Point(6, 6)
        Me.ChkWONo.Name = "ChkWONo"
        Me.ChkWONo.Size = New System.Drawing.Size(62, 17)
        Me.ChkWONo.TabIndex = 0
        Me.ChkWONo.Text = "WO No"
        Me.ChkWONo.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(350, 176)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(66, 22)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "&OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(7, 3)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(413, 171)
        Me.TabControl1.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.ChkRejectQty)
        Me.TabPage1.Controls.Add(Me.ChkGoodQty)
        Me.TabPage1.Controls.Add(Me.ChkProduceQty)
        Me.TabPage1.Controls.Add(Me.ChkOutputLocation)
        Me.TabPage1.Controls.Add(Me.ChkInputLocation)
        Me.TabPage1.Controls.Add(Me.ChkWorkHour)
        Me.TabPage1.Controls.Add(Me.ChkWONo)
        Me.TabPage1.Controls.Add(Me.ChkUnit)
        Me.TabPage1.Controls.Add(Me.ChkCodeNo)
        Me.TabPage1.Controls.Add(Me.ChkQty)
        Me.TabPage1.Controls.Add(Me.ChkDesc)
        Me.TabPage1.Controls.Add(Me.ChkRate)
        Me.TabPage1.Controls.Add(Me.ChkDayStart)
        Me.TabPage1.Controls.Add(Me.ChkShowProgress)
        Me.TabPage1.Controls.Add(Me.ChkDayFinish)
        Me.TabPage1.Controls.Add(Me.ChkShowProcess)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(405, 145)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Column Settings"
        '
        'ChkOutputLocation
        '
        Me.ChkOutputLocation.AutoSize = True
        Me.ChkOutputLocation.Location = New System.Drawing.Point(282, 6)
        Me.ChkOutputLocation.Name = "ChkOutputLocation"
        Me.ChkOutputLocation.Size = New System.Drawing.Size(102, 17)
        Me.ChkOutputLocation.TabIndex = 12
        Me.ChkOutputLocation.Text = "Output Location"
        Me.ChkOutputLocation.UseVisualStyleBackColor = True
        '
        'ChkInputLocation
        '
        Me.ChkInputLocation.AutoSize = True
        Me.ChkInputLocation.Location = New System.Drawing.Point(138, 121)
        Me.ChkInputLocation.Name = "ChkInputLocation"
        Me.ChkInputLocation.Size = New System.Drawing.Size(94, 17)
        Me.ChkInputLocation.TabIndex = 11
        Me.ChkInputLocation.Text = "Input Location"
        Me.ChkInputLocation.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage2.Controls.Add(Me.dgvItemConfig)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(411, 130)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Line Filter"
        '
        'dgvItemConfig
        '
        Me.dgvItemConfig.AllowUserToResizeColumns = False
        Me.dgvItemConfig.AllowUserToResizeRows = False
        Me.dgvItemConfig.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemConfig.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemConfig.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCode, Me.ColDesc, Me.ColRemarks, Me.ColPrimaryMasterLine})
        Me.dgvItemConfig.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemConfig.Location = New System.Drawing.Point(3, 3)
        Me.dgvItemConfig.Name = "dgvItemConfig"
        Me.dgvItemConfig.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemConfig.RowTemplate.Height = 24
        Me.dgvItemConfig.Size = New System.Drawing.Size(405, 124)
        Me.dgvItemConfig.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColCode
        '
        Me.ColCode.Frozen = True
        Me.ColCode.HeaderText = "Code"
        Me.ColCode.Name = "ColCode"
        Me.ColCode.Width = 57
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.Width = 85
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.ReadOnly = True
        Me.ColRemarks.Width = 74
        '
        'ColPrimaryMasterLine
        '
        Me.ColPrimaryMasterLine.HeaderText = "Primary"
        Me.ColPrimaryMasterLine.Name = "ColPrimaryMasterLine"
        Me.ColPrimaryMasterLine.ReadOnly = True
        Me.ColPrimaryMasterLine.Visible = False
        Me.ColPrimaryMasterLine.Width = 66
        '
        'BackgroundWorker1
        '
        '
        'ChkProduceQty
        '
        Me.ChkProduceQty.AutoSize = True
        Me.ChkProduceQty.Location = New System.Drawing.Point(282, 75)
        Me.ChkProduceQty.Name = "ChkProduceQty"
        Me.ChkProduceQty.Size = New System.Drawing.Size(118, 17)
        Me.ChkProduceQty.TabIndex = 15
        Me.ChkProduceQty.Text = "Actual Produce Qty"
        Me.ChkProduceQty.UseVisualStyleBackColor = True
        '
        'ChkGoodQty
        '
        Me.ChkGoodQty.AutoSize = True
        Me.ChkGoodQty.Location = New System.Drawing.Point(282, 29)
        Me.ChkGoodQty.Name = "ChkGoodQty"
        Me.ChkGoodQty.Size = New System.Drawing.Size(104, 17)
        Me.ChkGoodQty.TabIndex = 13
        Me.ChkGoodQty.Text = "Actual Good Qty"
        Me.ChkGoodQty.UseVisualStyleBackColor = True
        '
        'ChkRejectQty
        '
        Me.ChkRejectQty.AutoSize = True
        Me.ChkRejectQty.Location = New System.Drawing.Point(282, 52)
        Me.ChkRejectQty.Name = "ChkRejectQty"
        Me.ChkRejectQty.Size = New System.Drawing.Size(109, 17)
        Me.ChkRejectQty.TabIndex = 14
        Me.ChkRejectQty.Text = "Actual Reject Qty"
        Me.ChkRejectQty.UseVisualStyleBackColor = True
        '
        'FrmTrans_ProductionSchedule_NewSch2_ColumnSetting
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(427, 201)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Button1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(443, 215)
        Me.Name = "FrmTrans_ProductionSchedule_NewSch2_ColumnSetting"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Column Settings & Data Filter"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvItemConfig, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ChkDayFinish As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDayStart As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDesc As System.Windows.Forms.CheckBox
    Friend WithEvents ChkCodeNo As System.Windows.Forms.CheckBox
    Friend WithEvents ChkWONo As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ChkShowProcess As System.Windows.Forms.CheckBox
    Friend WithEvents ChkRate As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowProgress As System.Windows.Forms.CheckBox
    Friend WithEvents ChkWorkHour As System.Windows.Forms.CheckBox
    Friend WithEvents ChkUnit As System.Windows.Forms.CheckBox
    Friend WithEvents ChkQty As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvItemConfig As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryMasterLine As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChkOutputLocation As System.Windows.Forms.CheckBox
    Friend WithEvents ChkInputLocation As System.Windows.Forms.CheckBox
    Friend WithEvents ChkRejectQty As System.Windows.Forms.CheckBox
    Friend WithEvents ChkGoodQty As System.Windows.Forms.CheckBox
    Friend WithEvents ChkProduceQty As System.Windows.Forms.CheckBox
End Class
