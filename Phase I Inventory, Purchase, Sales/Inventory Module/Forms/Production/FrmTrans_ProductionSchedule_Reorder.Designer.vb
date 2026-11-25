<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_Reorder
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColModify = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colIdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colIdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMoNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWH = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStartDay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStartHour = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndDay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndHour = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProgress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPlanStartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.BtnOK = New System.Windows.Forms.Button
        Me.btnUp = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnUpJO = New System.Windows.Forms.Button
        Me.btnDownJO = New System.Windows.Forms.Button
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.LlbColumnsSetting = New System.Windows.Forms.LinkLabel
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColModify, Me.colIdentifierJO, Me.colIdentifierWO, Me.ColMoNo, Me.ColJONo, Me.ColProcessCode, Me.ColWONo, Me.ColCodeNo, Me.ColDesc, Me.ColRate, Me.ColQty, Me.ColUnit, Me.ColWH, Me.ColStartDay, Me.ColStartDate, Me.ColStartHour, Me.ColEndDay, Me.ColEndDate, Me.ColEndHour, Me.ColProgress, Me.ColPlanStartTime})
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.MenuHighlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvItem.DefaultCellStyle = DataGridViewCellStyle4
        Me.dgvItem.Location = New System.Drawing.Point(12, 28)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersVisible = False
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvItem.Size = New System.Drawing.Size(971, 565)
        Me.dgvItem.TabIndex = 0
        '
        'ColModify
        '
        Me.ColModify.Frozen = True
        Me.ColModify.HeaderText = "Modify"
        Me.ColModify.Name = "ColModify"
        Me.ColModify.Width = 55
        '
        'colIdentifierJO
        '
        Me.colIdentifierJO.HeaderText = "identifier jo"
        Me.colIdentifierJO.Name = "colIdentifierJO"
        Me.colIdentifierJO.ReadOnly = True
        Me.colIdentifierJO.Visible = False
        Me.colIdentifierJO.Width = 102
        '
        'colIdentifierWO
        '
        Me.colIdentifierWO.HeaderText = "identifier wo"
        Me.colIdentifierWO.Name = "colIdentifierWO"
        Me.colIdentifierWO.ReadOnly = True
        Me.colIdentifierWO.Visible = False
        Me.colIdentifierWO.Width = 108
        '
        'ColMoNo
        '
        Me.ColMoNo.Frozen = True
        Me.ColMoNo.HeaderText = "Mo No"
        Me.ColMoNo.Name = "ColMoNo"
        Me.ColMoNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMoNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMoNo.Width = 55
        '
        'ColJONo
        '
        Me.ColJONo.Frozen = True
        Me.ColJONo.HeaderText = "JO No"
        Me.ColJONo.Name = "ColJONo"
        Me.ColJONo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColJONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJONo.Width = 54
        '
        'ColProcessCode
        '
        Me.ColProcessCode.Frozen = True
        Me.ColProcessCode.HeaderText = "Process"
        Me.ColProcessCode.Name = "ColProcessCode"
        Me.ColProcessCode.Width = 84
        '
        'ColWONo
        '
        Me.ColWONo.Frozen = True
        Me.ColWONo.HeaderText = "WO No"
        Me.ColWONo.Name = "ColWONo"
        Me.ColWONo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWONo.Width = 60
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.Width = 88
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.Width = 104
        '
        'ColRate
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRate.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColRate.HeaderText = "Rate (Unit/Hour)"
        Me.ColRate.Name = "ColRate"
        Me.ColRate.Width = 137
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Width = 55
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 58
        '
        'ColWH
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColWH.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColWH.HeaderText = "W/H"
        Me.ColWH.Name = "ColWH"
        Me.ColWH.Width = 60
        '
        'ColStartDay
        '
        Me.ColStartDay.HeaderText = "Start"
        Me.ColStartDay.Name = "ColStartDay"
        Me.ColStartDay.Width = 63
        '
        'ColStartDate
        '
        Me.ColStartDate.HeaderText = ""
        Me.ColStartDate.Name = "ColStartDate"
        Me.ColStartDate.Width = 19
        '
        'ColStartHour
        '
        Me.ColStartHour.HeaderText = ""
        Me.ColStartHour.Name = "ColStartHour"
        Me.ColStartHour.Width = 19
        '
        'ColEndDay
        '
        Me.ColEndDay.HeaderText = "End"
        Me.ColEndDay.Name = "ColEndDay"
        Me.ColEndDay.Width = 58
        '
        'ColEndDate
        '
        Me.ColEndDate.HeaderText = ""
        Me.ColEndDate.Name = "ColEndDate"
        Me.ColEndDate.Width = 19
        '
        'ColEndHour
        '
        Me.ColEndHour.HeaderText = ""
        Me.ColEndHour.Name = "ColEndHour"
        Me.ColEndHour.Width = 19
        '
        'ColProgress
        '
        Me.ColProgress.HeaderText = "Progress"
        Me.ColProgress.Name = "ColProgress"
        Me.ColProgress.Width = 90
        '
        'ColPlanStartTime
        '
        Me.ColPlanStartTime.HeaderText = "Plan Start Time"
        Me.ColPlanStartTime.Name = "ColPlanStartTime"
        Me.ColPlanStartTime.ReadOnly = True
        Me.ColPlanStartTime.Visible = False
        Me.ColPlanStartTime.Width = 130
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnCancel.Location = New System.Drawing.Point(908, 597)
        Me.BtnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(75, 26)
        Me.BtnCancel.TabIndex = 1
        Me.BtnCancel.Text = "Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'BtnOK
        '
        Me.BtnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnOK.Location = New System.Drawing.Point(827, 597)
        Me.BtnOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(75, 26)
        Me.BtnOK.TabIndex = 2
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'btnUp
        '
        Me.btnUp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUp.Location = New System.Drawing.Point(12, 597)
        Me.btnUp.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnUp.Name = "btnUp"
        Me.btnUp.Size = New System.Drawing.Size(111, 26)
        Me.btnUp.TabIndex = 4
        Me.btnUp.Text = "Move Up"
        Me.btnUp.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.Location = New System.Drawing.Point(128, 597)
        Me.btnDown.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(111, 26)
        Me.btnDown.TabIndex = 3
        Me.btnDown.Text = "Move Down"
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'btnUpJO
        '
        Me.btnUpJO.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUpJO.Location = New System.Drawing.Point(244, 597)
        Me.btnUpJO.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnUpJO.Name = "btnUpJO"
        Me.btnUpJO.Size = New System.Drawing.Size(163, 26)
        Me.btnUpJO.TabIndex = 6
        Me.btnUpJO.Text = "Move Up (JO level)"
        Me.btnUpJO.UseVisualStyleBackColor = True
        '
        'btnDownJO
        '
        Me.btnDownJO.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDownJO.Location = New System.Drawing.Point(412, 597)
        Me.btnDownJO.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnDownJO.Name = "btnDownJO"
        Me.btnDownJO.Size = New System.Drawing.Size(163, 26)
        Me.btnDownJO.TabIndex = 5
        Me.btnDownJO.Text = "Move Down (JO level)"
        Me.btnDownJO.UseVisualStyleBackColor = True
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(777, 9)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(118, 17)
        Me.LinkLabel1.TabIndex = 7
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Check All Item JO"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(901, 9)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(82, 17)
        Me.LinkLabel2.TabIndex = 8
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Uncheck All"
        '
        'LlbColumnsSetting
        '
        Me.LlbColumnsSetting.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LlbColumnsSetting.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbColumnsSetting.AutoSize = True
        Me.LlbColumnsSetting.Location = New System.Drawing.Point(668, 9)
        Me.LlbColumnsSetting.Name = "LlbColumnsSetting"
        Me.LlbColumnsSetting.Size = New System.Drawing.Size(103, 17)
        Me.LlbColumnsSetting.TabIndex = 9
        Me.LlbColumnsSetting.TabStop = True
        Me.LlbColumnsSetting.Text = "Column Setting"
        Me.LlbColumnsSetting.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'BackgroundWorker1
        '
        '
        'Panel1
        '
        Me.Panel1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Location = New System.Drawing.Point(245, 286)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(504, 62)
        Me.Panel1.TabIndex = 22
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 34)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Label2"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(4, 9)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(4)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(497, 22)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.ProgressBar1.TabIndex = 0
        Me.ProgressBar1.Value = 25
        '
        'FrmTrans_ProductionSchedule_Reorder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(995, 634)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.LlbColumnsSetting)
        Me.Controls.Add(Me.LinkLabel2)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.btnUpJO)
        Me.Controls.Add(Me.btnDownJO)
        Me.Controls.Add(Me.btnUp)
        Me.Controls.Add(Me.btnDown)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.dgvItem)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1003, 671)
        Me.Name = "FrmTrans_ProductionSchedule_Reorder"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Reorder work schedule"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents btnUp As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnUpJO As System.Windows.Forms.Button
    Friend WithEvents btnDownJO As System.Windows.Forms.Button
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbColumnsSetting As System.Windows.Forms.LinkLabel
    Friend WithEvents ColModify As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colIdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colIdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMoNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWH As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStartDay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStartHour As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndDay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndHour As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProgress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPlanStartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
End Class
