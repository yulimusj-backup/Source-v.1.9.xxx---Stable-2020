<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterWorkUnit_AddSetupTime
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
        Me.dgvSetupTime = New System.Windows.Forms.DataGridView
        Me.ColSetupTime_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_WorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_ProcessCode1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_CodeNo1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_Desc1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_ProcessCode2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_CodeNo2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_Desc2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_SetupTimeDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_Set = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColSetupTime_PrimaryProcess1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_PrimaryCodeNo1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_PrimaryProses2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_PrimaryCodeNo2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_PrimarySetupTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_PrimaryWorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetupTime_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnSet = New System.Windows.Forms.Button
        Me.lblCheckAll = New System.Windows.Forms.LinkLabel
        Me.lblUncheckedAll = New System.Windows.Forms.LinkLabel
        Me.btnClose = New System.Windows.Forms.Button
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        CType(Me.dgvSetupTime, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvSetupTime
        '
        Me.dgvSetupTime.AllowUserToAddRows = False
        Me.dgvSetupTime.AllowUserToDeleteRows = False
        Me.dgvSetupTime.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSetupTime.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSetupTime.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSetupTime.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvSetupTime.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSetupTime_No, Me.ColSetupTime_WorkUnit, Me.ColSetupTime_ProcessCode1, Me.ColSetupTime_CodeNo1, Me.ColSetupTime_Desc1, Me.ColSetupTime_ProcessCode2, Me.ColSetupTime_CodeNo2, Me.ColSetupTime_Desc2, Me.ColSetupTime_Code, Me.ColSetupTime_SetupTimeDuration, Me.ColSetupTime_Remarks, Me.ColSetupTime_Set, Me.ColSetupTime_PrimaryProcess1, Me.ColSetupTime_PrimaryCodeNo1, Me.ColSetupTime_PrimaryProses2, Me.ColSetupTime_PrimaryCodeNo2, Me.ColSetupTime_PrimarySetupTime, Me.ColSetupTime_PrimaryWorkUnit, Me.ColSetupTime_Primary})
        Me.dgvSetupTime.Location = New System.Drawing.Point(15, 75)
        Me.dgvSetupTime.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dgvSetupTime.MultiSelect = False
        Me.dgvSetupTime.Name = "dgvSetupTime"
        Me.dgvSetupTime.RowTemplate.Height = 24
        Me.dgvSetupTime.Size = New System.Drawing.Size(1087, 700)
        Me.dgvSetupTime.TabIndex = 5
        '
        'ColSetupTime_No
        '
        Me.ColSetupTime_No.Frozen = True
        Me.ColSetupTime_No.HeaderText = "No"
        Me.ColSetupTime_No.Name = "ColSetupTime_No"
        Me.ColSetupTime_No.Width = 54
        '
        'ColSetupTime_WorkUnit
        '
        Me.ColSetupTime_WorkUnit.Frozen = True
        Me.ColSetupTime_WorkUnit.HeaderText = "Work Unit"
        Me.ColSetupTime_WorkUnit.Name = "ColSetupTime_WorkUnit"
        Me.ColSetupTime_WorkUnit.Width = 104
        '
        'ColSetupTime_ProcessCode1
        '
        Me.ColSetupTime_ProcessCode1.Frozen = True
        Me.ColSetupTime_ProcessCode1.HeaderText = "Process 1"
        Me.ColSetupTime_ProcessCode1.MaxInputLength = 50
        Me.ColSetupTime_ProcessCode1.Name = "ColSetupTime_ProcessCode1"
        Me.ColSetupTime_ProcessCode1.Width = 104
        '
        'ColSetupTime_CodeNo1
        '
        Me.ColSetupTime_CodeNo1.Frozen = True
        Me.ColSetupTime_CodeNo1.HeaderText = "Code No 1"
        Me.ColSetupTime_CodeNo1.MaxInputLength = 50
        Me.ColSetupTime_CodeNo1.Name = "ColSetupTime_CodeNo1"
        Me.ColSetupTime_CodeNo1.Width = 109
        '
        'ColSetupTime_Desc1
        '
        Me.ColSetupTime_Desc1.HeaderText = "Description 1"
        Me.ColSetupTime_Desc1.Name = "ColSetupTime_Desc1"
        Me.ColSetupTime_Desc1.Visible = False
        Me.ColSetupTime_Desc1.Width = 127
        '
        'ColSetupTime_ProcessCode2
        '
        Me.ColSetupTime_ProcessCode2.HeaderText = "Process 2"
        Me.ColSetupTime_ProcessCode2.MaxInputLength = 50
        Me.ColSetupTime_ProcessCode2.Name = "ColSetupTime_ProcessCode2"
        Me.ColSetupTime_ProcessCode2.Width = 104
        '
        'ColSetupTime_CodeNo2
        '
        Me.ColSetupTime_CodeNo2.HeaderText = "Code No 2"
        Me.ColSetupTime_CodeNo2.MaxInputLength = 150
        Me.ColSetupTime_CodeNo2.Name = "ColSetupTime_CodeNo2"
        Me.ColSetupTime_CodeNo2.Width = 109
        '
        'ColSetupTime_Desc2
        '
        Me.ColSetupTime_Desc2.HeaderText = "Description 2"
        Me.ColSetupTime_Desc2.Name = "ColSetupTime_Desc2"
        Me.ColSetupTime_Desc2.Visible = False
        Me.ColSetupTime_Desc2.Width = 127
        '
        'ColSetupTime_Code
        '
        Me.ColSetupTime_Code.HeaderText = "Setup Time"
        Me.ColSetupTime_Code.Name = "ColSetupTime_Code"
        Me.ColSetupTime_Code.Width = 115
        '
        'ColSetupTime_SetupTimeDuration
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSetupTime_SetupTimeDuration.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColSetupTime_SetupTimeDuration.HeaderText = "Duration"
        Me.ColSetupTime_SetupTimeDuration.Name = "ColSetupTime_SetupTimeDuration"
        Me.ColSetupTime_SetupTimeDuration.Width = 95
        '
        'ColSetupTime_Remarks
        '
        Me.ColSetupTime_Remarks.HeaderText = "Remarks"
        Me.ColSetupTime_Remarks.Name = "ColSetupTime_Remarks"
        Me.ColSetupTime_Remarks.Width = 98
        '
        'ColSetupTime_Set
        '
        Me.ColSetupTime_Set.HeaderText = "Set"
        Me.ColSetupTime_Set.Name = "ColSetupTime_Set"
        Me.ColSetupTime_Set.Width = 40
        '
        'ColSetupTime_PrimaryProcess1
        '
        Me.ColSetupTime_PrimaryProcess1.HeaderText = "Primary Process 1"
        Me.ColSetupTime_PrimaryProcess1.Name = "ColSetupTime_PrimaryProcess1"
        Me.ColSetupTime_PrimaryProcess1.Visible = False
        Me.ColSetupTime_PrimaryProcess1.Width = 160
        '
        'ColSetupTime_PrimaryCodeNo1
        '
        Me.ColSetupTime_PrimaryCodeNo1.HeaderText = "Primary Code No 1"
        Me.ColSetupTime_PrimaryCodeNo1.Name = "ColSetupTime_PrimaryCodeNo1"
        Me.ColSetupTime_PrimaryCodeNo1.Visible = False
        Me.ColSetupTime_PrimaryCodeNo1.Width = 165
        '
        'ColSetupTime_PrimaryProses2
        '
        Me.ColSetupTime_PrimaryProses2.HeaderText = "Primary Proses 2"
        Me.ColSetupTime_PrimaryProses2.Name = "ColSetupTime_PrimaryProses2"
        Me.ColSetupTime_PrimaryProses2.Visible = False
        Me.ColSetupTime_PrimaryProses2.Width = 152
        '
        'ColSetupTime_PrimaryCodeNo2
        '
        Me.ColSetupTime_PrimaryCodeNo2.HeaderText = "Primary Code No 2"
        Me.ColSetupTime_PrimaryCodeNo2.Name = "ColSetupTime_PrimaryCodeNo2"
        Me.ColSetupTime_PrimaryCodeNo2.Visible = False
        Me.ColSetupTime_PrimaryCodeNo2.Width = 165
        '
        'ColSetupTime_PrimarySetupTime
        '
        Me.ColSetupTime_PrimarySetupTime.HeaderText = "Primary Setup Time"
        Me.ColSetupTime_PrimarySetupTime.Name = "ColSetupTime_PrimarySetupTime"
        Me.ColSetupTime_PrimarySetupTime.Visible = False
        Me.ColSetupTime_PrimarySetupTime.Width = 171
        '
        'ColSetupTime_PrimaryWorkUnit
        '
        Me.ColSetupTime_PrimaryWorkUnit.HeaderText = "PrimaryWorkUnit"
        Me.ColSetupTime_PrimaryWorkUnit.Name = "ColSetupTime_PrimaryWorkUnit"
        Me.ColSetupTime_PrimaryWorkUnit.Visible = False
        Me.ColSetupTime_PrimaryWorkUnit.Width = 152
        '
        'ColSetupTime_Primary
        '
        Me.ColSetupTime_Primary.HeaderText = "Primary"
        Me.ColSetupTime_Primary.Name = "ColSetupTime_Primary"
        Me.ColSetupTime_Primary.Visible = False
        Me.ColSetupTime_Primary.Width = 86
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(457, 20)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Please specify setup time for the following process combination "
        '
        'btnSet
        '
        Me.btnSet.Location = New System.Drawing.Point(14, 38)
        Me.btnSet.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btnSet.Name = "btnSet"
        Me.btnSet.Size = New System.Drawing.Size(84, 29)
        Me.btnSet.TabIndex = 7
        Me.btnSet.Text = "Set"
        Me.btnSet.UseVisualStyleBackColor = True
        '
        'lblCheckAll
        '
        Me.lblCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblCheckAll.AutoSize = True
        Me.lblCheckAll.Location = New System.Drawing.Point(929, 50)
        Me.lblCheckAll.Name = "lblCheckAll"
        Me.lblCheckAll.Size = New System.Drawing.Size(75, 20)
        Me.lblCheckAll.TabIndex = 8
        Me.lblCheckAll.TabStop = True
        Me.lblCheckAll.Text = "Check All"
        '
        'lblUncheckedAll
        '
        Me.lblUncheckedAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUncheckedAll.AutoSize = True
        Me.lblUncheckedAll.Location = New System.Drawing.Point(1010, 50)
        Me.lblUncheckedAll.Name = "lblUncheckedAll"
        Me.lblUncheckedAll.Size = New System.Drawing.Size(93, 20)
        Me.lblUncheckedAll.TabIndex = 9
        Me.lblUncheckedAll.TabStop = True
        Me.lblUncheckedAll.Text = "Uncheck All"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(516, 381)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(84, 29)
        Me.btnClose.TabIndex = 10
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'BackgroundWorker3
        '
        '
        'FrmMasterWorkUnit_AddSetupTime
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1116, 791)
        Me.Controls.Add(Me.dgvSetupTime)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.lblUncheckedAll)
        Me.Controls.Add(Me.lblCheckAll)
        Me.Controls.Add(Me.btnSet)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1124, 827)
        Me.Name = "FrmMasterWorkUnit_AddSetupTime"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Setup Time"
        CType(Me.dgvSetupTime, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvSetupTime As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSet As System.Windows.Forms.Button
    Friend WithEvents lblCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents lblUncheckedAll As System.Windows.Forms.LinkLabel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColSetupTime_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_WorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_ProcessCode1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_CodeNo1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_Desc1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_ProcessCode2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_CodeNo2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_Desc2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_SetupTimeDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_Set As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColSetupTime_PrimaryProcess1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_PrimaryCodeNo1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_PrimaryProses2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_PrimaryCodeNo2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_PrimarySetupTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_PrimaryWorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetupTime_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
