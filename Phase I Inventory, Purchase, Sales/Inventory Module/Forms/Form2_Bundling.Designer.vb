<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2_Bundling
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
        Me.btnFindAccSumber = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtAccSumberCode = New System.Windows.Forms.TextBox
        Me.TxtAccSumberDesc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAccSumber = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAccTarget = New System.Windows.Forms.TextBox
        Me.TxtAccTargetDesc = New System.Windows.Forms.TextBox
        Me.TxtAccTargetCode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnFindAccTarget = New System.Windows.Forms.Button
        Me.TxtPrimaryLoc = New System.Windows.Forms.TextBox
        Me.TxtLocDesc = New System.Windows.Forms.TextBox
        Me.TxtLocCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnFindLoc = New System.Windows.Forms.Button
        Me.dtpBundling1 = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.dtpBundling2 = New System.Windows.Forms.DateTimePicker
        Me.btnStart = New System.Windows.Forms.Button
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.LblProgress = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvSumber = New System.Windows.Forms.DataGridView
        Me.ColSumber_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSumber_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSumber_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSumber_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSumber_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvTarget = New System.Windows.Forms.DataGridView
        Me.ColTarget_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTarget_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTarget_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTarget_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTarget_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnDeleteBundling = New System.Windows.Forms.Button
        Me.btnDeleteStock = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvSumber, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvTarget, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnFindAccSumber
        '
        Me.btnFindAccSumber.Location = New System.Drawing.Point(266, 12)
        Me.btnFindAccSumber.Name = "btnFindAccSumber"
        Me.btnFindAccSumber.Size = New System.Drawing.Size(75, 23)
        Me.btnFindAccSumber.TabIndex = 0
        Me.btnFindAccSumber.Text = "Find"
        Me.btnFindAccSumber.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Account Sumber"
        '
        'TxtAccSumberCode
        '
        Me.TxtAccSumberCode.Location = New System.Drawing.Point(132, 12)
        Me.TxtAccSumberCode.Name = "TxtAccSumberCode"
        Me.TxtAccSumberCode.Size = New System.Drawing.Size(128, 22)
        Me.TxtAccSumberCode.TabIndex = 2
        '
        'TxtAccSumberDesc
        '
        Me.TxtAccSumberDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccSumberDesc.Location = New System.Drawing.Point(347, 12)
        Me.TxtAccSumberDesc.Name = "TxtAccSumberDesc"
        Me.TxtAccSumberDesc.ReadOnly = True
        Me.TxtAccSumberDesc.Size = New System.Drawing.Size(298, 22)
        Me.TxtAccSumberDesc.TabIndex = 3
        '
        'TxtPrimaryAccSumber
        '
        Me.TxtPrimaryAccSumber.Location = New System.Drawing.Point(347, 12)
        Me.TxtPrimaryAccSumber.Name = "TxtPrimaryAccSumber"
        Me.TxtPrimaryAccSumber.Size = New System.Drawing.Size(46, 22)
        Me.TxtPrimaryAccSumber.TabIndex = 4
        Me.TxtPrimaryAccSumber.TabStop = False
        Me.TxtPrimaryAccSumber.Text = "-1"
        Me.TxtPrimaryAccSumber.Visible = False
        '
        'TxtPrimaryAccTarget
        '
        Me.TxtPrimaryAccTarget.Location = New System.Drawing.Point(347, 40)
        Me.TxtPrimaryAccTarget.Name = "TxtPrimaryAccTarget"
        Me.TxtPrimaryAccTarget.Size = New System.Drawing.Size(46, 22)
        Me.TxtPrimaryAccTarget.TabIndex = 9
        Me.TxtPrimaryAccTarget.TabStop = False
        Me.TxtPrimaryAccTarget.Text = "-1"
        Me.TxtPrimaryAccTarget.Visible = False
        '
        'TxtAccTargetDesc
        '
        Me.TxtAccTargetDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccTargetDesc.Location = New System.Drawing.Point(347, 40)
        Me.TxtAccTargetDesc.Name = "TxtAccTargetDesc"
        Me.TxtAccTargetDesc.ReadOnly = True
        Me.TxtAccTargetDesc.Size = New System.Drawing.Size(401, 22)
        Me.TxtAccTargetDesc.TabIndex = 8
        '
        'TxtAccTargetCode
        '
        Me.TxtAccTargetCode.Location = New System.Drawing.Point(132, 40)
        Me.TxtAccTargetCode.Name = "TxtAccTargetCode"
        Me.TxtAccTargetCode.Size = New System.Drawing.Size(128, 22)
        Me.TxtAccTargetCode.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(105, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Account Target"
        '
        'btnFindAccTarget
        '
        Me.btnFindAccTarget.Location = New System.Drawing.Point(266, 40)
        Me.btnFindAccTarget.Name = "btnFindAccTarget"
        Me.btnFindAccTarget.Size = New System.Drawing.Size(75, 23)
        Me.btnFindAccTarget.TabIndex = 5
        Me.btnFindAccTarget.Text = "Find"
        Me.btnFindAccTarget.UseVisualStyleBackColor = True
        '
        'TxtPrimaryLoc
        '
        Me.TxtPrimaryLoc.Location = New System.Drawing.Point(347, 68)
        Me.TxtPrimaryLoc.Name = "TxtPrimaryLoc"
        Me.TxtPrimaryLoc.Size = New System.Drawing.Size(46, 22)
        Me.TxtPrimaryLoc.TabIndex = 14
        Me.TxtPrimaryLoc.TabStop = False
        Me.TxtPrimaryLoc.Text = "-1"
        Me.TxtPrimaryLoc.Visible = False
        '
        'TxtLocDesc
        '
        Me.TxtLocDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLocDesc.Location = New System.Drawing.Point(347, 68)
        Me.TxtLocDesc.Name = "TxtLocDesc"
        Me.TxtLocDesc.ReadOnly = True
        Me.TxtLocDesc.Size = New System.Drawing.Size(401, 22)
        Me.TxtLocDesc.TabIndex = 13
        '
        'TxtLocCode
        '
        Me.TxtLocCode.Location = New System.Drawing.Point(132, 68)
        Me.TxtLocCode.Name = "TxtLocCode"
        Me.TxtLocCode.Size = New System.Drawing.Size(128, 22)
        Me.TxtLocCode.TabIndex = 12
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 71)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 17)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Lokasi"
        '
        'btnFindLoc
        '
        Me.btnFindLoc.Location = New System.Drawing.Point(266, 68)
        Me.btnFindLoc.Name = "btnFindLoc"
        Me.btnFindLoc.Size = New System.Drawing.Size(75, 23)
        Me.btnFindLoc.TabIndex = 10
        Me.btnFindLoc.Text = "Find"
        Me.btnFindLoc.UseVisualStyleBackColor = True
        '
        'dtpBundling1
        '
        Me.dtpBundling1.Location = New System.Drawing.Point(132, 97)
        Me.dtpBundling1.Name = "dtpBundling1"
        Me.dtpBundling1.Size = New System.Drawing.Size(209, 22)
        Me.dtpBundling1.TabIndex = 15
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 100)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(95, 17)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Tgl Bundling1"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 128)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(91, 17)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "TglBundling2"
        '
        'dtpBundling2
        '
        Me.dtpBundling2.Location = New System.Drawing.Point(132, 125)
        Me.dtpBundling2.Name = "dtpBundling2"
        Me.dtpBundling2.Size = New System.Drawing.Size(209, 22)
        Me.dtpBundling2.TabIndex = 17
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(347, 97)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 25)
        Me.btnStart.TabIndex = 19
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(132, 153)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(616, 23)
        Me.ProgressBar1.TabIndex = 20
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        '
        'LblProgress
        '
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(129, 179)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(49, 17)
        Me.LblProgress.TabIndex = 21
        Me.LblProgress.Text = "Ready"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 199)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(736, 407)
        Me.TabControl1.TabIndex = 22
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvSumber)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(728, 378)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Sumber"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvSumber
        '
        Me.dgvSumber.AllowUserToAddRows = False
        Me.dgvSumber.AllowUserToDeleteRows = False
        Me.dgvSumber.AllowUserToResizeColumns = False
        Me.dgvSumber.AllowUserToResizeRows = False
        Me.dgvSumber.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSumber.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSumber.ColumnHeadersHeight = 25
        Me.dgvSumber.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvSumber.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSumber_No, Me.ColSumber_CodeNo, Me.ColSumber_Desc, Me.ColSumber_Qty, Me.ColSumber_Unit})
        Me.dgvSumber.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSumber.Location = New System.Drawing.Point(3, 3)
        Me.dgvSumber.Name = "dgvSumber"
        Me.dgvSumber.ReadOnly = True
        Me.dgvSumber.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvSumber.RowTemplate.Height = 24
        Me.dgvSumber.Size = New System.Drawing.Size(722, 372)
        Me.dgvSumber.TabIndex = 0
        '
        'ColSumber_No
        '
        Me.ColSumber_No.HeaderText = "No"
        Me.ColSumber_No.Name = "ColSumber_No"
        Me.ColSumber_No.ReadOnly = True
        Me.ColSumber_No.Width = 51
        '
        'ColSumber_CodeNo
        '
        Me.ColSumber_CodeNo.HeaderText = "Code No"
        Me.ColSumber_CodeNo.Name = "ColSumber_CodeNo"
        Me.ColSumber_CodeNo.ReadOnly = True
        Me.ColSumber_CodeNo.Width = 88
        '
        'ColSumber_Desc
        '
        Me.ColSumber_Desc.HeaderText = "Desc"
        Me.ColSumber_Desc.Name = "ColSumber_Desc"
        Me.ColSumber_Desc.ReadOnly = True
        Me.ColSumber_Desc.Width = 65
        '
        'ColSumber_Qty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSumber_Qty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColSumber_Qty.HeaderText = "Qty"
        Me.ColSumber_Qty.Name = "ColSumber_Qty"
        Me.ColSumber_Qty.ReadOnly = True
        Me.ColSumber_Qty.Width = 55
        '
        'ColSumber_Unit
        '
        Me.ColSumber_Unit.HeaderText = "Unit"
        Me.ColSumber_Unit.Name = "ColSumber_Unit"
        Me.ColSumber_Unit.ReadOnly = True
        Me.ColSumber_Unit.Width = 58
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvTarget)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(728, 378)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Target"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvTarget
        '
        Me.dgvTarget.AllowUserToAddRows = False
        Me.dgvTarget.AllowUserToDeleteRows = False
        Me.dgvTarget.AllowUserToResizeColumns = False
        Me.dgvTarget.AllowUserToResizeRows = False
        Me.dgvTarget.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvTarget.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvTarget.ColumnHeadersHeight = 25
        Me.dgvTarget.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvTarget.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColTarget_No, Me.ColTarget_CodeNo, Me.ColTarget_Desc, Me.ColTarget_Qty, Me.ColTarget_Unit})
        Me.dgvTarget.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvTarget.Location = New System.Drawing.Point(3, 3)
        Me.dgvTarget.Name = "dgvTarget"
        Me.dgvTarget.ReadOnly = True
        Me.dgvTarget.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvTarget.RowTemplate.Height = 24
        Me.dgvTarget.Size = New System.Drawing.Size(722, 372)
        Me.dgvTarget.TabIndex = 1
        '
        'ColTarget_No
        '
        Me.ColTarget_No.HeaderText = "No"
        Me.ColTarget_No.Name = "ColTarget_No"
        Me.ColTarget_No.ReadOnly = True
        Me.ColTarget_No.Width = 51
        '
        'ColTarget_CodeNo
        '
        Me.ColTarget_CodeNo.HeaderText = "Code No"
        Me.ColTarget_CodeNo.Name = "ColTarget_CodeNo"
        Me.ColTarget_CodeNo.ReadOnly = True
        Me.ColTarget_CodeNo.Width = 88
        '
        'ColTarget_Desc
        '
        Me.ColTarget_Desc.HeaderText = "Desc"
        Me.ColTarget_Desc.Name = "ColTarget_Desc"
        Me.ColTarget_Desc.ReadOnly = True
        Me.ColTarget_Desc.Width = 65
        '
        'ColTarget_Qty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTarget_Qty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColTarget_Qty.HeaderText = "Qty"
        Me.ColTarget_Qty.Name = "ColTarget_Qty"
        Me.ColTarget_Qty.ReadOnly = True
        Me.ColTarget_Qty.Width = 55
        '
        'ColTarget_Unit
        '
        Me.ColTarget_Unit.HeaderText = "Unit"
        Me.ColTarget_Unit.Name = "ColTarget_Unit"
        Me.ColTarget_Unit.ReadOnly = True
        Me.ColTarget_Unit.Width = 58
        '
        'btnDeleteBundling
        '
        Me.btnDeleteBundling.Location = New System.Drawing.Point(428, 97)
        Me.btnDeleteBundling.Name = "btnDeleteBundling"
        Me.btnDeleteBundling.Size = New System.Drawing.Size(120, 25)
        Me.btnDeleteBundling.TabIndex = 23
        Me.btnDeleteBundling.Text = "Delete Bundling"
        Me.btnDeleteBundling.UseVisualStyleBackColor = True
        '
        'btnDeleteStock
        '
        Me.btnDeleteStock.Location = New System.Drawing.Point(651, 11)
        Me.btnDeleteStock.Name = "btnDeleteStock"
        Me.btnDeleteStock.Size = New System.Drawing.Size(97, 25)
        Me.btnDeleteStock.TabIndex = 24
        Me.btnDeleteStock.Text = "Del. Stock"
        Me.btnDeleteStock.UseVisualStyleBackColor = True
        '
        'Form2_Bundling
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(760, 618)
        Me.Controls.Add(Me.btnDeleteStock)
        Me.Controls.Add(Me.btnDeleteBundling)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.dtpBundling2)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.dtpBundling1)
        Me.Controls.Add(Me.TxtPrimaryLoc)
        Me.Controls.Add(Me.TxtLocDesc)
        Me.Controls.Add(Me.TxtLocCode)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnFindLoc)
        Me.Controls.Add(Me.TxtPrimaryAccTarget)
        Me.Controls.Add(Me.TxtAccTargetDesc)
        Me.Controls.Add(Me.TxtAccTargetCode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnFindAccTarget)
        Me.Controls.Add(Me.TxtPrimaryAccSumber)
        Me.Controls.Add(Me.TxtAccSumberDesc)
        Me.Controls.Add(Me.TxtAccSumberCode)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnFindAccSumber)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(768, 655)
        Me.Name = "Form2_Bundling"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Bundling By Account"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvSumber, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvTarget, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnFindAccSumber As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtAccSumberCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtAccSumberDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccSumber As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccTarget As System.Windows.Forms.TextBox
    Friend WithEvents TxtAccTargetDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtAccTargetCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFindAccTarget As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryLoc As System.Windows.Forms.TextBox
    Friend WithEvents TxtLocDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtLocCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnFindLoc As System.Windows.Forms.Button
    Friend WithEvents dtpBundling1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtpBundling2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvSumber As System.Windows.Forms.DataGridView
    Friend WithEvents ColSumber_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSumber_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSumber_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSumber_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSumber_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvTarget As System.Windows.Forms.DataGridView
    Friend WithEvents ColTarget_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTarget_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTarget_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTarget_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTarget_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnDeleteBundling As System.Windows.Forms.Button
    Friend WithEvents btnDeleteStock As System.Windows.Forms.Button
End Class
