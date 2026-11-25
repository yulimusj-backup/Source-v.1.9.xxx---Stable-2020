<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2_EditBundling
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColAcc_AccountCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAcc_AccDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAcc_PrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnStart = New System.Windows.Forms.Button
        Me.PrgBar = New System.Windows.Forms.ProgressBar
        Me.LblProgress = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryCodeNoBefore = New System.Windows.Forms.TextBox
        Me.TxtProductDescBefore = New System.Windows.Forms.TextBox
        Me.TxtCodeNoBefore = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnFindCodeNoBefore = New System.Windows.Forms.Button
        Me.TxtPrimaryCodeNoAfter = New System.Windows.Forms.TextBox
        Me.TxtProductDescAfter = New System.Windows.Forms.TextBox
        Me.TxtCodeNoAfter = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnFindCodeNoAfter = New System.Windows.Forms.Button
        Me.bgw = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.dgvItem)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 142)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(639, 460)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Inventory Account"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAcc_AccountCode, Me.ColAcc_AccDesc, Me.ColAcc_PrimaryAkun})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(3, 18)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(633, 439)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 5
        '
        'ColAcc_AccountCode
        '
        Me.ColAcc_AccountCode.HeaderText = "Account"
        Me.ColAcc_AccountCode.MaxInputLength = 50
        Me.ColAcc_AccountCode.Name = "ColAcc_AccountCode"
        Me.ColAcc_AccountCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAcc_AccountCode.Width = 65
        '
        'ColAcc_AccDesc
        '
        Me.ColAcc_AccDesc.HeaderText = "Description"
        Me.ColAcc_AccDesc.MaxInputLength = 150
        Me.ColAcc_AccDesc.Name = "ColAcc_AccDesc"
        Me.ColAcc_AccDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAcc_AccDesc.Width = 85
        '
        'ColAcc_PrimaryAkun
        '
        Me.ColAcc_PrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColAcc_PrimaryAkun.Name = "ColAcc_PrimaryAkun"
        Me.ColAcc_PrimaryAkun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAcc_PrimaryAkun.Visible = False
        Me.ColAcc_PrimaryAkun.Width = 94
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(12, 96)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 23)
        Me.btnStart.TabIndex = 6
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'PrgBar
        '
        Me.PrgBar.Location = New System.Drawing.Point(129, 96)
        Me.PrgBar.Name = "PrgBar"
        Me.PrgBar.Size = New System.Drawing.Size(522, 23)
        Me.PrgBar.TabIndex = 7
        '
        'LblProgress
        '
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(126, 122)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(51, 17)
        Me.LblProgress.TabIndex = 8
        Me.LblProgress.Text = "Label1"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(129, 68)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(145, 22)
        Me.dtpFrom.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 71)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 17)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "From"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(280, 71)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(25, 17)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "To"
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(311, 68)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(145, 22)
        Me.dtpTo.TabIndex = 12
        '
        'TxtPrimaryCodeNoBefore
        '
        Me.TxtPrimaryCodeNoBefore.Location = New System.Drawing.Point(344, 12)
        Me.TxtPrimaryCodeNoBefore.Name = "TxtPrimaryCodeNoBefore"
        Me.TxtPrimaryCodeNoBefore.Size = New System.Drawing.Size(46, 22)
        Me.TxtPrimaryCodeNoBefore.TabIndex = 17
        Me.TxtPrimaryCodeNoBefore.TabStop = False
        Me.TxtPrimaryCodeNoBefore.Text = "-1"
        Me.TxtPrimaryCodeNoBefore.Visible = False
        '
        'TxtProductDescBefore
        '
        Me.TxtProductDescBefore.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtProductDescBefore.Location = New System.Drawing.Point(344, 12)
        Me.TxtProductDescBefore.Name = "TxtProductDescBefore"
        Me.TxtProductDescBefore.ReadOnly = True
        Me.TxtProductDescBefore.Size = New System.Drawing.Size(307, 22)
        Me.TxtProductDescBefore.TabIndex = 16
        '
        'TxtCodeNoBefore
        '
        Me.TxtCodeNoBefore.Location = New System.Drawing.Point(129, 12)
        Me.TxtCodeNoBefore.Name = "TxtCodeNoBefore"
        Me.TxtCodeNoBefore.Size = New System.Drawing.Size(128, 22)
        Me.TxtCodeNoBefore.TabIndex = 15
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(109, 17)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Code No Before"
        '
        'btnFindCodeNoBefore
        '
        Me.btnFindCodeNoBefore.Location = New System.Drawing.Point(263, 12)
        Me.btnFindCodeNoBefore.Name = "btnFindCodeNoBefore"
        Me.btnFindCodeNoBefore.Size = New System.Drawing.Size(75, 23)
        Me.btnFindCodeNoBefore.TabIndex = 13
        Me.btnFindCodeNoBefore.Text = "Find"
        Me.btnFindCodeNoBefore.UseVisualStyleBackColor = True
        '
        'TxtPrimaryCodeNoAfter
        '
        Me.TxtPrimaryCodeNoAfter.Location = New System.Drawing.Point(344, 40)
        Me.TxtPrimaryCodeNoAfter.Name = "TxtPrimaryCodeNoAfter"
        Me.TxtPrimaryCodeNoAfter.Size = New System.Drawing.Size(46, 22)
        Me.TxtPrimaryCodeNoAfter.TabIndex = 22
        Me.TxtPrimaryCodeNoAfter.TabStop = False
        Me.TxtPrimaryCodeNoAfter.Text = "-1"
        Me.TxtPrimaryCodeNoAfter.Visible = False
        '
        'TxtProductDescAfter
        '
        Me.TxtProductDescAfter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtProductDescAfter.Location = New System.Drawing.Point(344, 40)
        Me.TxtProductDescAfter.Name = "TxtProductDescAfter"
        Me.TxtProductDescAfter.ReadOnly = True
        Me.TxtProductDescAfter.Size = New System.Drawing.Size(307, 22)
        Me.TxtProductDescAfter.TabIndex = 21
        '
        'TxtCodeNoAfter
        '
        Me.TxtCodeNoAfter.Location = New System.Drawing.Point(129, 40)
        Me.TxtCodeNoAfter.Name = "TxtCodeNoAfter"
        Me.TxtCodeNoAfter.Size = New System.Drawing.Size(128, 22)
        Me.TxtCodeNoAfter.TabIndex = 20
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(97, 17)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Code No After"
        '
        'btnFindCodeNoAfter
        '
        Me.btnFindCodeNoAfter.Location = New System.Drawing.Point(263, 40)
        Me.btnFindCodeNoAfter.Name = "btnFindCodeNoAfter"
        Me.btnFindCodeNoAfter.Size = New System.Drawing.Size(75, 23)
        Me.btnFindCodeNoAfter.TabIndex = 18
        Me.btnFindCodeNoAfter.Text = "Find"
        Me.btnFindCodeNoAfter.UseVisualStyleBackColor = True
        '
        'bgw
        '
        Me.bgw.WorkerReportsProgress = True
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        '
        'Form2_EditBundling
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(663, 614)
        Me.Controls.Add(Me.TxtPrimaryCodeNoAfter)
        Me.Controls.Add(Me.TxtProductDescAfter)
        Me.Controls.Add(Me.TxtCodeNoAfter)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnFindCodeNoAfter)
        Me.Controls.Add(Me.TxtPrimaryCodeNoBefore)
        Me.Controls.Add(Me.TxtProductDescBefore)
        Me.Controls.Add(Me.TxtCodeNoBefore)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnFindCodeNoBefore)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.PrgBar)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.GroupBox1)
        Me.MinimizeBox = False
        Me.Name = "Form2_EditBundling"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit Bundling"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColAcc_AccountCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAcc_AccDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAcc_PrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents PrgBar As System.Windows.Forms.ProgressBar
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryCodeNoBefore As System.Windows.Forms.TextBox
    Friend WithEvents TxtProductDescBefore As System.Windows.Forms.TextBox
    Friend WithEvents TxtCodeNoBefore As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnFindCodeNoBefore As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryCodeNoAfter As System.Windows.Forms.TextBox
    Friend WithEvents TxtProductDescAfter As System.Windows.Forms.TextBox
    Friend WithEvents TxtCodeNoAfter As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnFindCodeNoAfter As System.Windows.Forms.Button
    Friend WithEvents bgw As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
