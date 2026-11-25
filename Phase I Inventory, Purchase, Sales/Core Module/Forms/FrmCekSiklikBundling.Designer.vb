<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCekSiklikBundling
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTransactionNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.btnStart = New System.Windows.Forms.Button
        Me.LblPeriod = New System.Windows.Forms.Label
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.LblProgress = New System.Windows.Forms.Label
        Me.DataGridView2 = New System.Windows.Forms.DataGridView
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTransDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSudahHitung = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnRepair = New System.Windows.Forms.Button
        Me.btnDetemineManual = New System.Windows.Forms.Button
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCodeNo, Me.ColDesc, Me.ColPrimaryCodeNo, Me.ColTransactionNo})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 11)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(597, 228)
        Me.DataGridView1.TabIndex = 0
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.Width = 88
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.Width = 104
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "Primary"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.ReadOnly = True
        Me.ColPrimaryCodeNo.Visible = False
        '
        'ColTransactionNo
        '
        Me.ColTransactionNo.HeaderText = "Transaction No"
        Me.ColTransactionNo.Name = "ColTransactionNo"
        Me.ColTransactionNo.ReadOnly = True
        Me.ColTransactionNo.Width = 130
        '
        'dtpPeriod
        '
        Me.dtpPeriod.CustomFormat = "MMM yyyy"
        Me.dtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPeriod.Location = New System.Drawing.Point(109, 12)
        Me.dtpPeriod.Name = "dtpPeriod"
        Me.dtpPeriod.ShowUpDown = True
        Me.dtpPeriod.Size = New System.Drawing.Size(153, 22)
        Me.dtpPeriod.TabIndex = 1
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(268, 11)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 23)
        Me.btnStart.TabIndex = 2
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'LblPeriod
        '
        Me.LblPeriod.AutoSize = True
        Me.LblPeriod.Location = New System.Drawing.Point(12, 17)
        Me.LblPeriod.Name = "LblPeriod"
        Me.LblPeriod.Size = New System.Drawing.Size(49, 17)
        Me.LblPeriod.TabIndex = 3
        Me.LblPeriod.Text = "Period"
        '
        'BackgroundWorker1
        '
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(109, 40)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(500, 23)
        Me.ProgressBar1.TabIndex = 4
        '
        'LblProgress
        '
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(106, 66)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(0, 17)
        Me.LblProgress.TabIndex = 5
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDate, Me.ColDocumentNo, Me.ColTransDesc, Me.ColQtyIn, Me.ColQtyOut, Me.ColBalance, Me.ColSudahHitung})
        Me.DataGridView2.Location = New System.Drawing.Point(12, 245)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        Me.DataGridView2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView2.RowTemplate.Height = 24
        Me.DataGridView2.Size = New System.Drawing.Size(597, 318)
        Me.DataGridView2.TabIndex = 6
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDate.Width = 44
        '
        'ColDocumentNo
        '
        Me.ColDocumentNo.HeaderText = "Document No"
        Me.ColDocumentNo.Name = "ColDocumentNo"
        Me.ColDocumentNo.ReadOnly = True
        Me.ColDocumentNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColTransDesc
        '
        Me.ColTransDesc.HeaderText = "Description"
        Me.ColTransDesc.Name = "ColTransDesc"
        Me.ColTransDesc.ReadOnly = True
        Me.ColTransDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTransDesc.Width = 85
        '
        'ColQtyIn
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyIn.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQtyIn.HeaderText = "Qty In"
        Me.ColQtyIn.Name = "ColQtyIn"
        Me.ColQtyIn.ReadOnly = True
        Me.ColQtyIn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyIn.Width = 51
        '
        'ColQtyOut
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyOut.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyOut.HeaderText = "Qty Out"
        Me.ColQtyOut.Name = "ColQtyOut"
        Me.ColQtyOut.ReadOnly = True
        Me.ColQtyOut.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyOut.Width = 63
        '
        'ColBalance
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalance.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColBalance.HeaderText = "Balance"
        Me.ColBalance.Name = "ColBalance"
        Me.ColBalance.ReadOnly = True
        Me.ColBalance.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBalance.Width = 65
        '
        'ColSudahHitung
        '
        Me.ColSudahHitung.HeaderText = "Sudah Hitung"
        Me.ColSudahHitung.Name = "ColSudahHitung"
        Me.ColSudahHitung.ReadOnly = True
        Me.ColSudahHitung.Visible = False
        Me.ColSudahHitung.Width = 119
        '
        'btnRepair
        '
        Me.btnRepair.Location = New System.Drawing.Point(12, 11)
        Me.btnRepair.Name = "btnRepair"
        Me.btnRepair.Size = New System.Drawing.Size(125, 25)
        Me.btnRepair.TabIndex = 7
        Me.btnRepair.TabStop = False
        Me.btnRepair.Text = "Repair Data First"
        Me.btnRepair.UseVisualStyleBackColor = True
        Me.btnRepair.Visible = False
        '
        'btnDetemineManual
        '
        Me.btnDetemineManual.Location = New System.Drawing.Point(143, 11)
        Me.btnDetemineManual.Name = "btnDetemineManual"
        Me.btnDetemineManual.Size = New System.Drawing.Size(185, 25)
        Me.btnDetemineManual.TabIndex = 8
        Me.btnDetemineManual.TabStop = False
        Me.btnDetemineManual.Text = "Detemine cost manually"
        Me.btnDetemineManual.UseVisualStyleBackColor = True
        Me.btnDetemineManual.Visible = False
        '
        'FrmCekSiklikBundling
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(621, 575)
        Me.Controls.Add(Me.btnDetemineManual)
        Me.Controls.Add(Me.btnRepair)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.LblPeriod)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.dtpPeriod)
        Me.Controls.Add(Me.DataGridView1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(629, 612)
        Me.Name = "FrmCekSiklikBundling"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data check for cyclic bundling"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents LblPeriod As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTransactionNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTransDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSudahHitung As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btnRepair As System.Windows.Forms.Button
    Friend WithEvents btnDetemineManual As System.Windows.Forms.Button
End Class
