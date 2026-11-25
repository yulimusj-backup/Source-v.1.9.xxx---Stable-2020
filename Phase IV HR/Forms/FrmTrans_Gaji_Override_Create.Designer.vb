<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_Override_Create
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
        Me.btnDelete = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dtpTO = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnCreate = New System.Windows.Forms.Button
        Me.Prgb1 = New System.Windows.Forms.ProgressBar
        Me.LblProgress = New System.Windows.Forms.Label
        Me.bgw1 = New System.ComponentModel.BackgroundWorker
        Me.bgw2 = New System.ComponentModel.BackgroundWorker
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.Label3 = New System.Windows.Forms.Label
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.btnStart = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.dtpSalaryDate = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.btnSHow1 = New System.Windows.Forms.Button
        Me.dgvTaxCategory = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTaxCategory = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColResign = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.btnStart2 = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.dtpSalaryDate2 = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.bgw3 = New System.ComponentModel.BackgroundWorker
        Me.bgw4 = New System.ComponentModel.BackgroundWorker
        Me.btnShow2 = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvTaxCategory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(255, 33)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(99, 27)
        Me.btnDelete.TabIndex = 0
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dtpTO)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.dtpFrom)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 26)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(243, 83)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Salary Date"
        '
        'dtpTO
        '
        Me.dtpTO.CustomFormat = "dd MMM yyyy"
        Me.dtpTO.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTO.Location = New System.Drawing.Point(83, 49)
        Me.dtpTO.Name = "dtpTO"
        Me.dtpTO.Size = New System.Drawing.Size(151, 22)
        Me.dtpTO.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(25, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "To"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(83, 21)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(151, 22)
        Me.dtpFrom.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "From"
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(255, 66)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(99, 27)
        Me.btnCreate.TabIndex = 3
        Me.btnCreate.Text = "Create"
        Me.btnCreate.UseVisualStyleBackColor = True
        '
        'Prgb1
        '
        Me.Prgb1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Prgb1.Location = New System.Drawing.Point(12, 487)
        Me.Prgb1.Name = "Prgb1"
        Me.Prgb1.Size = New System.Drawing.Size(631, 27)
        Me.Prgb1.TabIndex = 4
        '
        'LblProgress
        '
        Me.LblProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(9, 520)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(51, 17)
        Me.LblProgress.TabIndex = 5
        Me.LblProgress.Text = "Label3"
        '
        'bgw1
        '
        '
        'bgw2
        '
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(631, 469)
        Me.TabControl1.TabIndex = 6
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Controls.Add(Me.btnDelete)
        Me.TabPage1.Controls.Add(Me.btnCreate)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(623, 440)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Based On Manual Ovveride"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(604, 17)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Membuat transaksi override berdasarkan manula override yang dilakukan transaksi s" & _
            "alary calc"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.btnStart)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.dtpSalaryDate)
        Me.TabPage2.Controls.Add(Me.Label5)
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(623, 440)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Based On Salary Calculation"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(250, 40)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 25)
        Me.btnStart.TabIndex = 9
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 43)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 17)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "Salary date"
        '
        'dtpSalaryDate
        '
        Me.dtpSalaryDate.CustomFormat = "dd MMM yyyy"
        Me.dtpSalaryDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSalaryDate.Location = New System.Drawing.Point(92, 40)
        Me.dtpSalaryDate.Name = "dtpSalaryDate"
        Me.dtpSalaryDate.Size = New System.Drawing.Size(152, 22)
        Me.dtpSalaryDate.TabIndex = 7
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 20)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(530, 17)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "hal ini dilakukan agar ketika membuat hitung ulang gaji infinity diakhir tahun mu" & _
            "dah"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 3)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(504, 17)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Membuat transaksi override berdasarkan semua komponent salary calculation,"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.btnShow2)
        Me.TabPage3.Controls.Add(Me.btnSHow1)
        Me.TabPage3.Controls.Add(Me.dgvTaxCategory)
        Me.TabPage3.Controls.Add(Me.btnStart2)
        Me.TabPage3.Controls.Add(Me.Label7)
        Me.TabPage3.Controls.Add(Me.dtpSalaryDate2)
        Me.TabPage3.Controls.Add(Me.Label8)
        Me.TabPage3.Controls.Add(Me.Label9)
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(623, 440)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Update Kategory Pajak"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'btnSHow1
        '
        Me.btnSHow1.Location = New System.Drawing.Point(331, 40)
        Me.btnSHow1.Name = "btnSHow1"
        Me.btnSHow1.Size = New System.Drawing.Size(75, 25)
        Me.btnSHow1.TabIndex = 16
        Me.btnSHow1.Text = "Show1"
        Me.btnSHow1.UseVisualStyleBackColor = True
        '
        'dgvTaxCategory
        '
        Me.dgvTaxCategory.AllowUserToAddRows = False
        Me.dgvTaxCategory.AllowUserToDeleteRows = False
        Me.dgvTaxCategory.AllowUserToResizeColumns = False
        Me.dgvTaxCategory.AllowUserToResizeRows = False
        Me.dgvTaxCategory.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvTaxCategory.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvTaxCategory.ColumnHeadersHeight = 25
        Me.dgvTaxCategory.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvTaxCategory.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCode, Me.ColName, Me.ColTaxCategory, Me.ColResign})
        Me.dgvTaxCategory.Location = New System.Drawing.Point(9, 71)
        Me.dgvTaxCategory.Name = "dgvTaxCategory"
        Me.dgvTaxCategory.ReadOnly = True
        Me.dgvTaxCategory.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvTaxCategory.RowTemplate.Height = 24
        Me.dgvTaxCategory.Size = New System.Drawing.Size(608, 363)
        Me.dgvTaxCategory.TabIndex = 15
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColCode
        '
        Me.ColCode.HeaderText = "Employee Code"
        Me.ColCode.Name = "ColCode"
        Me.ColCode.ReadOnly = True
        Me.ColCode.Width = 132
        '
        'ColName
        '
        Me.ColName.HeaderText = "Name"
        Me.ColName.Name = "ColName"
        Me.ColName.ReadOnly = True
        Me.ColName.Width = 70
        '
        'ColTaxCategory
        '
        Me.ColTaxCategory.HeaderText = "Tax Category"
        Me.ColTaxCategory.Name = "ColTaxCategory"
        Me.ColTaxCategory.ReadOnly = True
        Me.ColTaxCategory.Width = 117
        '
        'ColResign
        '
        Me.ColResign.HeaderText = "Resign"
        Me.ColResign.Name = "ColResign"
        Me.ColResign.ReadOnly = True
        Me.ColResign.Width = 58
        '
        'btnStart2
        '
        Me.btnStart2.Location = New System.Drawing.Point(250, 40)
        Me.btnStart2.Name = "btnStart2"
        Me.btnStart2.Size = New System.Drawing.Size(75, 25)
        Me.btnStart2.TabIndex = 14
        Me.btnStart2.Text = "Start"
        Me.btnStart2.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 43)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 17)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Salary date"
        '
        'dtpSalaryDate2
        '
        Me.dtpSalaryDate2.CustomFormat = "dd MMM yyyy"
        Me.dtpSalaryDate2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSalaryDate2.Location = New System.Drawing.Point(92, 40)
        Me.dtpSalaryDate2.Name = "dtpSalaryDate2"
        Me.dtpSalaryDate2.Size = New System.Drawing.Size(152, 22)
        Me.dtpSalaryDate2.TabIndex = 12
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 20)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(127, 17)
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "pada transaksi gaji"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 3)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(535, 17)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "Update kategory pajak pada master karyawan menjadi kategory pajak yang tercatat"
        '
        'bgw3
        '
        '
        'bgw4
        '
        '
        'btnShow2
        '
        Me.btnShow2.Location = New System.Drawing.Point(412, 41)
        Me.btnShow2.Name = "btnShow2"
        Me.btnShow2.Size = New System.Drawing.Size(75, 25)
        Me.btnShow2.TabIndex = 17
        Me.btnShow2.Text = "Show 2 "
        Me.btnShow2.UseVisualStyleBackColor = True
        '
        'FrmTrans_Gaji_Override_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(655, 546)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.Prgb1)
        Me.MinimumSize = New System.Drawing.Size(673, 209)
        Me.Name = "FrmTrans_Gaji_Override_Create"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Override Salary"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.dgvTaxCategory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpTO As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnCreate As System.Windows.Forms.Button
    Friend WithEvents Prgb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents bgw1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgw2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpSalaryDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents bgw3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents btnStart2 As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents dtpSalaryDate2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnSHow1 As System.Windows.Forms.Button
    Friend WithEvents dgvTaxCategory As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTaxCategory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColResign As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents bgw4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnShow2 As System.Windows.Forms.Button
End Class
