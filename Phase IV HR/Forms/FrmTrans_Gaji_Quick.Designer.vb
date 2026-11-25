<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_Quick
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
        Me.gboSalaryData = New System.Windows.Forms.GroupBox
        Me.lllSetDataSorting = New System.Windows.Forms.LinkLabel
        Me.CboTypeDateFilter = New System.Windows.Forms.ComboBox
        Me.btnEmail = New System.Windows.Forms.Button
        Me.LblAllEmail = New System.Windows.Forms.LinkLabel
        Me.LblAllPrint = New System.Windows.Forms.LinkLabel
        Me.BtnDelete = New System.Windows.Forms.Button
        Me.BtnPay = New System.Windows.Forms.Button
        Me.DGVSalaryCalculation = New System.Windows.Forms.DataGridView
        Me.BtnNew = New System.Windows.Forms.Button
        Me.BtnPrint = New System.Windows.Forms.Button
        Me.BtnShow = New System.Windows.Forms.Button
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.DtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.LblAllPaid = New System.Windows.Forms.LinkLabel
        Me.LblAllDelete = New System.Windows.Forms.LinkLabel
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocumentNo = New System.Windows.Forms.DataGridViewLinkColumn
        Me.COlDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeNIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeptCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPositionCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPosition = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEdit = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColChkPrint = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDelete = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPay = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimaryTransaksi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColChkEmail = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.gboSalaryData.SuspendLayout()
        CType(Me.DGVSalaryCalculation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gboSalaryData
        '
        Me.gboSalaryData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboSalaryData.Controls.Add(Me.lllSetDataSorting)
        Me.gboSalaryData.Controls.Add(Me.CboTypeDateFilter)
        Me.gboSalaryData.Controls.Add(Me.btnEmail)
        Me.gboSalaryData.Controls.Add(Me.LblAllEmail)
        Me.gboSalaryData.Controls.Add(Me.LblAllPrint)
        Me.gboSalaryData.Controls.Add(Me.BtnDelete)
        Me.gboSalaryData.Controls.Add(Me.BtnPay)
        Me.gboSalaryData.Controls.Add(Me.DGVSalaryCalculation)
        Me.gboSalaryData.Controls.Add(Me.BtnNew)
        Me.gboSalaryData.Controls.Add(Me.BtnPrint)
        Me.gboSalaryData.Controls.Add(Me.BtnShow)
        Me.gboSalaryData.Controls.Add(Me.DtpTo)
        Me.gboSalaryData.Controls.Add(Me.Label4)
        Me.gboSalaryData.Controls.Add(Me.DtpPeriod)
        Me.gboSalaryData.Controls.Add(Me.Label5)
        Me.gboSalaryData.Controls.Add(Me.Button2)
        Me.gboSalaryData.Controls.Add(Me.LblAllPaid)
        Me.gboSalaryData.Controls.Add(Me.LblAllDelete)
        Me.gboSalaryData.Location = New System.Drawing.Point(10, 10)
        Me.gboSalaryData.Margin = New System.Windows.Forms.Padding(2)
        Me.gboSalaryData.Name = "gboSalaryData"
        Me.gboSalaryData.Padding = New System.Windows.Forms.Padding(2)
        Me.gboSalaryData.Size = New System.Drawing.Size(846, 495)
        Me.gboSalaryData.TabIndex = 44
        Me.gboSalaryData.TabStop = False
        Me.gboSalaryData.Text = "Salary Data"
        '
        'lllSetDataSorting
        '
        Me.lllSetDataSorting.ActiveLinkColor = System.Drawing.Color.Blue
        Me.lllSetDataSorting.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lllSetDataSorting.AutoSize = True
        Me.lllSetDataSorting.Location = New System.Drawing.Point(756, 19)
        Me.lllSetDataSorting.Name = "lllSetDataSorting"
        Me.lllSetDataSorting.Size = New System.Drawing.Size(85, 13)
        Me.lllSetDataSorting.TabIndex = 57
        Me.lllSetDataSorting.TabStop = True
        Me.lllSetDataSorting.Text = "Set Data Sorting"
        Me.lllSetDataSorting.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'CboTypeDateFilter
        '
        Me.CboTypeDateFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboTypeDateFilter.FormattingEnabled = True
        Me.CboTypeDateFilter.Location = New System.Drawing.Point(7, 16)
        Me.CboTypeDateFilter.Name = "CboTypeDateFilter"
        Me.CboTypeDateFilter.Size = New System.Drawing.Size(121, 21)
        Me.CboTypeDateFilter.TabIndex = 0
        '
        'btnEmail
        '
        Me.btnEmail.Location = New System.Drawing.Point(626, 16)
        Me.btnEmail.Name = "btnEmail"
        Me.btnEmail.Size = New System.Drawing.Size(47, 20)
        Me.btnEmail.TabIndex = 56
        Me.btnEmail.Text = "Email"
        Me.btnEmail.UseVisualStyleBackColor = True
        '
        'LblAllEmail
        '
        Me.LblAllEmail.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblAllEmail.AutoSize = True
        Me.LblAllEmail.BackColor = System.Drawing.SystemColors.Control
        Me.LblAllEmail.Location = New System.Drawing.Point(530, 480)
        Me.LblAllEmail.Name = "LblAllEmail"
        Me.LblAllEmail.Size = New System.Drawing.Size(83, 13)
        Me.LblAllEmail.TabIndex = 55
        Me.LblAllEmail.TabStop = True
        Me.LblAllEmail.Text = "Email: Check All"
        '
        'LblAllPrint
        '
        Me.LblAllPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblAllPrint.AutoSize = True
        Me.LblAllPrint.BackColor = System.Drawing.SystemColors.Control
        Me.LblAllPrint.Location = New System.Drawing.Point(671, 480)
        Me.LblAllPrint.Name = "LblAllPrint"
        Me.LblAllPrint.Size = New System.Drawing.Size(76, 13)
        Me.LblAllPrint.TabIndex = 7
        Me.LblAllPrint.TabStop = True
        Me.LblAllPrint.Text = "Check All Print"
        '
        'BtnDelete
        '
        Me.BtnDelete.Location = New System.Drawing.Point(576, 16)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(47, 20)
        Me.BtnDelete.TabIndex = 6
        Me.BtnDelete.Text = "Delete"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'BtnPay
        '
        Me.BtnPay.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnPay.Location = New System.Drawing.Point(697, 16)
        Me.BtnPay.Name = "BtnPay"
        Me.BtnPay.Size = New System.Drawing.Size(47, 20)
        Me.BtnPay.TabIndex = 52
        Me.BtnPay.TabStop = False
        Me.BtnPay.Text = "Pay"
        Me.BtnPay.UseVisualStyleBackColor = True
        Me.BtnPay.Visible = False
        '
        'DGVSalaryCalculation
        '
        Me.DGVSalaryCalculation.AllowUserToAddRows = False
        Me.DGVSalaryCalculation.AllowUserToDeleteRows = False
        Me.DGVSalaryCalculation.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DGVSalaryCalculation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVSalaryCalculation.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DGVSalaryCalculation.ColumnHeadersHeight = 30
        Me.DGVSalaryCalculation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColDocumentNo, Me.COlDate, Me.ColEmployeeCode, Me.ColEmployeeName, Me.ColEmployeeNIK, Me.ColDeptCode, Me.ColDept, Me.ColJobPositionCode, Me.ColJobPosition, Me.ColStartDate, Me.ColEndDate, Me.ColEdit, Me.ColChkPrint, Me.ColDelete, Me.ColPay, Me.ColPrimaryTransaksi, Me.ColChkEmail})
        Me.DGVSalaryCalculation.Location = New System.Drawing.Point(6, 43)
        Me.DGVSalaryCalculation.Margin = New System.Windows.Forms.Padding(2)
        Me.DGVSalaryCalculation.Name = "DGVSalaryCalculation"
        Me.DGVSalaryCalculation.ReadOnly = True
        Me.DGVSalaryCalculation.RowTemplate.Height = 24
        Me.DGVSalaryCalculation.Size = New System.Drawing.Size(836, 433)
        Me.DGVSalaryCalculation.TabIndex = 6
        '
        'BtnNew
        '
        Me.BtnNew.Location = New System.Drawing.Point(526, 16)
        Me.BtnNew.Name = "BtnNew"
        Me.BtnNew.Size = New System.Drawing.Size(47, 20)
        Me.BtnNew.TabIndex = 5
        Me.BtnNew.Text = "New"
        Me.BtnNew.UseVisualStyleBackColor = True
        '
        'BtnPrint
        '
        Me.BtnPrint.Location = New System.Drawing.Point(476, 16)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(47, 20)
        Me.BtnPrint.TabIndex = 4
        Me.BtnPrint.Text = "Print"
        Me.BtnPrint.UseVisualStyleBackColor = True
        '
        'BtnShow
        '
        Me.BtnShow.Location = New System.Drawing.Point(426, 16)
        Me.BtnShow.Name = "BtnShow"
        Me.BtnShow.Size = New System.Drawing.Size(47, 20)
        Me.BtnShow.TabIndex = 3
        Me.BtnShow.Text = "Show"
        Me.BtnShow.UseVisualStyleBackColor = True
        '
        'DtpTo
        '
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(298, 16)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(122, 20)
        Me.DtpTo.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(261, 18)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 47
        Me.Label4.Text = "To"
        '
        'DtpPeriod
        '
        Me.DtpPeriod.CustomFormat = "dd MMM yyyy"
        Me.DtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpPeriod.Location = New System.Drawing.Point(134, 16)
        Me.DtpPeriod.Name = "DtpPeriod"
        Me.DtpPeriod.Size = New System.Drawing.Size(122, 20)
        Me.DtpPeriod.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(14, 20)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(37, 13)
        Me.Label5.TabIndex = 45
        Me.Label5.Text = "Period"
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Location = New System.Drawing.Point(412, 263)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(51, 20)
        Me.Button2.TabIndex = 42
        Me.Button2.Text = "Close"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'LblAllPaid
        '
        Me.LblAllPaid.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblAllPaid.AutoSize = True
        Me.LblAllPaid.BackColor = System.Drawing.SystemColors.Control
        Me.LblAllPaid.Location = New System.Drawing.Point(627, 480)
        Me.LblAllPaid.Name = "LblAllPaid"
        Me.LblAllPaid.Size = New System.Drawing.Size(79, 13)
        Me.LblAllPaid.TabIndex = 54
        Me.LblAllPaid.TabStop = True
        Me.LblAllPaid.Text = "Paid: Check All"
        Me.LblAllPaid.Visible = False
        '
        'LblAllDelete
        '
        Me.LblAllDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblAllDelete.AutoSize = True
        Me.LblAllDelete.BackColor = System.Drawing.SystemColors.Control
        Me.LblAllDelete.Location = New System.Drawing.Point(749, 480)
        Me.LblAllDelete.Name = "LblAllDelete"
        Me.LblAllDelete.Size = New System.Drawing.Size(89, 13)
        Me.LblAllDelete.TabIndex = 8
        Me.LblAllDelete.TabStop = True
        Me.LblAllDelete.Text = "Delete: Check All"
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 46
        '
        'ColDocumentNo
        '
        Me.ColDocumentNo.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColDocumentNo.HeaderText = "Document No"
        Me.ColDocumentNo.Name = "ColDocumentNo"
        Me.ColDocumentNo.ReadOnly = True
        Me.ColDocumentNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDocumentNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColDocumentNo.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColDocumentNo.Width = 98
        '
        'COlDate
        '
        Me.COlDate.HeaderText = "Date"
        Me.COlDate.Name = "COlDate"
        Me.COlDate.ReadOnly = True
        Me.COlDate.Width = 55
        '
        'ColEmployeeCode
        '
        Me.ColEmployeeCode.HeaderText = "Employee Code"
        Me.ColEmployeeCode.Name = "ColEmployeeCode"
        Me.ColEmployeeCode.ReadOnly = True
        Me.ColEmployeeCode.Visible = False
        Me.ColEmployeeCode.Width = 106
        '
        'ColEmployeeName
        '
        Me.ColEmployeeName.HeaderText = "Employee Name"
        Me.ColEmployeeName.Name = "ColEmployeeName"
        Me.ColEmployeeName.ReadOnly = True
        Me.ColEmployeeName.Width = 109
        '
        'ColEmployeeNIK
        '
        Me.ColEmployeeNIK.HeaderText = "Employee NIK"
        Me.ColEmployeeNIK.Name = "ColEmployeeNIK"
        Me.ColEmployeeNIK.ReadOnly = True
        Me.ColEmployeeNIK.Width = 99
        '
        'ColDeptCode
        '
        Me.ColDeptCode.HeaderText = "Department"
        Me.ColDeptCode.Name = "ColDeptCode"
        Me.ColDeptCode.ReadOnly = True
        Me.ColDeptCode.Width = 87
        '
        'ColDept
        '
        Me.ColDept.HeaderText = "Departement Description"
        Me.ColDept.Name = "ColDept"
        Me.ColDept.ReadOnly = True
        Me.ColDept.Width = 149
        '
        'ColJobPositionCode
        '
        Me.ColJobPositionCode.HeaderText = "Job Position"
        Me.ColJobPositionCode.Name = "ColJobPositionCode"
        Me.ColJobPositionCode.ReadOnly = True
        Me.ColJobPositionCode.Width = 89
        '
        'ColJobPosition
        '
        Me.ColJobPosition.HeaderText = "Job Position Description"
        Me.ColJobPosition.Name = "ColJobPosition"
        Me.ColJobPosition.ReadOnly = True
        Me.ColJobPosition.Width = 145
        '
        'ColStartDate
        '
        Me.ColStartDate.HeaderText = "Start Date"
        Me.ColStartDate.Name = "ColStartDate"
        Me.ColStartDate.ReadOnly = True
        Me.ColStartDate.Width = 80
        '
        'ColEndDate
        '
        Me.ColEndDate.HeaderText = "End Date"
        Me.ColEndDate.Name = "ColEndDate"
        Me.ColEndDate.ReadOnly = True
        Me.ColEndDate.Width = 77
        '
        'ColEdit
        '
        Me.ColEdit.HeaderText = "Edit"
        Me.ColEdit.Name = "ColEdit"
        Me.ColEdit.ReadOnly = True
        Me.ColEdit.Width = 31
        '
        'ColChkPrint
        '
        Me.ColChkPrint.HeaderText = "Print"
        Me.ColChkPrint.Name = "ColChkPrint"
        Me.ColChkPrint.ReadOnly = True
        Me.ColChkPrint.Width = 34
        '
        'ColDelete
        '
        Me.ColDelete.HeaderText = "Delete"
        Me.ColDelete.Name = "ColDelete"
        Me.ColDelete.ReadOnly = True
        Me.ColDelete.Width = 44
        '
        'ColPay
        '
        Me.ColPay.HeaderText = "Paid"
        Me.ColPay.Name = "ColPay"
        Me.ColPay.ReadOnly = True
        Me.ColPay.Visible = False
        Me.ColPay.Width = 34
        '
        'ColPrimaryTransaksi
        '
        Me.ColPrimaryTransaksi.HeaderText = "Primary Transaksi"
        Me.ColPrimaryTransaksi.Name = "ColPrimaryTransaksi"
        Me.ColPrimaryTransaksi.ReadOnly = True
        Me.ColPrimaryTransaksi.Visible = False
        Me.ColPrimaryTransaksi.Width = 115
        '
        'ColChkEmail
        '
        Me.ColChkEmail.HeaderText = "Email"
        Me.ColChkEmail.Name = "ColChkEmail"
        Me.ColChkEmail.ReadOnly = True
        Me.ColChkEmail.Width = 38
        '
        'FrmTrans_Gaji_Quick
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button2
        Me.ClientSize = New System.Drawing.Size(867, 515)
        Me.Controls.Add(Me.gboSalaryData)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(756, 552)
        Me.Name = "FrmTrans_Gaji_Quick"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quick Salary Calculation"
        Me.gboSalaryData.ResumeLayout(False)
        Me.gboSalaryData.PerformLayout()
        CType(Me.DGVSalaryCalculation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gboSalaryData As System.Windows.Forms.GroupBox
    Friend WithEvents BtnNew As System.Windows.Forms.Button
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
    Friend WithEvents BtnShow As System.Windows.Forms.Button
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents DGVSalaryCalculation As System.Windows.Forms.DataGridView
    Friend WithEvents BtnPay As System.Windows.Forms.Button
    Friend WithEvents LblAllPaid As System.Windows.Forms.LinkLabel
    Friend WithEvents LblAllPrint As System.Windows.Forms.LinkLabel
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents LblAllDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnEmail As System.Windows.Forms.Button
    Friend WithEvents LblAllEmail As System.Windows.Forms.LinkLabel
    Friend WithEvents CboTypeDateFilter As System.Windows.Forms.ComboBox
    Friend WithEvents lllSetDataSorting As System.Windows.Forms.LinkLabel
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocumentNo As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents COlDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeNIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeptCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPositionCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPosition As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEdit As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColChkPrint As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDelete As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPay As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimaryTransaksi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColChkEmail As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
