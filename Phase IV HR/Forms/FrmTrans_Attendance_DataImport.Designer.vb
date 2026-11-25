<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Attendance_DataImport
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
        Me.BtnFind = New System.Windows.Forms.Button
        Me.TxtFile = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.BtnImport = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.gboAttendanceData = New System.Windows.Forms.GroupBox
        Me.lllSetDataSorting = New System.Windows.Forms.LinkLabel
        Me.CboTypeDateFilter = New System.Windows.Forms.ComboBox
        Me.btnSendEmail = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.llbEmail_UncekAll = New System.Windows.Forms.LinkLabel
        Me.llbEmail_CekAll = New System.Windows.Forms.LinkLabel
        Me.Label7 = New System.Windows.Forms.Label
        Me.BtnDelete = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.LblUncheckedDelete = New System.Windows.Forms.LinkLabel
        Me.LblCheckedDelete = New System.Windows.Forms.LinkLabel
        Me.DGVAttendanceData = New System.Windows.Forms.DataGridView
        Me.LblUnchecked = New System.Windows.Forms.LinkLabel
        Me.LblChecked = New System.Windows.Forms.LinkLabel
        Me.BtnNew = New System.Windows.Forms.Button
        Me.BtnPrint = New System.Windows.Forms.Button
        Me.BtnShow = New System.Windows.Forms.Button
        Me.DtpTo_Attd = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.DtpPeriod_Attd = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.DtpTo_Import = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.DtpPeriod_Import = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.gboImportData = New System.Windows.Forms.GroupBox
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocumentNo = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.ColChkDelete = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColChkEmail = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimaryTransaksi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.gboAttendanceData.SuspendLayout()
        CType(Me.DGVAttendanceData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.gboImportData.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnFind
        '
        Me.BtnFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFind.Location = New System.Drawing.Point(309, 3)
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(42, 20)
        Me.BtnFind.TabIndex = 39
        Me.BtnFind.Text = "Find"
        Me.BtnFind.UseVisualStyleBackColor = True
        '
        'TxtFile
        '
        Me.TxtFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFile.Location = New System.Drawing.Point(88, 3)
        Me.TxtFile.MaxLength = 100
        Me.TxtFile.Name = "TxtFile"
        Me.TxtFile.Size = New System.Drawing.Size(217, 20)
        Me.TxtFile.TabIndex = 38
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(2, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(23, 13)
        Me.Label2.TabIndex = 40
        Me.Label2.Text = "File"
        '
        'BtnImport
        '
        Me.BtnImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnImport.Location = New System.Drawing.Point(89, 45)
        Me.BtnImport.Name = "BtnImport"
        Me.BtnImport.Size = New System.Drawing.Size(51, 20)
        Me.BtnImport.TabIndex = 41
        Me.BtnImport.Text = "Import"
        Me.BtnImport.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Location = New System.Drawing.Point(410, 263)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(51, 20)
        Me.Button2.TabIndex = 42
        Me.Button2.Text = "Close"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'gboAttendanceData
        '
        Me.gboAttendanceData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboAttendanceData.Controls.Add(Me.lllSetDataSorting)
        Me.gboAttendanceData.Controls.Add(Me.CboTypeDateFilter)
        Me.gboAttendanceData.Controls.Add(Me.btnSendEmail)
        Me.gboAttendanceData.Controls.Add(Me.Label8)
        Me.gboAttendanceData.Controls.Add(Me.llbEmail_UncekAll)
        Me.gboAttendanceData.Controls.Add(Me.llbEmail_CekAll)
        Me.gboAttendanceData.Controls.Add(Me.Label7)
        Me.gboAttendanceData.Controls.Add(Me.BtnDelete)
        Me.gboAttendanceData.Controls.Add(Me.Label6)
        Me.gboAttendanceData.Controls.Add(Me.LblUncheckedDelete)
        Me.gboAttendanceData.Controls.Add(Me.LblCheckedDelete)
        Me.gboAttendanceData.Controls.Add(Me.DGVAttendanceData)
        Me.gboAttendanceData.Controls.Add(Me.LblUnchecked)
        Me.gboAttendanceData.Controls.Add(Me.LblChecked)
        Me.gboAttendanceData.Controls.Add(Me.BtnNew)
        Me.gboAttendanceData.Controls.Add(Me.BtnPrint)
        Me.gboAttendanceData.Controls.Add(Me.BtnShow)
        Me.gboAttendanceData.Controls.Add(Me.DtpTo_Attd)
        Me.gboAttendanceData.Controls.Add(Me.Label4)
        Me.gboAttendanceData.Controls.Add(Me.DtpPeriod_Attd)
        Me.gboAttendanceData.Controls.Add(Me.Label5)
        Me.gboAttendanceData.Controls.Add(Me.Button2)
        Me.gboAttendanceData.Location = New System.Drawing.Point(6, 10)
        Me.gboAttendanceData.Margin = New System.Windows.Forms.Padding(2)
        Me.gboAttendanceData.Name = "gboAttendanceData"
        Me.gboAttendanceData.Padding = New System.Windows.Forms.Padding(2)
        Me.gboAttendanceData.Size = New System.Drawing.Size(844, 504)
        Me.gboAttendanceData.TabIndex = 43
        Me.gboAttendanceData.TabStop = False
        Me.gboAttendanceData.Text = "Attendance Data"
        '
        'lllSetDataSorting
        '
        Me.lllSetDataSorting.ActiveLinkColor = System.Drawing.Color.Blue
        Me.lllSetDataSorting.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lllSetDataSorting.AutoSize = True
        Me.lllSetDataSorting.Location = New System.Drawing.Point(752, 21)
        Me.lllSetDataSorting.Name = "lllSetDataSorting"
        Me.lllSetDataSorting.Size = New System.Drawing.Size(85, 13)
        Me.lllSetDataSorting.TabIndex = 61
        Me.lllSetDataSorting.TabStop = True
        Me.lllSetDataSorting.Text = "Set Data Sorting"
        Me.lllSetDataSorting.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'CboTypeDateFilter
        '
        Me.CboTypeDateFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboTypeDateFilter.FormattingEnabled = True
        Me.CboTypeDateFilter.Location = New System.Drawing.Point(12, 18)
        Me.CboTypeDateFilter.Name = "CboTypeDateFilter"
        Me.CboTypeDateFilter.Size = New System.Drawing.Size(121, 21)
        Me.CboTypeDateFilter.TabIndex = 45
        '
        'btnSendEmail
        '
        Me.btnSendEmail.Location = New System.Drawing.Point(628, 17)
        Me.btnSendEmail.Name = "btnSendEmail"
        Me.btnSendEmail.Size = New System.Drawing.Size(48, 20)
        Me.btnSendEmail.TabIndex = 60
        Me.btnSendEmail.Text = "Email"
        Me.btnSendEmail.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(240, 479)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(63, 13)
        Me.Label8.TabIndex = 59
        Me.Label8.Text = "Send Email:"
        '
        'llbEmail_UncekAll
        '
        Me.llbEmail_UncekAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbEmail_UncekAll.AutoSize = True
        Me.llbEmail_UncekAll.BackColor = System.Drawing.SystemColors.Control
        Me.llbEmail_UncekAll.Location = New System.Drawing.Point(360, 479)
        Me.llbEmail_UncekAll.Name = "llbEmail_UncekAll"
        Me.llbEmail_UncekAll.Size = New System.Drawing.Size(65, 13)
        Me.llbEmail_UncekAll.TabIndex = 58
        Me.llbEmail_UncekAll.TabStop = True
        Me.llbEmail_UncekAll.Text = "Uncheck All"
        '
        'llbEmail_CekAll
        '
        Me.llbEmail_CekAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbEmail_CekAll.AutoSize = True
        Me.llbEmail_CekAll.BackColor = System.Drawing.SystemColors.Control
        Me.llbEmail_CekAll.Location = New System.Drawing.Point(303, 479)
        Me.llbEmail_CekAll.Name = "llbEmail_CekAll"
        Me.llbEmail_CekAll.Size = New System.Drawing.Size(52, 13)
        Me.llbEmail_CekAll.TabIndex = 57
        Me.llbEmail_CekAll.TabStop = True
        Me.llbEmail_CekAll.Text = "Check All"
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(445, 479)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(58, 13)
        Me.Label7.TabIndex = 56
        Me.Label7.Text = "Print data :"
        '
        'BtnDelete
        '
        Me.BtnDelete.Location = New System.Drawing.Point(574, 17)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(48, 20)
        Me.BtnDelete.TabIndex = 51
        Me.BtnDelete.Text = "Delete"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(643, 479)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(68, 13)
        Me.Label6.TabIndex = 55
        Me.Label6.Text = "Delete data :"
        '
        'LblUncheckedDelete
        '
        Me.LblUncheckedDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUncheckedDelete.AutoSize = True
        Me.LblUncheckedDelete.BackColor = System.Drawing.SystemColors.Control
        Me.LblUncheckedDelete.Location = New System.Drawing.Point(774, 479)
        Me.LblUncheckedDelete.Name = "LblUncheckedDelete"
        Me.LblUncheckedDelete.Size = New System.Drawing.Size(65, 13)
        Me.LblUncheckedDelete.TabIndex = 54
        Me.LblUncheckedDelete.TabStop = True
        Me.LblUncheckedDelete.Text = "Uncheck All"
        '
        'LblCheckedDelete
        '
        Me.LblCheckedDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblCheckedDelete.AutoSize = True
        Me.LblCheckedDelete.BackColor = System.Drawing.SystemColors.Control
        Me.LblCheckedDelete.Location = New System.Drawing.Point(717, 479)
        Me.LblCheckedDelete.Name = "LblCheckedDelete"
        Me.LblCheckedDelete.Size = New System.Drawing.Size(52, 13)
        Me.LblCheckedDelete.TabIndex = 53
        Me.LblCheckedDelete.TabStop = True
        Me.LblCheckedDelete.Text = "Check All"
        '
        'DGVAttendanceData
        '
        Me.DGVAttendanceData.AllowUserToAddRows = False
        Me.DGVAttendanceData.AllowUserToDeleteRows = False
        Me.DGVAttendanceData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DGVAttendanceData.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVAttendanceData.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DGVAttendanceData.ColumnHeadersHeight = 30
        Me.DGVAttendanceData.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColDocumentNo, Me.ColDate, Me.ColEmployeeCode, Me.ColEmployeeName, Me.ColEmployeeNIK, Me.ColDeptCode, Me.ColDept, Me.ColJobPositionCode, Me.ColJobPosition, Me.ColStartDate, Me.ColEndDate, Me.ColEdit, Me.ColChkPrint, Me.ColChkDelete, Me.ColChkEmail, Me.ColPrimaryTransaksi})
        Me.DGVAttendanceData.Location = New System.Drawing.Point(4, 46)
        Me.DGVAttendanceData.Margin = New System.Windows.Forms.Padding(2)
        Me.DGVAttendanceData.Name = "DGVAttendanceData"
        Me.DGVAttendanceData.RowTemplate.Height = 24
        Me.DGVAttendanceData.Size = New System.Drawing.Size(835, 431)
        Me.DGVAttendanceData.TabIndex = 0
        '
        'LblUnchecked
        '
        Me.LblUnchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnchecked.AutoSize = True
        Me.LblUnchecked.BackColor = System.Drawing.SystemColors.Control
        Me.LblUnchecked.Location = New System.Drawing.Point(566, 479)
        Me.LblUnchecked.Name = "LblUnchecked"
        Me.LblUnchecked.Size = New System.Drawing.Size(65, 13)
        Me.LblUnchecked.TabIndex = 52
        Me.LblUnchecked.TabStop = True
        Me.LblUnchecked.Text = "Uncheck All"
        '
        'LblChecked
        '
        Me.LblChecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblChecked.AutoSize = True
        Me.LblChecked.BackColor = System.Drawing.SystemColors.Control
        Me.LblChecked.Location = New System.Drawing.Point(509, 479)
        Me.LblChecked.Name = "LblChecked"
        Me.LblChecked.Size = New System.Drawing.Size(52, 13)
        Me.LblChecked.TabIndex = 51
        Me.LblChecked.TabStop = True
        Me.LblChecked.Text = "Check All"
        '
        'BtnNew
        '
        Me.BtnNew.Location = New System.Drawing.Point(526, 17)
        Me.BtnNew.Name = "BtnNew"
        Me.BtnNew.Size = New System.Drawing.Size(42, 20)
        Me.BtnNew.TabIndex = 50
        Me.BtnNew.Text = "New"
        Me.BtnNew.UseVisualStyleBackColor = True
        '
        'BtnPrint
        '
        Me.BtnPrint.Location = New System.Drawing.Point(478, 17)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(42, 20)
        Me.BtnPrint.TabIndex = 49
        Me.BtnPrint.Text = "Print"
        Me.BtnPrint.UseVisualStyleBackColor = True
        '
        'BtnShow
        '
        Me.BtnShow.Location = New System.Drawing.Point(430, 17)
        Me.BtnShow.Name = "BtnShow"
        Me.BtnShow.Size = New System.Drawing.Size(42, 20)
        Me.BtnShow.TabIndex = 41
        Me.BtnShow.Text = "Show"
        Me.BtnShow.UseVisualStyleBackColor = True
        '
        'DtpTo_Attd
        '
        Me.DtpTo_Attd.CustomFormat = "dd MMM yyyy"
        Me.DtpTo_Attd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo_Attd.Location = New System.Drawing.Point(303, 18)
        Me.DtpTo_Attd.Name = "DtpTo_Attd"
        Me.DtpTo_Attd.Size = New System.Drawing.Size(122, 20)
        Me.DtpTo_Attd.TabIndex = 48
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(266, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 47
        Me.Label4.Text = "To"
        '
        'DtpPeriod_Attd
        '
        Me.DtpPeriod_Attd.CustomFormat = "dd MMM yyyy"
        Me.DtpPeriod_Attd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpPeriod_Attd.Location = New System.Drawing.Point(139, 18)
        Me.DtpPeriod_Attd.Name = "DtpPeriod_Attd"
        Me.DtpPeriod_Attd.Size = New System.Drawing.Size(122, 20)
        Me.DtpPeriod_Attd.TabIndex = 46
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
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(2, 17)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFind)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtFile)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.DtpTo_Import)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.DtpPeriod_Import)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Size = New System.Drawing.Size(719, 26)
        Me.SplitContainer1.SplitterDistance = 355
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 44
        '
        'DtpTo_Import
        '
        Me.DtpTo_Import.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTo_Import.CustomFormat = "dd MMM yyyy"
        Me.DtpTo_Import.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo_Import.Location = New System.Drawing.Point(242, 3)
        Me.DtpTo_Import.Name = "DtpTo_Import"
        Me.DtpTo_Import.Size = New System.Drawing.Size(120, 20)
        Me.DtpTo_Import.TabIndex = 44
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(206, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 43
        Me.Label3.Text = "To"
        '
        'DtpPeriod_Import
        '
        Me.DtpPeriod_Import.CustomFormat = "dd MMM yyyy"
        Me.DtpPeriod_Import.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpPeriod_Import.Location = New System.Drawing.Point(79, 3)
        Me.DtpPeriod_Import.Name = "DtpPeriod_Import"
        Me.DtpPeriod_Import.Size = New System.Drawing.Size(122, 20)
        Me.DtpPeriod_Import.TabIndex = 42
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 41
        Me.Label1.Text = "Period"
        '
        'gboImportData
        '
        Me.gboImportData.Controls.Add(Me.SplitContainer1)
        Me.gboImportData.Controls.Add(Me.BtnImport)
        Me.gboImportData.Location = New System.Drawing.Point(9, 10)
        Me.gboImportData.Margin = New System.Windows.Forms.Padding(2)
        Me.gboImportData.Name = "gboImportData"
        Me.gboImportData.Padding = New System.Windows.Forms.Padding(2)
        Me.gboImportData.Size = New System.Drawing.Size(728, 70)
        Me.gboImportData.TabIndex = 45
        Me.gboImportData.TabStop = False
        Me.gboImportData.Text = "Data Import"
        Me.gboImportData.Visible = False
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
        Me.ColDocumentNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDocumentNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColDocumentNo.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColDocumentNo.Width = 98
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.Width = 55
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
        DataGridViewCellStyle1.Format = "dd MMM yyyy"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.ColStartDate.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColStartDate.HeaderText = "Start Date"
        Me.ColStartDate.Name = "ColStartDate"
        Me.ColStartDate.ReadOnly = True
        Me.ColStartDate.Width = 80
        '
        'ColEndDate
        '
        DataGridViewCellStyle2.Format = "dd MMM yyyy"
        Me.ColEndDate.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColEndDate.HeaderText = "End Date"
        Me.ColEndDate.Name = "ColEndDate"
        Me.ColEndDate.ReadOnly = True
        Me.ColEndDate.Width = 77
        '
        'ColEdit
        '
        Me.ColEdit.HeaderText = "Edit"
        Me.ColEdit.Name = "ColEdit"
        Me.ColEdit.Width = 31
        '
        'ColChkPrint
        '
        Me.ColChkPrint.HeaderText = "Print"
        Me.ColChkPrint.Name = "ColChkPrint"
        Me.ColChkPrint.Width = 34
        '
        'ColChkDelete
        '
        Me.ColChkDelete.HeaderText = "Delete"
        Me.ColChkDelete.Name = "ColChkDelete"
        Me.ColChkDelete.Width = 44
        '
        'ColChkEmail
        '
        Me.ColChkEmail.HeaderText = "Email"
        Me.ColChkEmail.Name = "ColChkEmail"
        Me.ColChkEmail.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColChkEmail.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColChkEmail.Width = 57
        '
        'ColPrimaryTransaksi
        '
        Me.ColPrimaryTransaksi.HeaderText = "Primary Transaksi"
        Me.ColPrimaryTransaksi.Name = "ColPrimaryTransaksi"
        Me.ColPrimaryTransaksi.Visible = False
        Me.ColPrimaryTransaksi.Width = 115
        '
        'FrmTrans_Attendance_DataImport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button2
        Me.ClientSize = New System.Drawing.Size(855, 515)
        Me.Controls.Add(Me.gboAttendanceData)
        Me.Controls.Add(Me.gboImportData)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(758, 552)
        Me.Name = "FrmTrans_Attendance_DataImport"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Attendance Data Import"
        Me.gboAttendanceData.ResumeLayout(False)
        Me.gboAttendanceData.PerformLayout()
        CType(Me.DGVAttendanceData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.gboImportData.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnFind As System.Windows.Forms.Button
    Friend WithEvents TxtFile As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BtnImport As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents gboAttendanceData As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DtpTo_Import As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DtpPeriod_Import As System.Windows.Forms.DateTimePicker
    Friend WithEvents DGVAttendanceData As System.Windows.Forms.DataGridView
    Friend WithEvents gboImportData As System.Windows.Forms.GroupBox
    Friend WithEvents BtnShow As System.Windows.Forms.Button
    Friend WithEvents DtpTo_Attd As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DtpPeriod_Attd As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
    Friend WithEvents BtnNew As System.Windows.Forms.Button
    Friend WithEvents LblUnchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents LblChecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents LblUncheckedDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents LblCheckedDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents llbEmail_UncekAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbEmail_CekAll As System.Windows.Forms.LinkLabel
    Friend WithEvents btnSendEmail As System.Windows.Forms.Button
    Friend WithEvents CboTypeDateFilter As System.Windows.Forms.ComboBox
    Friend WithEvents lllSetDataSorting As System.Windows.Forms.LinkLabel
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocumentNo As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
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
    Friend WithEvents ColChkDelete As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColChkEmail As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimaryTransaksi As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
