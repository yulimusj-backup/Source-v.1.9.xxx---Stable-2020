<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Attendance_Log
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
        Me.BtnIndexSettings = New System.Windows.Forms.Button
        Me.BtnShowFromFile = New System.Windows.Forms.Button
        Me.ChkListServer = New System.Windows.Forms.CheckedListBox
        Me.BtnShowLog = New System.Windows.Forms.Button
        Me.BtnDelete = New System.Windows.Forms.Button
        Me.ChkDeleteBeforeImport = New System.Windows.Forms.CheckBox
        Me.BtnImport = New System.Windows.Forms.Button
        Me.BtnLoad = New System.Windows.Forms.Button
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFingerprintID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColServerName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTerminalID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryEmployee = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryServer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BtnClose = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.GroupBox1.SuspendLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.BtnIndexSettings)
        Me.GroupBox1.Controls.Add(Me.BtnShowFromFile)
        Me.GroupBox1.Controls.Add(Me.ChkListServer)
        Me.GroupBox1.Controls.Add(Me.BtnShowLog)
        Me.GroupBox1.Controls.Add(Me.BtnDelete)
        Me.GroupBox1.Controls.Add(Me.ChkDeleteBeforeImport)
        Me.GroupBox1.Controls.Add(Me.BtnImport)
        Me.GroupBox1.Controls.Add(Me.BtnLoad)
        Me.GroupBox1.Controls.Add(Me.DtpTo)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.DtpFrom)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(911, 67)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'BtnIndexSettings
        '
        Me.BtnIndexSettings.Location = New System.Drawing.Point(564, 39)
        Me.BtnIndexSettings.Name = "BtnIndexSettings"
        Me.BtnIndexSettings.Size = New System.Drawing.Size(24, 23)
        Me.BtnIndexSettings.TabIndex = 9
        Me.BtnIndexSettings.Text = "..."
        Me.BtnIndexSettings.UseVisualStyleBackColor = True
        '
        'BtnShowFromFile
        '
        Me.BtnShowFromFile.Location = New System.Drawing.Point(471, 39)
        Me.BtnShowFromFile.Name = "BtnShowFromFile"
        Me.BtnShowFromFile.Size = New System.Drawing.Size(91, 23)
        Me.BtnShowFromFile.TabIndex = 8
        Me.BtnShowFromFile.Text = "&Show From File"
        Me.BtnShowFromFile.UseVisualStyleBackColor = True
        '
        'ChkListServer
        '
        Me.ChkListServer.FormattingEnabled = True
        Me.ChkListServer.Location = New System.Drawing.Point(200, 9)
        Me.ChkListServer.Name = "ChkListServer"
        Me.ChkListServer.Size = New System.Drawing.Size(265, 49)
        Me.ChkListServer.TabIndex = 7
        '
        'BtnShowLog
        '
        Me.BtnShowLog.Location = New System.Drawing.Point(594, 11)
        Me.BtnShowLog.Name = "BtnShowLog"
        Me.BtnShowLog.Size = New System.Drawing.Size(108, 23)
        Me.BtnShowLog.TabIndex = 6
        Me.BtnShowLog.Text = "Show log"
        Me.BtnShowLog.UseVisualStyleBackColor = True
        '
        'BtnDelete
        '
        Me.BtnDelete.Location = New System.Drawing.Point(708, 11)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(108, 23)
        Me.BtnDelete.TabIndex = 5
        Me.BtnDelete.Text = "Delete log"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'ChkDeleteBeforeImport
        '
        Me.ChkDeleteBeforeImport.AutoSize = True
        Me.ChkDeleteBeforeImport.Location = New System.Drawing.Point(708, 43)
        Me.ChkDeleteBeforeImport.Name = "ChkDeleteBeforeImport"
        Me.ChkDeleteBeforeImport.Size = New System.Drawing.Size(133, 17)
        Me.ChkDeleteBeforeImport.TabIndex = 2
        Me.ChkDeleteBeforeImport.Text = "Delete log before save"
        Me.ChkDeleteBeforeImport.UseVisualStyleBackColor = True
        '
        'BtnImport
        '
        Me.BtnImport.Location = New System.Drawing.Point(594, 39)
        Me.BtnImport.Name = "BtnImport"
        Me.BtnImport.Size = New System.Drawing.Size(108, 23)
        Me.BtnImport.TabIndex = 4
        Me.BtnImport.Text = "Save to log"
        Me.BtnImport.UseVisualStyleBackColor = True
        '
        'BtnLoad
        '
        Me.BtnLoad.Location = New System.Drawing.Point(471, 10)
        Me.BtnLoad.Name = "BtnLoad"
        Me.BtnLoad.Size = New System.Drawing.Size(117, 23)
        Me.BtnLoad.TabIndex = 3
        Me.BtnLoad.Text = "Show From Server"
        Me.BtnLoad.UseVisualStyleBackColor = True
        '
        'DtpTo
        '
        Me.DtpTo.CustomFormat = "dd MMMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(52, 35)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(142, 20)
        Me.DtpTo.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To"
        '
        'DtpFrom
        '
        Me.DtpFrom.CustomFormat = "dd MMMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(52, 11)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(142, 20)
        Me.DtpFrom.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From"
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.AllowUserToDeleteRows = False
        Me.DgvItem.AllowUserToResizeRows = False
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColEmployeeCode, Me.ColEmployeeName, Me.ColFingerprintID, Me.ColServerName, Me.ColLocation, Me.ColTerminalID, Me.ColTime, Me.ColPrimaryEmployee, Me.ColPrimaryServer})
        Me.DgvItem.Location = New System.Drawing.Point(12, 90)
        Me.DgvItem.MultiSelect = False
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowHeadersVisible = False
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.Size = New System.Drawing.Size(911, 409)
        Me.DgvItem.TabIndex = 0
        '
        'ColNo
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.ColNo.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 46
        '
        'ColEmployeeCode
        '
        Me.ColEmployeeCode.HeaderText = "Employee Code"
        Me.ColEmployeeCode.Name = "ColEmployeeCode"
        Me.ColEmployeeCode.ReadOnly = True
        Me.ColEmployeeCode.Width = 106
        '
        'ColEmployeeName
        '
        Me.ColEmployeeName.HeaderText = "Employee Name"
        Me.ColEmployeeName.Name = "ColEmployeeName"
        Me.ColEmployeeName.ReadOnly = True
        Me.ColEmployeeName.Width = 109
        '
        'ColFingerprintID
        '
        Me.ColFingerprintID.HeaderText = "Fingerprint ID"
        Me.ColFingerprintID.Name = "ColFingerprintID"
        Me.ColFingerprintID.ReadOnly = True
        Me.ColFingerprintID.Width = 95
        '
        'ColServerName
        '
        Me.ColServerName.HeaderText = "Server Name"
        Me.ColServerName.Name = "ColServerName"
        Me.ColServerName.ReadOnly = True
        Me.ColServerName.Width = 94
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Sever Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.Width = 104
        '
        'ColTerminalID
        '
        Me.ColTerminalID.HeaderText = "Terminal ID"
        Me.ColTerminalID.Name = "ColTerminalID"
        Me.ColTerminalID.ReadOnly = True
        Me.ColTerminalID.Width = 86
        '
        'ColTime
        '
        Me.ColTime.HeaderText = "Time"
        Me.ColTime.Name = "ColTime"
        Me.ColTime.ReadOnly = True
        Me.ColTime.Width = 55
        '
        'ColPrimaryEmployee
        '
        Me.ColPrimaryEmployee.HeaderText = "Primary Employee"
        Me.ColPrimaryEmployee.Name = "ColPrimaryEmployee"
        Me.ColPrimaryEmployee.ReadOnly = True
        Me.ColPrimaryEmployee.Visible = False
        Me.ColPrimaryEmployee.Width = 115
        '
        'ColPrimaryServer
        '
        Me.ColPrimaryServer.HeaderText = "Primary Server"
        Me.ColPrimaryServer.Name = "ColPrimaryServer"
        Me.ColPrimaryServer.ReadOnly = True
        Me.ColPrimaryServer.Visible = False
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(833, 456)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 6
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 74)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Showing: "
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "SDF File|*.sdf"
        '
        'FrmTrans_Attendance_Log
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(934, 511)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.DgvItem)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BtnClose)
        Me.Name = "FrmTrans_Attendance_Log"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Import Attendance Log"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BtnLoad As System.Windows.Forms.Button
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BtnImport As System.Windows.Forms.Button
    Friend WithEvents ChkDeleteBeforeImport As System.Windows.Forms.CheckBox
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents BtnShowLog As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ChkListServer As System.Windows.Forms.CheckedListBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFingerprintID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColServerName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTerminalID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryEmployee As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryServer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BtnIndexSettings As System.Windows.Forms.Button
    Friend WithEvents BtnShowFromFile As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
End Class
