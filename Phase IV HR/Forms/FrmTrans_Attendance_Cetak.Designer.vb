<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Attendance_Cetak
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.gboFormat = New System.Windows.Forms.GroupBox
        Me.BtnPrintAll = New System.Windows.Forms.Button
        Me.dtpStart = New System.Windows.Forms.DateTimePicker
        Me.DtpEnd = New System.Windows.Forms.DateTimePicker
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtPrimaryAttendance = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.BtnNext = New System.Windows.Forms.Button
        Me.BtnPrevious = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtIndex = New System.Windows.Forms.TextBox
        Me.RboA5 = New System.Windows.Forms.RadioButton
        Me.RboA4 = New System.Windows.Forms.RadioButton
        Me.TxtNoAttendance = New System.Windows.Forms.TextBox
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.CboFormat = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.gboDetail.SuspendLayout()
        Me.gboFormat.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(421, 270)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(43, 23)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crv1.Location = New System.Drawing.Point(0, 0)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(734, 512)
        Me.crv1.TabIndex = 0
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.CboFormat)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.gboFormat)
        Me.gboDetail.Location = New System.Drawing.Point(12, 192)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(532, 308)
        Me.gboDetail.TabIndex = 1
        Me.gboDetail.TabStop = False
        Me.gboDetail.Visible = False
        '
        'gboFormat
        '
        Me.gboFormat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboFormat.Controls.Add(Me.BtnPrintAll)
        Me.gboFormat.Controls.Add(Me.dtpStart)
        Me.gboFormat.Controls.Add(Me.DtpEnd)
        Me.gboFormat.Controls.Add(Me.Label6)
        Me.gboFormat.Controls.Add(Me.Label7)
        Me.gboFormat.Controls.Add(Me.TxtPrimaryAttendance)
        Me.gboFormat.Controls.Add(Me.Label2)
        Me.gboFormat.Controls.Add(Me.BtnNext)
        Me.gboFormat.Controls.Add(Me.BtnPrevious)
        Me.gboFormat.Controls.Add(Me.Label1)
        Me.gboFormat.Controls.Add(Me.TxtIndex)
        Me.gboFormat.Controls.Add(Me.RboA5)
        Me.gboFormat.Controls.Add(Me.RboA4)
        Me.gboFormat.Controls.Add(Me.TxtNoAttendance)
        Me.gboFormat.Location = New System.Drawing.Point(0, 106)
        Me.gboFormat.Name = "gboFormat"
        Me.gboFormat.Size = New System.Drawing.Size(529, 136)
        Me.gboFormat.TabIndex = 0
        Me.gboFormat.TabStop = False
        Me.gboFormat.Visible = False
        '
        'BtnPrintAll
        '
        Me.BtnPrintAll.Location = New System.Drawing.Point(246, 75)
        Me.BtnPrintAll.Name = "BtnPrintAll"
        Me.BtnPrintAll.Size = New System.Drawing.Size(59, 23)
        Me.BtnPrintAll.TabIndex = 49
        Me.BtnPrintAll.Text = "Print All"
        Me.BtnPrintAll.UseVisualStyleBackColor = True
        '
        'dtpStart
        '
        Me.dtpStart.CustomFormat = "dd MMM yyyy"
        Me.dtpStart.Enabled = False
        Me.dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStart.Location = New System.Drawing.Point(93, 46)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.Size = New System.Drawing.Size(147, 20)
        Me.dtpStart.TabIndex = 1
        '
        'DtpEnd
        '
        Me.DtpEnd.CustomFormat = "dd MMM yyyy"
        Me.DtpEnd.Enabled = False
        Me.DtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpEnd.Location = New System.Drawing.Point(93, 75)
        Me.DtpEnd.Name = "DtpEnd"
        Me.DtpEnd.Size = New System.Drawing.Size(147, 20)
        Me.DtpEnd.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 46)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(29, 13)
        Me.Label6.TabIndex = 47
        Me.Label6.Text = "Start"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 71)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(26, 13)
        Me.Label7.TabIndex = 48
        Me.Label7.Text = "End"
        '
        'TxtPrimaryAttendance
        '
        Me.TxtPrimaryAttendance.Location = New System.Drawing.Point(341, 22)
        Me.TxtPrimaryAttendance.Name = "TxtPrimaryAttendance"
        Me.TxtPrimaryAttendance.Size = New System.Drawing.Size(24, 20)
        Me.TxtPrimaryAttendance.TabIndex = 15
        Me.TxtPrimaryAttendance.TabStop = False
        Me.TxtPrimaryAttendance.Text = "-1"
        Me.TxtPrimaryAttendance.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 97)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Format"
        '
        'BtnNext
        '
        Me.BtnNext.Location = New System.Drawing.Point(246, 45)
        Me.BtnNext.Name = "BtnNext"
        Me.BtnNext.Size = New System.Drawing.Size(59, 23)
        Me.BtnNext.TabIndex = 6
        Me.BtnNext.Text = "Next"
        Me.BtnNext.UseVisualStyleBackColor = True
        '
        'BtnPrevious
        '
        Me.BtnPrevious.Location = New System.Drawing.Point(246, 20)
        Me.BtnPrevious.Name = "BtnPrevious"
        Me.BtnPrevious.Size = New System.Drawing.Size(59, 23)
        Me.BtnPrevious.TabIndex = 5
        Me.BtnPrevious.Text = "Previous"
        Me.BtnPrevious.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Employee Code"
        '
        'TxtIndex
        '
        Me.TxtIndex.Location = New System.Drawing.Point(311, 23)
        Me.TxtIndex.Name = "TxtIndex"
        Me.TxtIndex.Size = New System.Drawing.Size(24, 20)
        Me.TxtIndex.TabIndex = 10
        Me.TxtIndex.TabStop = False
        Me.TxtIndex.Text = "0"
        Me.TxtIndex.Visible = False
        '
        'RboA5
        '
        Me.RboA5.AutoSize = True
        Me.RboA5.Checked = True
        Me.RboA5.Location = New System.Drawing.Point(93, 97)
        Me.RboA5.Name = "RboA5"
        Me.RboA5.Size = New System.Drawing.Size(38, 17)
        Me.RboA5.TabIndex = 3
        Me.RboA5.TabStop = True
        Me.RboA5.Text = "A5"
        Me.RboA5.UseVisualStyleBackColor = True
        '
        'RboA4
        '
        Me.RboA4.AutoSize = True
        Me.RboA4.Location = New System.Drawing.Point(153, 97)
        Me.RboA4.Name = "RboA4"
        Me.RboA4.Size = New System.Drawing.Size(38, 17)
        Me.RboA4.TabIndex = 4
        Me.RboA4.Text = "A4"
        Me.RboA4.UseVisualStyleBackColor = True
        '
        'TxtNoAttendance
        '
        Me.TxtNoAttendance.Location = New System.Drawing.Point(93, 19)
        Me.TxtNoAttendance.Name = "TxtNoAttendance"
        Me.TxtNoAttendance.ReadOnly = True
        Me.TxtNoAttendance.Size = New System.Drawing.Size(147, 20)
        Me.TxtNoAttendance.TabIndex = 0
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'CboFormat
        '
        Me.CboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFormat.FormattingEnabled = True
        Me.CboFormat.Items.AddRange(New Object() {"A5", "A4"})
        Me.CboFormat.Location = New System.Drawing.Point(70, 79)
        Me.CboFormat.Name = "CboFormat"
        Me.CboFormat.Size = New System.Drawing.Size(121, 21)
        Me.CboFormat.TabIndex = 50
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 82)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 51
        Me.Label3.Text = "Format"
        '
        'FrmTrans_Attendance_Cetak
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(734, 512)
        Me.Controls.Add(Me.crv1)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnClose)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(750, 550)
        Me.Name = "FrmTrans_Attendance_Cetak"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Detail"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.gboFormat.ResumeLayout(False)
        Me.gboFormat.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents RboA5 As System.Windows.Forms.RadioButton
    Friend WithEvents RboA4 As System.Windows.Forms.RadioButton
    Friend WithEvents gboFormat As System.Windows.Forms.GroupBox
    Friend WithEvents TxtIndex As System.Windows.Forms.TextBox
    Friend WithEvents TxtNoAttendance As System.Windows.Forms.TextBox
    Friend WithEvents BtnNext As System.Windows.Forms.Button
    Friend WithEvents BtnPrevious As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryAttendance As System.Windows.Forms.TextBox
    Friend WithEvents dtpStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents BtnPrintAll As System.Windows.Forms.Button
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents CboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
