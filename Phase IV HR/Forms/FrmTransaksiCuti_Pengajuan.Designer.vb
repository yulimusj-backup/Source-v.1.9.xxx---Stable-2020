<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiCuti_Pengajuan
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
        Me.components = New System.ComponentModel.Container
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtDocNumber = New System.Windows.Forms.TextBox
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.btnFind = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TxtReason = New System.Windows.Forms.TextBox
        Me.TxtPrimaryLeave = New System.Windows.Forms.TextBox
        Me.TblmastercutiBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DseTrans_Cuti_Pengajuan = New HR.DseTrans_Cuti_Pengajuan
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtPrimaryPosition = New System.Windows.Forms.TextBox
        Me.TxtPosition = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.DtpDate = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtEmployeeName = New System.Windows.Forms.TextBox
        Me.BtnFindEmployeeCode = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtEmployeeCode = New System.Windows.Forms.TextBox
        Me.TxtPrimaryEmployee = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboLeave = New System.Windows.Forms.ComboBox
        Me.DtpEnd = New System.Windows.Forms.DateTimePicker
        Me.llbDuration = New System.Windows.Forms.LinkLabel
        Me.GboApproval = New System.Windows.Forms.GroupBox
        Me.llbNotes = New System.Windows.Forms.LinkLabel
        Me.DtpApp = New System.Windows.Forms.DateTimePicker
        Me.TxtAppBy = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.TxtStatus = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.dtpStart = New System.Windows.Forms.DateTimePicker
        Me.llbBalance = New System.Windows.Forms.LinkLabel
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtDuration = New System.Windows.Forms.TextBox
        Me.TxtBalance = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.ChkHalfDayStart = New System.Windows.Forms.CheckBox
        Me.ChkHalfDayEnd = New System.Windows.Forms.CheckBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.Tblmaster_cutiTableAdapter = New HR.DseTrans_Cuti_PengajuanTableAdapters.tblmaster_cutiTableAdapter
        Me.llbGoToApproval = New System.Windows.Forms.LinkLabel
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.TblmastercutiBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DseTrans_Cuti_Pengajuan, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GboApproval.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(605, 342)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(101, 22)
        Me.TxtReadOnly.TabIndex = 4
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(541, 385)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(107, 32)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(312, 385)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(107, 32)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(356, 337)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(107, 32)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 7)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 17)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Document No"
        '
        'TxtDocNumber
        '
        Me.TxtDocNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDocNumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDocNumber.Location = New System.Drawing.Point(124, 4)
        Me.TxtDocNumber.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtDocNumber.MaxLength = 50
        Me.TxtDocNumber.Name = "TxtDocNumber"
        Me.TxtDocNumber.ReadOnly = True
        Me.TxtDocNumber.Size = New System.Drawing.Size(324, 22)
        Me.TxtDocNumber.TabIndex = 0
        Me.TxtDocNumber.TabStop = False
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(470, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 441)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip2.Size = New System.Drawing.Size(960, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 54
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(470, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(471, 337)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(107, 32)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(470, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 466)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(960, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 53
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(470, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(656, 385)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(107, 32)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(427, 385)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(107, 32)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(8, 4)
        Me.TxtPrimary.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(36, 22)
        Me.TxtPrimary.TabIndex = 0
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(197, 385)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(107, 32)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.GroupBox1)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.TxtReadOnly)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(12, 1)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Size = New System.Drawing.Size(935, 377)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TxtReason)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryLeave)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 206)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(919, 124)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Reason"
        '
        'TxtReason
        '
        Me.TxtReason.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtReason.Location = New System.Drawing.Point(4, 19)
        Me.TxtReason.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtReason.MaxLength = 100
        Me.TxtReason.Multiline = True
        Me.TxtReason.Name = "TxtReason"
        Me.TxtReason.Size = New System.Drawing.Size(911, 101)
        Me.TxtReason.TabIndex = 0
        '
        'TxtPrimaryLeave
        '
        Me.TxtPrimaryLeave.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.TblmastercutiBindingSource, "primary", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.TxtPrimaryLeave.Location = New System.Drawing.Point(9, 23)
        Me.TxtPrimaryLeave.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryLeave.Name = "TxtPrimaryLeave"
        Me.TxtPrimaryLeave.Size = New System.Drawing.Size(36, 22)
        Me.TxtPrimaryLeave.TabIndex = 40
        Me.TxtPrimaryLeave.TabStop = False
        Me.TxtPrimaryLeave.Text = "-1"
        '
        'TblmastercutiBindingSource
        '
        Me.TblmastercutiBindingSource.DataMember = "tblmaster_cuti"
        Me.TblmastercutiBindingSource.DataSource = Me.DseTrans_Cuti_Pengajuan
        '
        'DseTrans_Cuti_Pengajuan
        '
        Me.DseTrans_Cuti_Pengajuan.DataSetName = "DseTrans_Cuti_Pengajuan"
        Me.DseTrans_Cuti_Pengajuan.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(8, 14)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryPosition)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPosition)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtEmployeeName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindEmployeeCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtEmployeeCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryEmployee)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDocNumber)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboLeave)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.DtpEnd)
        Me.SplitContainer1.Panel2.Controls.Add(Me.llbDuration)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GboApproval)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dtpStart)
        Me.SplitContainer1.Panel2.Controls.Add(Me.llbBalance)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDuration)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtBalance)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel2.Controls.Add(Me.ChkHalfDayStart)
        Me.SplitContainer1.Panel2.Controls.Add(Me.ChkHalfDayEnd)
        Me.SplitContainer1.Size = New System.Drawing.Size(919, 190)
        Me.SplitContainer1.SplitterDistance = 454
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'TxtPrimaryPosition
        '
        Me.TxtPrimaryPosition.Location = New System.Drawing.Point(83, 129)
        Me.TxtPrimaryPosition.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryPosition.Name = "TxtPrimaryPosition"
        Me.TxtPrimaryPosition.Size = New System.Drawing.Size(36, 22)
        Me.TxtPrimaryPosition.TabIndex = 53
        Me.TxtPrimaryPosition.TabStop = False
        Me.TxtPrimaryPosition.Text = "-1"
        Me.TxtPrimaryPosition.Visible = False
        '
        'TxtPosition
        '
        Me.TxtPosition.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPosition.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPosition.Location = New System.Drawing.Point(124, 129)
        Me.TxtPosition.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPosition.MaxLength = 50
        Me.TxtPosition.Name = "TxtPosition"
        Me.TxtPosition.ReadOnly = True
        Me.TxtPosition.Size = New System.Drawing.Size(324, 22)
        Me.TxtPosition.TabIndex = 4
        Me.TxtPosition.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 133)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 17)
        Me.Label4.TabIndex = 52
        Me.Label4.Text = "Position"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(5, 43)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(38, 17)
        Me.Label11.TabIndex = 50
        Me.Label11.Text = "Date"
        '
        'DtpDate
        '
        Me.DtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDate.CustomFormat = "dd MMM yyyy"
        Me.DtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDate.Location = New System.Drawing.Point(124, 36)
        Me.DtpDate.Margin = New System.Windows.Forms.Padding(4)
        Me.DtpDate.Name = "DtpDate"
        Me.DtpDate.Size = New System.Drawing.Size(324, 22)
        Me.DtpDate.TabIndex = 0
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(4, 165)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(84, 17)
        Me.Label5.TabIndex = 41
        Me.Label5.Text = "Leave Code"
        '
        'TxtEmployeeName
        '
        Me.TxtEmployeeName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmployeeName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtEmployeeName.Location = New System.Drawing.Point(124, 100)
        Me.TxtEmployeeName.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtEmployeeName.MaxLength = 50
        Me.TxtEmployeeName.Name = "TxtEmployeeName"
        Me.TxtEmployeeName.ReadOnly = True
        Me.TxtEmployeeName.Size = New System.Drawing.Size(324, 22)
        Me.TxtEmployeeName.TabIndex = 3
        Me.TxtEmployeeName.TabStop = False
        '
        'BtnFindEmployeeCode
        '
        Me.BtnFindEmployeeCode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindEmployeeCode.Location = New System.Drawing.Point(394, 68)
        Me.BtnFindEmployeeCode.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnFindEmployeeCode.Name = "BtnFindEmployeeCode"
        Me.BtnFindEmployeeCode.Size = New System.Drawing.Size(56, 25)
        Me.BtnFindEmployeeCode.TabIndex = 2
        Me.BtnFindEmployeeCode.TabStop = False
        Me.BtnFindEmployeeCode.Text = "Find"
        Me.BtnFindEmployeeCode.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 103)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(111, 17)
        Me.Label3.TabIndex = 38
        Me.Label3.Text = "Employee Name"
        '
        'TxtEmployeeCode
        '
        Me.TxtEmployeeCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmployeeCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtEmployeeCode.Location = New System.Drawing.Point(124, 68)
        Me.TxtEmployeeCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtEmployeeCode.MaxLength = 100
        Me.TxtEmployeeCode.Name = "TxtEmployeeCode"
        Me.TxtEmployeeCode.Size = New System.Drawing.Size(260, 22)
        Me.TxtEmployeeCode.TabIndex = 1
        '
        'TxtPrimaryEmployee
        '
        Me.TxtPrimaryEmployee.Location = New System.Drawing.Point(83, 68)
        Me.TxtPrimaryEmployee.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryEmployee.Name = "TxtPrimaryEmployee"
        Me.TxtPrimaryEmployee.Size = New System.Drawing.Size(36, 22)
        Me.TxtPrimaryEmployee.TabIndex = 35
        Me.TxtPrimaryEmployee.TabStop = False
        Me.TxtPrimaryEmployee.Text = "-1"
        Me.TxtPrimaryEmployee.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 71)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(107, 17)
        Me.Label2.TabIndex = 37
        Me.Label2.Text = "Employee Code"
        '
        'cboLeave
        '
        Me.cboLeave.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboLeave.DataSource = Me.TblmastercutiBindingSource
        Me.cboLeave.DisplayMember = "CodeDesc"
        Me.cboLeave.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLeave.FormattingEnabled = True
        Me.cboLeave.Location = New System.Drawing.Point(124, 161)
        Me.cboLeave.Margin = New System.Windows.Forms.Padding(4)
        Me.cboLeave.Name = "cboLeave"
        Me.cboLeave.Size = New System.Drawing.Size(324, 24)
        Me.cboLeave.TabIndex = 2
        Me.cboLeave.ValueMember = "primary"
        '
        'DtpEnd
        '
        Me.DtpEnd.CustomFormat = "dd MMM yyyy"
        Me.DtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpEnd.Location = New System.Drawing.Point(103, 128)
        Me.DtpEnd.Margin = New System.Windows.Forms.Padding(4)
        Me.DtpEnd.Name = "DtpEnd"
        Me.DtpEnd.Size = New System.Drawing.Size(252, 22)
        Me.DtpEnd.TabIndex = 7
        '
        'llbDuration
        '
        Me.llbDuration.AutoSize = True
        Me.llbDuration.Location = New System.Drawing.Point(12, 164)
        Me.llbDuration.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbDuration.Name = "llbDuration"
        Me.llbDuration.Size = New System.Drawing.Size(62, 17)
        Me.llbDuration.TabIndex = 9
        Me.llbDuration.TabStop = True
        Me.llbDuration.Text = "Duration"
        '
        'GboApproval
        '
        Me.GboApproval.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GboApproval.Controls.Add(Me.llbNotes)
        Me.GboApproval.Controls.Add(Me.DtpApp)
        Me.GboApproval.Controls.Add(Me.TxtAppBy)
        Me.GboApproval.Controls.Add(Me.Label12)
        Me.GboApproval.Controls.Add(Me.TxtStatus)
        Me.GboApproval.Controls.Add(Me.Label13)
        Me.GboApproval.Location = New System.Drawing.Point(4, 10)
        Me.GboApproval.Margin = New System.Windows.Forms.Padding(4)
        Me.GboApproval.Name = "GboApproval"
        Me.GboApproval.Padding = New System.Windows.Forms.Padding(4)
        Me.GboApproval.Size = New System.Drawing.Size(452, 82)
        Me.GboApproval.TabIndex = 0
        Me.GboApproval.TabStop = False
        Me.GboApproval.Text = "Approval"
        '
        'llbNotes
        '
        Me.llbNotes.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbNotes.AutoSize = True
        Me.llbNotes.Location = New System.Drawing.Point(394, 26)
        Me.llbNotes.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbNotes.Name = "llbNotes"
        Me.llbNotes.Size = New System.Drawing.Size(45, 17)
        Me.llbNotes.TabIndex = 3
        Me.llbNotes.TabStop = True
        Me.llbNotes.Text = "Notes"
        '
        'DtpApp
        '
        Me.DtpApp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpApp.CustomFormat = "dd MMM yyyy"
        Me.DtpApp.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpApp.Location = New System.Drawing.Point(289, 53)
        Me.DtpApp.Margin = New System.Windows.Forms.Padding(4)
        Me.DtpApp.Name = "DtpApp"
        Me.DtpApp.Size = New System.Drawing.Size(153, 22)
        Me.DtpApp.TabIndex = 4
        '
        'TxtAppBy
        '
        Me.TxtAppBy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAppBy.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAppBy.Location = New System.Drawing.Point(99, 53)
        Me.TxtAppBy.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAppBy.MaxLength = 50
        Me.TxtAppBy.Name = "TxtAppBy"
        Me.TxtAppBy.ReadOnly = True
        Me.TxtAppBy.Size = New System.Drawing.Size(178, 22)
        Me.TxtAppBy.TabIndex = 1
        Me.TxtAppBy.TabStop = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(8, 57)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(81, 17)
        Me.Label12.TabIndex = 56
        Me.Label12.Text = "Approve By"
        '
        'TxtStatus
        '
        Me.TxtStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtStatus.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtStatus.Location = New System.Drawing.Point(99, 23)
        Me.TxtStatus.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtStatus.MaxLength = 50
        Me.TxtStatus.Name = "TxtStatus"
        Me.TxtStatus.ReadOnly = True
        Me.TxtStatus.Size = New System.Drawing.Size(286, 22)
        Me.TxtStatus.TabIndex = 0
        Me.TxtStatus.TabStop = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(8, 27)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(48, 17)
        Me.Label13.TabIndex = 54
        Me.Label13.Text = "Status"
        '
        'dtpStart
        '
        Me.dtpStart.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpStart.CustomFormat = "dd MMM yyyy"
        Me.dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStart.Location = New System.Drawing.Point(103, 96)
        Me.dtpStart.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.Size = New System.Drawing.Size(254, 22)
        Me.dtpStart.TabIndex = 5
        '
        'llbBalance
        '
        Me.llbBalance.AutoSize = True
        Me.llbBalance.Location = New System.Drawing.Point(221, 164)
        Me.llbBalance.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbBalance.Name = "llbBalance"
        Me.llbBalance.Size = New System.Drawing.Size(59, 17)
        Me.llbBalance.TabIndex = 10
        Me.llbBalance.TabStop = True
        Me.llbBalance.Text = "Balance"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 103)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(38, 17)
        Me.Label6.TabIndex = 43
        Me.Label6.Text = "Start"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(12, 135)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(33, 17)
        Me.Label7.TabIndex = 44
        Me.Label7.Text = "End"
        '
        'Label10
        '
        Me.Label10.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(398, 164)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(40, 17)
        Me.Label10.TabIndex = 10
        Me.Label10.Text = "Days"
        '
        'TxtDuration
        '
        Me.TxtDuration.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDuration.Location = New System.Drawing.Point(103, 160)
        Me.TxtDuration.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtDuration.MaxLength = 50
        Me.TxtDuration.Name = "TxtDuration"
        Me.TxtDuration.ReadOnly = True
        Me.TxtDuration.Size = New System.Drawing.Size(44, 22)
        Me.TxtDuration.TabIndex = 6
        Me.TxtDuration.TabStop = False
        Me.TxtDuration.Text = "0"
        Me.TxtDuration.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtBalance
        '
        Me.TxtBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBalance.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBalance.Location = New System.Drawing.Point(291, 160)
        Me.TxtBalance.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtBalance.MaxLength = 50
        Me.TxtBalance.Name = "TxtBalance"
        Me.TxtBalance.ReadOnly = True
        Me.TxtBalance.Size = New System.Drawing.Size(98, 22)
        Me.TxtBalance.TabIndex = 9
        Me.TxtBalance.TabStop = False
        Me.TxtBalance.Text = "0"
        Me.TxtBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(156, 164)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(40, 17)
        Me.Label9.TabIndex = 7
        Me.Label9.Text = "Days"
        '
        'ChkHalfDayStart
        '
        Me.ChkHalfDayStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ChkHalfDayStart.AutoSize = True
        Me.ChkHalfDayStart.Location = New System.Drawing.Point(370, 98)
        Me.ChkHalfDayStart.Margin = New System.Windows.Forms.Padding(4)
        Me.ChkHalfDayStart.Name = "ChkHalfDayStart"
        Me.ChkHalfDayStart.Size = New System.Drawing.Size(84, 21)
        Me.ChkHalfDayStart.TabIndex = 6
        Me.ChkHalfDayStart.Text = "Half Day"
        Me.ChkHalfDayStart.UseVisualStyleBackColor = True
        '
        'ChkHalfDayEnd
        '
        Me.ChkHalfDayEnd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ChkHalfDayEnd.AutoSize = True
        Me.ChkHalfDayEnd.Location = New System.Drawing.Point(370, 129)
        Me.ChkHalfDayEnd.Margin = New System.Windows.Forms.Padding(4)
        Me.ChkHalfDayEnd.Name = "ChkHalfDayEnd"
        Me.ChkHalfDayEnd.Size = New System.Drawing.Size(84, 21)
        Me.ChkHalfDayEnd.TabIndex = 8
        Me.ChkHalfDayEnd.Text = "Half Day"
        Me.ChkHalfDayEnd.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'Tblmaster_cutiTableAdapter
        '
        Me.Tblmaster_cutiTableAdapter.ClearBeforeFill = True
        '
        'llbGoToApproval
        '
        Me.llbGoToApproval.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbGoToApproval.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbGoToApproval.AutoSize = True
        Me.llbGoToApproval.Location = New System.Drawing.Point(9, 393)
        Me.llbGoToApproval.Name = "llbGoToApproval"
        Me.llbGoToApproval.Size = New System.Drawing.Size(151, 17)
        Me.llbGoToApproval.TabIndex = 55
        Me.llbGoToApproval.TabStop = True
        Me.llbGoToApproval.Text = "Go To Leave Approval"
        Me.llbGoToApproval.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'FrmTransaksiCuti_Pengajuan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(960, 491)
        Me.Controls.Add(Me.llbGoToApproval)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.gboDetail)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimumSize = New System.Drawing.Size(968, 528)
        Me.Name = "FrmTransaksiCuti_Pengajuan"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Leave Application"
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.TblmastercutiBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DseTrans_Cuti_Pengajuan, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.GboApproval.ResumeLayout(False)
        Me.GboApproval.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtDocNumber As System.Windows.Forms.TextBox
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents llbBalance As System.Windows.Forms.LinkLabel
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtBalance As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ChkHalfDayEnd As System.Windows.Forms.CheckBox
    Friend WithEvents ChkHalfDayStart As System.Windows.Forms.CheckBox
    Friend WithEvents TxtDuration As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryLeave As System.Windows.Forms.TextBox
    Friend WithEvents cboLeave As System.Windows.Forms.ComboBox
    Friend WithEvents TxtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindEmployeeCode As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtEmployeeCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryEmployee As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents DtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPosition As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GboApproval As System.Windows.Forms.GroupBox
    Friend WithEvents TxtAppBy As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtStatus As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents DtpApp As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryPosition As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtReason As System.Windows.Forms.TextBox
    Friend WithEvents llbNotes As System.Windows.Forms.LinkLabel
    Friend WithEvents llbDuration As System.Windows.Forms.LinkLabel
    Friend WithEvents DtpEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Public WithEvents TblmastercutiBindingSource As System.Windows.Forms.BindingSource
    Public WithEvents DseTrans_Cuti_Pengajuan As HR.DseTrans_Cuti_Pengajuan
    Public WithEvents Tblmaster_cutiTableAdapter As HR.DseTrans_Cuti_PengajuanTableAdapters.tblmaster_cutiTableAdapter
    Friend WithEvents llbGoToApproval As System.Windows.Forms.LinkLabel
End Class
