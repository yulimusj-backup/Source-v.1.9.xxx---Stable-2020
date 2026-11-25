<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterKaryawan
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.TxtPrefix = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.TxtSuffix = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.TxtLastName = New System.Windows.Forms.TextBox
        Me.Label27 = New System.Windows.Forms.Label
        Me.TxtMiddleName = New System.Windows.Forms.TextBox
        Me.Label28 = New System.Windows.Forms.Label
        Me.TxtFirstName = New System.Windows.Forms.TextBox
        Me.dgvFamily = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHubKeluarga = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColIsDependent = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColEducation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSex = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColGolDarah = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColContactNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBirthPlace = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBirthday = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage6 = New System.Windows.Forms.TabPage
        Me.DGVKompetensi = New System.Windows.Forms.DataGridView
        Me.ColCompetenceNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCompetence_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCempetence_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCompetence_StandarValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCompetence_Value = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCompetence_NextCertificationDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCompetence_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCompetence_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCempetence_DataType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage8 = New System.Windows.Forms.TabPage
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.DGVKontrak = New System.Windows.Forms.DataGridView
        Me.ColContract_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColContract_Type = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColContract_DocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColContract_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColContract_Start = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColContract_End = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColContract_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColContract_PrimaryType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColContract_IdentifierContract = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DGVAttanchmentKontrak = New System.Windows.Forms.DataGridView
        Me.ColKontrakAtt_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKontrakAtt_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKontrakAtt_LocationFile = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKontrakAtt_Preview = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColKontrakAtt_BtnOpenDelete = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColKontrakAtt_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage9 = New System.Windows.Forms.TabPage
        Me.DGVAttanchment = New System.Windows.Forms.DataGridView
        Me.ColAtt_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAtt_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAtt_LocationFile = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAtt_Preview = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColAtt_BtnOpenDelete = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColAtt_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.chkloses_tax_obligations = New System.Windows.Forms.CheckBox
        Me.lblResignDate = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.DtpResignDate = New System.Windows.Forms.DateTimePicker
        Me.dtpAppoinmentDate = New System.Windows.Forms.DateTimePicker
        Me.Label43 = New System.Windows.Forms.Label
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.CboEmployeeStatus = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.BtnFindPosition = New System.Windows.Forms.Button
        Me.cboStatus = New System.Windows.Forms.ComboBox
        Me.TxtPicturePath = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TextBox11 = New System.Windows.Forms.TextBox
        Me.TxtSalaryFormula = New System.Windows.Forms.TextBox
        Me.TxtPrimaryEmployee = New System.Windows.Forms.TextBox
        Me.TxtPrimaryPosition = New System.Windows.Forms.TextBox
        Me.TxtPositionCode = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.DtpAdmision = New System.Windows.Forms.DateTimePicker
        Me.btnFindDepartemen = New System.Windows.Forms.Button
        Me.TxtPrimaryDepartemen = New System.Windows.Forms.TextBox
        Me.TxtDepartemen = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtNIK = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.TxtFullName = New System.Windows.Forms.TextBox
        Me.TxtEmployeeCode = New System.Windows.Forms.TextBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.Label41 = New System.Windows.Forms.Label
        Me.Label40 = New System.Windows.Forms.Label
        Me.Label36 = New System.Windows.Forms.Label
        Me.Label35 = New System.Windows.Forms.Label
        Me.TxtPersonal_Mother_Name = New System.Windows.Forms.TextBox
        Me.Label34 = New System.Windows.Forms.Label
        Me.TxtPersonal_Phone = New System.Windows.Forms.TextBox
        Me.Label33 = New System.Windows.Forms.Label
        Me.TxtPersonal_Email = New System.Windows.Forms.TextBox
        Me.TxtPersonal_Nationality = New System.Windows.Forms.TextBox
        Me.Label32 = New System.Windows.Forms.Label
        Me.CboPersonal_BloodType = New System.Windows.Forms.ComboBox
        Me.Label31 = New System.Windows.Forms.Label
        Me.LblDateOfMarriage = New System.Windows.Forms.Label
        Me.DtpPersonal_MTdate = New System.Windows.Forms.DateTimePicker
        Me.CboPersonal_MartialStatus = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.TxtPersonal_Education = New System.Windows.Forms.TextBox
        Me.TxtPersonal_BirthPlace = New System.Windows.Forms.TextBox
        Me.Label30 = New System.Windows.Forms.Label
        Me.DtpPersonal_BPDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPersonal_Mother_Address = New System.Windows.Forms.TextBox
        Me.CboPersonal_Sex = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.CboPersonal_Religion = New System.Windows.Forms.ComboBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.Label51 = New System.Windows.Forms.Label
        Me.dtpFirstTaxPeriod = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryCountryCode = New System.Windows.Forms.TextBox
        Me.Label42 = New System.Windows.Forms.Label
        Me.CboCountryDesc = New System.Windows.Forms.ComboBox
        Me.CboForeignEmployee = New System.Windows.Forms.ComboBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.ChkNoNPWP = New System.Windows.Forms.CheckBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.TxtMedicalCenter = New System.Windows.Forms.TextBox
        Me.TxtLegal_KPJ_No = New System.Windows.Forms.TextBox
        Me.TxtLegal_ID_Card_No = New System.Windows.Forms.TextBox
        Me.Label29 = New System.Windows.Forms.Label
        Me.TxtLegal_ID_Address = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.TxtLegal_Astek_No = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.DtpLegal_Expired = New System.Windows.Forms.DateTimePicker
        Me.TxtLegal_NPWP_Address = New System.Windows.Forms.TextBox
        Me.TxtLegal_Current_Address = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtPersonal_NPWP = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.CboLegal_Category = New System.Windows.Forms.ComboBox
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.Label45 = New System.Windows.Forms.Label
        Me.TxtBank_Code = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label37 = New System.Windows.Forms.Label
        Me.TxtBank_Name = New System.Windows.Forms.TextBox
        Me.Label38 = New System.Windows.Forms.Label
        Me.TxtBank_Acc_No = New System.Windows.Forms.TextBox
        Me.Label39 = New System.Windows.Forms.Label
        Me.TxtBank_Acc_Holder = New System.Windows.Forms.TextBox
        Me.TxtBank_Branch = New System.Windows.Forms.TextBox
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.TxtNotes = New System.Windows.Forms.TextBox
        Me.TabPage7 = New System.Windows.Forms.TabPage
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.ChkOverrideOTOrder = New System.Windows.Forms.CheckBox
        Me.chkSendNotifEmailOvertimeOrder = New System.Windows.Forms.CheckBox
        Me.chkUseOvertimeOrder = New System.Windows.Forms.CheckBox
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.chkOverrideCalculateOTTolerance = New System.Windows.Forms.CheckBox
        Me.chkOverrideCalculateEarlyOutTolerance = New System.Windows.Forms.CheckBox
        Me.chkOverrideCalculateLateTolerance = New System.Windows.Forms.CheckBox
        Me.chkOverrideCalculateEarlyOTTolerance = New System.Windows.Forms.CheckBox
        Me.Label50 = New System.Windows.Forms.Label
        Me.txtMaxValueOT = New System.Windows.Forms.TextBox
        Me.cboRoundingTypeOT = New System.Windows.Forms.ComboBox
        Me.txtRoundingOT = New System.Windows.Forms.TextBox
        Me.txtCalculateAfterOT = New System.Windows.Forms.TextBox
        Me.txtMaxValueEarlyOut = New System.Windows.Forms.TextBox
        Me.cboRoundingTypeEarlyOut = New System.Windows.Forms.ComboBox
        Me.txtRoundingEarlyOut = New System.Windows.Forms.TextBox
        Me.txtCalculateAfterEarlyOut = New System.Windows.Forms.TextBox
        Me.txtMaxValueLate = New System.Windows.Forms.TextBox
        Me.cboRoundingTypeLate = New System.Windows.Forms.ComboBox
        Me.txtRoundingLate = New System.Windows.Forms.TextBox
        Me.txtCalculateAfterLate = New System.Windows.Forms.TextBox
        Me.Label46 = New System.Windows.Forms.Label
        Me.Label47 = New System.Windows.Forms.Label
        Me.Label48 = New System.Windows.Forms.Label
        Me.txtMaxValueEarlyOT = New System.Windows.Forms.TextBox
        Me.cboRoundingTypeEarlyOT = New System.Windows.Forms.ComboBox
        Me.txtRoundingEarlyOT = New System.Windows.Forms.TextBox
        Me.txtCalculateAfterEarlyOT = New System.Windows.Forms.TextBox
        Me.Label49 = New System.Windows.Forms.Label
        Me.CboIsCalculateOT = New System.Windows.Forms.ComboBox
        Me.CboIsCalculateEarlyOut = New System.Windows.Forms.ComboBox
        Me.ChkOverrideOT = New System.Windows.Forms.CheckBox
        Me.ChkOverrideEarlyOut = New System.Windows.Forms.CheckBox
        Me.CboIsCalculateLate = New System.Windows.Forms.ComboBox
        Me.CboIsCalculateEarlyOT = New System.Windows.Forms.ComboBox
        Me.ChkOverrideLate = New System.Windows.Forms.CheckBox
        Me.ChkOverrideEarlyOT = New System.Windows.Forms.CheckBox
        Me.TxtFingerprintID = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ChkDoesnNotClockOut = New System.Windows.Forms.CheckBox
        Me.ChkOverrideClockOut = New System.Windows.Forms.CheckBox
        Me.ChkDoesnNotClockIn = New System.Windows.Forms.CheckBox
        Me.ChkOverrideClockIn = New System.Windows.Forms.CheckBox
        Me.Label44 = New System.Windows.Forms.Label
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnViewWorkShift = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStripStatusLabel4 = New System.Windows.Forms.ToolStripStatusLabel
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.BgwDeleteRowDgv = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.OpenFileDialog2 = New System.Windows.Forms.OpenFileDialog
        Me.llbCopy = New System.Windows.Forms.LinkLabel
        Me.gboDetail.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvFamily, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        CType(Me.DGVKompetensi, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage8.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.DGVKontrak, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGVAttanchmentKontrak, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage9.SuspendLayout()
        CType(Me.DGVAttanchment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.TabPage7.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(505, 412)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(675, 452)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 6
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(417, 452)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(245, 452)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TabControl2)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Location = New System.Drawing.Point(5, 2)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(994, 444)
        Me.gboDetail.TabIndex = 7
        Me.gboDetail.TabStop = False
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(409, 412)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 26)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Controls.Add(Me.TabPage6)
        Me.TabControl2.Controls.Add(Me.TabPage8)
        Me.TabControl2.Controls.Add(Me.TabPage9)
        Me.TabControl2.Location = New System.Drawing.Point(3, 281)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(985, 125)
        Me.TabControl2.TabIndex = 95
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.GroupBox2)
        Me.TabPage5.Controls.Add(Me.dgvFamily)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage5.Size = New System.Drawing.Size(977, 99)
        Me.TabPage5.TabIndex = 0
        Me.TabPage5.Text = "Family Member"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label20)
        Me.GroupBox2.Controls.Add(Me.TxtPrefix)
        Me.GroupBox2.Controls.Add(Me.Label25)
        Me.GroupBox2.Controls.Add(Me.TxtSuffix)
        Me.GroupBox2.Controls.Add(Me.Label26)
        Me.GroupBox2.Controls.Add(Me.TxtLastName)
        Me.GroupBox2.Controls.Add(Me.Label27)
        Me.GroupBox2.Controls.Add(Me.TxtMiddleName)
        Me.GroupBox2.Controls.Add(Me.Label28)
        Me.GroupBox2.Controls.Add(Me.TxtFirstName)
        Me.GroupBox2.Location = New System.Drawing.Point(642, 16)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(200, 174)
        Me.GroupBox2.TabIndex = 96
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Detail Name"
        Me.GroupBox2.Visible = False
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(3, 16)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(33, 13)
        Me.Label20.TabIndex = 50
        Me.Label20.Text = "Prefix"
        '
        'TxtPrefix
        '
        Me.TxtPrefix.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrefix.Location = New System.Drawing.Point(83, 13)
        Me.TxtPrefix.MaxLength = 100
        Me.TxtPrefix.Name = "TxtPrefix"
        Me.TxtPrefix.Size = New System.Drawing.Size(213, 20)
        Me.TxtPrefix.TabIndex = 49
        Me.TxtPrefix.TabStop = False
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(3, 120)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(33, 13)
        Me.Label25.TabIndex = 48
        Me.Label25.Text = "Suffix"
        '
        'TxtSuffix
        '
        Me.TxtSuffix.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSuffix.Location = New System.Drawing.Point(83, 117)
        Me.TxtSuffix.MaxLength = 100
        Me.TxtSuffix.Name = "TxtSuffix"
        Me.TxtSuffix.Size = New System.Drawing.Size(213, 20)
        Me.TxtSuffix.TabIndex = 47
        Me.TxtSuffix.TabStop = False
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(3, 94)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(27, 13)
        Me.Label26.TabIndex = 46
        Me.Label26.Text = "Last"
        '
        'TxtLastName
        '
        Me.TxtLastName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLastName.Location = New System.Drawing.Point(83, 91)
        Me.TxtLastName.MaxLength = 100
        Me.TxtLastName.Name = "TxtLastName"
        Me.TxtLastName.Size = New System.Drawing.Size(213, 20)
        Me.TxtLastName.TabIndex = 45
        Me.TxtLastName.TabStop = False
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(3, 68)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(38, 13)
        Me.Label27.TabIndex = 44
        Me.Label27.Text = "Middle"
        '
        'TxtMiddleName
        '
        Me.TxtMiddleName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMiddleName.Location = New System.Drawing.Point(83, 65)
        Me.TxtMiddleName.MaxLength = 100
        Me.TxtMiddleName.Name = "TxtMiddleName"
        Me.TxtMiddleName.Size = New System.Drawing.Size(213, 20)
        Me.TxtMiddleName.TabIndex = 43
        Me.TxtMiddleName.TabStop = False
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(3, 42)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(26, 13)
        Me.Label28.TabIndex = 42
        Me.Label28.Text = "First"
        '
        'TxtFirstName
        '
        Me.TxtFirstName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFirstName.Location = New System.Drawing.Point(83, 39)
        Me.TxtFirstName.MaxLength = 100
        Me.TxtFirstName.Name = "TxtFirstName"
        Me.TxtFirstName.Size = New System.Drawing.Size(213, 20)
        Me.TxtFirstName.TabIndex = 41
        Me.TxtFirstName.TabStop = False
        '
        'dgvFamily
        '
        Me.dgvFamily.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvFamily.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColIdentifier, Me.ColName, Me.ColHubKeluarga, Me.ColIsDependent, Me.ColEducation, Me.ColSex, Me.ColGolDarah, Me.ColContactNo, Me.ColBirthPlace, Me.ColBirthday, Me.ColRemarks})
        Me.dgvFamily.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvFamily.Location = New System.Drawing.Point(3, 3)
        Me.dgvFamily.MultiSelect = False
        Me.dgvFamily.Name = "dgvFamily"
        Me.dgvFamily.RowTemplate.Height = 24
        Me.dgvFamily.Size = New System.Drawing.Size(971, 93)
        Me.dgvFamily.TabIndex = 95
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 46
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.Visible = False
        Me.ColIdentifier.Width = 87
        '
        'ColName
        '
        Me.ColName.HeaderText = "Name"
        Me.ColName.MaxInputLength = 250
        Me.ColName.Name = "ColName"
        Me.ColName.Width = 60
        '
        'ColHubKeluarga
        '
        Me.ColHubKeluarga.HeaderText = "Relationship"
        Me.ColHubKeluarga.Items.AddRange(New Object() {"Husband", "Wife", "Son", "Daughter", "Brother", "Sister", "Father", "Mother", "Uncle", "Aunt", "Other"})
        Me.ColHubKeluarga.Name = "ColHubKeluarga"
        Me.ColHubKeluarga.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColHubKeluarga.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColHubKeluarga.Width = 90
        '
        'ColIsDependent
        '
        Me.ColIsDependent.HeaderText = "Dependent"
        Me.ColIsDependent.Name = "ColIsDependent"
        Me.ColIsDependent.Width = 66
        '
        'ColEducation
        '
        Me.ColEducation.HeaderText = "Education"
        Me.ColEducation.MaxInputLength = 250
        Me.ColEducation.Name = "ColEducation"
        Me.ColEducation.Width = 80
        '
        'ColSex
        '
        Me.ColSex.HeaderText = "Sex"
        Me.ColSex.Items.AddRange(New Object() {"Male", "Female"})
        Me.ColSex.Name = "ColSex"
        Me.ColSex.Width = 39
        '
        'ColGolDarah
        '
        Me.ColGolDarah.HeaderText = "Blood Type"
        Me.ColGolDarah.Items.AddRange(New Object() {"A", "B", "AB", "O"})
        Me.ColGolDarah.Name = "ColGolDarah"
        Me.ColGolDarah.Width = 67
        '
        'ColContactNo
        '
        Me.ColContactNo.HeaderText = "Phone No"
        Me.ColContactNo.MaxInputLength = 250
        Me.ColContactNo.Name = "ColContactNo"
        Me.ColContactNo.Width = 80
        '
        'ColBirthPlace
        '
        Me.ColBirthPlace.HeaderText = "Birth Place"
        Me.ColBirthPlace.MaxInputLength = 250
        Me.ColBirthPlace.Name = "ColBirthPlace"
        Me.ColBirthPlace.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColBirthPlace.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBirthPlace.Width = 64
        '
        'ColBirthday
        '
        DataGridViewCellStyle1.Format = "d"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.ColBirthday.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColBirthday.HeaderText = "Birthday"
        Me.ColBirthday.MaxInputLength = 250
        Me.ColBirthday.Name = "ColBirthday"
        Me.ColBirthday.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColBirthday.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBirthday.Width = 51
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.MaxInputLength = 250
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.DGVKompetensi)
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage6.Size = New System.Drawing.Size(977, 100)
        Me.TabPage6.TabIndex = 1
        Me.TabPage6.Text = "Competence"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'DGVKompetensi
        '
        Me.DGVKompetensi.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVKompetensi.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCompetenceNo, Me.ColCompetence_Code, Me.ColCempetence_Description, Me.ColCompetence_StandarValue, Me.ColCompetence_Value, Me.ColCompetence_NextCertificationDate, Me.ColCompetence_Remarks, Me.ColCompetence_Primary, Me.ColCempetence_DataType})
        Me.DGVKompetensi.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVKompetensi.Location = New System.Drawing.Point(3, 3)
        Me.DGVKompetensi.MultiSelect = False
        Me.DGVKompetensi.Name = "DGVKompetensi"
        Me.DGVKompetensi.RowTemplate.Height = 24
        Me.DGVKompetensi.Size = New System.Drawing.Size(971, 94)
        Me.DGVKompetensi.TabIndex = 96
        '
        'ColCompetenceNo
        '
        Me.ColCompetenceNo.HeaderText = "No"
        Me.ColCompetenceNo.Name = "ColCompetenceNo"
        Me.ColCompetenceNo.Width = 46
        '
        'ColCompetence_Code
        '
        Me.ColCompetence_Code.HeaderText = "Code"
        Me.ColCompetence_Code.Name = "ColCompetence_Code"
        Me.ColCompetence_Code.Width = 57
        '
        'ColCempetence_Description
        '
        Me.ColCempetence_Description.HeaderText = "Description"
        Me.ColCempetence_Description.Name = "ColCempetence_Description"
        Me.ColCempetence_Description.Width = 85
        '
        'ColCompetence_StandarValue
        '
        Me.ColCompetence_StandarValue.HeaderText = "Standar Value"
        Me.ColCompetence_StandarValue.Name = "ColCompetence_StandarValue"
        Me.ColCompetence_StandarValue.Width = 99
        '
        'ColCompetence_Value
        '
        Me.ColCompetence_Value.HeaderText = "Value"
        Me.ColCompetence_Value.Name = "ColCompetence_Value"
        Me.ColCompetence_Value.Width = 59
        '
        'ColCompetence_NextCertificationDate
        '
        Me.ColCompetence_NextCertificationDate.HeaderText = "Next Certification Date"
        Me.ColCompetence_NextCertificationDate.Name = "ColCompetence_NextCertificationDate"
        Me.ColCompetence_NextCertificationDate.Width = 138
        '
        'ColCompetence_Remarks
        '
        Me.ColCompetence_Remarks.HeaderText = "Remarks"
        Me.ColCompetence_Remarks.Name = "ColCompetence_Remarks"
        Me.ColCompetence_Remarks.Width = 74
        '
        'ColCompetence_Primary
        '
        Me.ColCompetence_Primary.HeaderText = "Primary"
        Me.ColCompetence_Primary.Name = "ColCompetence_Primary"
        Me.ColCompetence_Primary.Visible = False
        Me.ColCompetence_Primary.Width = 81
        '
        'ColCempetence_DataType
        '
        Me.ColCempetence_DataType.HeaderText = "DataType"
        Me.ColCempetence_DataType.Name = "ColCempetence_DataType"
        Me.ColCempetence_DataType.Visible = False
        Me.ColCempetence_DataType.Width = 95
        '
        'TabPage8
        '
        Me.TabPage8.Controls.Add(Me.SplitContainer2)
        Me.TabPage8.Location = New System.Drawing.Point(4, 22)
        Me.TabPage8.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage8.Size = New System.Drawing.Size(977, 100)
        Me.TabPage8.TabIndex = 2
        Me.TabPage8.Text = "Work Contract"
        Me.TabPage8.UseVisualStyleBackColor = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(2, 2)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.DGVKontrak)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.DGVAttanchmentKontrak)
        Me.SplitContainer2.Size = New System.Drawing.Size(973, 96)
        Me.SplitContainer2.SplitterDistance = 482
        Me.SplitContainer2.SplitterWidth = 3
        Me.SplitContainer2.TabIndex = 1
        '
        'DGVKontrak
        '
        Me.DGVKontrak.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVKontrak.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColContract_No, Me.ColContract_Type, Me.ColContract_DocNo, Me.ColContract_Date, Me.ColContract_Start, Me.ColContract_End, Me.ColContract_Remarks, Me.ColContract_PrimaryType, Me.ColContract_IdentifierContract})
        Me.DGVKontrak.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVKontrak.Location = New System.Drawing.Point(0, 0)
        Me.DGVKontrak.MultiSelect = False
        Me.DGVKontrak.Name = "DGVKontrak"
        Me.DGVKontrak.RowTemplate.Height = 24
        Me.DGVKontrak.Size = New System.Drawing.Size(482, 96)
        Me.DGVKontrak.TabIndex = 97
        '
        'ColContract_No
        '
        Me.ColContract_No.HeaderText = "No"
        Me.ColContract_No.Name = "ColContract_No"
        Me.ColContract_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColContract_No.Width = 27
        '
        'ColContract_Type
        '
        Me.ColContract_Type.HeaderText = "Type"
        Me.ColContract_Type.Name = "ColContract_Type"
        Me.ColContract_Type.Width = 39
        '
        'ColContract_DocNo
        '
        Me.ColContract_DocNo.HeaderText = "Document No"
        Me.ColContract_DocNo.Name = "ColContract_DocNo"
        Me.ColContract_DocNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColContract_DocNo.Width = 79
        '
        'ColContract_Date
        '
        DataGridViewCellStyle2.Format = "dd MMM yyyy"
        Me.ColContract_Date.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColContract_Date.HeaderText = "Date"
        Me.ColContract_Date.Name = "ColContract_Date"
        Me.ColContract_Date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColContract_Date.Width = 36
        '
        'ColContract_Start
        '
        DataGridViewCellStyle3.Format = "dd MMM yyyy"
        Me.ColContract_Start.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColContract_Start.HeaderText = "Start"
        Me.ColContract_Start.Name = "ColContract_Start"
        Me.ColContract_Start.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColContract_Start.Width = 35
        '
        'ColContract_End
        '
        DataGridViewCellStyle4.Format = "dd MMM yyyy"
        Me.ColContract_End.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColContract_End.HeaderText = "End"
        Me.ColContract_End.Name = "ColContract_End"
        Me.ColContract_End.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColContract_End.Width = 32
        '
        'ColContract_Remarks
        '
        Me.ColContract_Remarks.HeaderText = "Remarks"
        Me.ColContract_Remarks.Name = "ColContract_Remarks"
        Me.ColContract_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColContract_Remarks.Width = 55
        '
        'ColContract_PrimaryType
        '
        Me.ColContract_PrimaryType.HeaderText = "Primary Type"
        Me.ColContract_PrimaryType.Name = "ColContract_PrimaryType"
        Me.ColContract_PrimaryType.Visible = False
        Me.ColContract_PrimaryType.Width = 93
        '
        'ColContract_IdentifierContract
        '
        Me.ColContract_IdentifierContract.HeaderText = "Identifier Contract"
        Me.ColContract_IdentifierContract.Name = "ColContract_IdentifierContract"
        Me.ColContract_IdentifierContract.Visible = False
        Me.ColContract_IdentifierContract.Width = 115
        '
        'DGVAttanchmentKontrak
        '
        Me.DGVAttanchmentKontrak.AllowUserToResizeColumns = False
        Me.DGVAttanchmentKontrak.AllowUserToResizeRows = False
        Me.DGVAttanchmentKontrak.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVAttanchmentKontrak.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DGVAttanchmentKontrak.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColKontrakAtt_No, Me.ColKontrakAtt_Description, Me.ColKontrakAtt_LocationFile, Me.ColKontrakAtt_Preview, Me.ColKontrakAtt_BtnOpenDelete, Me.ColKontrakAtt_Remarks})
        Me.DGVAttanchmentKontrak.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVAttanchmentKontrak.Location = New System.Drawing.Point(0, 0)
        Me.DGVAttanchmentKontrak.MultiSelect = False
        Me.DGVAttanchmentKontrak.Name = "DGVAttanchmentKontrak"
        Me.DGVAttanchmentKontrak.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DGVAttanchmentKontrak.RowTemplate.Height = 24
        Me.DGVAttanchmentKontrak.Size = New System.Drawing.Size(488, 96)
        Me.DGVAttanchmentKontrak.TabIndex = 2
        '
        'ColKontrakAtt_No
        '
        Me.ColKontrakAtt_No.HeaderText = "No"
        Me.ColKontrakAtt_No.Name = "ColKontrakAtt_No"
        Me.ColKontrakAtt_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColKontrakAtt_No.Width = 27
        '
        'ColKontrakAtt_Description
        '
        Me.ColKontrakAtt_Description.HeaderText = "Description"
        Me.ColKontrakAtt_Description.Name = "ColKontrakAtt_Description"
        Me.ColKontrakAtt_Description.Width = 85
        '
        'ColKontrakAtt_LocationFile
        '
        Me.ColKontrakAtt_LocationFile.HeaderText = "File Path"
        Me.ColKontrakAtt_LocationFile.Name = "ColKontrakAtt_LocationFile"
        Me.ColKontrakAtt_LocationFile.Width = 73
        '
        'ColKontrakAtt_Preview
        '
        Me.ColKontrakAtt_Preview.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle5.NullValue = "-"
        Me.ColKontrakAtt_Preview.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColKontrakAtt_Preview.HeaderText = ""
        Me.ColKontrakAtt_Preview.Name = "ColKontrakAtt_Preview"
        Me.ColKontrakAtt_Preview.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColKontrakAtt_Preview.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColKontrakAtt_Preview.Width = 21
        '
        'ColKontrakAtt_BtnOpenDelete
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle6.NullValue = "add"
        Me.ColKontrakAtt_BtnOpenDelete.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColKontrakAtt_BtnOpenDelete.HeaderText = ""
        Me.ColKontrakAtt_BtnOpenDelete.Name = "ColKontrakAtt_BtnOpenDelete"
        Me.ColKontrakAtt_BtnOpenDelete.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColKontrakAtt_BtnOpenDelete.Text = ""
        Me.ColKontrakAtt_BtnOpenDelete.Width = 36
        '
        'ColKontrakAtt_Remarks
        '
        Me.ColKontrakAtt_Remarks.HeaderText = "Remarks"
        Me.ColKontrakAtt_Remarks.Name = "ColKontrakAtt_Remarks"
        Me.ColKontrakAtt_Remarks.Width = 74
        '
        'TabPage9
        '
        Me.TabPage9.Controls.Add(Me.DGVAttanchment)
        Me.TabPage9.Location = New System.Drawing.Point(4, 22)
        Me.TabPage9.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage9.Name = "TabPage9"
        Me.TabPage9.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage9.Size = New System.Drawing.Size(977, 100)
        Me.TabPage9.TabIndex = 3
        Me.TabPage9.Text = "Attachment"
        Me.TabPage9.UseVisualStyleBackColor = True
        '
        'DGVAttanchment
        '
        Me.DGVAttanchment.AllowUserToResizeColumns = False
        Me.DGVAttanchment.AllowUserToResizeRows = False
        Me.DGVAttanchment.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVAttanchment.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DGVAttanchment.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAtt_No, Me.ColAtt_Description, Me.ColAtt_LocationFile, Me.ColAtt_Preview, Me.ColAtt_BtnOpenDelete, Me.ColAtt_Remarks})
        Me.DGVAttanchment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVAttanchment.Location = New System.Drawing.Point(2, 2)
        Me.DGVAttanchment.MultiSelect = False
        Me.DGVAttanchment.Name = "DGVAttanchment"
        Me.DGVAttanchment.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DGVAttanchment.RowTemplate.Height = 24
        Me.DGVAttanchment.Size = New System.Drawing.Size(973, 96)
        Me.DGVAttanchment.TabIndex = 4
        '
        'ColAtt_No
        '
        Me.ColAtt_No.HeaderText = "No"
        Me.ColAtt_No.Name = "ColAtt_No"
        Me.ColAtt_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAtt_No.Width = 27
        '
        'ColAtt_Description
        '
        Me.ColAtt_Description.HeaderText = "Description"
        Me.ColAtt_Description.Name = "ColAtt_Description"
        Me.ColAtt_Description.Width = 85
        '
        'ColAtt_LocationFile
        '
        Me.ColAtt_LocationFile.HeaderText = "File Path"
        Me.ColAtt_LocationFile.Name = "ColAtt_LocationFile"
        Me.ColAtt_LocationFile.Width = 73
        '
        'ColAtt_Preview
        '
        Me.ColAtt_Preview.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle7.NullValue = "-"
        Me.ColAtt_Preview.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColAtt_Preview.HeaderText = ""
        Me.ColAtt_Preview.Name = "ColAtt_Preview"
        Me.ColAtt_Preview.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColAtt_Preview.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColAtt_Preview.Width = 21
        '
        'ColAtt_BtnOpenDelete
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle8.NullValue = "add"
        Me.ColAtt_BtnOpenDelete.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColAtt_BtnOpenDelete.HeaderText = ""
        Me.ColAtt_BtnOpenDelete.Name = "ColAtt_BtnOpenDelete"
        Me.ColAtt_BtnOpenDelete.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColAtt_BtnOpenDelete.Text = ""
        Me.ColAtt_BtnOpenDelete.Width = 36
        '
        'ColAtt_Remarks
        '
        Me.ColAtt_Remarks.HeaderText = "Remarks"
        Me.ColAtt_Remarks.Name = "ColAtt_Remarks"
        Me.ColAtt_Remarks.Width = 74
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(9, 10)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkloses_tax_obligations)
        Me.SplitContainer1.Panel1.Controls.Add(Me.lblResignDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label21)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpResignDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpAppoinmentDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label43)
        Me.SplitContainer1.Panel1.Controls.Add(Me.PictureBox3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.PictureBox2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label24)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboEmployeeStatus)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LinkLabel2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindPosition)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboStatus)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPicturePath)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TextBox11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtSalaryFormula)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryEmployee)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryPosition)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPositionCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label19)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpAdmision)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindDepartemen)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryDepartemen)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDepartemen)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtNIK)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtReadOnly)
        Me.SplitContainer1.Panel1.Controls.Add(Me.PictureBox1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtFullName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtEmployeeCode)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TabControl1)
        Me.SplitContainer1.Size = New System.Drawing.Size(982, 265)
        Me.SplitContainer1.SplitterDistance = 409
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'chkloses_tax_obligations
        '
        Me.chkloses_tax_obligations.AutoSize = True
        Me.chkloses_tax_obligations.Location = New System.Drawing.Point(237, 239)
        Me.chkloses_tax_obligations.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkloses_tax_obligations.Name = "chkloses_tax_obligations"
        Me.chkloses_tax_obligations.Size = New System.Drawing.Size(175, 17)
        Me.chkloses_tax_obligations.TabIndex = 106
        Me.chkloses_tax_obligations.Text = "Loses subjective tax obligations"
        Me.chkloses_tax_obligations.UseVisualStyleBackColor = True
        '
        'lblResignDate
        '
        Me.lblResignDate.AutoSize = True
        Me.lblResignDate.Location = New System.Drawing.Point(145, 218)
        Me.lblResignDate.Name = "lblResignDate"
        Me.lblResignDate.Size = New System.Drawing.Size(66, 13)
        Me.lblResignDate.TabIndex = 105
        Me.lblResignDate.Text = "Resign Date"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(145, 147)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(76, 13)
        Me.Label21.TabIndex = 79
        Me.Label21.Text = "Salary Formula"
        '
        'DtpResignDate
        '
        Me.DtpResignDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpResignDate.CustomFormat = "dd MMM yyyy"
        Me.DtpResignDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpResignDate.Location = New System.Drawing.Point(237, 215)
        Me.DtpResignDate.Name = "DtpResignDate"
        Me.DtpResignDate.Size = New System.Drawing.Size(167, 20)
        Me.DtpResignDate.TabIndex = 13
        '
        'dtpAppoinmentDate
        '
        Me.dtpAppoinmentDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpAppoinmentDate.CustomFormat = "dd MMM yyyy"
        Me.dtpAppoinmentDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpAppoinmentDate.Location = New System.Drawing.Point(237, 73)
        Me.dtpAppoinmentDate.Name = "dtpAppoinmentDate"
        Me.dtpAppoinmentDate.Size = New System.Drawing.Size(167, 20)
        Me.dtpAppoinmentDate.TabIndex = 5
        '
        'Label43
        '
        Me.Label43.AutoSize = True
        Me.Label43.Location = New System.Drawing.Point(145, 76)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(92, 13)
        Me.Label43.TabIndex = 103
        Me.Label43.Text = "Appointment Date"
        '
        'PictureBox3
        '
        Me.PictureBox3.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox3.Location = New System.Drawing.Point(77, 188)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(27, 24)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox3.TabIndex = 102
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox2.Location = New System.Drawing.Point(44, 188)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(27, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox2.TabIndex = 101
        Me.PictureBox2.TabStop = False
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(145, 170)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(86, 13)
        Me.Label24.TabIndex = 98
        Me.Label24.Text = "Employee Status"
        '
        'CboEmployeeStatus
        '
        Me.CboEmployeeStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboEmployeeStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboEmployeeStatus.FormattingEnabled = True
        Me.CboEmployeeStatus.Items.AddRange(New Object() {"Probation", "Contract", "Jobholder"})
        Me.CboEmployeeStatus.Location = New System.Drawing.Point(237, 167)
        Me.CboEmployeeStatus.Name = "CboEmployeeStatus"
        Me.CboEmployeeStatus.Size = New System.Drawing.Size(167, 21)
        Me.CboEmployeeStatus.TabIndex = 11
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(145, 194)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(37, 13)
        Me.Label9.TabIndex = 55
        Me.Label9.Text = "Status"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(379, 6)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(25, 13)
        Me.LinkLabel2.TabIndex = 2
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Edit"
        '
        'BtnFindPosition
        '
        Me.BtnFindPosition.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindPosition.Location = New System.Drawing.Point(367, 119)
        Me.BtnFindPosition.Name = "BtnFindPosition"
        Me.BtnFindPosition.Size = New System.Drawing.Size(35, 20)
        Me.BtnFindPosition.TabIndex = 9
        Me.BtnFindPosition.TabStop = False
        Me.BtnFindPosition.Text = "Find"
        Me.BtnFindPosition.UseVisualStyleBackColor = True
        '
        'cboStatus
        '
        Me.cboStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStatus.FormattingEnabled = True
        Me.cboStatus.Items.AddRange(New Object() {"Active", "Not Active"})
        Me.cboStatus.Location = New System.Drawing.Point(237, 191)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Size = New System.Drawing.Size(167, 21)
        Me.cboStatus.TabIndex = 12
        '
        'TxtPicturePath
        '
        Me.TxtPicturePath.Location = New System.Drawing.Point(18, 85)
        Me.TxtPicturePath.Name = "TxtPicturePath"
        Me.TxtPicturePath.Size = New System.Drawing.Size(101, 20)
        Me.TxtPicturePath.TabIndex = 91
        Me.TxtPicturePath.TabStop = False
        Me.TxtPicturePath.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(145, 99)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(62, 13)
        Me.Label4.TabIndex = 56
        Me.Label4.Text = "Department"
        '
        'TextBox11
        '
        Me.TextBox11.Location = New System.Drawing.Point(206, 154)
        Me.TextBox11.Name = "TextBox11"
        Me.TextBox11.Size = New System.Drawing.Size(25, 20)
        Me.TextBox11.TabIndex = 80
        Me.TextBox11.TabStop = False
        Me.TextBox11.Text = "-1"
        Me.TextBox11.Visible = False
        '
        'TxtSalaryFormula
        '
        Me.TxtSalaryFormula.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSalaryFormula.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSalaryFormula.Location = New System.Drawing.Point(237, 144)
        Me.TxtSalaryFormula.MaxLength = 150
        Me.TxtSalaryFormula.Name = "TxtSalaryFormula"
        Me.TxtSalaryFormula.ReadOnly = True
        Me.TxtSalaryFormula.Size = New System.Drawing.Size(167, 20)
        Me.TxtSalaryFormula.TabIndex = 10
        Me.TxtSalaryFormula.TabStop = False
        '
        'TxtPrimaryEmployee
        '
        Me.TxtPrimaryEmployee.Location = New System.Drawing.Point(183, 7)
        Me.TxtPrimaryEmployee.Name = "TxtPrimaryEmployee"
        Me.TxtPrimaryEmployee.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryEmployee.TabIndex = 0
        Me.TxtPrimaryEmployee.TabStop = False
        Me.TxtPrimaryEmployee.Text = "-1"
        Me.TxtPrimaryEmployee.Visible = False
        '
        'TxtPrimaryPosition
        '
        Me.TxtPrimaryPosition.Location = New System.Drawing.Point(206, 128)
        Me.TxtPrimaryPosition.Name = "TxtPrimaryPosition"
        Me.TxtPrimaryPosition.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryPosition.TabIndex = 83
        Me.TxtPrimaryPosition.TabStop = False
        Me.TxtPrimaryPosition.Text = "-1"
        Me.TxtPrimaryPosition.Visible = False
        '
        'TxtPositionCode
        '
        Me.TxtPositionCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPositionCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPositionCode.Location = New System.Drawing.Point(237, 120)
        Me.TxtPositionCode.MaxLength = 150
        Me.TxtPositionCode.Name = "TxtPositionCode"
        Me.TxtPositionCode.Size = New System.Drawing.Size(125, 20)
        Me.TxtPositionCode.TabIndex = 8
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(145, 123)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(44, 13)
        Me.Label19.TabIndex = 82
        Me.Label19.Text = "Position"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(145, 53)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(75, 13)
        Me.Label11.TabIndex = 67
        Me.Label11.Text = "Admision Date"
        '
        'DtpAdmision
        '
        Me.DtpAdmision.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpAdmision.CustomFormat = "dd MMM yyyy"
        Me.DtpAdmision.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpAdmision.Location = New System.Drawing.Point(237, 50)
        Me.DtpAdmision.Name = "DtpAdmision"
        Me.DtpAdmision.Size = New System.Drawing.Size(167, 20)
        Me.DtpAdmision.TabIndex = 4
        '
        'btnFindDepartemen
        '
        Me.btnFindDepartemen.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindDepartemen.Location = New System.Drawing.Point(367, 95)
        Me.btnFindDepartemen.Name = "btnFindDepartemen"
        Me.btnFindDepartemen.Size = New System.Drawing.Size(35, 20)
        Me.btnFindDepartemen.TabIndex = 7
        Me.btnFindDepartemen.TabStop = False
        Me.btnFindDepartemen.Text = "Find"
        Me.btnFindDepartemen.UseVisualStyleBackColor = True
        '
        'TxtPrimaryDepartemen
        '
        Me.TxtPrimaryDepartemen.Location = New System.Drawing.Point(206, 102)
        Me.TxtPrimaryDepartemen.Name = "TxtPrimaryDepartemen"
        Me.TxtPrimaryDepartemen.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryDepartemen.TabIndex = 57
        Me.TxtPrimaryDepartemen.TabStop = False
        Me.TxtPrimaryDepartemen.Text = "-1"
        Me.TxtPrimaryDepartemen.Visible = False
        '
        'TxtDepartemen
        '
        Me.TxtDepartemen.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDepartemen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDepartemen.Location = New System.Drawing.Point(237, 96)
        Me.TxtDepartemen.MaxLength = 150
        Me.TxtDepartemen.Name = "TxtDepartemen"
        Me.TxtDepartemen.Size = New System.Drawing.Size(125, 20)
        Me.TxtDepartemen.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(145, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Code"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(145, 30)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(25, 13)
        Me.Label3.TabIndex = 53
        Me.Label3.Text = "NIK"
        '
        'TxtNIK
        '
        Me.TxtNIK.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNIK.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNIK.Location = New System.Drawing.Point(237, 27)
        Me.TxtNIK.MaxLength = 150
        Me.TxtNIK.Name = "TxtNIK"
        Me.TxtNIK.Size = New System.Drawing.Size(167, 20)
        Me.TxtNIK.TabIndex = 3
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(180, 30)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(39, 20)
        Me.TxtReadOnly.TabIndex = 51
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Location = New System.Drawing.Point(3, 28)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(137, 152)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 34
        Me.PictureBox1.TabStop = False
        '
        'TxtFullName
        '
        Me.TxtFullName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFullName.Location = New System.Drawing.Point(237, 3)
        Me.TxtFullName.MaxLength = 100
        Me.TxtFullName.Name = "TxtFullName"
        Me.TxtFullName.ReadOnly = True
        Me.TxtFullName.Size = New System.Drawing.Size(137, 20)
        Me.TxtFullName.TabIndex = 1
        Me.TxtFullName.TabStop = False
        '
        'TxtEmployeeCode
        '
        Me.TxtEmployeeCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtEmployeeCode.Location = New System.Drawing.Point(44, 3)
        Me.TxtEmployeeCode.MaxLength = 150
        Me.TxtEmployeeCode.Name = "TxtEmployeeCode"
        Me.TxtEmployeeCode.Size = New System.Drawing.Size(96, 20)
        Me.TxtEmployeeCode.TabIndex = 0
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage7)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(569, 265)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Label41)
        Me.TabPage1.Controls.Add(Me.Label40)
        Me.TabPage1.Controls.Add(Me.Label36)
        Me.TabPage1.Controls.Add(Me.Label35)
        Me.TabPage1.Controls.Add(Me.TxtPersonal_Mother_Name)
        Me.TabPage1.Controls.Add(Me.Label34)
        Me.TabPage1.Controls.Add(Me.TxtPersonal_Phone)
        Me.TabPage1.Controls.Add(Me.Label33)
        Me.TabPage1.Controls.Add(Me.TxtPersonal_Email)
        Me.TabPage1.Controls.Add(Me.TxtPersonal_Nationality)
        Me.TabPage1.Controls.Add(Me.Label32)
        Me.TabPage1.Controls.Add(Me.CboPersonal_BloodType)
        Me.TabPage1.Controls.Add(Me.Label31)
        Me.TabPage1.Controls.Add(Me.LblDateOfMarriage)
        Me.TabPage1.Controls.Add(Me.DtpPersonal_MTdate)
        Me.TabPage1.Controls.Add(Me.CboPersonal_MartialStatus)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.Label15)
        Me.TabPage1.Controls.Add(Me.TxtPersonal_Education)
        Me.TabPage1.Controls.Add(Me.TxtPersonal_BirthPlace)
        Me.TabPage1.Controls.Add(Me.Label30)
        Me.TabPage1.Controls.Add(Me.DtpPersonal_BPDate)
        Me.TabPage1.Controls.Add(Me.TxtPersonal_Mother_Address)
        Me.TabPage1.Controls.Add(Me.CboPersonal_Sex)
        Me.TabPage1.Controls.Add(Me.Label10)
        Me.TabPage1.Controls.Add(Me.Label12)
        Me.TabPage1.Controls.Add(Me.CboPersonal_Religion)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage1.Size = New System.Drawing.Size(561, 239)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Personal"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Label41
        '
        Me.Label41.AutoSize = True
        Me.Label41.Location = New System.Drawing.Point(235, 88)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(30, 13)
        Me.Label41.TabIndex = 121
        Me.Label41.Text = "Date"
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Location = New System.Drawing.Point(235, 10)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(30, 13)
        Me.Label40.TabIndex = 120
        Me.Label40.Text = "Date"
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(6, 192)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(88, 13)
        Me.Label36.TabIndex = 119
        Me.Label36.Text = "Mother's Address"
        Me.Label36.Visible = False
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(6, 166)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(78, 13)
        Me.Label35.TabIndex = 118
        Me.Label35.Text = "Mother's Name"
        Me.Label35.Visible = False
        '
        'TxtPersonal_Mother_Name
        '
        Me.TxtPersonal_Mother_Name.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPersonal_Mother_Name.Location = New System.Drawing.Point(100, 163)
        Me.TxtPersonal_Mother_Name.MaxLength = 150
        Me.TxtPersonal_Mother_Name.Name = "TxtPersonal_Mother_Name"
        Me.TxtPersonal_Mother_Name.Size = New System.Drawing.Size(429, 20)
        Me.TxtPersonal_Mother_Name.TabIndex = 11
        Me.TxtPersonal_Mother_Name.TabStop = False
        Me.TxtPersonal_Mother_Name.Visible = False
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(6, 141)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(32, 13)
        Me.Label34.TabIndex = 116
        Me.Label34.Text = "Email"
        '
        'TxtPersonal_Phone
        '
        Me.TxtPersonal_Phone.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPersonal_Phone.Location = New System.Drawing.Point(297, 138)
        Me.TxtPersonal_Phone.MaxLength = 150
        Me.TxtPersonal_Phone.Name = "TxtPersonal_Phone"
        Me.TxtPersonal_Phone.Size = New System.Drawing.Size(231, 20)
        Me.TxtPersonal_Phone.TabIndex = 10
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(235, 141)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(38, 13)
        Me.Label33.TabIndex = 114
        Me.Label33.Text = "Phone"
        '
        'TxtPersonal_Email
        '
        Me.TxtPersonal_Email.Location = New System.Drawing.Point(100, 138)
        Me.TxtPersonal_Email.MaxLength = 150
        Me.TxtPersonal_Email.Name = "TxtPersonal_Email"
        Me.TxtPersonal_Email.Size = New System.Drawing.Size(129, 20)
        Me.TxtPersonal_Email.TabIndex = 9
        '
        'TxtPersonal_Nationality
        '
        Me.TxtPersonal_Nationality.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPersonal_Nationality.Location = New System.Drawing.Point(297, 60)
        Me.TxtPersonal_Nationality.MaxLength = 150
        Me.TxtPersonal_Nationality.Name = "TxtPersonal_Nationality"
        Me.TxtPersonal_Nationality.Size = New System.Drawing.Size(231, 20)
        Me.TxtPersonal_Nationality.TabIndex = 5
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(235, 63)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(56, 13)
        Me.Label32.TabIndex = 111
        Me.Label32.Text = "Nationality"
        '
        'CboPersonal_BloodType
        '
        Me.CboPersonal_BloodType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPersonal_BloodType.FormattingEnabled = True
        Me.CboPersonal_BloodType.Items.AddRange(New Object() {"A", "B", "AB", "O"})
        Me.CboPersonal_BloodType.Location = New System.Drawing.Point(100, 60)
        Me.CboPersonal_BloodType.Name = "CboPersonal_BloodType"
        Me.CboPersonal_BloodType.Size = New System.Drawing.Size(129, 21)
        Me.CboPersonal_BloodType.TabIndex = 4
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(6, 64)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(61, 13)
        Me.Label31.TabIndex = 110
        Me.Label31.Text = "Blood Type"
        '
        'LblDateOfMarriage
        '
        Me.LblDateOfMarriage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblDateOfMarriage.AutoSize = True
        Me.LblDateOfMarriage.Location = New System.Drawing.Point(297, 88)
        Me.LblDateOfMarriage.Name = "LblDateOfMarriage"
        Me.LblDateOfMarriage.Size = New System.Drawing.Size(24, 13)
        Me.LblDateOfMarriage.TabIndex = 108
        Me.LblDateOfMarriage.Text = "n/a"
        '
        'DtpPersonal_MTdate
        '
        Me.DtpPersonal_MTdate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpPersonal_MTdate.CustomFormat = "dd MMM yyyy"
        Me.DtpPersonal_MTdate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpPersonal_MTdate.Location = New System.Drawing.Point(297, 86)
        Me.DtpPersonal_MTdate.Name = "DtpPersonal_MTdate"
        Me.DtpPersonal_MTdate.Size = New System.Drawing.Size(231, 20)
        Me.DtpPersonal_MTdate.TabIndex = 7
        '
        'CboPersonal_MartialStatus
        '
        Me.CboPersonal_MartialStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPersonal_MartialStatus.FormattingEnabled = True
        Me.CboPersonal_MartialStatus.Items.AddRange(New Object() {"Single", "Married"})
        Me.CboPersonal_MartialStatus.Location = New System.Drawing.Point(100, 85)
        Me.CboPersonal_MartialStatus.Name = "CboPersonal_MartialStatus"
        Me.CboPersonal_MartialStatus.Size = New System.Drawing.Size(129, 21)
        Me.CboPersonal_MartialStatus.TabIndex = 6
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 88)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(71, 13)
        Me.Label8.TabIndex = 105
        Me.Label8.Text = "Marital Status"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(6, 115)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(55, 13)
        Me.Label15.TabIndex = 104
        Me.Label15.Text = "Education"
        '
        'TxtPersonal_Education
        '
        Me.TxtPersonal_Education.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPersonal_Education.Location = New System.Drawing.Point(100, 112)
        Me.TxtPersonal_Education.MaxLength = 150
        Me.TxtPersonal_Education.Name = "TxtPersonal_Education"
        Me.TxtPersonal_Education.Size = New System.Drawing.Size(429, 20)
        Me.TxtPersonal_Education.TabIndex = 8
        '
        'TxtPersonal_BirthPlace
        '
        Me.TxtPersonal_BirthPlace.Location = New System.Drawing.Point(100, 6)
        Me.TxtPersonal_BirthPlace.MaxLength = 150
        Me.TxtPersonal_BirthPlace.Name = "TxtPersonal_BirthPlace"
        Me.TxtPersonal_BirthPlace.Size = New System.Drawing.Size(129, 20)
        Me.TxtPersonal_BirthPlace.TabIndex = 0
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(6, 10)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(58, 13)
        Me.Label30.TabIndex = 100
        Me.Label30.Text = "Birth Place"
        '
        'DtpPersonal_BPDate
        '
        Me.DtpPersonal_BPDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpPersonal_BPDate.CustomFormat = "dd MMM yyyy"
        Me.DtpPersonal_BPDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpPersonal_BPDate.Location = New System.Drawing.Point(297, 6)
        Me.DtpPersonal_BPDate.Name = "DtpPersonal_BPDate"
        Me.DtpPersonal_BPDate.Size = New System.Drawing.Size(231, 20)
        Me.DtpPersonal_BPDate.TabIndex = 1
        '
        'TxtPersonal_Mother_Address
        '
        Me.TxtPersonal_Mother_Address.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPersonal_Mother_Address.Location = New System.Drawing.Point(100, 189)
        Me.TxtPersonal_Mother_Address.MaxLength = 150
        Me.TxtPersonal_Mother_Address.Name = "TxtPersonal_Mother_Address"
        Me.TxtPersonal_Mother_Address.Size = New System.Drawing.Size(429, 20)
        Me.TxtPersonal_Mother_Address.TabIndex = 12
        Me.TxtPersonal_Mother_Address.TabStop = False
        Me.TxtPersonal_Mother_Address.Visible = False
        '
        'CboPersonal_Sex
        '
        Me.CboPersonal_Sex.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPersonal_Sex.FormattingEnabled = True
        Me.CboPersonal_Sex.Items.AddRange(New Object() {"Male", "Female"})
        Me.CboPersonal_Sex.Location = New System.Drawing.Point(100, 33)
        Me.CboPersonal_Sex.Name = "CboPersonal_Sex"
        Me.CboPersonal_Sex.Size = New System.Drawing.Size(129, 21)
        Me.CboPersonal_Sex.TabIndex = 2
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 37)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(25, 13)
        Me.Label10.TabIndex = 65
        Me.Label10.Text = "Sex"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(235, 36)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(45, 13)
        Me.Label12.TabIndex = 69
        Me.Label12.Text = "Religion"
        '
        'CboPersonal_Religion
        '
        Me.CboPersonal_Religion.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboPersonal_Religion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPersonal_Religion.FormattingEnabled = True
        Me.CboPersonal_Religion.Items.AddRange(New Object() {"Islam", "Katolik", "Kristen", "Hindu", "Buddha", "Konghucu", "Others"})
        Me.CboPersonal_Religion.Location = New System.Drawing.Point(297, 34)
        Me.CboPersonal_Religion.Name = "CboPersonal_Religion"
        Me.CboPersonal_Religion.Size = New System.Drawing.Size(231, 21)
        Me.CboPersonal_Religion.TabIndex = 3
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label51)
        Me.TabPage2.Controls.Add(Me.dtpFirstTaxPeriod)
        Me.TabPage2.Controls.Add(Me.TxtPrimaryCountryCode)
        Me.TabPage2.Controls.Add(Me.Label42)
        Me.TabPage2.Controls.Add(Me.CboCountryDesc)
        Me.TabPage2.Controls.Add(Me.CboForeignEmployee)
        Me.TabPage2.Controls.Add(Me.Label16)
        Me.TabPage2.Controls.Add(Me.ChkNoNPWP)
        Me.TabPage2.Controls.Add(Me.Label14)
        Me.TabPage2.Controls.Add(Me.TxtMedicalCenter)
        Me.TabPage2.Controls.Add(Me.TxtLegal_KPJ_No)
        Me.TabPage2.Controls.Add(Me.TxtLegal_ID_Card_No)
        Me.TabPage2.Controls.Add(Me.Label29)
        Me.TabPage2.Controls.Add(Me.TxtLegal_ID_Address)
        Me.TabPage2.Controls.Add(Me.Label13)
        Me.TabPage2.Controls.Add(Me.Label18)
        Me.TabPage2.Controls.Add(Me.TxtLegal_Astek_No)
        Me.TabPage2.Controls.Add(Me.Label23)
        Me.TabPage2.Controls.Add(Me.Label22)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.DtpLegal_Expired)
        Me.TabPage2.Controls.Add(Me.TxtLegal_NPWP_Address)
        Me.TabPage2.Controls.Add(Me.TxtLegal_Current_Address)
        Me.TabPage2.Controls.Add(Me.Label5)
        Me.TabPage2.Controls.Add(Me.TxtPersonal_NPWP)
        Me.TabPage2.Controls.Add(Me.Label17)
        Me.TabPage2.Controls.Add(Me.CboLegal_Category)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage2.Size = New System.Drawing.Size(560, 239)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Legal"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label51
        '
        Me.Label51.AutoSize = True
        Me.Label51.Location = New System.Drawing.Point(6, 188)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(80, 13)
        Me.Label51.TabIndex = 109
        Me.Label51.Text = "First Tax Period"
        '
        'dtpFirstTaxPeriod
        '
        Me.dtpFirstTaxPeriod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFirstTaxPeriod.CustomFormat = "MMM yyyy"
        Me.dtpFirstTaxPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFirstTaxPeriod.Location = New System.Drawing.Point(98, 185)
        Me.dtpFirstTaxPeriod.Name = "dtpFirstTaxPeriod"
        Me.dtpFirstTaxPeriod.ShowUpDown = True
        Me.dtpFirstTaxPeriod.Size = New System.Drawing.Size(425, 20)
        Me.dtpFirstTaxPeriod.TabIndex = 108
        '
        'TxtPrimaryCountryCode
        '
        Me.TxtPrimaryCountryCode.Location = New System.Drawing.Point(272, 132)
        Me.TxtPrimaryCountryCode.Name = "TxtPrimaryCountryCode"
        Me.TxtPrimaryCountryCode.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCountryCode.TabIndex = 107
        Me.TxtPrimaryCountryCode.TabStop = False
        Me.TxtPrimaryCountryCode.Text = "-1"
        Me.TxtPrimaryCountryCode.Visible = False
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(217, 137)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(43, 13)
        Me.Label42.TabIndex = 106
        Me.Label42.Text = "Country"
        '
        'CboCountryDesc
        '
        Me.CboCountryDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboCountryDesc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboCountryDesc.FormattingEnabled = True
        Me.CboCountryDesc.Location = New System.Drawing.Point(304, 132)
        Me.CboCountryDesc.Name = "CboCountryDesc"
        Me.CboCountryDesc.Size = New System.Drawing.Size(219, 21)
        Me.CboCountryDesc.TabIndex = 9
        '
        'CboForeignEmployee
        '
        Me.CboForeignEmployee.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboForeignEmployee.FormattingEnabled = True
        Me.CboForeignEmployee.Items.AddRange(New Object() {"No", "Yes"})
        Me.CboForeignEmployee.Location = New System.Drawing.Point(98, 132)
        Me.CboForeignEmployee.Name = "CboForeignEmployee"
        Me.CboForeignEmployee.Size = New System.Drawing.Size(113, 21)
        Me.CboForeignEmployee.TabIndex = 8
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(3, 135)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(90, 13)
        Me.Label16.TabIndex = 103
        Me.Label16.Text = "Foreign employee"
        '
        'ChkNoNPWP
        '
        Me.ChkNoNPWP.AutoSize = True
        Me.ChkNoNPWP.Location = New System.Drawing.Point(9, 86)
        Me.ChkNoNPWP.Name = "ChkNoNPWP"
        Me.ChkNoNPWP.Size = New System.Drawing.Size(76, 17)
        Me.ChkNoNPWP.TabIndex = 4
        Me.ChkNoNPWP.Text = "NPWP No"
        Me.ChkNoNPWP.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 188)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(78, 13)
        Me.Label14.TabIndex = 102
        Me.Label14.Text = "Medical Center"
        '
        'TxtMedicalCenter
        '
        Me.TxtMedicalCenter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMedicalCenter.Location = New System.Drawing.Point(98, 185)
        Me.TxtMedicalCenter.MaxLength = 150
        Me.TxtMedicalCenter.Name = "TxtMedicalCenter"
        Me.TxtMedicalCenter.Size = New System.Drawing.Size(425, 20)
        Me.TxtMedicalCenter.TabIndex = 101
        Me.TxtMedicalCenter.TabStop = False
        Me.TxtMedicalCenter.Visible = False
        '
        'TxtLegal_KPJ_No
        '
        Me.TxtLegal_KPJ_No.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLegal_KPJ_No.Location = New System.Drawing.Point(304, 159)
        Me.TxtLegal_KPJ_No.MaxLength = 150
        Me.TxtLegal_KPJ_No.Name = "TxtLegal_KPJ_No"
        Me.TxtLegal_KPJ_No.Size = New System.Drawing.Size(219, 20)
        Me.TxtLegal_KPJ_No.TabIndex = 9
        '
        'TxtLegal_ID_Card_No
        '
        Me.TxtLegal_ID_Card_No.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtLegal_ID_Card_No.Location = New System.Drawing.Point(98, 6)
        Me.TxtLegal_ID_Card_No.MaxLength = 150
        Me.TxtLegal_ID_Card_No.Name = "TxtLegal_ID_Card_No"
        Me.TxtLegal_ID_Card_No.Size = New System.Drawing.Size(113, 20)
        Me.TxtLegal_ID_Card_No.TabIndex = 0
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(217, 162)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(75, 13)
        Me.Label29.TabIndex = 100
        Me.Label29.Text = "Help BPJS No"
        '
        'TxtLegal_ID_Address
        '
        Me.TxtLegal_ID_Address.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLegal_ID_Address.Location = New System.Drawing.Point(98, 32)
        Me.TxtLegal_ID_Address.MaxLength = 150
        Me.TxtLegal_ID_Address.Name = "TxtLegal_ID_Address"
        Me.TxtLegal_ID_Address.Size = New System.Drawing.Size(425, 20)
        Me.TxtLegal_ID_Address.TabIndex = 2
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(6, 35)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(59, 13)
        Me.Label13.TabIndex = 71
        Me.Label13.Text = "ID Address"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(217, 9)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(42, 13)
        Me.Label18.TabIndex = 79
        Me.Label18.Text = "Expired"
        '
        'TxtLegal_Astek_No
        '
        Me.TxtLegal_Astek_No.Location = New System.Drawing.Point(98, 159)
        Me.TxtLegal_Astek_No.MaxLength = 150
        Me.TxtLegal_Astek_No.Name = "TxtLegal_Astek_No"
        Me.TxtLegal_Astek_No.Size = New System.Drawing.Size(113, 20)
        Me.TxtLegal_Astek_No.TabIndex = 8
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(6, 61)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(82, 13)
        Me.Label23.TabIndex = 96
        Me.Label23.Text = "Current Address"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(6, 111)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(81, 13)
        Me.Label22.TabIndex = 94
        Me.Label22.Text = "NPWP Address"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 162)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(50, 13)
        Me.Label6.TabIndex = 60
        Me.Label6.Text = "BPJS No"
        '
        'DtpLegal_Expired
        '
        Me.DtpLegal_Expired.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpLegal_Expired.CustomFormat = "dd MMM yyyy"
        Me.DtpLegal_Expired.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpLegal_Expired.Location = New System.Drawing.Point(272, 6)
        Me.DtpLegal_Expired.Name = "DtpLegal_Expired"
        Me.DtpLegal_Expired.Size = New System.Drawing.Size(251, 20)
        Me.DtpLegal_Expired.TabIndex = 1
        '
        'TxtLegal_NPWP_Address
        '
        Me.TxtLegal_NPWP_Address.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLegal_NPWP_Address.Location = New System.Drawing.Point(98, 108)
        Me.TxtLegal_NPWP_Address.MaxLength = 150
        Me.TxtLegal_NPWP_Address.Name = "TxtLegal_NPWP_Address"
        Me.TxtLegal_NPWP_Address.Size = New System.Drawing.Size(425, 20)
        Me.TxtLegal_NPWP_Address.TabIndex = 7
        '
        'TxtLegal_Current_Address
        '
        Me.TxtLegal_Current_Address.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLegal_Current_Address.Location = New System.Drawing.Point(98, 58)
        Me.TxtLegal_Current_Address.MaxLength = 150
        Me.TxtLegal_Current_Address.Name = "TxtLegal_Current_Address"
        Me.TxtLegal_Current_Address.Size = New System.Drawing.Size(425, 20)
        Me.TxtLegal_Current_Address.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 9)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 13)
        Me.Label5.TabIndex = 61
        Me.Label5.Text = "ID Card No"
        '
        'TxtPersonal_NPWP
        '
        Me.TxtPersonal_NPWP.Location = New System.Drawing.Point(98, 84)
        Me.TxtPersonal_NPWP.MaxLength = 150
        Me.TxtPersonal_NPWP.Name = "TxtPersonal_NPWP"
        Me.TxtPersonal_NPWP.Size = New System.Drawing.Size(113, 20)
        Me.TxtPersonal_NPWP.TabIndex = 5
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(217, 87)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(49, 13)
        Me.Label17.TabIndex = 77
        Me.Label17.Text = "Category"
        '
        'CboLegal_Category
        '
        Me.CboLegal_Category.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboLegal_Category.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboLegal_Category.FormattingEnabled = True
        Me.CboLegal_Category.Items.AddRange(New Object() {"TK0", "TK1", "TK2", "TK3", "K0", "K1", "K2", "K3"})
        Me.CboLegal_Category.Location = New System.Drawing.Point(272, 84)
        Me.CboLegal_Category.Name = "CboLegal_Category"
        Me.CboLegal_Category.Size = New System.Drawing.Size(251, 21)
        Me.CboLegal_Category.TabIndex = 6
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.Label45)
        Me.TabPage3.Controls.Add(Me.TxtBank_Code)
        Me.TabPage3.Controls.Add(Me.Label7)
        Me.TabPage3.Controls.Add(Me.Label37)
        Me.TabPage3.Controls.Add(Me.TxtBank_Name)
        Me.TabPage3.Controls.Add(Me.Label38)
        Me.TabPage3.Controls.Add(Me.TxtBank_Acc_No)
        Me.TabPage3.Controls.Add(Me.Label39)
        Me.TabPage3.Controls.Add(Me.TxtBank_Acc_Holder)
        Me.TabPage3.Controls.Add(Me.TxtBank_Branch)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage3.Size = New System.Drawing.Size(560, 239)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Bank"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'Label45
        '
        Me.Label45.AutoSize = True
        Me.Label45.Location = New System.Drawing.Point(6, 60)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(60, 13)
        Me.Label45.TabIndex = 132
        Me.Label45.Text = "Bank Code"
        '
        'TxtBank_Code
        '
        Me.TxtBank_Code.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBank_Code.Location = New System.Drawing.Point(100, 57)
        Me.TxtBank_Code.MaxLength = 150
        Me.TxtBank_Code.Name = "TxtBank_Code"
        Me.TxtBank_Code.Size = New System.Drawing.Size(427, 20)
        Me.TxtBank_Code.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 112)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(41, 13)
        Me.Label7.TabIndex = 130
        Me.Label7.Text = "Branch"
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Location = New System.Drawing.Point(6, 87)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(63, 13)
        Me.Label37.TabIndex = 129
        Me.Label37.Text = "Bank Name"
        '
        'TxtBank_Name
        '
        Me.TxtBank_Name.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBank_Name.Location = New System.Drawing.Point(100, 83)
        Me.TxtBank_Name.MaxLength = 150
        Me.TxtBank_Name.Name = "TxtBank_Name"
        Me.TxtBank_Name.Size = New System.Drawing.Size(427, 20)
        Me.TxtBank_Name.TabIndex = 3
        '
        'Label38
        '
        Me.Label38.AutoSize = True
        Me.Label38.Location = New System.Drawing.Point(6, 35)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(64, 13)
        Me.Label38.TabIndex = 127
        Me.Label38.Text = "Account No"
        '
        'TxtBank_Acc_No
        '
        Me.TxtBank_Acc_No.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBank_Acc_No.Location = New System.Drawing.Point(100, 32)
        Me.TxtBank_Acc_No.MaxLength = 150
        Me.TxtBank_Acc_No.Name = "TxtBank_Acc_No"
        Me.TxtBank_Acc_No.Size = New System.Drawing.Size(427, 20)
        Me.TxtBank_Acc_No.TabIndex = 1
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Location = New System.Drawing.Point(6, 9)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(81, 13)
        Me.Label39.TabIndex = 123
        Me.Label39.Text = "Account Holder"
        '
        'TxtBank_Acc_Holder
        '
        Me.TxtBank_Acc_Holder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBank_Acc_Holder.Location = New System.Drawing.Point(100, 6)
        Me.TxtBank_Acc_Holder.MaxLength = 150
        Me.TxtBank_Acc_Holder.Name = "TxtBank_Acc_Holder"
        Me.TxtBank_Acc_Holder.Size = New System.Drawing.Size(427, 20)
        Me.TxtBank_Acc_Holder.TabIndex = 0
        '
        'TxtBank_Branch
        '
        Me.TxtBank_Branch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBank_Branch.Location = New System.Drawing.Point(100, 109)
        Me.TxtBank_Branch.MaxLength = 150
        Me.TxtBank_Branch.Name = "TxtBank_Branch"
        Me.TxtBank_Branch.Size = New System.Drawing.Size(427, 20)
        Me.TxtBank_Branch.TabIndex = 4
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.TxtNotes)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage4.Size = New System.Drawing.Size(560, 239)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Notes"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'TxtNotes
        '
        Me.TxtNotes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtNotes.Location = New System.Drawing.Point(3, 3)
        Me.TxtNotes.MaxLength = 150
        Me.TxtNotes.Multiline = True
        Me.TxtNotes.Name = "TxtNotes"
        Me.TxtNotes.Size = New System.Drawing.Size(554, 233)
        Me.TxtNotes.TabIndex = 0
        '
        'TabPage7
        '
        Me.TabPage7.Controls.Add(Me.GroupBox3)
        Me.TabPage7.Controls.Add(Me.TxtFingerprintID)
        Me.TabPage7.Controls.Add(Me.GroupBox1)
        Me.TabPage7.Controls.Add(Me.Label44)
        Me.TabPage7.Location = New System.Drawing.Point(4, 22)
        Me.TabPage7.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage7.Size = New System.Drawing.Size(560, 239)
        Me.TabPage7.TabIndex = 4
        Me.TabPage7.Text = "Attendance Settings"
        Me.TabPage7.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.ChkOverrideOTOrder)
        Me.GroupBox3.Controls.Add(Me.chkSendNotifEmailOvertimeOrder)
        Me.GroupBox3.Controls.Add(Me.chkUseOvertimeOrder)
        Me.GroupBox3.Controls.Add(Me.LinkLabel1)
        Me.GroupBox3.Controls.Add(Me.chkOverrideCalculateOTTolerance)
        Me.GroupBox3.Controls.Add(Me.chkOverrideCalculateEarlyOutTolerance)
        Me.GroupBox3.Controls.Add(Me.chkOverrideCalculateLateTolerance)
        Me.GroupBox3.Controls.Add(Me.chkOverrideCalculateEarlyOTTolerance)
        Me.GroupBox3.Controls.Add(Me.Label50)
        Me.GroupBox3.Controls.Add(Me.txtMaxValueOT)
        Me.GroupBox3.Controls.Add(Me.cboRoundingTypeOT)
        Me.GroupBox3.Controls.Add(Me.txtRoundingOT)
        Me.GroupBox3.Controls.Add(Me.txtCalculateAfterOT)
        Me.GroupBox3.Controls.Add(Me.txtMaxValueEarlyOut)
        Me.GroupBox3.Controls.Add(Me.cboRoundingTypeEarlyOut)
        Me.GroupBox3.Controls.Add(Me.txtRoundingEarlyOut)
        Me.GroupBox3.Controls.Add(Me.txtCalculateAfterEarlyOut)
        Me.GroupBox3.Controls.Add(Me.txtMaxValueLate)
        Me.GroupBox3.Controls.Add(Me.cboRoundingTypeLate)
        Me.GroupBox3.Controls.Add(Me.txtRoundingLate)
        Me.GroupBox3.Controls.Add(Me.txtCalculateAfterLate)
        Me.GroupBox3.Controls.Add(Me.Label46)
        Me.GroupBox3.Controls.Add(Me.Label47)
        Me.GroupBox3.Controls.Add(Me.Label48)
        Me.GroupBox3.Controls.Add(Me.txtMaxValueEarlyOT)
        Me.GroupBox3.Controls.Add(Me.cboRoundingTypeEarlyOT)
        Me.GroupBox3.Controls.Add(Me.txtRoundingEarlyOT)
        Me.GroupBox3.Controls.Add(Me.txtCalculateAfterEarlyOT)
        Me.GroupBox3.Controls.Add(Me.Label49)
        Me.GroupBox3.Controls.Add(Me.CboIsCalculateOT)
        Me.GroupBox3.Controls.Add(Me.CboIsCalculateEarlyOut)
        Me.GroupBox3.Controls.Add(Me.ChkOverrideOT)
        Me.GroupBox3.Controls.Add(Me.ChkOverrideEarlyOut)
        Me.GroupBox3.Controls.Add(Me.CboIsCalculateLate)
        Me.GroupBox3.Controls.Add(Me.CboIsCalculateEarlyOT)
        Me.GroupBox3.Controls.Add(Me.ChkOverrideLate)
        Me.GroupBox3.Controls.Add(Me.ChkOverrideEarlyOT)
        Me.GroupBox3.Location = New System.Drawing.Point(5, 75)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(555, 162)
        Me.GroupBox3.TabIndex = 134
        Me.GroupBox3.TabStop = False
        '
        'ChkOverrideOTOrder
        '
        Me.ChkOverrideOTOrder.AutoSize = True
        Me.ChkOverrideOTOrder.Location = New System.Drawing.Point(6, 138)
        Me.ChkOverrideOTOrder.Name = "ChkOverrideOTOrder"
        Me.ChkOverrideOTOrder.Size = New System.Drawing.Size(113, 17)
        Me.ChkOverrideOTOrder.TabIndex = 212
        Me.ChkOverrideOTOrder.Text = "Override OT Order"
        Me.ChkOverrideOTOrder.UseVisualStyleBackColor = True
        '
        'chkSendNotifEmailOvertimeOrder
        '
        Me.chkSendNotifEmailOvertimeOrder.AutoSize = True
        Me.chkSendNotifEmailOvertimeOrder.Location = New System.Drawing.Point(265, 139)
        Me.chkSendNotifEmailOvertimeOrder.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkSendNotifEmailOvertimeOrder.Name = "chkSendNotifEmailOvertimeOrder"
        Me.chkSendNotifEmailOvertimeOrder.Size = New System.Drawing.Size(280, 17)
        Me.chkSendNotifEmailOvertimeOrder.TabIndex = 214
        Me.chkSendNotifEmailOvertimeOrder.Text = "Send notification email for overtime order authorization"
        Me.chkSendNotifEmailOvertimeOrder.UseVisualStyleBackColor = True
        '
        'chkUseOvertimeOrder
        '
        Me.chkUseOvertimeOrder.AutoSize = True
        Me.chkUseOvertimeOrder.Location = New System.Drawing.Point(136, 139)
        Me.chkUseOvertimeOrder.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkUseOvertimeOrder.Name = "chkUseOvertimeOrder"
        Me.chkUseOvertimeOrder.Size = New System.Drawing.Size(115, 17)
        Me.chkUseOvertimeOrder.TabIndex = 213
        Me.chkUseOvertimeOrder.Text = "Use overtime order"
        Me.chkUseOvertimeOrder.UseVisualStyleBackColor = True
        '
        'LinkLabel1
        '
        Me.LinkLabel1.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(434, 16)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(91, 13)
        Me.LinkLabel1.TabIndex = 212
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Advance Settings"
        Me.LinkLabel1.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'chkOverrideCalculateOTTolerance
        '
        Me.chkOverrideCalculateOTTolerance.AutoSize = True
        Me.chkOverrideCalculateOTTolerance.Checked = True
        Me.chkOverrideCalculateOTTolerance.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOverrideCalculateOTTolerance.Location = New System.Drawing.Point(436, 112)
        Me.chkOverrideCalculateOTTolerance.Name = "chkOverrideCalculateOTTolerance"
        Me.chkOverrideCalculateOTTolerance.Size = New System.Drawing.Size(121, 17)
        Me.chkOverrideCalculateOTTolerance.TabIndex = 211
        Me.chkOverrideCalculateOTTolerance.Text = "Calculate Tolerance"
        Me.chkOverrideCalculateOTTolerance.UseVisualStyleBackColor = True
        '
        'chkOverrideCalculateEarlyOutTolerance
        '
        Me.chkOverrideCalculateEarlyOutTolerance.AutoSize = True
        Me.chkOverrideCalculateEarlyOutTolerance.Checked = True
        Me.chkOverrideCalculateEarlyOutTolerance.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOverrideCalculateEarlyOutTolerance.Location = New System.Drawing.Point(436, 89)
        Me.chkOverrideCalculateEarlyOutTolerance.Name = "chkOverrideCalculateEarlyOutTolerance"
        Me.chkOverrideCalculateEarlyOutTolerance.Size = New System.Drawing.Size(121, 17)
        Me.chkOverrideCalculateEarlyOutTolerance.TabIndex = 210
        Me.chkOverrideCalculateEarlyOutTolerance.Text = "Calculate Tolerance"
        Me.chkOverrideCalculateEarlyOutTolerance.UseVisualStyleBackColor = True
        '
        'chkOverrideCalculateLateTolerance
        '
        Me.chkOverrideCalculateLateTolerance.AutoSize = True
        Me.chkOverrideCalculateLateTolerance.Checked = True
        Me.chkOverrideCalculateLateTolerance.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOverrideCalculateLateTolerance.Location = New System.Drawing.Point(436, 66)
        Me.chkOverrideCalculateLateTolerance.Name = "chkOverrideCalculateLateTolerance"
        Me.chkOverrideCalculateLateTolerance.Size = New System.Drawing.Size(121, 17)
        Me.chkOverrideCalculateLateTolerance.TabIndex = 209
        Me.chkOverrideCalculateLateTolerance.Text = "Calculate Tolerance"
        Me.chkOverrideCalculateLateTolerance.UseVisualStyleBackColor = True
        '
        'chkOverrideCalculateEarlyOTTolerance
        '
        Me.chkOverrideCalculateEarlyOTTolerance.AutoSize = True
        Me.chkOverrideCalculateEarlyOTTolerance.Checked = True
        Me.chkOverrideCalculateEarlyOTTolerance.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOverrideCalculateEarlyOTTolerance.Location = New System.Drawing.Point(436, 43)
        Me.chkOverrideCalculateEarlyOTTolerance.Name = "chkOverrideCalculateEarlyOTTolerance"
        Me.chkOverrideCalculateEarlyOTTolerance.Size = New System.Drawing.Size(121, 17)
        Me.chkOverrideCalculateEarlyOTTolerance.TabIndex = 208
        Me.chkOverrideCalculateEarlyOTTolerance.Text = "Calculate Tolerance"
        Me.chkOverrideCalculateEarlyOTTolerance.UseVisualStyleBackColor = True
        '
        'Label50
        '
        Me.Label50.AutoSize = True
        Me.Label50.Location = New System.Drawing.Point(139, 16)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(51, 13)
        Me.Label50.TabIndex = 207
        Me.Label50.Text = "Calculate"
        '
        'txtMaxValueOT
        '
        Me.txtMaxValueOT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMaxValueOT.Location = New System.Drawing.Point(389, 110)
        Me.txtMaxValueOT.MaxLength = 150
        Me.txtMaxValueOT.Name = "txtMaxValueOT"
        Me.txtMaxValueOT.Size = New System.Drawing.Size(41, 20)
        Me.txtMaxValueOT.TabIndex = 157
        Me.txtMaxValueOT.Text = "00.00"
        Me.txtMaxValueOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboRoundingTypeOT
        '
        Me.cboRoundingTypeOT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingTypeOT.FormattingEnabled = True
        Me.cboRoundingTypeOT.Items.AddRange(New Object() {"No Rounding", "Round Up", "Round down", "Middle"})
        Me.cboRoundingTypeOT.Location = New System.Drawing.Point(245, 110)
        Me.cboRoundingTypeOT.Name = "cboRoundingTypeOT"
        Me.cboRoundingTypeOT.Size = New System.Drawing.Size(91, 21)
        Me.cboRoundingTypeOT.TabIndex = 155
        '
        'txtRoundingOT
        '
        Me.txtRoundingOT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtRoundingOT.Location = New System.Drawing.Point(342, 110)
        Me.txtRoundingOT.MaxLength = 150
        Me.txtRoundingOT.Name = "txtRoundingOT"
        Me.txtRoundingOT.Size = New System.Drawing.Size(41, 20)
        Me.txtRoundingOT.TabIndex = 156
        Me.txtRoundingOT.Text = "00.00"
        Me.txtRoundingOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCalculateAfterOT
        '
        Me.txtCalculateAfterOT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCalculateAfterOT.Location = New System.Drawing.Point(198, 110)
        Me.txtCalculateAfterOT.MaxLength = 150
        Me.txtCalculateAfterOT.Name = "txtCalculateAfterOT"
        Me.txtCalculateAfterOT.Size = New System.Drawing.Size(41, 20)
        Me.txtCalculateAfterOT.TabIndex = 154
        Me.txtCalculateAfterOT.Text = "00.00"
        Me.txtCalculateAfterOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtMaxValueEarlyOut
        '
        Me.txtMaxValueEarlyOut.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMaxValueEarlyOut.Location = New System.Drawing.Point(389, 87)
        Me.txtMaxValueEarlyOut.MaxLength = 150
        Me.txtMaxValueEarlyOut.Name = "txtMaxValueEarlyOut"
        Me.txtMaxValueEarlyOut.Size = New System.Drawing.Size(41, 20)
        Me.txtMaxValueEarlyOut.TabIndex = 151
        Me.txtMaxValueEarlyOut.Text = "00.00"
        Me.txtMaxValueEarlyOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboRoundingTypeEarlyOut
        '
        Me.cboRoundingTypeEarlyOut.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingTypeEarlyOut.FormattingEnabled = True
        Me.cboRoundingTypeEarlyOut.Items.AddRange(New Object() {"No Rounding", "Round Up", "Round down", "Middle"})
        Me.cboRoundingTypeEarlyOut.Location = New System.Drawing.Point(245, 87)
        Me.cboRoundingTypeEarlyOut.Name = "cboRoundingTypeEarlyOut"
        Me.cboRoundingTypeEarlyOut.Size = New System.Drawing.Size(91, 21)
        Me.cboRoundingTypeEarlyOut.TabIndex = 149
        '
        'txtRoundingEarlyOut
        '
        Me.txtRoundingEarlyOut.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtRoundingEarlyOut.Location = New System.Drawing.Point(342, 87)
        Me.txtRoundingEarlyOut.MaxLength = 150
        Me.txtRoundingEarlyOut.Name = "txtRoundingEarlyOut"
        Me.txtRoundingEarlyOut.Size = New System.Drawing.Size(41, 20)
        Me.txtRoundingEarlyOut.TabIndex = 150
        Me.txtRoundingEarlyOut.Text = "00.00"
        Me.txtRoundingEarlyOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCalculateAfterEarlyOut
        '
        Me.txtCalculateAfterEarlyOut.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCalculateAfterEarlyOut.Location = New System.Drawing.Point(198, 87)
        Me.txtCalculateAfterEarlyOut.MaxLength = 150
        Me.txtCalculateAfterEarlyOut.Name = "txtCalculateAfterEarlyOut"
        Me.txtCalculateAfterEarlyOut.Size = New System.Drawing.Size(41, 20)
        Me.txtCalculateAfterEarlyOut.TabIndex = 148
        Me.txtCalculateAfterEarlyOut.Text = "00.00"
        Me.txtCalculateAfterEarlyOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtMaxValueLate
        '
        Me.txtMaxValueLate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMaxValueLate.Location = New System.Drawing.Point(389, 64)
        Me.txtMaxValueLate.MaxLength = 150
        Me.txtMaxValueLate.Name = "txtMaxValueLate"
        Me.txtMaxValueLate.Size = New System.Drawing.Size(41, 20)
        Me.txtMaxValueLate.TabIndex = 145
        Me.txtMaxValueLate.Text = "00.00"
        Me.txtMaxValueLate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboRoundingTypeLate
        '
        Me.cboRoundingTypeLate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingTypeLate.FormattingEnabled = True
        Me.cboRoundingTypeLate.Items.AddRange(New Object() {"No Rounding", "Round Up", "Round down", "Middle"})
        Me.cboRoundingTypeLate.Location = New System.Drawing.Point(245, 64)
        Me.cboRoundingTypeLate.Name = "cboRoundingTypeLate"
        Me.cboRoundingTypeLate.Size = New System.Drawing.Size(91, 21)
        Me.cboRoundingTypeLate.TabIndex = 143
        '
        'txtRoundingLate
        '
        Me.txtRoundingLate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtRoundingLate.Location = New System.Drawing.Point(342, 64)
        Me.txtRoundingLate.MaxLength = 150
        Me.txtRoundingLate.Name = "txtRoundingLate"
        Me.txtRoundingLate.Size = New System.Drawing.Size(41, 20)
        Me.txtRoundingLate.TabIndex = 144
        Me.txtRoundingLate.Text = "00.00"
        Me.txtRoundingLate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCalculateAfterLate
        '
        Me.txtCalculateAfterLate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCalculateAfterLate.Location = New System.Drawing.Point(198, 64)
        Me.txtCalculateAfterLate.MaxLength = 150
        Me.txtCalculateAfterLate.Name = "txtCalculateAfterLate"
        Me.txtCalculateAfterLate.Size = New System.Drawing.Size(41, 20)
        Me.txtCalculateAfterLate.TabIndex = 142
        Me.txtCalculateAfterLate.Text = "00.00"
        Me.txtCalculateAfterLate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label46
        '
        Me.Label46.Location = New System.Drawing.Point(391, 9)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(37, 26)
        Me.Label46.TabIndex = 194
        Me.Label46.Text = "Max value"
        Me.Label46.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label47
        '
        Me.Label47.Location = New System.Drawing.Point(332, 9)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(60, 26)
        Me.Label47.TabIndex = 193
        Me.Label47.Text = "Rounding value"
        Me.Label47.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label48
        '
        Me.Label48.Location = New System.Drawing.Point(252, 16)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(76, 13)
        Me.Label48.TabIndex = 192
        Me.Label48.Text = "Rounding type"
        '
        'txtMaxValueEarlyOT
        '
        Me.txtMaxValueEarlyOT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMaxValueEarlyOT.Location = New System.Drawing.Point(389, 41)
        Me.txtMaxValueEarlyOT.MaxLength = 150
        Me.txtMaxValueEarlyOT.Name = "txtMaxValueEarlyOT"
        Me.txtMaxValueEarlyOT.Size = New System.Drawing.Size(41, 20)
        Me.txtMaxValueEarlyOT.TabIndex = 139
        Me.txtMaxValueEarlyOT.Text = "00.00"
        Me.txtMaxValueEarlyOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboRoundingTypeEarlyOT
        '
        Me.cboRoundingTypeEarlyOT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingTypeEarlyOT.FormattingEnabled = True
        Me.cboRoundingTypeEarlyOT.Items.AddRange(New Object() {"No Rounding", "Round Up", "Round down", "Middle"})
        Me.cboRoundingTypeEarlyOT.Location = New System.Drawing.Point(245, 41)
        Me.cboRoundingTypeEarlyOT.Name = "cboRoundingTypeEarlyOT"
        Me.cboRoundingTypeEarlyOT.Size = New System.Drawing.Size(91, 21)
        Me.cboRoundingTypeEarlyOT.TabIndex = 137
        '
        'txtRoundingEarlyOT
        '
        Me.txtRoundingEarlyOT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtRoundingEarlyOT.Location = New System.Drawing.Point(342, 41)
        Me.txtRoundingEarlyOT.MaxLength = 150
        Me.txtRoundingEarlyOT.Name = "txtRoundingEarlyOT"
        Me.txtRoundingEarlyOT.Size = New System.Drawing.Size(41, 20)
        Me.txtRoundingEarlyOT.TabIndex = 138
        Me.txtRoundingEarlyOT.Text = "00.00"
        Me.txtRoundingEarlyOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCalculateAfterEarlyOT
        '
        Me.txtCalculateAfterEarlyOT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCalculateAfterEarlyOT.Location = New System.Drawing.Point(198, 41)
        Me.txtCalculateAfterEarlyOT.MaxLength = 150
        Me.txtCalculateAfterEarlyOT.Name = "txtCalculateAfterEarlyOT"
        Me.txtCalculateAfterEarlyOT.Size = New System.Drawing.Size(41, 20)
        Me.txtCalculateAfterEarlyOT.TabIndex = 136
        Me.txtCalculateAfterEarlyOT.Text = "00.00"
        Me.txtCalculateAfterEarlyOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label49
        '
        Me.Label49.Location = New System.Drawing.Point(191, 9)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(55, 26)
        Me.Label49.TabIndex = 188
        Me.Label49.Text = "Calculate after"
        Me.Label49.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CboIsCalculateOT
        '
        Me.CboIsCalculateOT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboIsCalculateOT.FormattingEnabled = True
        Me.CboIsCalculateOT.Items.AddRange(New Object() {"Yes", "No"})
        Me.CboIsCalculateOT.Location = New System.Drawing.Point(137, 110)
        Me.CboIsCalculateOT.Name = "CboIsCalculateOT"
        Me.CboIsCalculateOT.Size = New System.Drawing.Size(55, 21)
        Me.CboIsCalculateOT.TabIndex = 153
        '
        'CboIsCalculateEarlyOut
        '
        Me.CboIsCalculateEarlyOut.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboIsCalculateEarlyOut.FormattingEnabled = True
        Me.CboIsCalculateEarlyOut.Items.AddRange(New Object() {"Yes", "No"})
        Me.CboIsCalculateEarlyOut.Location = New System.Drawing.Point(137, 87)
        Me.CboIsCalculateEarlyOut.Name = "CboIsCalculateEarlyOut"
        Me.CboIsCalculateEarlyOut.Size = New System.Drawing.Size(55, 21)
        Me.CboIsCalculateEarlyOut.TabIndex = 147
        '
        'ChkOverrideOT
        '
        Me.ChkOverrideOT.AutoSize = True
        Me.ChkOverrideOT.Location = New System.Drawing.Point(6, 112)
        Me.ChkOverrideOT.Name = "ChkOverrideOT"
        Me.ChkOverrideOT.Size = New System.Drawing.Size(104, 17)
        Me.ChkOverrideOT.TabIndex = 152
        Me.ChkOverrideOT.Text = "Override OT rule"
        Me.ChkOverrideOT.UseVisualStyleBackColor = True
        '
        'ChkOverrideEarlyOut
        '
        Me.ChkOverrideEarlyOut.AutoSize = True
        Me.ChkOverrideEarlyOut.Location = New System.Drawing.Point(6, 89)
        Me.ChkOverrideEarlyOut.Name = "ChkOverrideEarlyOut"
        Me.ChkOverrideEarlyOut.Size = New System.Drawing.Size(129, 17)
        Me.ChkOverrideEarlyOut.TabIndex = 146
        Me.ChkOverrideEarlyOut.Text = "Override early out rule"
        Me.ChkOverrideEarlyOut.UseVisualStyleBackColor = True
        '
        'CboIsCalculateLate
        '
        Me.CboIsCalculateLate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboIsCalculateLate.FormattingEnabled = True
        Me.CboIsCalculateLate.Items.AddRange(New Object() {"Yes", "No"})
        Me.CboIsCalculateLate.Location = New System.Drawing.Point(137, 64)
        Me.CboIsCalculateLate.Name = "CboIsCalculateLate"
        Me.CboIsCalculateLate.Size = New System.Drawing.Size(55, 21)
        Me.CboIsCalculateLate.TabIndex = 141
        '
        'CboIsCalculateEarlyOT
        '
        Me.CboIsCalculateEarlyOT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboIsCalculateEarlyOT.FormattingEnabled = True
        Me.CboIsCalculateEarlyOT.Items.AddRange(New Object() {"Yes", "No"})
        Me.CboIsCalculateEarlyOT.Location = New System.Drawing.Point(137, 41)
        Me.CboIsCalculateEarlyOT.Name = "CboIsCalculateEarlyOT"
        Me.CboIsCalculateEarlyOT.Size = New System.Drawing.Size(55, 21)
        Me.CboIsCalculateEarlyOT.TabIndex = 135
        '
        'ChkOverrideLate
        '
        Me.ChkOverrideLate.AutoSize = True
        Me.ChkOverrideLate.Location = New System.Drawing.Point(6, 66)
        Me.ChkOverrideLate.Name = "ChkOverrideLate"
        Me.ChkOverrideLate.Size = New System.Drawing.Size(106, 17)
        Me.ChkOverrideLate.TabIndex = 140
        Me.ChkOverrideLate.Text = "Override late rule"
        Me.ChkOverrideLate.UseVisualStyleBackColor = True
        '
        'ChkOverrideEarlyOT
        '
        Me.ChkOverrideEarlyOT.AutoSize = True
        Me.ChkOverrideEarlyOT.Location = New System.Drawing.Point(6, 43)
        Me.ChkOverrideEarlyOT.Name = "ChkOverrideEarlyOT"
        Me.ChkOverrideEarlyOT.Size = New System.Drawing.Size(129, 17)
        Me.ChkOverrideEarlyOT.TabIndex = 131
        Me.ChkOverrideEarlyOT.Text = "Override early OT rule"
        Me.ChkOverrideEarlyOT.UseVisualStyleBackColor = True
        '
        'TxtFingerprintID
        '
        Me.TxtFingerprintID.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFingerprintID.Location = New System.Drawing.Point(106, 7)
        Me.TxtFingerprintID.MaxLength = 150
        Me.TxtFingerprintID.Name = "TxtFingerprintID"
        Me.TxtFingerprintID.Size = New System.Drawing.Size(453, 20)
        Me.TxtFingerprintID.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.ChkDoesnNotClockOut)
        Me.GroupBox1.Controls.Add(Me.ChkOverrideClockOut)
        Me.GroupBox1.Controls.Add(Me.ChkDoesnNotClockIn)
        Me.GroupBox1.Controls.Add(Me.ChkOverrideClockIn)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 23)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(555, 80)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'ChkDoesnNotClockOut
        '
        Me.ChkDoesnNotClockOut.AutoSize = True
        Me.ChkDoesnNotClockOut.Location = New System.Drawing.Point(144, 30)
        Me.ChkDoesnNotClockOut.Name = "ChkDoesnNotClockOut"
        Me.ChkDoesnNotClockOut.Size = New System.Drawing.Size(172, 17)
        Me.ChkDoesnNotClockOut.TabIndex = 3
        Me.ChkDoesnNotClockOut.Text = "This person does not clock out"
        Me.ChkDoesnNotClockOut.UseVisualStyleBackColor = True
        '
        'ChkOverrideClockOut
        '
        Me.ChkOverrideClockOut.AutoSize = True
        Me.ChkOverrideClockOut.Location = New System.Drawing.Point(8, 30)
        Me.ChkOverrideClockOut.Name = "ChkOverrideClockOut"
        Me.ChkOverrideClockOut.Size = New System.Drawing.Size(133, 17)
        Me.ChkOverrideClockOut.TabIndex = 2
        Me.ChkOverrideClockOut.Text = "Override clock out rule"
        Me.ChkOverrideClockOut.UseVisualStyleBackColor = True
        '
        'ChkDoesnNotClockIn
        '
        Me.ChkDoesnNotClockIn.AutoSize = True
        Me.ChkDoesnNotClockIn.Location = New System.Drawing.Point(144, 9)
        Me.ChkDoesnNotClockIn.Name = "ChkDoesnNotClockIn"
        Me.ChkDoesnNotClockIn.Size = New System.Drawing.Size(165, 17)
        Me.ChkDoesnNotClockIn.TabIndex = 1
        Me.ChkDoesnNotClockIn.Text = "This person does not clock in"
        Me.ChkDoesnNotClockIn.UseVisualStyleBackColor = True
        '
        'ChkOverrideClockIn
        '
        Me.ChkOverrideClockIn.AutoSize = True
        Me.ChkOverrideClockIn.Location = New System.Drawing.Point(8, 9)
        Me.ChkOverrideClockIn.Name = "ChkOverrideClockIn"
        Me.ChkOverrideClockIn.Size = New System.Drawing.Size(126, 17)
        Me.ChkOverrideClockIn.TabIndex = 0
        Me.ChkOverrideClockIn.Text = "Override clock in rule"
        Me.ChkOverrideClockIn.UseVisualStyleBackColor = True
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Location = New System.Drawing.Point(5, 10)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(70, 13)
        Me.Label44.TabIndex = 132
        Me.Label44.Text = "Fingerprint ID"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(503, 452)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(331, 452)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnPrint.Location = New System.Drawing.Point(589, 452)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnViewWorkShift
        '
        Me.btnViewWorkShift.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnViewWorkShift.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnViewWorkShift.Location = New System.Drawing.Point(675, 452)
        Me.btnViewWorkShift.Name = "btnViewWorkShift"
        Me.btnViewWorkShift.Size = New System.Drawing.Size(80, 26)
        Me.btnViewWorkShift.TabIndex = 5
        Me.btnViewWorkShift.Text = "Workshift"
        Me.btnViewWorkShift.UseVisualStyleBackColor = True
        Me.btnViewWorkShift.Visible = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2, Me.ToolStripStatusLabel3, Me.ToolStripStatusLabel4})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 487)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1004, 22)
        Me.StatusStrip1.TabIndex = 41
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(247, 17)
        Me.ToolStripStatusLabel1.Spring = True
        Me.ToolStripStatusLabel1.Text = "Created By: "
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(247, 17)
        Me.ToolStripStatusLabel2.Spring = True
        Me.ToolStripStatusLabel2.Text = "Created Date:"
        '
        'ToolStripStatusLabel3
        '
        Me.ToolStripStatusLabel3.Name = "ToolStripStatusLabel3"
        Me.ToolStripStatusLabel3.Size = New System.Drawing.Size(247, 17)
        Me.ToolStripStatusLabel3.Spring = True
        Me.ToolStripStatusLabel3.Text = "Modified By: "
        '
        'ToolStripStatusLabel4
        '
        Me.ToolStripStatusLabel4.Name = "ToolStripStatusLabel4"
        Me.ToolStripStatusLabel4.Size = New System.Drawing.Size(247, 17)
        Me.ToolStripStatusLabel4.Spring = True
        Me.ToolStripStatusLabel4.Text = "Modified Date: "
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "Picture Files|*.jpg;*.jpeg;*.bmp;*.png"
        Me.OpenFileDialog1.Title = "Find Picture"
        '
        'BgwDeleteRowDgv
        '
        '
        'BackgroundWorker1
        '
        '
        'OpenFileDialog2
        '
        Me.OpenFileDialog2.Title = "Please select a file"
        '
        'llbCopy
        '
        Me.llbCopy.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCopy.AutoSize = True
        Me.llbCopy.Location = New System.Drawing.Point(965, 459)
        Me.llbCopy.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbCopy.Name = "llbCopy"
        Me.llbCopy.Size = New System.Drawing.Size(31, 13)
        Me.llbCopy.TabIndex = 42
        Me.llbCopy.TabStop = True
        Me.llbCopy.Text = "Copy"
        '
        'FrmMasterKaryawan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(1004, 509)
        Me.Controls.Add(Me.llbCopy)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnViewWorkShift)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1020, 547)
        Me.Name = "FrmMasterKaryawan"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Employee Data"
        Me.gboDetail.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage5.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dgvFamily, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        CType(Me.DGVKompetensi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage8.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.DGVKontrak, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGVAttanchmentKontrak, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage9.ResumeLayout(False)
        CType(Me.DGVAttanchment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.TabPage7.ResumeLayout(False)
        Me.TabPage7.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryEmployee As System.Windows.Forms.TextBox
    Friend WithEvents TxtFullName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtEmployeeCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnFindDepartemen As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryDepartemen As System.Windows.Forms.TextBox
    Friend WithEvents TxtDepartemen As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtNIK As System.Windows.Forms.TextBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents cboStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents DtpAdmision As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryPosition As System.Windows.Forms.TextBox
    Friend WithEvents TxtPositionCode As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents TextBox11 As System.Windows.Forms.TextBox
    Friend WithEvents TxtSalaryFormula As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtPicturePath As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindPosition As System.Windows.Forms.Button
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents CboEmployeeStatus As System.Windows.Forms.ComboBox
    Friend WithEvents btnViewWorkShift As System.Windows.Forms.Button
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TxtPersonal_Mother_Address As System.Windows.Forms.TextBox
    Friend WithEvents CboPersonal_Sex As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents CboPersonal_Religion As System.Windows.Forms.ComboBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TxtLegal_KPJ_No As System.Windows.Forms.TextBox
    Friend WithEvents TxtLegal_ID_Card_No As System.Windows.Forms.TextBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents TxtLegal_ID_Address As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents TxtLegal_Astek_No As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents DtpLegal_Expired As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtLegal_NPWP_Address As System.Windows.Forms.TextBox
    Friend WithEvents TxtLegal_Current_Address As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtPersonal_NPWP As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents CboLegal_Category As System.Windows.Forms.ComboBox
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents TxtPersonal_Phone As System.Windows.Forms.TextBox
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents TxtPersonal_Email As System.Windows.Forms.TextBox
    Friend WithEvents TxtPersonal_Nationality As System.Windows.Forms.TextBox
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents CboPersonal_BloodType As System.Windows.Forms.ComboBox
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents LblDateOfMarriage As System.Windows.Forms.Label
    Friend WithEvents DtpPersonal_MTdate As System.Windows.Forms.DateTimePicker
    Friend WithEvents CboPersonal_MartialStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TxtPersonal_Education As System.Windows.Forms.TextBox
    Friend WithEvents TxtPersonal_BirthPlace As System.Windows.Forms.TextBox
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents DtpPersonal_BPDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents TxtPersonal_Mother_Name As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents TxtBank_Name As System.Windows.Forms.TextBox
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents TxtBank_Acc_No As System.Windows.Forms.TextBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents TxtBank_Acc_Holder As System.Windows.Forms.TextBox
    Friend WithEvents TxtBank_Branch As System.Windows.Forms.TextBox
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents TxtNotes As System.Windows.Forms.TextBox
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel3 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel4 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents TxtPrefix As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents TxtSuffix As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents TxtLastName As System.Windows.Forms.TextBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents TxtMiddleName As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents TxtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents dgvFamily As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents DGVKompetensi As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHubKeluarga As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColIsDependent As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColEducation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSex As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColGolDarah As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColContactNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBirthPlace As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBirthday As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TxtMedicalCenter As System.Windows.Forms.TextBox
    Friend WithEvents ColCompetenceNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCompetence_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCempetence_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCompetence_StandarValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCompetence_Value As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCompetence_NextCertificationDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCompetence_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCompetence_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCempetence_DataType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BgwDeleteRowDgv As System.ComponentModel.BackgroundWorker
    Friend WithEvents ChkNoNPWP As System.Windows.Forms.CheckBox
    Friend WithEvents CboForeignEmployee As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryCountryCode As System.Windows.Forms.TextBox
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents CboCountryDesc As System.Windows.Forms.ComboBox
    Friend WithEvents dtpAppoinmentDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
    Friend WithEvents TxtFingerprintID As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkDoesnNotClockOut As System.Windows.Forms.CheckBox
    Friend WithEvents ChkOverrideClockOut As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDoesnNotClockIn As System.Windows.Forms.CheckBox
    Friend WithEvents ChkOverrideClockIn As System.Windows.Forms.CheckBox
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents CboIsCalculateOT As System.Windows.Forms.ComboBox
    Friend WithEvents CboIsCalculateEarlyOut As System.Windows.Forms.ComboBox
    Friend WithEvents ChkOverrideOT As System.Windows.Forms.CheckBox
    Friend WithEvents ChkOverrideEarlyOut As System.Windows.Forms.CheckBox
    Friend WithEvents CboIsCalculateLate As System.Windows.Forms.ComboBox
    Friend WithEvents CboIsCalculateEarlyOT As System.Windows.Forms.ComboBox
    Friend WithEvents ChkOverrideLate As System.Windows.Forms.CheckBox
    Friend WithEvents ChkOverrideEarlyOT As System.Windows.Forms.CheckBox
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents TxtBank_Code As System.Windows.Forms.TextBox
    Friend WithEvents txtMaxValueOT As System.Windows.Forms.TextBox
    Friend WithEvents cboRoundingTypeOT As System.Windows.Forms.ComboBox
    Friend WithEvents txtRoundingOT As System.Windows.Forms.TextBox
    Friend WithEvents txtCalculateAfterOT As System.Windows.Forms.TextBox
    Friend WithEvents txtMaxValueEarlyOut As System.Windows.Forms.TextBox
    Friend WithEvents cboRoundingTypeEarlyOut As System.Windows.Forms.ComboBox
    Friend WithEvents txtRoundingEarlyOut As System.Windows.Forms.TextBox
    Friend WithEvents txtCalculateAfterEarlyOut As System.Windows.Forms.TextBox
    Friend WithEvents txtMaxValueLate As System.Windows.Forms.TextBox
    Friend WithEvents cboRoundingTypeLate As System.Windows.Forms.ComboBox
    Friend WithEvents txtRoundingLate As System.Windows.Forms.TextBox
    Friend WithEvents txtCalculateAfterLate As System.Windows.Forms.TextBox
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents Label48 As System.Windows.Forms.Label
    Friend WithEvents txtMaxValueEarlyOT As System.Windows.Forms.TextBox
    Friend WithEvents cboRoundingTypeEarlyOT As System.Windows.Forms.ComboBox
    Friend WithEvents txtRoundingEarlyOT As System.Windows.Forms.TextBox
    Friend WithEvents txtCalculateAfterEarlyOT As System.Windows.Forms.TextBox
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents Label50 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents DtpResignDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblResignDate As System.Windows.Forms.Label
    Friend WithEvents chkOverrideCalculateOTTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents chkOverrideCalculateEarlyOutTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents chkOverrideCalculateLateTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents chkOverrideCalculateEarlyOTTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents Label51 As System.Windows.Forms.Label
    Friend WithEvents dtpFirstTaxPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents TabPage8 As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents DGVKontrak As System.Windows.Forms.DataGridView
    Public WithEvents DGVAttanchmentKontrak As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage9 As System.Windows.Forms.TabPage
    Public WithEvents DGVAttanchment As System.Windows.Forms.DataGridView
    Friend WithEvents ColKontrakAtt_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKontrakAtt_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKontrakAtt_LocationFile As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKontrakAtt_Preview As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColKontrakAtt_BtnOpenDelete As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColKontrakAtt_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAtt_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAtt_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAtt_LocationFile As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAtt_Preview As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColAtt_BtnOpenDelete As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColAtt_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OpenFileDialog2 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ColContract_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColContract_Type As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColContract_DocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColContract_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColContract_Start As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColContract_End As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColContract_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColContract_PrimaryType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColContract_IdentifierContract As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkSendNotifEmailOvertimeOrder As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseOvertimeOrder As System.Windows.Forms.CheckBox
    Friend WithEvents ChkOverrideOTOrder As System.Windows.Forms.CheckBox
    Friend WithEvents llbCopy As System.Windows.Forms.LinkLabel
    Friend WithEvents chkloses_tax_obligations As System.Windows.Forms.CheckBox
End Class
