<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterSalaryFormula_OtherDeduction
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnDelete = New System.Windows.Forms.Button
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.TxtPrimaryOther = New System.Windows.Forms.TextBox
        Me.btnFind = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.gboHeader = New System.Windows.Forms.GroupBox
        Me.TxtRoundingUnit = New System.Windows.Forms.TextBox
        Me.CboRoundingType = New System.Windows.Forms.ComboBox
        Me.DtpYearlyDate = New System.Windows.Forms.DateTimePicker
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtMonthlyMinValue = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.TxtMonthlyMaxValue = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtYearlyMinValue = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtYearlyMaxValue = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.cboCalculateOn = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cboCalculatePeriod = New System.Windows.Forms.ComboBox
        Me.BtnEditFormula = New System.Windows.Forms.Button
        Me.TxtFormula = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtPrimaryReffered = New System.Windows.Forms.TextBox
        Me.TxtRefferedCode = New System.Windows.Forms.TextBox
        Me.BtnFindReffered = New System.Windows.Forms.Button
        Me.TxtRefferedDesc = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtPrimarySlip = New System.Windows.Forms.TextBox
        Me.TxtSlipCode = New System.Windows.Forms.TextBox
        Me.BtnFindSlip = New System.Windows.Forms.Button
        Me.TxtSlipDesc = New System.Windows.Forms.TextBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtDescription = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPosition = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryEmployee = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnNew = New System.Windows.Forms.Button
        Me.BgwHapus = New System.ComponentModel.BackgroundWorker
        Me.StatusStrip1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.gboHeader.SuspendLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(453, 553)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(107, 32)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(388, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 602)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(797, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 45
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(388, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(388, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 627)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip2.Size = New System.Drawing.Size(797, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 46
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(388, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(404, 505)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(107, 32)
        Me.btnCancel.TabIndex = 15
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(568, 553)
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
        Me.btnEdit.Location = New System.Drawing.Point(339, 553)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(107, 32)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'TxtPrimaryOther
        '
        Me.TxtPrimaryOther.Location = New System.Drawing.Point(104, 23)
        Me.TxtPrimaryOther.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryOther.Name = "TxtPrimaryOther"
        Me.TxtPrimaryOther.Size = New System.Drawing.Size(56, 22)
        Me.TxtPrimaryOther.TabIndex = 0
        Me.TxtPrimaryOther.TabStop = False
        Me.TxtPrimaryOther.Text = "-1"
        Me.TxtPrimaryOther.Visible = False
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(109, 553)
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
        Me.gboDetail.Controls.Add(Me.gboHeader)
        Me.gboDetail.Controls.Add(Me.cboCalculateOn)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.cboCalculatePeriod)
        Me.gboDetail.Controls.Add(Me.BtnEditFormula)
        Me.gboDetail.Controls.Add(Me.TxtFormula)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryReffered)
        Me.gboDetail.Controls.Add(Me.TxtRefferedCode)
        Me.gboDetail.Controls.Add(Me.BtnFindReffered)
        Me.gboDetail.Controls.Add(Me.TxtRefferedDesc)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.TxtPrimarySlip)
        Me.gboDetail.Controls.Add(Me.TxtSlipCode)
        Me.gboDetail.Controls.Add(Me.BtnFindSlip)
        Me.gboDetail.Controls.Add(Me.TxtSlipDesc)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryOther)
        Me.gboDetail.Controls.Add(Me.TxtDescription)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.TxtCode)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.DgvItem)
        Me.gboDetail.Location = New System.Drawing.Point(12, 1)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Size = New System.Drawing.Size(757, 544)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'gboHeader
        '
        Me.gboHeader.Controls.Add(Me.TxtRoundingUnit)
        Me.gboHeader.Controls.Add(Me.CboRoundingType)
        Me.gboHeader.Controls.Add(Me.DtpYearlyDate)
        Me.gboHeader.Controls.Add(Me.Label11)
        Me.gboHeader.Controls.Add(Me.TxtMonthlyMinValue)
        Me.gboHeader.Controls.Add(Me.Label12)
        Me.gboHeader.Controls.Add(Me.TxtMonthlyMaxValue)
        Me.gboHeader.Controls.Add(Me.Label13)
        Me.gboHeader.Controls.Add(Me.Label9)
        Me.gboHeader.Controls.Add(Me.TxtYearlyMinValue)
        Me.gboHeader.Controls.Add(Me.Label10)
        Me.gboHeader.Controls.Add(Me.TxtYearlyMaxValue)
        Me.gboHeader.Controls.Add(Me.Label8)
        Me.gboHeader.Controls.Add(Me.Label7)
        Me.gboHeader.Location = New System.Drawing.Point(12, 270)
        Me.gboHeader.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.gboHeader.Name = "gboHeader"
        Me.gboHeader.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.gboHeader.Size = New System.Drawing.Size(387, 229)
        Me.gboHeader.TabIndex = 13
        Me.gboHeader.TabStop = False
        Me.gboHeader.Visible = False
        '
        'TxtRoundingUnit
        '
        Me.TxtRoundingUnit.Location = New System.Drawing.Point(168, 15)
        Me.TxtRoundingUnit.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtRoundingUnit.MaxLength = 100
        Me.TxtRoundingUnit.Name = "TxtRoundingUnit"
        Me.TxtRoundingUnit.Size = New System.Drawing.Size(211, 22)
        Me.TxtRoundingUnit.TabIndex = 52
        Me.TxtRoundingUnit.TabStop = False
        Me.TxtRoundingUnit.Text = "1"
        '
        'CboRoundingType
        '
        Me.CboRoundingType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboRoundingType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboRoundingType.FormattingEnabled = True
        Me.CboRoundingType.Items.AddRange(New Object() {"No Rounding", "Round Down", "Round Middlepoint", "Round Up"})
        Me.CboRoundingType.Location = New System.Drawing.Point(168, 49)
        Me.CboRoundingType.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CboRoundingType.Name = "CboRoundingType"
        Me.CboRoundingType.Size = New System.Drawing.Size(211, 24)
        Me.CboRoundingType.TabIndex = 51
        Me.CboRoundingType.TabStop = False
        '
        'DtpYearlyDate
        '
        Me.DtpYearlyDate.Location = New System.Drawing.Point(168, 137)
        Me.DtpYearlyDate.Margin = New System.Windows.Forms.Padding(4)
        Me.DtpYearlyDate.Name = "DtpYearlyDate"
        Me.DtpYearlyDate.Size = New System.Drawing.Size(211, 22)
        Me.DtpYearlyDate.TabIndex = 49
        Me.DtpYearlyDate.TabStop = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(7, 198)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(123, 17)
        Me.Label11.TabIndex = 47
        Me.Label11.Text = "Monthly Min Value"
        '
        'TxtMonthlyMinValue
        '
        Me.TxtMonthlyMinValue.Location = New System.Drawing.Point(168, 199)
        Me.TxtMonthlyMinValue.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMonthlyMinValue.MaxLength = 100
        Me.TxtMonthlyMinValue.Name = "TxtMonthlyMinValue"
        Me.TxtMonthlyMinValue.Size = New System.Drawing.Size(211, 22)
        Me.TxtMonthlyMinValue.TabIndex = 48
        Me.TxtMonthlyMinValue.TabStop = False
        Me.TxtMonthlyMinValue.Text = "0"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(7, 167)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(118, 17)
        Me.Label12.TabIndex = 45
        Me.Label12.Text = "Montly Max Value"
        '
        'TxtMonthlyMaxValue
        '
        Me.TxtMonthlyMaxValue.Location = New System.Drawing.Point(168, 169)
        Me.TxtMonthlyMaxValue.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtMonthlyMaxValue.MaxLength = 100
        Me.TxtMonthlyMaxValue.Name = "TxtMonthlyMaxValue"
        Me.TxtMonthlyMaxValue.Size = New System.Drawing.Size(211, 22)
        Me.TxtMonthlyMaxValue.TabIndex = 46
        Me.TxtMonthlyMaxValue.TabStop = False
        Me.TxtMonthlyMaxValue.Text = "0"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(7, 138)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(82, 17)
        Me.Label13.TabIndex = 43
        Me.Label13.Text = "Yearly Date"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(7, 108)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(114, 17)
        Me.Label9.TabIndex = 41
        Me.Label9.Text = "Yearly Min Value"
        '
        'TxtYearlyMinValue
        '
        Me.TxtYearlyMinValue.Location = New System.Drawing.Point(168, 110)
        Me.TxtYearlyMinValue.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtYearlyMinValue.MaxLength = 100
        Me.TxtYearlyMinValue.Name = "TxtYearlyMinValue"
        Me.TxtYearlyMinValue.Size = New System.Drawing.Size(211, 22)
        Me.TxtYearlyMinValue.TabIndex = 42
        Me.TxtYearlyMinValue.TabStop = False
        Me.TxtYearlyMinValue.Text = "0"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(7, 78)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(117, 17)
        Me.Label10.TabIndex = 39
        Me.Label10.Text = "Yearly Max Value"
        '
        'TxtYearlyMaxValue
        '
        Me.TxtYearlyMaxValue.Location = New System.Drawing.Point(168, 81)
        Me.TxtYearlyMaxValue.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtYearlyMaxValue.MaxLength = 100
        Me.TxtYearlyMaxValue.Name = "TxtYearlyMaxValue"
        Me.TxtYearlyMaxValue.Size = New System.Drawing.Size(211, 22)
        Me.TxtYearlyMaxValue.TabIndex = 40
        Me.TxtYearlyMaxValue.TabStop = False
        Me.TxtYearlyMaxValue.Text = "0"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(7, 48)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(105, 17)
        Me.Label8.TabIndex = 37
        Me.Label8.Text = "Rounding Type"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(7, 18)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(98, 17)
        Me.Label7.TabIndex = 35
        Me.Label7.Text = "Rounding Unit"
        '
        'cboCalculateOn
        '
        Me.cboCalculateOn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboCalculateOn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCalculateOn.FormattingEnabled = True
        Me.cboCalculateOn.Items.AddRange(New Object() {"Jan", "Feb", "Mar", "Apr", "May", "June", "Jul", "Agt", "Sept", "Oct", "Nov", "Dec"})
        Me.cboCalculateOn.Location = New System.Drawing.Point(404, 145)
        Me.cboCalculateOn.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cboCalculateOn.Name = "cboCalculateOn"
        Me.cboCalculateOn.Size = New System.Drawing.Size(345, 24)
        Me.cboCalculateOn.TabIndex = 9
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(8, 148)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(122, 17)
        Me.Label6.TabIndex = 45
        Me.Label6.Text = "Calculation Period"
        '
        'cboCalculatePeriod
        '
        Me.cboCalculatePeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCalculatePeriod.FormattingEnabled = True
        Me.cboCalculatePeriod.Items.AddRange(New Object() {"Daily", "Monthly", "Yearly"})
        Me.cboCalculatePeriod.Location = New System.Drawing.Point(169, 145)
        Me.cboCalculatePeriod.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cboCalculatePeriod.Name = "cboCalculatePeriod"
        Me.cboCalculatePeriod.Size = New System.Drawing.Size(228, 24)
        Me.cboCalculatePeriod.TabIndex = 8
        '
        'BtnEditFormula
        '
        Me.BtnEditFormula.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnEditFormula.Location = New System.Drawing.Point(669, 174)
        Me.BtnEditFormula.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnEditFormula.Name = "BtnEditFormula"
        Me.BtnEditFormula.Size = New System.Drawing.Size(81, 26)
        Me.BtnEditFormula.TabIndex = 11
        Me.BtnEditFormula.Text = "Edit"
        Me.BtnEditFormula.UseVisualStyleBackColor = True
        '
        'TxtFormula
        '
        Me.TxtFormula.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFormula.Location = New System.Drawing.Point(169, 175)
        Me.TxtFormula.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtFormula.MaxLength = 100
        Me.TxtFormula.Name = "TxtFormula"
        Me.TxtFormula.ReadOnly = True
        Me.TxtFormula.Size = New System.Drawing.Size(492, 22)
        Me.TxtFormula.TabIndex = 10
        Me.TxtFormula.TabStop = False
        Me.TxtFormula.Text = "0"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 175)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(59, 17)
        Me.Label5.TabIndex = 42
        Me.Label5.Text = "Formula"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 114)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(119, 17)
        Me.Label4.TabIndex = 37
        Me.Label4.Text = "Referred Account"
        '
        'TxtPrimaryReffered
        '
        Me.TxtPrimaryReffered.Location = New System.Drawing.Point(104, 116)
        Me.TxtPrimaryReffered.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryReffered.Name = "TxtPrimaryReffered"
        Me.TxtPrimaryReffered.Size = New System.Drawing.Size(56, 22)
        Me.TxtPrimaryReffered.TabIndex = 40
        Me.TxtPrimaryReffered.TabStop = False
        Me.TxtPrimaryReffered.Text = "-1"
        Me.TxtPrimaryReffered.Visible = False
        '
        'TxtRefferedCode
        '
        Me.TxtRefferedCode.Location = New System.Drawing.Point(169, 116)
        Me.TxtRefferedCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtRefferedCode.MaxLength = 100
        Me.TxtRefferedCode.Name = "TxtRefferedCode"
        Me.TxtRefferedCode.Size = New System.Drawing.Size(164, 22)
        Me.TxtRefferedCode.TabIndex = 5
        '
        'BtnFindReffered
        '
        Me.BtnFindReffered.Location = New System.Drawing.Point(340, 116)
        Me.BtnFindReffered.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnFindReffered.Name = "BtnFindReffered"
        Me.BtnFindReffered.Size = New System.Drawing.Size(57, 25)
        Me.BtnFindReffered.TabIndex = 6
        Me.BtnFindReffered.TabStop = False
        Me.BtnFindReffered.Text = "Find"
        Me.BtnFindReffered.UseVisualStyleBackColor = True
        '
        'TxtRefferedDesc
        '
        Me.TxtRefferedDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRefferedDesc.Location = New System.Drawing.Point(404, 114)
        Me.TxtRefferedDesc.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtRefferedDesc.MaxLength = 100
        Me.TxtRefferedDesc.Name = "TxtRefferedDesc"
        Me.TxtRefferedDesc.ReadOnly = True
        Me.TxtRefferedDesc.Size = New System.Drawing.Size(345, 22)
        Me.TxtRefferedDesc.TabIndex = 7
        Me.TxtRefferedDesc.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 85)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(150, 17)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "Salary Slip Description"
        '
        'TxtPrimarySlip
        '
        Me.TxtPrimarySlip.Location = New System.Drawing.Point(104, 86)
        Me.TxtPrimarySlip.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimarySlip.Name = "TxtPrimarySlip"
        Me.TxtPrimarySlip.Size = New System.Drawing.Size(56, 22)
        Me.TxtPrimarySlip.TabIndex = 35
        Me.TxtPrimarySlip.TabStop = False
        Me.TxtPrimarySlip.Text = "-1"
        Me.TxtPrimarySlip.Visible = False
        '
        'TxtSlipCode
        '
        Me.TxtSlipCode.Location = New System.Drawing.Point(169, 86)
        Me.TxtSlipCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSlipCode.MaxLength = 100
        Me.TxtSlipCode.Name = "TxtSlipCode"
        Me.TxtSlipCode.Size = New System.Drawing.Size(164, 22)
        Me.TxtSlipCode.TabIndex = 2
        '
        'BtnFindSlip
        '
        Me.BtnFindSlip.Location = New System.Drawing.Point(340, 85)
        Me.BtnFindSlip.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnFindSlip.Name = "BtnFindSlip"
        Me.BtnFindSlip.Size = New System.Drawing.Size(57, 26)
        Me.BtnFindSlip.TabIndex = 3
        Me.BtnFindSlip.TabStop = False
        Me.BtnFindSlip.Text = "Find"
        Me.BtnFindSlip.UseVisualStyleBackColor = True
        '
        'TxtSlipDesc
        '
        Me.TxtSlipDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSlipDesc.Location = New System.Drawing.Point(404, 85)
        Me.TxtSlipDesc.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSlipDesc.MaxLength = 100
        Me.TxtSlipDesc.Name = "TxtSlipDesc"
        Me.TxtSlipDesc.ReadOnly = True
        Me.TxtSlipDesc.Size = New System.Drawing.Size(345, 22)
        Me.TxtSlipDesc.TabIndex = 4
        Me.TxtSlipDesc.TabStop = False
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(249, 505)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(107, 32)
        Me.btnOK.TabIndex = 14
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtDescription
        '
        Me.TxtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescription.Location = New System.Drawing.Point(169, 55)
        Me.TxtDescription.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtDescription.MaxLength = 100
        Me.TxtDescription.Name = "TxtDescription"
        Me.TxtDescription.Size = New System.Drawing.Size(581, 22)
        Me.TxtDescription.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 55)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(79, 17)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Description"
        '
        'TxtCode
        '
        Me.TxtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCode.Location = New System.Drawing.Point(169, 23)
        Me.TxtCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCode.MaxLength = 50
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(581, 22)
        Me.TxtCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 23)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 17)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Code"
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToResizeColumns = False
        Me.DgvItem.AllowUserToResizeRows = False
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColEmployeeCode, Me.ColEmployeeName, Me.ColNIK, Me.ColDept, Me.ColPosition, Me.ColRemarks, Me.ColPrimaryEmployee})
        Me.DgvItem.Location = New System.Drawing.Point(7, 206)
        Me.DgvItem.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.Size = New System.Drawing.Size(743, 293)
        Me.DgvItem.TabIndex = 12
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColEmployeeCode
        '
        Me.ColEmployeeCode.HeaderText = "Code"
        Me.ColEmployeeCode.Name = "ColEmployeeCode"
        Me.ColEmployeeCode.Width = 66
        '
        'ColEmployeeName
        '
        Me.ColEmployeeName.HeaderText = "Name"
        Me.ColEmployeeName.Name = "ColEmployeeName"
        Me.ColEmployeeName.ReadOnly = True
        Me.ColEmployeeName.Width = 70
        '
        'ColNIK
        '
        Me.ColNIK.HeaderText = "NIK"
        Me.ColNIK.Name = "ColNIK"
        Me.ColNIK.ReadOnly = True
        Me.ColNIK.Width = 55
        '
        'ColDept
        '
        Me.ColDept.HeaderText = "Departement"
        Me.ColDept.Name = "ColDept"
        Me.ColDept.ReadOnly = True
        Me.ColDept.Width = 115
        '
        'ColPosition
        '
        Me.ColPosition.HeaderText = "Position"
        Me.ColPosition.Name = "ColPosition"
        Me.ColPosition.ReadOnly = True
        Me.ColPosition.Width = 83
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 89
        '
        'ColPrimaryEmployee
        '
        DataGridViewCellStyle3.NullValue = "-1"
        Me.ColPrimaryEmployee.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColPrimaryEmployee.HeaderText = "Primary Employee"
        Me.ColPrimaryEmployee.Name = "ColPrimaryEmployee"
        Me.ColPrimaryEmployee.Visible = False
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(224, 553)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(107, 32)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'BgwHapus
        '
        '
        'FrmMasterSalaryFormula_OtherDeduction
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(797, 652)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnNew)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(805, 394)
        Me.Name = "FrmMasterSalaryFormula_OtherDeduction"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Other Deduction"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.gboHeader.ResumeLayout(False)
        Me.gboHeader.PerformLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryOther As System.Windows.Forms.TextBox
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimarySlip As System.Windows.Forms.TextBox
    Friend WithEvents TxtSlipCode As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindSlip As System.Windows.Forms.Button
    Friend WithEvents TxtSlipDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryReffered As System.Windows.Forms.TextBox
    Friend WithEvents TxtRefferedCode As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindReffered As System.Windows.Forms.Button
    Friend WithEvents TxtRefferedDesc As System.Windows.Forms.TextBox
    Friend WithEvents BtnEditFormula As System.Windows.Forms.Button
    Friend WithEvents TxtFormula As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboCalculatePeriod As System.Windows.Forms.ComboBox
    Friend WithEvents cboCalculateOn As System.Windows.Forms.ComboBox
    Friend WithEvents gboHeader As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TxtMonthlyMinValue As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtMonthlyMaxValue As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtYearlyMinValue As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtYearlyMaxValue As System.Windows.Forms.TextBox
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BgwHapus As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPosition As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryEmployee As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DtpYearlyDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents CboRoundingType As System.Windows.Forms.ComboBox
    Friend WithEvents TxtRoundingUnit As System.Windows.Forms.TextBox
End Class
