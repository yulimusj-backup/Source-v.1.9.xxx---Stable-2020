<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiLoan
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiLoan))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtTotalAmount = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtRatePersen = New System.Windows.Forms.TextBox
        Me.BtnInformation = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtInstallmentTo = New System.Windows.Forms.TextBox
        Me.CboLoanType = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.BtnFindEmployee = New System.Windows.Forms.Button
        Me.TxtEmployeeName = New System.Windows.Forms.TextBox
        Me.TxtEmployeeCode = New System.Windows.Forms.TextBox
        Me.DtpFirst = New System.Windows.Forms.DateTimePicker
        Me.LblNo = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtPrimaryLoan = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtDocumentNo = New System.Windows.Forms.TextBox
        Me.TxtPrimaryEmployee = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtPrimaryReceivableAcc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCashAcc = New System.Windows.Forms.TextBox
        Me.BtnFindReceivable = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtReceivableAccount = New System.Windows.Forms.TextBox
        Me.CboCashAccount = New System.Windows.Forms.ComboBox
        Me.TxtDescCash = New System.Windows.Forms.TextBox
        Me.TxtDescRec = New System.Windows.Forms.TextBox
        Me.txtJournalDesc = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtJournalNo = New System.Windows.Forms.TextBox
        Me.CboStatus = New System.Windows.Forms.ComboBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColIntallmentPeriod = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDateSchedule = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInterest = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBase = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInstallmentAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPaymentDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPaymentAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TxtPrimaryJurnal = New System.Windows.Forms.TextBox
        Me.BtnPrint = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(444, 535)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(358, 535)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(186, 535)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(272, 535)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 575)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(884, 22)
        Me.StatusStrip1.TabIndex = 14
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(616, 535)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(440, 496)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 27
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(324, 496)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 26
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(72, 53)
        Me.TxtReadOnly.MaxLength = 30
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(29, 20)
        Me.TxtReadOnly.TabIndex = 4
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(126, 55)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(282, 43)
        Me.TxtRemark.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 58)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 65
        Me.Label6.Text = "Remarks"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.SplitContainer2)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryJurnal)
        Me.gboDetail.Location = New System.Drawing.Point(4, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(869, 528)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(12, 11)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label12)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer2.Panel1.Controls.Add(Me.SplitContainer1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.BtnFindEmployee)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtEmployeeName)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtEmployeeCode)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DtpFirst)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtReadOnly)
        Me.SplitContainer2.Panel1.Controls.Add(Me.LblNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimaryLoan)
        Me.SplitContainer2.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtDocumentNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimaryEmployee)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label14)
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtJournalDesc)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRemark)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtJournalNo)
        Me.SplitContainer2.Panel2.Controls.Add(Me.CboStatus)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label13)
        Me.SplitContainer2.Size = New System.Drawing.Size(851, 183)
        Me.SplitContainer2.SplitterDistance = 436
        Me.SplitContainer2.TabIndex = 1
        Me.SplitContainer2.TabStop = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(9, 107)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(84, 13)
        Me.Label8.TabIndex = 82
        Me.Label8.Text = "Employee Name"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(9, 57)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(96, 13)
        Me.Label12.TabIndex = 85
        Me.Label12.Text = "First Payment Date"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 82)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(81, 13)
        Me.Label7.TabIndex = 79
        Me.Label7.Text = "Employee Code"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 127)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtTotalAmount)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRatePersen)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnInformation)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtInstallmentTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.CboLoanType)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Size = New System.Drawing.Size(430, 48)
        Me.SplitContainer1.SplitterDistance = 226
        Me.SplitContainer1.TabIndex = 11
        Me.SplitContainer1.TabStop = False
        '
        'TxtTotalAmount
        '
        Me.TxtTotalAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalAmount.Location = New System.Drawing.Point(104, 0)
        Me.TxtTotalAmount.MaxLength = 30
        Me.TxtTotalAmount.Name = "TxtTotalAmount"
        Me.TxtTotalAmount.Size = New System.Drawing.Size(119, 20)
        Me.TxtTotalAmount.TabIndex = 0
        Me.TxtTotalAmount.Text = "0"
        Me.TxtTotalAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 28)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 13)
        Me.Label4.TabIndex = 73
        Me.Label4.Text = "% Rate Yearly"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 13)
        Me.Label2.TabIndex = 71
        Me.Label2.Text = "Loan Amount"
        '
        'TxtRatePersen
        '
        Me.TxtRatePersen.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRatePersen.Location = New System.Drawing.Point(105, 25)
        Me.TxtRatePersen.MaxLength = 30
        Me.TxtRatePersen.Name = "TxtRatePersen"
        Me.TxtRatePersen.Size = New System.Drawing.Size(118, 20)
        Me.TxtRatePersen.TabIndex = 1
        Me.TxtRatePersen.Text = "0"
        Me.TxtRatePersen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'BtnInformation
        '
        Me.BtnInformation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInformation.BackgroundImage = CType(resources.GetObject("BtnInformation.BackgroundImage"), System.Drawing.Image)
        Me.BtnInformation.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.BtnInformation.Location = New System.Drawing.Point(167, 0)
        Me.BtnInformation.Name = "BtnInformation"
        Me.BtnInformation.Size = New System.Drawing.Size(33, 22)
        Me.BtnInformation.TabIndex = 13
        Me.BtnInformation.TabStop = False
        Me.BtnInformation.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 28)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(73, 13)
        Me.Label5.TabIndex = 76
        Me.Label5.Text = "Installment To"
        '
        'TxtInstallmentTo
        '
        Me.TxtInstallmentTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtInstallmentTo.Location = New System.Drawing.Point(85, 25)
        Me.TxtInstallmentTo.MaxLength = 30
        Me.TxtInstallmentTo.Name = "TxtInstallmentTo"
        Me.TxtInstallmentTo.Size = New System.Drawing.Size(115, 20)
        Me.TxtInstallmentTo.TabIndex = 1
        Me.TxtInstallmentTo.Text = "0"
        Me.TxtInstallmentTo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'CboLoanType
        '
        Me.CboLoanType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboLoanType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboLoanType.FormattingEnabled = True
        Me.CboLoanType.Items.AddRange(New Object() {"No Rate", "Annuity", "Flat", "Effective"})
        Me.CboLoanType.Location = New System.Drawing.Point(85, 1)
        Me.CboLoanType.Name = "CboLoanType"
        Me.CboLoanType.Size = New System.Drawing.Size(80, 21)
        Me.CboLoanType.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 13)
        Me.Label1.TabIndex = 70
        Me.Label1.Text = "Loan Type"
        '
        'BtnFindEmployee
        '
        Me.BtnFindEmployee.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindEmployee.Location = New System.Drawing.Point(391, 76)
        Me.BtnFindEmployee.Name = "BtnFindEmployee"
        Me.BtnFindEmployee.Size = New System.Drawing.Size(42, 23)
        Me.BtnFindEmployee.TabIndex = 9
        Me.BtnFindEmployee.TabStop = False
        Me.BtnFindEmployee.Text = "Find"
        Me.BtnFindEmployee.UseVisualStyleBackColor = True
        '
        'TxtEmployeeName
        '
        Me.TxtEmployeeName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmployeeName.Location = New System.Drawing.Point(107, 104)
        Me.TxtEmployeeName.MaxLength = 30
        Me.TxtEmployeeName.Name = "TxtEmployeeName"
        Me.TxtEmployeeName.ReadOnly = True
        Me.TxtEmployeeName.Size = New System.Drawing.Size(326, 20)
        Me.TxtEmployeeName.TabIndex = 10
        Me.TxtEmployeeName.TabStop = False
        '
        'TxtEmployeeCode
        '
        Me.TxtEmployeeCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmployeeCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtEmployeeCode.Location = New System.Drawing.Point(107, 78)
        Me.TxtEmployeeCode.MaxLength = 150
        Me.TxtEmployeeCode.Name = "TxtEmployeeCode"
        Me.TxtEmployeeCode.Size = New System.Drawing.Size(278, 20)
        Me.TxtEmployeeCode.TabIndex = 8
        '
        'DtpFirst
        '
        Me.DtpFirst.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpFirst.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpFirst.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFirst.Location = New System.Drawing.Point(107, 53)
        Me.DtpFirst.Name = "DtpFirst"
        Me.DtpFirst.Size = New System.Drawing.Size(326, 20)
        Me.DtpFirst.TabIndex = 7
        '
        'LblNo
        '
        Me.LblNo.AutoSize = True
        Me.LblNo.Location = New System.Drawing.Point(9, 7)
        Me.LblNo.Name = "LblNo"
        Me.LblNo.Size = New System.Drawing.Size(73, 13)
        Me.LblNo.TabIndex = 45
        Me.LblNo.Text = "Document No"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 31)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Date"
        '
        'TxtPrimaryLoan
        '
        Me.TxtPrimaryLoan.Location = New System.Drawing.Point(107, 3)
        Me.TxtPrimaryLoan.Name = "TxtPrimaryLoan"
        Me.TxtPrimaryLoan.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryLoan.TabIndex = 0
        Me.TxtPrimaryLoan.TabStop = False
        Me.TxtPrimaryLoan.Text = "-1"
        Me.TxtPrimaryLoan.Visible = False
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(107, 27)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(326, 20)
        Me.dtpDate.TabIndex = 6
        '
        'TxtDocumentNo
        '
        Me.TxtDocumentNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDocumentNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDocumentNo.Location = New System.Drawing.Point(107, 3)
        Me.TxtDocumentNo.Name = "TxtDocumentNo"
        Me.TxtDocumentNo.ReadOnly = True
        Me.TxtDocumentNo.Size = New System.Drawing.Size(326, 20)
        Me.TxtDocumentNo.TabIndex = 2
        Me.TxtDocumentNo.TabStop = False
        '
        'TxtPrimaryEmployee
        '
        Me.TxtPrimaryEmployee.Location = New System.Drawing.Point(80, 80)
        Me.TxtPrimaryEmployee.Name = "TxtPrimaryEmployee"
        Me.TxtPrimaryEmployee.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryEmployee.TabIndex = 83
        Me.TxtPrimaryEmployee.TabStop = False
        Me.TxtPrimaryEmployee.Text = "-1"
        Me.TxtPrimaryEmployee.Visible = False
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(9, 157)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(69, 13)
        Me.Label14.TabIndex = 89
        Me.Label14.Text = "Journal Desc"
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer4.Location = New System.Drawing.Point(6, 2)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtPrimaryReceivableAcc)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtPrimaryCashAcc)
        Me.SplitContainer4.Panel1.Controls.Add(Me.BtnFindReceivable)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtReceivableAccount)
        Me.SplitContainer4.Panel1.Controls.Add(Me.CboCashAccount)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtDescCash)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtDescRec)
        Me.SplitContainer4.Size = New System.Drawing.Size(402, 50)
        Me.SplitContainer4.SplitterDistance = 234
        Me.SplitContainer4.TabIndex = 0
        Me.SplitContainer4.TabStop = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(3, 6)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(94, 13)
        Me.Label10.TabIndex = 88
        Me.Label10.Text = "Cash Out Account"
        '
        'TxtPrimaryReceivableAcc
        '
        Me.TxtPrimaryReceivableAcc.Location = New System.Drawing.Point(119, 27)
        Me.TxtPrimaryReceivableAcc.Name = "TxtPrimaryReceivableAcc"
        Me.TxtPrimaryReceivableAcc.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryReceivableAcc.TabIndex = 19
        Me.TxtPrimaryReceivableAcc.TabStop = False
        Me.TxtPrimaryReceivableAcc.Text = "-1"
        Me.TxtPrimaryReceivableAcc.Visible = False
        '
        'TxtPrimaryCashAcc
        '
        Me.TxtPrimaryCashAcc.Location = New System.Drawing.Point(119, 3)
        Me.TxtPrimaryCashAcc.Name = "TxtPrimaryCashAcc"
        Me.TxtPrimaryCashAcc.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCashAcc.TabIndex = 16
        Me.TxtPrimaryCashAcc.TabStop = False
        Me.TxtPrimaryCashAcc.Text = "-1"
        Me.TxtPrimaryCashAcc.Visible = False
        '
        'BtnFindReceivable
        '
        Me.BtnFindReceivable.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindReceivable.Location = New System.Drawing.Point(189, 25)
        Me.BtnFindReceivable.Name = "BtnFindReceivable"
        Me.BtnFindReceivable.Size = New System.Drawing.Size(42, 23)
        Me.BtnFindReceivable.TabIndex = 1
        Me.BtnFindReceivable.TabStop = False
        Me.BtnFindReceivable.Text = "Find"
        Me.BtnFindReceivable.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 27)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(104, 13)
        Me.Label11.TabIndex = 93
        Me.Label11.Text = "Receivable Account"
        '
        'TxtReceivableAccount
        '
        Me.TxtReceivableAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReceivableAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtReceivableAccount.Location = New System.Drawing.Point(120, 27)
        Me.TxtReceivableAccount.MaxLength = 150
        Me.TxtReceivableAccount.Name = "TxtReceivableAccount"
        Me.TxtReceivableAccount.Size = New System.Drawing.Size(64, 20)
        Me.TxtReceivableAccount.TabIndex = 20
        '
        'CboCashAccount
        '
        Me.CboCashAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboCashAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboCashAccount.FormattingEnabled = True
        Me.CboCashAccount.Location = New System.Drawing.Point(119, 3)
        Me.CboCashAccount.Name = "CboCashAccount"
        Me.CboCashAccount.Size = New System.Drawing.Size(112, 21)
        Me.CboCashAccount.TabIndex = 0
        '
        'TxtDescCash
        '
        Me.TxtDescCash.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescCash.Location = New System.Drawing.Point(7, 0)
        Me.TxtDescCash.MaxLength = 30
        Me.TxtDescCash.Name = "TxtDescCash"
        Me.TxtDescCash.ReadOnly = True
        Me.TxtDescCash.Size = New System.Drawing.Size(157, 20)
        Me.TxtDescCash.TabIndex = 0
        Me.TxtDescCash.TabStop = False
        '
        'TxtDescRec
        '
        Me.TxtDescRec.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescRec.Location = New System.Drawing.Point(7, 26)
        Me.TxtDescRec.MaxLength = 30
        Me.TxtDescRec.Name = "TxtDescRec"
        Me.TxtDescRec.ReadOnly = True
        Me.TxtDescRec.Size = New System.Drawing.Size(157, 20)
        Me.TxtDescRec.TabIndex = 1
        Me.TxtDescRec.TabStop = False
        '
        'txtJournalDesc
        '
        Me.txtJournalDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtJournalDesc.Location = New System.Drawing.Point(125, 154)
        Me.txtJournalDesc.MaxLength = 150
        Me.txtJournalDesc.Name = "txtJournalDesc"
        Me.txtJournalDesc.Size = New System.Drawing.Size(283, 20)
        Me.txtJournalDesc.TabIndex = 4
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(9, 106)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(37, 13)
        Me.Label9.TabIndex = 85
        Me.Label9.Text = "Status"
        '
        'txtJournalNo
        '
        Me.txtJournalNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtJournalNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtJournalNo.Location = New System.Drawing.Point(125, 129)
        Me.txtJournalNo.MaxLength = 150
        Me.txtJournalNo.Name = "txtJournalNo"
        Me.txtJournalNo.Size = New System.Drawing.Size(283, 20)
        Me.txtJournalNo.TabIndex = 3
        '
        'CboStatus
        '
        Me.CboStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboStatus.Enabled = False
        Me.CboStatus.FormattingEnabled = True
        Me.CboStatus.Items.AddRange(New Object() {"On Going", "Finished"})
        Me.CboStatus.Location = New System.Drawing.Point(126, 102)
        Me.CboStatus.Name = "CboStatus"
        Me.CboStatus.Size = New System.Drawing.Size(282, 21)
        Me.CboStatus.TabIndex = 2
        Me.CboStatus.TabStop = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(9, 132)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(58, 13)
        Me.Label13.TabIndex = 87
        Me.Label13.Text = "Journal No"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColIntallmentPeriod, Me.ColDateSchedule, Me.ColInterest, Me.ColBase, Me.ColInstallmentAmount, Me.ColBalance, Me.ColPaymentDate, Me.ColPaymentAmount, Me.ColRemarks, Me.ColIdentifier})
        Me.dgvItem.Location = New System.Drawing.Point(12, 200)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.Size = New System.Drawing.Size(851, 290)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 25
        Me.dgvItem.TabStop = False
        '
        'ColIntallmentPeriod
        '
        Me.ColIntallmentPeriod.Frozen = True
        Me.ColIntallmentPeriod.HeaderText = "Installment Period"
        Me.ColIntallmentPeriod.Name = "ColIntallmentPeriod"
        Me.ColIntallmentPeriod.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColIntallmentPeriod.Width = 96
        '
        'ColDateSchedule
        '
        DataGridViewCellStyle2.Format = "dd MMM yyyy"
        Me.ColDateSchedule.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColDateSchedule.Frozen = True
        Me.ColDateSchedule.HeaderText = "Date Schedule"
        Me.ColDateSchedule.Name = "ColDateSchedule"
        Me.ColDateSchedule.Width = 103
        '
        'ColInterest
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.NullValue = "0"
        Me.ColInterest.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColInterest.Frozen = True
        Me.ColInterest.HeaderText = "Interest"
        Me.ColInterest.Name = "ColInterest"
        Me.ColInterest.Width = 67
        '
        'ColBase
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.NullValue = "0"
        Me.ColBase.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColBase.Frozen = True
        Me.ColBase.HeaderText = "Base"
        Me.ColBase.Name = "ColBase"
        Me.ColBase.Width = 56
        '
        'ColInstallmentAmount
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.NullValue = "0"
        Me.ColInstallmentAmount.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColInstallmentAmount.Frozen = True
        Me.ColInstallmentAmount.HeaderText = "Installment Amount"
        Me.ColInstallmentAmount.Name = "ColInstallmentAmount"
        Me.ColInstallmentAmount.Width = 121
        '
        'ColBalance
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle6.NullValue = "0"
        Me.ColBalance.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColBalance.Frozen = True
        Me.ColBalance.HeaderText = "Balance"
        Me.ColBalance.Name = "ColBalance"
        Me.ColBalance.Width = 71
        '
        'ColPaymentDate
        '
        DataGridViewCellStyle7.Format = "mm DDD yyyy"
        DataGridViewCellStyle7.NullValue = "-"
        Me.ColPaymentDate.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColPaymentDate.Frozen = True
        Me.ColPaymentDate.HeaderText = "Payment Date"
        Me.ColPaymentDate.Name = "ColPaymentDate"
        Me.ColPaymentDate.Width = 99
        '
        'ColPaymentAmount
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle8.NullValue = "0"
        Me.ColPaymentAmount.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColPaymentAmount.Frozen = True
        Me.ColPaymentAmount.HeaderText = "Payment Amount"
        Me.ColPaymentAmount.Name = "ColPaymentAmount"
        Me.ColPaymentAmount.Width = 112
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.Visible = False
        Me.ColIdentifier.Width = 72
        '
        'TxtPrimaryJurnal
        '
        Me.TxtPrimaryJurnal.Location = New System.Drawing.Point(385, 281)
        Me.TxtPrimaryJurnal.Name = "TxtPrimaryJurnal"
        Me.TxtPrimaryJurnal.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryJurnal.TabIndex = 1
        Me.TxtPrimaryJurnal.TabStop = False
        Me.TxtPrimaryJurnal.Text = "-1"
        Me.TxtPrimaryJurnal.Visible = False
        '
        'BtnPrint
        '
        Me.BtnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnPrint.Location = New System.Drawing.Point(530, 535)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(80, 26)
        Me.BtnPrint.TabIndex = 4
        Me.BtnPrint.Text = "&Print"
        Me.BtnPrint.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'FrmTransaksiLoan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(884, 597)
        Me.Controls.Add(Me.BtnPrint)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.gboDetail)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 628)
        Me.Name = "FrmTransaksiLoan"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Loan Data"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.Panel2.PerformLayout()
        Me.SplitContainer4.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TxtPrimaryLoan As System.Windows.Forms.TextBox
    Friend WithEvents TxtDocumentNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LblNo As System.Windows.Forms.Label
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Public WithEvents btnNew As System.Windows.Forms.Button
    Public WithEvents btnOK As System.Windows.Forms.Button
    Public WithEvents TxtRemark As System.Windows.Forms.TextBox
    Public WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CboLoanType As System.Windows.Forms.ComboBox
    Friend WithEvents TxtTotalAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtRatePersen As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryEmployee As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindEmployee As System.Windows.Forms.Button
    Public WithEvents TxtEmployeeCode As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtInstallmentTo As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents CboStatus As System.Windows.Forms.ComboBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtPrimaryCashAcc As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescCash As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryReceivableAcc As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescRec As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindReceivable As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents TxtReceivableAccount As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents DtpFirst As System.Windows.Forms.DateTimePicker
    Friend WithEvents BtnInformation As System.Windows.Forms.Button
    Friend WithEvents CboCashAccount As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPrimaryJurnal As System.Windows.Forms.TextBox
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
    Friend WithEvents ColIntallmentPeriod As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDateSchedule As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInterest As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBase As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInstallmentAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPaymentDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPaymentAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents txtJournalNo As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents txtJournalDesc As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
