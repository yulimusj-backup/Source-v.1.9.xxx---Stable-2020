<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiLoan_Payment
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiLoan_Payment))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
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
        Me.Label6 = New System.Windows.Forms.Label
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.GboItem = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtPrimaryRev = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtInterest = New System.Windows.Forms.TextBox
        Me.TxtBase = New System.Windows.Forms.TextBox
        Me.TxtPrimaryLoan = New System.Windows.Forms.TextBox
        Me.TxtInstallmentAt = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtIdentifier = New System.Windows.Forms.TextBox
        Me.CboLoanNumber = New System.Windows.Forms.ComboBox
        Me.TxtRemarksItem = New System.Windows.Forms.TextBox
        Me.BtnInput = New System.Windows.Forms.Button
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.TxtPayment = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.TxtTotalInstallment = New System.Windows.Forms.TextBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.TxtEmployeeCode1 = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtDescCash = New System.Windows.Forms.TextBox
        Me.TxtPrimarySalary = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCashAcc = New System.Windows.Forms.TextBox
        Me.LblNo = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtPrimaryLoanPayment = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtEmployeeName = New System.Windows.Forms.TextBox
        Me.CboCashAccount = New System.Windows.Forms.ComboBox
        Me.BtnFindEmployee = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtDocumentNo = New System.Windows.Forms.TextBox
        Me.TxtPrimaryEmployee = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.TxtPrimaryLoss = New System.Windows.Forms.TextBox
        Me.TxtPrimaryGainAcc = New System.Windows.Forms.TextBox
        Me.BtnFindLoss = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtLossAccount = New System.Windows.Forms.TextBox
        Me.BtnFindGain = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtGainAccount = New System.Windows.Forms.TextBox
        Me.TxtDescLoss = New System.Windows.Forms.TextBox
        Me.TxtDescGain = New System.Windows.Forms.TextBox
        Me.TxtJournalDesc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryJurnal = New System.Windows.Forms.TextBox
        Me.txtJournalNo = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLoanNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInstallementAt = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBase = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInterest = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalInstallment = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPayment = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLoan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryRev = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BtnPrint = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.GboItem.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
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
        Me.btnDelete.Location = New System.Drawing.Point(448, 535)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(352, 535)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(160, 535)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(256, 535)
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
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 570)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(884, 22)
        Me.StatusStrip1.TabIndex = 14
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(620, 535)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(444, 496)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(328, 496)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(72, 26)
        Me.TxtReadOnly.MaxLength = 30
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(29, 20)
        Me.TxtReadOnly.TabIndex = 4
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 56)
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
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.GboItem)
        Me.gboDetail.Controls.Add(Me.SplitContainer2)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(4, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(877, 528)
        Me.gboDetail.TabIndex = 6
        Me.gboDetail.TabStop = False
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = CType(resources.GetObject("btnUP.BackgroundImage"), System.Drawing.Image)
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(12, 496)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 6
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = CType(resources.GetObject("btnDown.BackgroundImage"), System.Drawing.Image)
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(46, 496)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 7
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'GboItem
        '
        Me.GboItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GboItem.Controls.Add(Me.SplitContainer1)
        Me.GboItem.Location = New System.Drawing.Point(13, 140)
        Me.GboItem.Name = "GboItem"
        Me.GboItem.Size = New System.Drawing.Size(858, 122)
        Me.GboItem.TabIndex = 1
        Me.GboItem.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryRev)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtInterest)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtBase)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryLoan)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtInstallmentAt)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtIdentifier)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboLoanNumber)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarksItem)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnInput)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label14)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label13)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPayment)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label12)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtTotalInstallment)
        Me.SplitContainer1.Size = New System.Drawing.Size(852, 103)
        Me.SplitContainer1.SplitterDistance = 439
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'TxtPrimaryRev
        '
        Me.TxtPrimaryRev.Location = New System.Drawing.Point(72, 51)
        Me.TxtPrimaryRev.Name = "TxtPrimaryRev"
        Me.TxtPrimaryRev.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryRev.TabIndex = 90
        Me.TxtPrimaryRev.TabStop = False
        Me.TxtPrimaryRev.Text = "-1"
        Me.TxtPrimaryRev.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 59)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(31, 13)
        Me.Label5.TabIndex = 87
        Me.Label5.Text = "Base"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(5, 80)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(42, 13)
        Me.Label9.TabIndex = 89
        Me.Label9.Text = "Interest"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 13)
        Me.Label1.TabIndex = 70
        Me.Label1.Text = "Loan Number"
        '
        'TxtInterest
        '
        Me.TxtInterest.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtInterest.Location = New System.Drawing.Point(101, 80)
        Me.TxtInterest.MaxLength = 30
        Me.TxtInterest.Name = "TxtInterest"
        Me.TxtInterest.ReadOnly = True
        Me.TxtInterest.Size = New System.Drawing.Size(333, 20)
        Me.TxtInterest.TabIndex = 3
        Me.TxtInterest.TabStop = False
        Me.TxtInterest.Text = "0"
        Me.TxtInterest.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtBase
        '
        Me.TxtBase.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBase.Location = New System.Drawing.Point(101, 56)
        Me.TxtBase.MaxLength = 30
        Me.TxtBase.Name = "TxtBase"
        Me.TxtBase.ReadOnly = True
        Me.TxtBase.Size = New System.Drawing.Size(333, 20)
        Me.TxtBase.TabIndex = 2
        Me.TxtBase.TabStop = False
        Me.TxtBase.Text = "0"
        Me.TxtBase.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtPrimaryLoan
        '
        Me.TxtPrimaryLoan.Location = New System.Drawing.Point(72, 3)
        Me.TxtPrimaryLoan.Name = "TxtPrimaryLoan"
        Me.TxtPrimaryLoan.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryLoan.TabIndex = 84
        Me.TxtPrimaryLoan.TabStop = False
        Me.TxtPrimaryLoan.Text = "-1"
        Me.TxtPrimaryLoan.Visible = False
        '
        'TxtInstallmentAt
        '
        Me.TxtInstallmentAt.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtInstallmentAt.Location = New System.Drawing.Point(101, 30)
        Me.TxtInstallmentAt.MaxLength = 30
        Me.TxtInstallmentAt.Name = "TxtInstallmentAt"
        Me.TxtInstallmentAt.ReadOnly = True
        Me.TxtInstallmentAt.Size = New System.Drawing.Size(333, 20)
        Me.TxtInstallmentAt.TabIndex = 1
        Me.TxtInstallmentAt.TabStop = False
        Me.TxtInstallmentAt.Text = "0"
        Me.TxtInstallmentAt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 33)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 13)
        Me.Label4.TabIndex = 85
        Me.Label4.Text = "Installment At"
        '
        'TxtIdentifier
        '
        Me.TxtIdentifier.Location = New System.Drawing.Point(72, 27)
        Me.TxtIdentifier.Name = "TxtIdentifier"
        Me.TxtIdentifier.Size = New System.Drawing.Size(25, 20)
        Me.TxtIdentifier.TabIndex = 85
        Me.TxtIdentifier.TabStop = False
        Me.TxtIdentifier.Visible = False
        '
        'CboLoanNumber
        '
        Me.CboLoanNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboLoanNumber.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboLoanNumber.FormattingEnabled = True
        Me.CboLoanNumber.Location = New System.Drawing.Point(101, 3)
        Me.CboLoanNumber.Name = "CboLoanNumber"
        Me.CboLoanNumber.Size = New System.Drawing.Size(333, 21)
        Me.CboLoanNumber.TabIndex = 0
        '
        'TxtRemarksItem
        '
        Me.TxtRemarksItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarksItem.Location = New System.Drawing.Point(105, 55)
        Me.TxtRemarksItem.MaxLength = 100
        Me.TxtRemarksItem.Name = "TxtRemarksItem"
        Me.TxtRemarksItem.Size = New System.Drawing.Size(301, 20)
        Me.TxtRemarksItem.TabIndex = 2
        '
        'BtnInput
        '
        Me.BtnInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInput.Location = New System.Drawing.Point(333, 80)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(73, 23)
        Me.BtnInput.TabIndex = 3
        Me.BtnInput.Text = "Input"
        Me.BtnInput.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(14, 58)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(49, 13)
        Me.Label14.TabIndex = 73
        Me.Label14.Text = "Remarks"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(14, 34)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(48, 13)
        Me.Label13.TabIndex = 93
        Me.Label13.Text = "Payment"
        '
        'TxtPayment
        '
        Me.TxtPayment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPayment.Location = New System.Drawing.Point(105, 29)
        Me.TxtPayment.MaxLength = 30
        Me.TxtPayment.Name = "TxtPayment"
        Me.TxtPayment.Size = New System.Drawing.Size(301, 20)
        Me.TxtPayment.TabIndex = 1
        Me.TxtPayment.Text = "0"
        Me.TxtPayment.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(14, 8)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(84, 13)
        Me.Label12.TabIndex = 91
        Me.Label12.Text = "Total Installment"
        '
        'TxtTotalInstallment
        '
        Me.TxtTotalInstallment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalInstallment.Location = New System.Drawing.Point(105, 3)
        Me.TxtTotalInstallment.MaxLength = 30
        Me.TxtTotalInstallment.Name = "TxtTotalInstallment"
        Me.TxtTotalInstallment.ReadOnly = True
        Me.TxtTotalInstallment.Size = New System.Drawing.Size(301, 20)
        Me.TxtTotalInstallment.TabIndex = 0
        Me.TxtTotalInstallment.TabStop = False
        Me.TxtTotalInstallment.Text = "0"
        Me.TxtTotalInstallment.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtEmployeeCode1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtDescCash)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimarySalary)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtReadOnly)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimaryCashAcc)
        Me.SplitContainer2.Panel1.Controls.Add(Me.LblNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimaryLoanPayment)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer2.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtEmployeeName)
        Me.SplitContainer2.Panel1.Controls.Add(Me.CboCashAccount)
        Me.SplitContainer2.Panel1.Controls.Add(Me.BtnFindEmployee)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtDocumentNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimaryEmployee)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label16)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRemark)
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtJournalDesc)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtPrimaryJurnal)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtJournalNo)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label15)
        Me.SplitContainer2.Size = New System.Drawing.Size(859, 127)
        Me.SplitContainer2.SplitterDistance = 441
        Me.SplitContainer2.TabIndex = 0
        Me.SplitContainer2.TabStop = False
        '
        'TxtEmployeeCode1
        '
        Me.TxtEmployeeCode1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmployeeCode1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtEmployeeCode1.Location = New System.Drawing.Point(107, 52)
        Me.TxtEmployeeCode1.MaxLength = 150
        Me.TxtEmployeeCode1.Name = "TxtEmployeeCode1"
        Me.TxtEmployeeCode1.Size = New System.Drawing.Size(285, 20)
        Me.TxtEmployeeCode1.TabIndex = 4
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(9, 105)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(86, 13)
        Me.Label10.TabIndex = 88
        Me.Label10.Text = "Cash In Account"
        '
        'TxtDescCash
        '
        Me.TxtDescCash.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescCash.Location = New System.Drawing.Point(239, 102)
        Me.TxtDescCash.MaxLength = 30
        Me.TxtDescCash.Name = "TxtDescCash"
        Me.TxtDescCash.ReadOnly = True
        Me.TxtDescCash.Size = New System.Drawing.Size(199, 20)
        Me.TxtDescCash.TabIndex = 0
        Me.TxtDescCash.TabStop = False
        '
        'TxtPrimarySalary
        '
        Me.TxtPrimarySalary.Location = New System.Drawing.Point(138, 4)
        Me.TxtPrimarySalary.Name = "TxtPrimarySalary"
        Me.TxtPrimarySalary.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimarySalary.TabIndex = 48
        Me.TxtPrimarySalary.TabStop = False
        Me.TxtPrimarySalary.Text = "-1"
        Me.TxtPrimarySalary.Visible = False
        '
        'TxtPrimaryCashAcc
        '
        Me.TxtPrimaryCashAcc.Location = New System.Drawing.Point(107, 105)
        Me.TxtPrimaryCashAcc.Name = "TxtPrimaryCashAcc"
        Me.TxtPrimaryCashAcc.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCashAcc.TabIndex = 0
        Me.TxtPrimaryCashAcc.TabStop = False
        Me.TxtPrimaryCashAcc.Text = "-1"
        Me.TxtPrimaryCashAcc.Visible = False
        '
        'LblNo
        '
        Me.LblNo.AutoSize = True
        Me.LblNo.Location = New System.Drawing.Point(9, 3)
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
        'TxtPrimaryLoanPayment
        '
        Me.TxtPrimaryLoanPayment.Location = New System.Drawing.Point(107, 3)
        Me.TxtPrimaryLoanPayment.Name = "TxtPrimaryLoanPayment"
        Me.TxtPrimaryLoanPayment.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryLoanPayment.TabIndex = 47
        Me.TxtPrimaryLoanPayment.TabStop = False
        Me.TxtPrimaryLoanPayment.Text = "-1"
        Me.TxtPrimaryLoanPayment.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(9, 80)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(84, 13)
        Me.Label8.TabIndex = 82
        Me.Label8.Text = "Employee Name"
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(105, 27)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(333, 20)
        Me.dtpDate.TabIndex = 3
        '
        'TxtEmployeeName
        '
        Me.TxtEmployeeName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmployeeName.Location = New System.Drawing.Point(107, 77)
        Me.TxtEmployeeName.MaxLength = 30
        Me.TxtEmployeeName.Name = "TxtEmployeeName"
        Me.TxtEmployeeName.ReadOnly = True
        Me.TxtEmployeeName.Size = New System.Drawing.Size(331, 20)
        Me.TxtEmployeeName.TabIndex = 6
        Me.TxtEmployeeName.TabStop = False
        '
        'CboCashAccount
        '
        Me.CboCashAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboCashAccount.FormattingEnabled = True
        Me.CboCashAccount.Location = New System.Drawing.Point(107, 102)
        Me.CboCashAccount.Name = "CboCashAccount"
        Me.CboCashAccount.Size = New System.Drawing.Size(126, 21)
        Me.CboCashAccount.TabIndex = 7
        '
        'BtnFindEmployee
        '
        Me.BtnFindEmployee.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindEmployee.Location = New System.Drawing.Point(396, 51)
        Me.BtnFindEmployee.Name = "BtnFindEmployee"
        Me.BtnFindEmployee.Size = New System.Drawing.Size(42, 23)
        Me.BtnFindEmployee.TabIndex = 5
        Me.BtnFindEmployee.TabStop = False
        Me.BtnFindEmployee.Text = "Find"
        Me.BtnFindEmployee.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 54)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(81, 13)
        Me.Label7.TabIndex = 79
        Me.Label7.Text = "Employee Code"
        '
        'TxtDocumentNo
        '
        Me.TxtDocumentNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDocumentNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDocumentNo.Location = New System.Drawing.Point(105, 3)
        Me.TxtDocumentNo.Name = "TxtDocumentNo"
        Me.TxtDocumentNo.ReadOnly = True
        Me.TxtDocumentNo.Size = New System.Drawing.Size(333, 20)
        Me.TxtDocumentNo.TabIndex = 0
        Me.TxtDocumentNo.TabStop = False
        '
        'TxtPrimaryEmployee
        '
        Me.TxtPrimaryEmployee.Location = New System.Drawing.Point(76, 52)
        Me.TxtPrimaryEmployee.Name = "TxtPrimaryEmployee"
        Me.TxtPrimaryEmployee.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryEmployee.TabIndex = 83
        Me.TxtPrimaryEmployee.TabStop = False
        Me.TxtPrimaryEmployee.Text = "-1"
        Me.TxtPrimaryEmployee.Visible = False
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(9, 105)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(69, 13)
        Me.Label16.TabIndex = 88
        Me.Label16.Text = "Journal Desc"
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(107, 53)
        Me.TxtRemark.MaxLength = 100
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(301, 19)
        Me.TxtRemark.TabIndex = 1
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer4.Location = New System.Drawing.Point(6, 3)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtPrimaryLoss)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtPrimaryGainAcc)
        Me.SplitContainer4.Panel1.Controls.Add(Me.BtnFindLoss)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtLossAccount)
        Me.SplitContainer4.Panel1.Controls.Add(Me.BtnFindGain)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtGainAccount)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtDescLoss)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtDescGain)
        Me.SplitContainer4.Size = New System.Drawing.Size(405, 50)
        Me.SplitContainer4.SplitterDistance = 247
        Me.SplitContainer4.TabIndex = 0
        Me.SplitContainer4.TabStop = False
        '
        'TxtPrimaryLoss
        '
        Me.TxtPrimaryLoss.Location = New System.Drawing.Point(119, 26)
        Me.TxtPrimaryLoss.Name = "TxtPrimaryLoss"
        Me.TxtPrimaryLoss.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryLoss.TabIndex = 2
        Me.TxtPrimaryLoss.TabStop = False
        Me.TxtPrimaryLoss.Text = "-1"
        Me.TxtPrimaryLoss.Visible = False
        '
        'TxtPrimaryGainAcc
        '
        Me.TxtPrimaryGainAcc.Location = New System.Drawing.Point(119, 2)
        Me.TxtPrimaryGainAcc.Name = "TxtPrimaryGainAcc"
        Me.TxtPrimaryGainAcc.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryGainAcc.TabIndex = 1
        Me.TxtPrimaryGainAcc.TabStop = False
        Me.TxtPrimaryGainAcc.Text = "-1"
        Me.TxtPrimaryGainAcc.Visible = False
        '
        'BtnFindLoss
        '
        Me.BtnFindLoss.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindLoss.Location = New System.Drawing.Point(203, 23)
        Me.BtnFindLoss.Name = "BtnFindLoss"
        Me.BtnFindLoss.Size = New System.Drawing.Size(42, 23)
        Me.BtnFindLoss.TabIndex = 98
        Me.BtnFindLoss.TabStop = False
        Me.BtnFindLoss.Text = "Find"
        Me.BtnFindLoss.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 13)
        Me.Label2.TabIndex = 97
        Me.Label2.Text = "Loss Account"
        '
        'TxtLossAccount
        '
        Me.TxtLossAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLossAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtLossAccount.Location = New System.Drawing.Point(101, 25)
        Me.TxtLossAccount.MaxLength = 150
        Me.TxtLossAccount.Name = "TxtLossAccount"
        Me.TxtLossAccount.Size = New System.Drawing.Size(96, 20)
        Me.TxtLossAccount.TabIndex = 2
        '
        'BtnFindGain
        '
        Me.BtnFindGain.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindGain.Location = New System.Drawing.Point(203, -1)
        Me.BtnFindGain.Name = "BtnFindGain"
        Me.BtnFindGain.Size = New System.Drawing.Size(42, 23)
        Me.BtnFindGain.TabIndex = 94
        Me.BtnFindGain.TabStop = False
        Me.BtnFindGain.Text = "Find"
        Me.BtnFindGain.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 6)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(72, 13)
        Me.Label11.TabIndex = 93
        Me.Label11.Text = "Gain Account"
        '
        'TxtGainAccount
        '
        Me.TxtGainAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtGainAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtGainAccount.Location = New System.Drawing.Point(101, 0)
        Me.TxtGainAccount.MaxLength = 150
        Me.TxtGainAccount.Name = "TxtGainAccount"
        Me.TxtGainAccount.Size = New System.Drawing.Size(96, 20)
        Me.TxtGainAccount.TabIndex = 1
        '
        'TxtDescLoss
        '
        Me.TxtDescLoss.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescLoss.Location = New System.Drawing.Point(3, 25)
        Me.TxtDescLoss.MaxLength = 30
        Me.TxtDescLoss.Name = "TxtDescLoss"
        Me.TxtDescLoss.ReadOnly = True
        Me.TxtDescLoss.Size = New System.Drawing.Size(148, 20)
        Me.TxtDescLoss.TabIndex = 2
        Me.TxtDescLoss.TabStop = False
        '
        'TxtDescGain
        '
        Me.TxtDescGain.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescGain.Location = New System.Drawing.Point(3, 1)
        Me.TxtDescGain.MaxLength = 30
        Me.TxtDescGain.Name = "TxtDescGain"
        Me.TxtDescGain.ReadOnly = True
        Me.TxtDescGain.Size = New System.Drawing.Size(148, 20)
        Me.TxtDescGain.TabIndex = 1
        Me.TxtDescGain.TabStop = False
        '
        'TxtJournalDesc
        '
        Me.TxtJournalDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtJournalDesc.Location = New System.Drawing.Point(107, 101)
        Me.TxtJournalDesc.MaxLength = 150
        Me.TxtJournalDesc.Name = "TxtJournalDesc"
        Me.TxtJournalDesc.Size = New System.Drawing.Size(301, 20)
        Me.TxtJournalDesc.TabIndex = 3
        '
        'TxtPrimaryJurnal
        '
        Me.TxtPrimaryJurnal.Location = New System.Drawing.Point(76, 76)
        Me.TxtPrimaryJurnal.Name = "TxtPrimaryJurnal"
        Me.TxtPrimaryJurnal.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryJurnal.TabIndex = 84
        Me.TxtPrimaryJurnal.TabStop = False
        Me.TxtPrimaryJurnal.Text = "-1"
        Me.TxtPrimaryJurnal.Visible = False
        '
        'txtJournalNo
        '
        Me.txtJournalNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtJournalNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtJournalNo.Location = New System.Drawing.Point(107, 76)
        Me.txtJournalNo.MaxLength = 150
        Me.txtJournalNo.Name = "txtJournalNo"
        Me.txtJournalNo.Size = New System.Drawing.Size(301, 20)
        Me.txtJournalNo.TabIndex = 2
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(9, 80)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(58, 13)
        Me.Label15.TabIndex = 86
        Me.Label15.Text = "Journal No"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColLoanNo, Me.ColInstallementAt, Me.ColBase, Me.ColInterest, Me.ColTotalInstallment, Me.ColPayment, Me.ColRemarks, Me.ColPrimaryLoan, Me.ColIdentifier, Me.ColPrimaryRev})
        Me.dgvItem.Location = New System.Drawing.Point(12, 268)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.Size = New System.Drawing.Size(859, 222)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 2
        Me.dgvItem.TabStop = False
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 46
        '
        'ColLoanNo
        '
        Me.ColLoanNo.Frozen = True
        Me.ColLoanNo.HeaderText = "Loan Number"
        Me.ColLoanNo.Name = "ColLoanNo"
        Me.ColLoanNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLoanNo.Width = 77
        '
        'ColInstallementAt
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle2.NullValue = "0"
        Me.ColInstallementAt.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColInstallementAt.Frozen = True
        Me.ColInstallementAt.HeaderText = "Installment At"
        Me.ColInstallementAt.Name = "ColInstallementAt"
        Me.ColInstallementAt.Width = 95
        '
        'ColBase
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.NullValue = "0"
        Me.ColBase.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColBase.Frozen = True
        Me.ColBase.HeaderText = "Base"
        Me.ColBase.Name = "ColBase"
        Me.ColBase.Width = 56
        '
        'ColInterest
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.NullValue = "0"
        Me.ColInterest.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColInterest.Frozen = True
        Me.ColInterest.HeaderText = "Interest"
        Me.ColInterest.Name = "ColInterest"
        Me.ColInterest.Width = 67
        '
        'ColTotalInstallment
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.NullValue = "0"
        Me.ColTotalInstallment.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColTotalInstallment.Frozen = True
        Me.ColTotalInstallment.HeaderText = "Total Installment"
        Me.ColTotalInstallment.Name = "ColTotalInstallment"
        Me.ColTotalInstallment.Width = 109
        '
        'ColPayment
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle6.NullValue = "0"
        Me.ColPayment.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColPayment.HeaderText = "Payment"
        Me.ColPayment.Name = "ColPayment"
        Me.ColPayment.Width = 73
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'ColPrimaryLoan
        '
        Me.ColPrimaryLoan.HeaderText = "Primary Loan"
        Me.ColPrimaryLoan.Name = "ColPrimaryLoan"
        Me.ColPrimaryLoan.Visible = False
        Me.ColPrimaryLoan.Width = 93
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.Visible = False
        Me.ColIdentifier.Width = 72
        '
        'ColPrimaryRev
        '
        Me.ColPrimaryRev.HeaderText = "Primary Rev"
        Me.ColPrimaryRev.Name = "ColPrimaryRev"
        Me.ColPrimaryRev.Visible = False
        Me.ColPrimaryRev.Width = 89
        '
        'BtnPrint
        '
        Me.BtnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnPrint.Location = New System.Drawing.Point(534, 535)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(80, 26)
        Me.BtnPrint.TabIndex = 4
        Me.BtnPrint.Text = "&Print"
        Me.BtnPrint.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'FrmTransaksiLoan_Payment
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(884, 592)
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
        Me.Name = "FrmTransaksiLoan_Payment"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Loan Payment"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.GboItem.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
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
    Friend WithEvents TxtPrimaryLoanPayment As System.Windows.Forms.TextBox
    Friend WithEvents TxtDocumentNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LblNo As System.Windows.Forms.Label
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Public WithEvents btnNew As System.Windows.Forms.Button
    Public WithEvents btnOK As System.Windows.Forms.Button
    Public WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CboLoanNumber As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPrimaryEmployee As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindEmployee As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtPrimaryCashAcc As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescCash As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryGainAcc As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescGain As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindGain As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents TxtGainAccount As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtPrimaryLoss As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BtnFindLoss As System.Windows.Forms.Button
    Public WithEvents TxtLossAccount As System.Windows.Forms.TextBox
    Friend WithEvents GboItem As System.Windows.Forms.GroupBox
    Friend WithEvents TxtDescLoss As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryLoan As System.Windows.Forms.TextBox
    Friend WithEvents TxtIdentifier As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtBase As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtInstallmentAt As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtInterest As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TxtPayment As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalInstallment As System.Windows.Forms.TextBox
    Friend WithEvents BtnInput As System.Windows.Forms.Button
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents TxtPrimarySalary As System.Windows.Forms.TextBox
    Friend WithEvents CboCashAccount As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPrimaryJurnal As System.Windows.Forms.TextBox
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
    Public WithEvents TxtEmployeeCode1 As System.Windows.Forms.TextBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLoanNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInstallementAt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBase As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInterest As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalInstallment As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPayment As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLoan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryRev As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtPrimaryRev As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemarksItem As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Public WithEvents txtJournalNo As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Public WithEvents TxtJournalDesc As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
