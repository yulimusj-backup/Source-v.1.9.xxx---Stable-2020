<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_Quick2
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
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.RboAll = New System.Windows.Forms.RadioButton
        Me.RboSelected = New System.Windows.Forms.RadioButton
        Me.dtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.DtpDate = New System.Windows.Forms.DateTimePicker
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.chkDirectlySave = New System.Windows.Forms.CheckBox
        Me.cboTaxCalculation = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.GboLoanPayment = New System.Windows.Forms.GroupBox
        Me.TxtPrimaryCashAcc = New System.Windows.Forms.TextBox
        Me.TxtAccDesc = New System.Windows.Forms.TextBox
        Me.CboCashAccount = New System.Windows.Forms.ComboBox
        Me.TxtPrimaryLoss = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtPrimaryGainAcc = New System.Windows.Forms.TextBox
        Me.BtnFindGain = New System.Windows.Forms.Button
        Me.BtnFindLoss = New System.Windows.Forms.Button
        Me.TxtLossAccount = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtGainAccount = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtPrimaryTaxExpAcc = New System.Windows.Forms.TextBox
        Me.TxtExpAccDesc = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.BtnFindTaxAcc = New System.Windows.Forms.Button
        Me.TxtExpAcc = New System.Windows.Forms.TextBox
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.chkJobPosition = New System.Windows.Forms.CheckBox
        Me.Lbl_JP_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_JP_checked = New System.Windows.Forms.LinkLabel
        Me.ClbPosition = New System.Windows.Forms.CheckedListBox
        Me.gboDepartment = New System.Windows.Forms.GroupBox
        Me.ChkDepartment = New System.Windows.Forms.CheckBox
        Me.Lbl_Dept_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_Dept_Checked = New System.Windows.Forms.LinkLabel
        Me.ClbDepartment = New System.Windows.Forms.CheckedListBox
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.gboSalaryFormula = New System.Windows.Forms.GroupBox
        Me.chkSalaryFormula = New System.Windows.Forms.CheckBox
        Me.Lbl_SF_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_SF_Checked = New System.Windows.Forms.LinkLabel
        Me.clbSalaryFormula = New System.Windows.Forms.CheckedListBox
        Me.gboEmployee = New System.Windows.Forms.GroupBox
        Me.dgvEmployee = New System.Windows.Forms.DataGridView
        Me.ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChkEmployee = New System.Windows.Forms.CheckBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.LblProgress = New System.Windows.Forms.Label
        Me.BtnPause = New System.Windows.Forms.Button
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GboLoanPayment.SuspendLayout()
        Me.gboItem.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.gboDepartment.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.gboSalaryFormula.SuspendLayout()
        Me.gboEmployee.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(536, 573)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 19)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(536, 573)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 19)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "Start"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 0)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "What would you like to do?"
        '
        'RboAll
        '
        Me.RboAll.AutoSize = True
        Me.RboAll.Checked = True
        Me.RboAll.Location = New System.Drawing.Point(102, 20)
        Me.RboAll.Margin = New System.Windows.Forms.Padding(2)
        Me.RboAll.Name = "RboAll"
        Me.RboAll.Size = New System.Drawing.Size(175, 17)
        Me.RboAll.TabIndex = 1
        Me.RboAll.TabStop = True
        Me.RboAll.Text = "Calculate salary for all employee"
        Me.RboAll.UseVisualStyleBackColor = True
        '
        'RboSelected
        '
        Me.RboSelected.AutoSize = True
        Me.RboSelected.Location = New System.Drawing.Point(43, 96)
        Me.RboSelected.Margin = New System.Windows.Forms.Padding(2)
        Me.RboSelected.Name = "RboSelected"
        Me.RboSelected.Size = New System.Drawing.Size(205, 17)
        Me.RboSelected.TabIndex = 2
        Me.RboSelected.Text = "Calculate salary for selected employee"
        Me.RboSelected.UseVisualStyleBackColor = True
        Me.RboSelected.Visible = False
        '
        'dtpPeriod
        '
        Me.dtpPeriod.CustomFormat = "dd MMM yyyy"
        Me.dtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPeriod.Location = New System.Drawing.Point(90, 70)
        Me.dtpPeriod.Name = "dtpPeriod"
        Me.dtpPeriod.Size = New System.Drawing.Size(187, 20)
        Me.dtpPeriod.TabIndex = 5
        '
        'DtpTo
        '
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(90, 96)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(187, 20)
        Me.DtpTo.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 102)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 49
        Me.Label4.Text = "To"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(7, 76)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(37, 13)
        Me.Label5.TabIndex = 48
        Me.Label5.Text = "Period"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 51
        Me.Label2.Text = "Date"
        '
        'DtpDate
        '
        Me.DtpDate.CustomFormat = "dd MMM yyyy"
        Me.DtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDate.Location = New System.Drawing.Point(90, 20)
        Me.DtpDate.Name = "DtpDate"
        Me.DtpDate.Size = New System.Drawing.Size(188, 20)
        Me.DtpDate.TabIndex = 3
        '
        'gboDetail
        '
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Location = New System.Drawing.Point(15, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(580, 142)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkDirectlySave)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpPeriod)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpTo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RboAll)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RboSelected)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboTaxCalculation)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GboLoanPayment)
        Me.SplitContainer1.Size = New System.Drawing.Size(574, 123)
        Me.SplitContainer1.SplitterDistance = 286
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'chkDirectlySave
        '
        Me.chkDirectlySave.AutoSize = True
        Me.chkDirectlySave.Location = New System.Drawing.Point(2, -1)
        Me.chkDirectlySave.Margin = New System.Windows.Forms.Padding(2)
        Me.chkDirectlySave.Name = "chkDirectlySave"
        Me.chkDirectlySave.Size = New System.Drawing.Size(146, 17)
        Me.chkDirectlySave.TabIndex = 52
        Me.chkDirectlySave.Text = "Directly save to database"
        Me.chkDirectlySave.UseVisualStyleBackColor = True
        '
        'cboTaxCalculation
        '
        Me.cboTaxCalculation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboTaxCalculation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTaxCalculation.FormattingEnabled = True
        Me.cboTaxCalculation.Items.AddRange(New Object() {"Gross Up", "Netto", "Gross", "Manual Tax Calc"})
        Me.cboTaxCalculation.Location = New System.Drawing.Point(90, 44)
        Me.cboTaxCalculation.Margin = New System.Windows.Forms.Padding(2)
        Me.cboTaxCalculation.Name = "cboTaxCalculation"
        Me.cboTaxCalculation.Size = New System.Drawing.Size(186, 21)
        Me.cboTaxCalculation.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(5, 47)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 13)
        Me.Label6.TabIndex = 46
        Me.Label6.Text = "Tax Calculation"
        '
        'GboLoanPayment
        '
        Me.GboLoanPayment.Controls.Add(Me.TxtPrimaryCashAcc)
        Me.GboLoanPayment.Controls.Add(Me.TxtAccDesc)
        Me.GboLoanPayment.Controls.Add(Me.CboCashAccount)
        Me.GboLoanPayment.Controls.Add(Me.TxtPrimaryLoss)
        Me.GboLoanPayment.Controls.Add(Me.Label10)
        Me.GboLoanPayment.Controls.Add(Me.Label11)
        Me.GboLoanPayment.Controls.Add(Me.TxtPrimaryGainAcc)
        Me.GboLoanPayment.Controls.Add(Me.BtnFindGain)
        Me.GboLoanPayment.Controls.Add(Me.BtnFindLoss)
        Me.GboLoanPayment.Controls.Add(Me.TxtLossAccount)
        Me.GboLoanPayment.Controls.Add(Me.Label8)
        Me.GboLoanPayment.Controls.Add(Me.TxtGainAccount)
        Me.GboLoanPayment.Location = New System.Drawing.Point(3, 3)
        Me.GboLoanPayment.Name = "GboLoanPayment"
        Me.GboLoanPayment.Size = New System.Drawing.Size(271, 117)
        Me.GboLoanPayment.TabIndex = 111
        Me.GboLoanPayment.TabStop = False
        Me.GboLoanPayment.Text = "Loan Payment"
        '
        'TxtPrimaryCashAcc
        '
        Me.TxtPrimaryCashAcc.Location = New System.Drawing.Point(129, 10)
        Me.TxtPrimaryCashAcc.Name = "TxtPrimaryCashAcc"
        Me.TxtPrimaryCashAcc.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCashAcc.TabIndex = 99
        Me.TxtPrimaryCashAcc.TabStop = False
        Me.TxtPrimaryCashAcc.Text = "-1"
        Me.TxtPrimaryCashAcc.Visible = False
        '
        'TxtAccDesc
        '
        Me.TxtAccDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAccDesc.Location = New System.Drawing.Point(185, 10)
        Me.TxtAccDesc.MaxLength = 150
        Me.TxtAccDesc.Name = "TxtAccDesc"
        Me.TxtAccDesc.ReadOnly = True
        Me.TxtAccDesc.Size = New System.Drawing.Size(80, 20)
        Me.TxtAccDesc.TabIndex = 110
        Me.TxtAccDesc.TabStop = False
        '
        'CboCashAccount
        '
        Me.CboCashAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboCashAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboCashAccount.FormattingEnabled = True
        Me.CboCashAccount.Location = New System.Drawing.Point(93, 10)
        Me.CboCashAccount.Name = "CboCashAccount"
        Me.CboCashAccount.Size = New System.Drawing.Size(86, 21)
        Me.CboCashAccount.TabIndex = 7
        '
        'TxtPrimaryLoss
        '
        Me.TxtPrimaryLoss.Location = New System.Drawing.Point(129, 58)
        Me.TxtPrimaryLoss.Name = "TxtPrimaryLoss"
        Me.TxtPrimaryLoss.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryLoss.TabIndex = 104
        Me.TxtPrimaryLoss.TabStop = False
        Me.TxtPrimaryLoss.Text = "-1"
        Me.TxtPrimaryLoss.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(1, 17)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(86, 13)
        Me.Label10.TabIndex = 105
        Me.Label10.Text = "Cash In Account"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(1, 37)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(72, 13)
        Me.Label11.TabIndex = 106
        Me.Label11.Text = "Gain Account"
        '
        'TxtPrimaryGainAcc
        '
        Me.TxtPrimaryGainAcc.Location = New System.Drawing.Point(129, 34)
        Me.TxtPrimaryGainAcc.Name = "TxtPrimaryGainAcc"
        Me.TxtPrimaryGainAcc.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryGainAcc.TabIndex = 101
        Me.TxtPrimaryGainAcc.TabStop = False
        Me.TxtPrimaryGainAcc.Text = "-1"
        Me.TxtPrimaryGainAcc.Visible = False
        '
        'BtnFindGain
        '
        Me.BtnFindGain.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindGain.Location = New System.Drawing.Point(223, 32)
        Me.BtnFindGain.Name = "BtnFindGain"
        Me.BtnFindGain.Size = New System.Drawing.Size(42, 23)
        Me.BtnFindGain.TabIndex = 107
        Me.BtnFindGain.TabStop = False
        Me.BtnFindGain.Text = "Find"
        Me.BtnFindGain.UseVisualStyleBackColor = True
        '
        'BtnFindLoss
        '
        Me.BtnFindLoss.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindLoss.Location = New System.Drawing.Point(223, 58)
        Me.BtnFindLoss.Name = "BtnFindLoss"
        Me.BtnFindLoss.Size = New System.Drawing.Size(42, 23)
        Me.BtnFindLoss.TabIndex = 109
        Me.BtnFindLoss.TabStop = False
        Me.BtnFindLoss.Text = "Find"
        Me.BtnFindLoss.UseVisualStyleBackColor = True
        '
        'TxtLossAccount
        '
        Me.TxtLossAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLossAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtLossAccount.Location = New System.Drawing.Point(93, 59)
        Me.TxtLossAccount.MaxLength = 150
        Me.TxtLossAccount.Name = "TxtLossAccount"
        Me.TxtLossAccount.Size = New System.Drawing.Size(128, 20)
        Me.TxtLossAccount.TabIndex = 9
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(1, 58)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(72, 13)
        Me.Label8.TabIndex = 108
        Me.Label8.Text = "Loss Account"
        '
        'TxtGainAccount
        '
        Me.TxtGainAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtGainAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtGainAccount.Location = New System.Drawing.Point(93, 34)
        Me.TxtGainAccount.MaxLength = 150
        Me.TxtGainAccount.Name = "TxtGainAccount"
        Me.TxtGainAccount.Size = New System.Drawing.Size(128, 20)
        Me.TxtGainAccount.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(21, 201)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 13)
        Me.Label3.TabIndex = 63
        Me.Label3.Text = "Tax Description"
        Me.Label3.Visible = False
        '
        'TxtPrimaryTaxExpAcc
        '
        Me.TxtPrimaryTaxExpAcc.Location = New System.Drawing.Point(123, 178)
        Me.TxtPrimaryTaxExpAcc.Name = "TxtPrimaryTaxExpAcc"
        Me.TxtPrimaryTaxExpAcc.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimaryTaxExpAcc.TabIndex = 63
        Me.TxtPrimaryTaxExpAcc.TabStop = False
        Me.TxtPrimaryTaxExpAcc.Text = "-1"
        Me.TxtPrimaryTaxExpAcc.Visible = False
        '
        'TxtExpAccDesc
        '
        Me.TxtExpAccDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtExpAccDesc.Location = New System.Drawing.Point(123, 198)
        Me.TxtExpAccDesc.MaxLength = 50
        Me.TxtExpAccDesc.Name = "TxtExpAccDesc"
        Me.TxtExpAccDesc.ReadOnly = True
        Me.TxtExpAccDesc.Size = New System.Drawing.Size(176, 20)
        Me.TxtExpAccDesc.TabIndex = 5
        Me.TxtExpAccDesc.TabStop = False
        Me.TxtExpAccDesc.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(23, 177)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(89, 13)
        Me.Label7.TabIndex = 62
        Me.Label7.Text = "Tax Exp Account"
        Me.Label7.Visible = False
        '
        'BtnFindTaxAcc
        '
        Me.BtnFindTaxAcc.Location = New System.Drawing.Point(254, 178)
        Me.BtnFindTaxAcc.Name = "BtnFindTaxAcc"
        Me.BtnFindTaxAcc.Size = New System.Drawing.Size(42, 20)
        Me.BtnFindTaxAcc.TabIndex = 60
        Me.BtnFindTaxAcc.TabStop = False
        Me.BtnFindTaxAcc.Text = "Find"
        Me.BtnFindTaxAcc.UseVisualStyleBackColor = True
        Me.BtnFindTaxAcc.Visible = False
        '
        'TxtExpAcc
        '
        Me.TxtExpAcc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtExpAcc.Location = New System.Drawing.Point(123, 178)
        Me.TxtExpAcc.MaxLength = 100
        Me.TxtExpAcc.Name = "TxtExpAcc"
        Me.TxtExpAcc.Size = New System.Drawing.Size(128, 20)
        Me.TxtExpAcc.TabIndex = 4
        Me.TxtExpAcc.Visible = False
        '
        'gboItem
        '
        Me.gboItem.Controls.Add(Me.SplitContainer3)
        Me.gboItem.Location = New System.Drawing.Point(15, 145)
        Me.gboItem.Margin = New System.Windows.Forms.Padding(2)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Padding = New System.Windows.Forms.Padding(2)
        Me.gboItem.Size = New System.Drawing.Size(580, 423)
        Me.gboItem.TabIndex = 1
        Me.gboItem.TabStop = False
        Me.gboItem.Text = "Data Filter"
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Location = New System.Drawing.Point(5, 18)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer3.Size = New System.Drawing.Size(569, 400)
        Me.SplitContainer3.SplitterDistance = 281
        Me.SplitContainer3.TabIndex = 0
        Me.SplitContainer3.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox3)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.gboDepartment)
        Me.SplitContainer2.Size = New System.Drawing.Size(281, 400)
        Me.SplitContainer2.SplitterDistance = 192
        Me.SplitContainer2.TabIndex = 0
        Me.SplitContainer2.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.chkJobPosition)
        Me.GroupBox3.Controls.Add(Me.Lbl_JP_Unchecked)
        Me.GroupBox3.Controls.Add(Me.Lbl_JP_checked)
        Me.GroupBox3.Controls.Add(Me.ClbPosition)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(281, 192)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        '
        'chkJobPosition
        '
        Me.chkJobPosition.AutoSize = True
        Me.chkJobPosition.Location = New System.Drawing.Point(6, -1)
        Me.chkJobPosition.Name = "chkJobPosition"
        Me.chkJobPosition.Size = New System.Drawing.Size(83, 17)
        Me.chkJobPosition.TabIndex = 0
        Me.chkJobPosition.Text = "Job Position"
        Me.chkJobPosition.UseVisualStyleBackColor = True
        '
        'Lbl_JP_Unchecked
        '
        Me.Lbl_JP_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_JP_Unchecked.AutoSize = True
        Me.Lbl_JP_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_JP_Unchecked.Location = New System.Drawing.Point(210, 0)
        Me.Lbl_JP_Unchecked.Name = "Lbl_JP_Unchecked"
        Me.Lbl_JP_Unchecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_JP_Unchecked.TabIndex = 2
        Me.Lbl_JP_Unchecked.TabStop = True
        Me.Lbl_JP_Unchecked.Text = "Uncheck All"
        '
        'Lbl_JP_checked
        '
        Me.Lbl_JP_checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_JP_checked.AutoSize = True
        Me.Lbl_JP_checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_JP_checked.Location = New System.Drawing.Point(152, 0)
        Me.Lbl_JP_checked.Name = "Lbl_JP_checked"
        Me.Lbl_JP_checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_JP_checked.TabIndex = 1
        Me.Lbl_JP_checked.TabStop = True
        Me.Lbl_JP_checked.Text = "Check All"
        '
        'ClbPosition
        '
        Me.ClbPosition.CheckOnClick = True
        Me.ClbPosition.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ClbPosition.FormattingEnabled = True
        Me.ClbPosition.Location = New System.Drawing.Point(3, 16)
        Me.ClbPosition.Name = "ClbPosition"
        Me.ClbPosition.Size = New System.Drawing.Size(275, 169)
        Me.ClbPosition.TabIndex = 3
        '
        'gboDepartment
        '
        Me.gboDepartment.Controls.Add(Me.ChkDepartment)
        Me.gboDepartment.Controls.Add(Me.Lbl_Dept_Unchecked)
        Me.gboDepartment.Controls.Add(Me.Lbl_Dept_Checked)
        Me.gboDepartment.Controls.Add(Me.ClbDepartment)
        Me.gboDepartment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboDepartment.Location = New System.Drawing.Point(0, 0)
        Me.gboDepartment.Name = "gboDepartment"
        Me.gboDepartment.Size = New System.Drawing.Size(281, 204)
        Me.gboDepartment.TabIndex = 0
        Me.gboDepartment.TabStop = False
        '
        'ChkDepartment
        '
        Me.ChkDepartment.AutoSize = True
        Me.ChkDepartment.Location = New System.Drawing.Point(6, -1)
        Me.ChkDepartment.Name = "ChkDepartment"
        Me.ChkDepartment.Size = New System.Drawing.Size(81, 17)
        Me.ChkDepartment.TabIndex = 0
        Me.ChkDepartment.Text = "Department"
        Me.ChkDepartment.UseVisualStyleBackColor = True
        '
        'Lbl_Dept_Unchecked
        '
        Me.Lbl_Dept_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Dept_Unchecked.AutoSize = True
        Me.Lbl_Dept_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Dept_Unchecked.Location = New System.Drawing.Point(210, 0)
        Me.Lbl_Dept_Unchecked.Name = "Lbl_Dept_Unchecked"
        Me.Lbl_Dept_Unchecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_Dept_Unchecked.TabIndex = 2
        Me.Lbl_Dept_Unchecked.TabStop = True
        Me.Lbl_Dept_Unchecked.Text = "Uncheck All"
        '
        'Lbl_Dept_Checked
        '
        Me.Lbl_Dept_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Dept_Checked.AutoSize = True
        Me.Lbl_Dept_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Dept_Checked.Location = New System.Drawing.Point(152, 0)
        Me.Lbl_Dept_Checked.Name = "Lbl_Dept_Checked"
        Me.Lbl_Dept_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_Dept_Checked.TabIndex = 1
        Me.Lbl_Dept_Checked.TabStop = True
        Me.Lbl_Dept_Checked.Text = "Check All"
        '
        'ClbDepartment
        '
        Me.ClbDepartment.CheckOnClick = True
        Me.ClbDepartment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ClbDepartment.FormattingEnabled = True
        Me.ClbDepartment.Location = New System.Drawing.Point(3, 16)
        Me.ClbDepartment.Name = "ClbDepartment"
        Me.ClbDepartment.Size = New System.Drawing.Size(275, 184)
        Me.ClbDepartment.TabIndex = 3
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Name = "SplitContainer4"
        Me.SplitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.gboSalaryFormula)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.gboEmployee)
        Me.SplitContainer4.Size = New System.Drawing.Size(284, 400)
        Me.SplitContainer4.SplitterDistance = 191
        Me.SplitContainer4.TabIndex = 0
        Me.SplitContainer4.TabStop = False
        '
        'gboSalaryFormula
        '
        Me.gboSalaryFormula.Controls.Add(Me.chkSalaryFormula)
        Me.gboSalaryFormula.Controls.Add(Me.Lbl_SF_Unchecked)
        Me.gboSalaryFormula.Controls.Add(Me.Lbl_SF_Checked)
        Me.gboSalaryFormula.Controls.Add(Me.clbSalaryFormula)
        Me.gboSalaryFormula.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboSalaryFormula.Location = New System.Drawing.Point(0, 0)
        Me.gboSalaryFormula.Name = "gboSalaryFormula"
        Me.gboSalaryFormula.Size = New System.Drawing.Size(284, 191)
        Me.gboSalaryFormula.TabIndex = 0
        Me.gboSalaryFormula.TabStop = False
        '
        'chkSalaryFormula
        '
        Me.chkSalaryFormula.AutoSize = True
        Me.chkSalaryFormula.Location = New System.Drawing.Point(6, -1)
        Me.chkSalaryFormula.Name = "chkSalaryFormula"
        Me.chkSalaryFormula.Size = New System.Drawing.Size(95, 17)
        Me.chkSalaryFormula.TabIndex = 0
        Me.chkSalaryFormula.Text = "Salary Formula"
        Me.chkSalaryFormula.UseVisualStyleBackColor = True
        '
        'Lbl_SF_Unchecked
        '
        Me.Lbl_SF_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_SF_Unchecked.AutoSize = True
        Me.Lbl_SF_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_SF_Unchecked.Location = New System.Drawing.Point(214, 0)
        Me.Lbl_SF_Unchecked.Name = "Lbl_SF_Unchecked"
        Me.Lbl_SF_Unchecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_SF_Unchecked.TabIndex = 2
        Me.Lbl_SF_Unchecked.TabStop = True
        Me.Lbl_SF_Unchecked.Text = "Uncheck All"
        '
        'Lbl_SF_Checked
        '
        Me.Lbl_SF_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_SF_Checked.AutoSize = True
        Me.Lbl_SF_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_SF_Checked.Location = New System.Drawing.Point(155, 0)
        Me.Lbl_SF_Checked.Name = "Lbl_SF_Checked"
        Me.Lbl_SF_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_SF_Checked.TabIndex = 1
        Me.Lbl_SF_Checked.TabStop = True
        Me.Lbl_SF_Checked.Text = "Check All"
        '
        'clbSalaryFormula
        '
        Me.clbSalaryFormula.CheckOnClick = True
        Me.clbSalaryFormula.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbSalaryFormula.FormattingEnabled = True
        Me.clbSalaryFormula.Location = New System.Drawing.Point(3, 16)
        Me.clbSalaryFormula.Name = "clbSalaryFormula"
        Me.clbSalaryFormula.Size = New System.Drawing.Size(278, 169)
        Me.clbSalaryFormula.TabIndex = 3
        '
        'gboEmployee
        '
        Me.gboEmployee.Controls.Add(Me.dgvEmployee)
        Me.gboEmployee.Controls.Add(Me.ChkEmployee)
        Me.gboEmployee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboEmployee.Location = New System.Drawing.Point(0, 0)
        Me.gboEmployee.Name = "gboEmployee"
        Me.gboEmployee.Size = New System.Drawing.Size(284, 205)
        Me.gboEmployee.TabIndex = 0
        Me.gboEmployee.TabStop = False
        '
        'dgvEmployee
        '
        Me.dgvEmployee.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCode, Me.ColName, Me.ColNIK, Me.ColPrimaryCode})
        Me.dgvEmployee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvEmployee.Location = New System.Drawing.Point(3, 16)
        Me.dgvEmployee.Name = "dgvEmployee"
        Me.dgvEmployee.RowTemplate.Height = 24
        Me.dgvEmployee.Size = New System.Drawing.Size(278, 186)
        Me.dgvEmployee.TabIndex = 0
        '
        'ColCode
        '
        Me.ColCode.HeaderText = "Code"
        Me.ColCode.Name = "ColCode"
        Me.ColCode.Width = 57
        '
        'ColName
        '
        Me.ColName.HeaderText = "Name"
        Me.ColName.Name = "ColName"
        Me.ColName.Width = 60
        '
        'ColNIK
        '
        Me.ColNIK.HeaderText = "NIK"
        Me.ColNIK.Name = "ColNIK"
        Me.ColNIK.Width = 50
        '
        'ColPrimaryCode
        '
        Me.ColPrimaryCode.HeaderText = "Primary"
        Me.ColPrimaryCode.Name = "ColPrimaryCode"
        Me.ColPrimaryCode.Visible = False
        Me.ColPrimaryCode.Width = 66
        '
        'ChkEmployee
        '
        Me.ChkEmployee.AutoSize = True
        Me.ChkEmployee.Location = New System.Drawing.Point(6, -1)
        Me.ChkEmployee.Name = "ChkEmployee"
        Me.ChkEmployee.Size = New System.Drawing.Size(72, 17)
        Me.ChkEmployee.TabIndex = 1
        Me.ChkEmployee.Text = "Employee"
        Me.ChkEmployee.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(23, 573)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(508, 18)
        Me.ProgressBar1.TabIndex = 64
        '
        'LblProgress
        '
        Me.LblProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(23, 594)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(39, 13)
        Me.LblProgress.TabIndex = 65
        Me.LblProgress.Text = "Label9"
        '
        'BtnPause
        '
        Me.BtnPause.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnPause.Location = New System.Drawing.Point(536, 573)
        Me.BtnPause.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnPause.Name = "BtnPause"
        Me.BtnPause.Size = New System.Drawing.Size(56, 19)
        Me.BtnPause.TabIndex = 66
        Me.BtnPause.Text = "Pause"
        Me.BtnPause.UseVisualStyleBackColor = True
        '
        'BackgroundWorker2
        '
        '
        'FrmTrans_Gaji_Quick2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(607, 608)
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.gboItem)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TxtPrimaryTaxExpAcc)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.TxtExpAccDesc)
        Me.Controls.Add(Me.TxtExpAcc)
        Me.Controls.Add(Me.BtnFindTaxAcc)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.BtnPause)
        Me.Controls.Add(Me.btnCancel)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(623, 646)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(623, 646)
        Me.Name = "FrmTrans_Gaji_Quick2"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data Filter And Calculate Salary"
        Me.gboDetail.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GboLoanPayment.ResumeLayout(False)
        Me.GboLoanPayment.PerformLayout()
        Me.gboItem.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.gboDepartment.ResumeLayout(False)
        Me.gboDepartment.PerformLayout()
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.ResumeLayout(False)
        Me.gboSalaryFormula.ResumeLayout(False)
        Me.gboSalaryFormula.PerformLayout()
        Me.gboEmployee.ResumeLayout(False)
        Me.gboEmployee.PerformLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents RboAll As System.Windows.Forms.RadioButton
    Friend WithEvents RboSelected As System.Windows.Forms.RadioButton
    Friend WithEvents dtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents chkJobPosition As System.Windows.Forms.CheckBox
    Friend WithEvents Lbl_JP_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_JP_checked As System.Windows.Forms.LinkLabel
    Friend WithEvents ClbPosition As System.Windows.Forms.CheckedListBox
    Friend WithEvents gboDepartment As System.Windows.Forms.GroupBox
    Friend WithEvents ChkDepartment As System.Windows.Forms.CheckBox
    Friend WithEvents Lbl_Dept_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_Dept_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents ClbDepartment As System.Windows.Forms.CheckedListBox
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents gboSalaryFormula As System.Windows.Forms.GroupBox
    Friend WithEvents chkSalaryFormula As System.Windows.Forms.CheckBox
    Friend WithEvents Lbl_SF_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_SF_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents clbSalaryFormula As System.Windows.Forms.CheckedListBox
    Friend WithEvents gboEmployee As System.Windows.Forms.GroupBox
    Friend WithEvents dgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChkEmployee As System.Windows.Forms.CheckBox
    Friend WithEvents TxtExpAccDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryTaxExpAcc As System.Windows.Forms.TextBox
    Friend WithEvents TxtExpAcc As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindTaxAcc As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboTaxCalculation As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TxtPrimaryLoss As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryGainAcc As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindLoss As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryCashAcc As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents TxtLossAccount As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindGain As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents TxtGainAccount As System.Windows.Forms.TextBox
    Friend WithEvents CboCashAccount As System.Windows.Forms.ComboBox
    Public WithEvents TxtAccDesc As System.Windows.Forms.TextBox
    Friend WithEvents GboLoanPayment As System.Windows.Forms.GroupBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents BtnPause As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents chkDirectlySave As System.Windows.Forms.CheckBox
End Class
