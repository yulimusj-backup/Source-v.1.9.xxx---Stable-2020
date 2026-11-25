<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporan_ProfitNLoss
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
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.dtpFrom_I = New System.Windows.Forms.DateTimePicker
        Me.gboPeriod_II = New System.Windows.Forms.GroupBox
        Me.dtpFrom_II = New System.Windows.Forms.DateTimePicker
        Me.dtpTo_II = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.gboPeriod_I = New System.Windows.Forms.GroupBox
        Me.dtpTo_I = New System.Windows.Forms.DateTimePicker
        Me.LblPeriodFrom = New System.Windows.Forms.Label
        Me.LblPeriodTo = New System.Windows.Forms.Label
        Me.cboFormatLaporan = New System.Windows.Forms.ComboBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.chkUseTotal = New System.Windows.Forms.CheckBox
        Me.chkShowLvl4 = New System.Windows.Forms.CheckBox
        Me.chkShowLvl3 = New System.Windows.Forms.CheckBox
        Me.chkShowLvl2 = New System.Windows.Forms.CheckBox
        Me.chkHideZero = New System.Windows.Forms.CheckBox
        Me.chkHideZeroInactiveAccount = New System.Windows.Forms.CheckBox
        Me.chkShowAccountCode = New System.Windows.Forms.CheckBox
        Me.TxtDeptCode = New System.Windows.Forms.TextBox
        Me.btnFindDept = New System.Windows.Forms.Button
        Me.TxtPrimaryDept = New System.Windows.Forms.TextBox
        Me.TxtJobCode = New System.Windows.Forms.TextBox
        Me.btnFindJob = New System.Windows.Forms.Button
        Me.TxtPrimaryJob = New System.Windows.Forms.TextBox
        Me.LblDept = New System.Windows.Forms.Label
        Me.LblJob = New System.Windows.Forms.Label
        Me.chkShowCOGS = New System.Windows.Forms.CheckBox
        Me.chkShowIncome = New System.Windows.Forms.CheckBox
        Me.chkShowExpense = New System.Windows.Forms.CheckBox
        Me.chkShowProfit = New System.Windows.Forms.CheckBox
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.chkShowLvl8 = New System.Windows.Forms.CheckBox
        Me.chkShowLvl7 = New System.Windows.Forms.CheckBox
        Me.chkShowLvl6 = New System.Windows.Forms.CheckBox
        Me.chkShowLvl5 = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.gboPeriod_II.SuspendLayout()
        Me.gboPeriod_I.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.EnableDrillDown = False
        Me.crv1.Location = New System.Drawing.Point(330, 12)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(542, 538)
        Me.crv1.TabIndex = 0
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'dtpFrom_I
        '
        Me.dtpFrom_I.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom_I.CustomFormat = "MMM yyyy"
        Me.dtpFrom_I.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom_I.Location = New System.Drawing.Point(42, 19)
        Me.dtpFrom_I.Name = "dtpFrom_I"
        Me.dtpFrom_I.ShowUpDown = True
        Me.dtpFrom_I.Size = New System.Drawing.Size(98, 20)
        Me.dtpFrom_I.TabIndex = 0
        '
        'gboPeriod_II
        '
        Me.gboPeriod_II.Controls.Add(Me.dtpFrom_II)
        Me.gboPeriod_II.Controls.Add(Me.dtpTo_II)
        Me.gboPeriod_II.Controls.Add(Me.Label2)
        Me.gboPeriod_II.Controls.Add(Me.Label3)
        Me.gboPeriod_II.Location = New System.Drawing.Point(160, 46)
        Me.gboPeriod_II.Name = "gboPeriod_II"
        Me.gboPeriod_II.Size = New System.Drawing.Size(148, 76)
        Me.gboPeriod_II.TabIndex = 2
        Me.gboPeriod_II.TabStop = False
        Me.gboPeriod_II.Text = "Period II"
        '
        'dtpFrom_II
        '
        Me.dtpFrom_II.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom_II.CustomFormat = "MMM yyyy"
        Me.dtpFrom_II.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom_II.Location = New System.Drawing.Point(42, 19)
        Me.dtpFrom_II.Name = "dtpFrom_II"
        Me.dtpFrom_II.ShowUpDown = True
        Me.dtpFrom_II.Size = New System.Drawing.Size(100, 20)
        Me.dtpFrom_II.TabIndex = 0
        '
        'dtpTo_II
        '
        Me.dtpTo_II.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo_II.CustomFormat = "MMM yyyy"
        Me.dtpTo_II.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo_II.Location = New System.Drawing.Point(42, 45)
        Me.dtpTo_II.Name = "dtpTo_II"
        Me.dtpTo_II.ShowUpDown = True
        Me.dtpTo_II.Size = New System.Drawing.Size(100, 20)
        Me.dtpTo_II.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "From"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 50)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "To"
        '
        'gboPeriod_I
        '
        Me.gboPeriod_I.Controls.Add(Me.dtpFrom_I)
        Me.gboPeriod_I.Controls.Add(Me.dtpTo_I)
        Me.gboPeriod_I.Controls.Add(Me.LblPeriodFrom)
        Me.gboPeriod_I.Controls.Add(Me.LblPeriodTo)
        Me.gboPeriod_I.Location = New System.Drawing.Point(6, 46)
        Me.gboPeriod_I.Name = "gboPeriod_I"
        Me.gboPeriod_I.Size = New System.Drawing.Size(148, 76)
        Me.gboPeriod_I.TabIndex = 1
        Me.gboPeriod_I.TabStop = False
        Me.gboPeriod_I.Text = "Period I"
        '
        'dtpTo_I
        '
        Me.dtpTo_I.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo_I.CustomFormat = "MMM yyyy"
        Me.dtpTo_I.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo_I.Location = New System.Drawing.Point(42, 45)
        Me.dtpTo_I.Name = "dtpTo_I"
        Me.dtpTo_I.ShowUpDown = True
        Me.dtpTo_I.Size = New System.Drawing.Size(98, 20)
        Me.dtpTo_I.TabIndex = 1
        '
        'LblPeriodFrom
        '
        Me.LblPeriodFrom.AutoSize = True
        Me.LblPeriodFrom.Location = New System.Drawing.Point(6, 24)
        Me.LblPeriodFrom.Name = "LblPeriodFrom"
        Me.LblPeriodFrom.Size = New System.Drawing.Size(30, 13)
        Me.LblPeriodFrom.TabIndex = 30
        Me.LblPeriodFrom.Text = "From"
        '
        'LblPeriodTo
        '
        Me.LblPeriodTo.AutoSize = True
        Me.LblPeriodTo.Location = New System.Drawing.Point(6, 50)
        Me.LblPeriodTo.Name = "LblPeriodTo"
        Me.LblPeriodTo.Size = New System.Drawing.Size(20, 13)
        Me.LblPeriodTo.TabIndex = 32
        Me.LblPeriodTo.Text = "To"
        '
        'cboFormatLaporan
        '
        Me.cboFormatLaporan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboFormatLaporan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormatLaporan.FormattingEnabled = True
        Me.cboFormatLaporan.Items.AddRange(New Object() {"Profit & Loss By Period", "Profit & Loss 2 Periods Comparison", "Profit & Loss Vs Budget Comparison", "Profit & Loss 2 Periods Year To Date", "Profit & Loss 6 Months Period", "Profit & Loss 6 Months Period Vs Budget", "Profit & Loss 1 Year Period", "Profit & Loss 1 Year Period Vs Budget", "Profit & Loss Quarter Yearly", "Monthly Recap Chart"})
        Me.cboFormatLaporan.Location = New System.Drawing.Point(6, 19)
        Me.cboFormatLaporan.Name = "cboFormatLaporan"
        Me.cboFormatLaporan.Size = New System.Drawing.Size(302, 21)
        Me.cboFormatLaporan.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(226, 334)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 23)
        Me.btnClose.TabIndex = 13
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(144, 334)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(76, 23)
        Me.btnShow.TabIndex = 12
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'chkUseTotal
        '
        Me.chkUseTotal.AutoSize = True
        Me.chkUseTotal.Location = New System.Drawing.Point(15, 128)
        Me.chkUseTotal.Name = "chkUseTotal"
        Me.chkUseTotal.Size = New System.Drawing.Size(72, 17)
        Me.chkUseTotal.TabIndex = 3
        Me.chkUseTotal.Text = "Use Total"
        Me.chkUseTotal.UseVisualStyleBackColor = True
        '
        'chkShowLvl4
        '
        Me.chkShowLvl4.AutoSize = True
        Me.chkShowLvl4.Checked = True
        Me.chkShowLvl4.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl4.Location = New System.Drawing.Point(15, 279)
        Me.chkShowLvl4.Name = "chkShowLvl4"
        Me.chkShowLvl4.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl4.TabIndex = 14
        Me.chkShowLvl4.Text = "Show level 4 accounts"
        Me.chkShowLvl4.UseVisualStyleBackColor = True
        '
        'chkShowLvl3
        '
        Me.chkShowLvl3.AutoSize = True
        Me.chkShowLvl3.Checked = True
        Me.chkShowLvl3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl3.Location = New System.Drawing.Point(15, 256)
        Me.chkShowLvl3.Name = "chkShowLvl3"
        Me.chkShowLvl3.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl3.TabIndex = 13
        Me.chkShowLvl3.Text = "Show level 3 accounts"
        Me.chkShowLvl3.UseVisualStyleBackColor = True
        '
        'chkShowLvl2
        '
        Me.chkShowLvl2.AutoSize = True
        Me.chkShowLvl2.Checked = True
        Me.chkShowLvl2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl2.Location = New System.Drawing.Point(15, 233)
        Me.chkShowLvl2.Name = "chkShowLvl2"
        Me.chkShowLvl2.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl2.TabIndex = 12
        Me.chkShowLvl2.Text = "Show level 2 accounts"
        Me.chkShowLvl2.UseVisualStyleBackColor = True
        '
        'chkHideZero
        '
        Me.chkHideZero.AutoSize = True
        Me.chkHideZero.Location = New System.Drawing.Point(15, 197)
        Me.chkHideZero.Name = "chkHideZero"
        Me.chkHideZero.Size = New System.Drawing.Size(151, 17)
        Me.chkHideZero.TabIndex = 6
        Me.chkHideZero.Text = "Hide zero amount account"
        Me.chkHideZero.UseVisualStyleBackColor = True
        '
        'chkHideZeroInactiveAccount
        '
        Me.chkHideZeroInactiveAccount.AutoSize = True
        Me.chkHideZeroInactiveAccount.Location = New System.Drawing.Point(15, 174)
        Me.chkHideZeroInactiveAccount.Name = "chkHideZeroInactiveAccount"
        Me.chkHideZeroInactiveAccount.Size = New System.Drawing.Size(153, 17)
        Me.chkHideZeroInactiveAccount.TabIndex = 5
        Me.chkHideZeroInactiveAccount.Text = "Hide zero inactive account"
        Me.chkHideZeroInactiveAccount.UseVisualStyleBackColor = True
        '
        'chkShowAccountCode
        '
        Me.chkShowAccountCode.AutoSize = True
        Me.chkShowAccountCode.Checked = True
        Me.chkShowAccountCode.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowAccountCode.Location = New System.Drawing.Point(15, 151)
        Me.chkShowAccountCode.Name = "chkShowAccountCode"
        Me.chkShowAccountCode.Size = New System.Drawing.Size(124, 17)
        Me.chkShowAccountCode.TabIndex = 4
        Me.chkShowAccountCode.Text = "Show Account Code"
        Me.chkShowAccountCode.UseVisualStyleBackColor = True
        '
        'TxtDeptCode
        '
        Me.TxtDeptCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDeptCode.Location = New System.Drawing.Point(86, 88)
        Me.TxtDeptCode.Name = "TxtDeptCode"
        Me.TxtDeptCode.Size = New System.Drawing.Size(120, 20)
        Me.TxtDeptCode.TabIndex = 5
        '
        'btnFindDept
        '
        Me.btnFindDept.Location = New System.Drawing.Point(212, 87)
        Me.btnFindDept.Name = "btnFindDept"
        Me.btnFindDept.Size = New System.Drawing.Size(37, 21)
        Me.btnFindDept.TabIndex = 6
        Me.btnFindDept.TabStop = False
        Me.btnFindDept.Text = "Find"
        Me.btnFindDept.UseVisualStyleBackColor = True
        '
        'TxtPrimaryDept
        '
        Me.TxtPrimaryDept.Location = New System.Drawing.Point(60, 87)
        Me.TxtPrimaryDept.Name = "TxtPrimaryDept"
        Me.TxtPrimaryDept.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryDept.TabIndex = 57
        Me.TxtPrimaryDept.Text = "-1"
        Me.TxtPrimaryDept.Visible = False
        '
        'TxtJobCode
        '
        Me.TxtJobCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtJobCode.Location = New System.Drawing.Point(86, 111)
        Me.TxtJobCode.Name = "TxtJobCode"
        Me.TxtJobCode.Size = New System.Drawing.Size(120, 20)
        Me.TxtJobCode.TabIndex = 7
        '
        'btnFindJob
        '
        Me.btnFindJob.Location = New System.Drawing.Point(212, 110)
        Me.btnFindJob.Name = "btnFindJob"
        Me.btnFindJob.Size = New System.Drawing.Size(37, 21)
        Me.btnFindJob.TabIndex = 8
        Me.btnFindJob.TabStop = False
        Me.btnFindJob.Text = "Find"
        Me.btnFindJob.UseVisualStyleBackColor = True
        '
        'TxtPrimaryJob
        '
        Me.TxtPrimaryJob.Location = New System.Drawing.Point(60, 110)
        Me.TxtPrimaryJob.Name = "TxtPrimaryJob"
        Me.TxtPrimaryJob.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryJob.TabIndex = 59
        Me.TxtPrimaryJob.Text = "-1"
        Me.TxtPrimaryJob.Visible = False
        '
        'LblDept
        '
        Me.LblDept.AutoSize = True
        Me.LblDept.Location = New System.Drawing.Point(12, 91)
        Me.LblDept.Name = "LblDept"
        Me.LblDept.Size = New System.Drawing.Size(68, 13)
        Me.LblDept.TabIndex = 56
        Me.LblDept.Text = "Departement"
        '
        'LblJob
        '
        Me.LblJob.AutoSize = True
        Me.LblJob.Location = New System.Drawing.Point(12, 114)
        Me.LblJob.Name = "LblJob"
        Me.LblJob.Size = New System.Drawing.Size(24, 13)
        Me.LblJob.TabIndex = 58
        Me.LblJob.Text = "Job"
        '
        'chkShowCOGS
        '
        Me.chkShowCOGS.AutoSize = True
        Me.chkShowCOGS.Enabled = False
        Me.chkShowCOGS.Location = New System.Drawing.Point(169, 151)
        Me.chkShowCOGS.Name = "chkShowCOGS"
        Me.chkShowCOGS.Size = New System.Drawing.Size(86, 17)
        Me.chkShowCOGS.TabIndex = 8
        Me.chkShowCOGS.TabStop = False
        Me.chkShowCOGS.Text = "Show COGS"
        Me.chkShowCOGS.UseVisualStyleBackColor = True
        '
        'chkShowIncome
        '
        Me.chkShowIncome.AutoSize = True
        Me.chkShowIncome.Enabled = False
        Me.chkShowIncome.Location = New System.Drawing.Point(169, 128)
        Me.chkShowIncome.Name = "chkShowIncome"
        Me.chkShowIncome.Size = New System.Drawing.Size(91, 17)
        Me.chkShowIncome.TabIndex = 7
        Me.chkShowIncome.TabStop = False
        Me.chkShowIncome.Text = "Show Income"
        Me.chkShowIncome.UseVisualStyleBackColor = True
        '
        'chkShowExpense
        '
        Me.chkShowExpense.AutoSize = True
        Me.chkShowExpense.Enabled = False
        Me.chkShowExpense.Location = New System.Drawing.Point(169, 173)
        Me.chkShowExpense.Name = "chkShowExpense"
        Me.chkShowExpense.Size = New System.Drawing.Size(97, 17)
        Me.chkShowExpense.TabIndex = 9
        Me.chkShowExpense.TabStop = False
        Me.chkShowExpense.Text = "Show Expense"
        Me.chkShowExpense.UseVisualStyleBackColor = True
        '
        'chkShowProfit
        '
        Me.chkShowProfit.AutoSize = True
        Me.chkShowProfit.Enabled = False
        Me.chkShowProfit.Location = New System.Drawing.Point(169, 196)
        Me.chkShowProfit.Name = "chkShowProfit"
        Me.chkShowProfit.Size = New System.Drawing.Size(80, 17)
        Me.chkShowProfit.TabIndex = 10
        Me.chkShowProfit.TabStop = False
        Me.chkShowProfit.Text = "Show Profit"
        Me.chkShowProfit.UseVisualStyleBackColor = True
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(12, 339)
        Me.LinkLabel1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(70, 13)
        Me.LinkLabel1.TabIndex = 11
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Detail Setting"
        Me.LinkLabel1.Visible = False
        '
        'chkShowLvl8
        '
        Me.chkShowLvl8.AutoSize = True
        Me.chkShowLvl8.Checked = True
        Me.chkShowLvl8.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl8.Location = New System.Drawing.Point(169, 279)
        Me.chkShowLvl8.Name = "chkShowLvl8"
        Me.chkShowLvl8.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl8.TabIndex = 64
        Me.chkShowLvl8.Text = "Show level 8 accounts"
        Me.chkShowLvl8.UseVisualStyleBackColor = True
        '
        'chkShowLvl7
        '
        Me.chkShowLvl7.AutoSize = True
        Me.chkShowLvl7.Checked = True
        Me.chkShowLvl7.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl7.Location = New System.Drawing.Point(169, 256)
        Me.chkShowLvl7.Name = "chkShowLvl7"
        Me.chkShowLvl7.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl7.TabIndex = 63
        Me.chkShowLvl7.Text = "Show level 7 accounts"
        Me.chkShowLvl7.UseVisualStyleBackColor = True
        '
        'chkShowLvl6
        '
        Me.chkShowLvl6.AutoSize = True
        Me.chkShowLvl6.Checked = True
        Me.chkShowLvl6.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl6.Location = New System.Drawing.Point(169, 233)
        Me.chkShowLvl6.Name = "chkShowLvl6"
        Me.chkShowLvl6.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl6.TabIndex = 62
        Me.chkShowLvl6.Text = "Show level 6 accounts"
        Me.chkShowLvl6.UseVisualStyleBackColor = True
        '
        'chkShowLvl5
        '
        Me.chkShowLvl5.AutoSize = True
        Me.chkShowLvl5.Checked = True
        Me.chkShowLvl5.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl5.Location = New System.Drawing.Point(15, 302)
        Me.chkShowLvl5.Name = "chkShowLvl5"
        Me.chkShowLvl5.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl5.TabIndex = 61
        Me.chkShowLvl5.Text = "Show level 5 accounts"
        Me.chkShowLvl5.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cboFormatLaporan)
        Me.GroupBox2.Controls.Add(Me.chkShowLvl8)
        Me.GroupBox2.Controls.Add(Me.gboPeriod_I)
        Me.GroupBox2.Controls.Add(Me.chkShowLvl7)
        Me.GroupBox2.Controls.Add(Me.gboPeriod_II)
        Me.GroupBox2.Controls.Add(Me.chkShowLvl6)
        Me.GroupBox2.Controls.Add(Me.LinkLabel1)
        Me.GroupBox2.Controls.Add(Me.chkShowLvl5)
        Me.GroupBox2.Controls.Add(Me.chkHideZeroInactiveAccount)
        Me.GroupBox2.Controls.Add(Me.chkShowExpense)
        Me.GroupBox2.Controls.Add(Me.chkHideZero)
        Me.GroupBox2.Controls.Add(Me.chkShowProfit)
        Me.GroupBox2.Controls.Add(Me.chkShowIncome)
        Me.GroupBox2.Controls.Add(Me.chkUseTotal)
        Me.GroupBox2.Controls.Add(Me.chkShowCOGS)
        Me.GroupBox2.Controls.Add(Me.chkShowAccountCode)
        Me.GroupBox2.Controls.Add(Me.btnClose)
        Me.GroupBox2.Controls.Add(Me.chkShowLvl4)
        Me.GroupBox2.Controls.Add(Me.btnShow)
        Me.GroupBox2.Controls.Add(Me.chkShowLvl3)
        Me.GroupBox2.Controls.Add(Me.chkShowLvl2)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(316, 368)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Filter By"
        '
        'FrmLaporan_ProfitNLoss
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(884, 562)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.LblJob)
        Me.Controls.Add(Me.LblDept)
        Me.Controls.Add(Me.TxtPrimaryJob)
        Me.Controls.Add(Me.btnFindJob)
        Me.Controls.Add(Me.TxtJobCode)
        Me.Controls.Add(Me.TxtPrimaryDept)
        Me.Controls.Add(Me.btnFindDept)
        Me.Controls.Add(Me.TxtDeptCode)
        Me.Controls.Add(Me.crv1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 600)
        Me.Name = "FrmLaporan_ProfitNLoss"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Profit & Loss"
        Me.gboPeriod_II.ResumeLayout(False)
        Me.gboPeriod_II.PerformLayout()
        Me.gboPeriod_I.ResumeLayout(False)
        Me.gboPeriod_I.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents dtpFrom_I As System.Windows.Forms.DateTimePicker
    Friend WithEvents gboPeriod_II As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFrom_II As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpTo_II As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents gboPeriod_I As System.Windows.Forms.GroupBox
    Friend WithEvents dtpTo_I As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblPeriodFrom As System.Windows.Forms.Label
    Friend WithEvents LblPeriodTo As System.Windows.Forms.Label
    Friend WithEvents cboFormatLaporan As System.Windows.Forms.ComboBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents chkUseTotal As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl4 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl3 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkHideZero As System.Windows.Forms.CheckBox
    Friend WithEvents chkHideZeroInactiveAccount As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowAccountCode As System.Windows.Forms.CheckBox
    Friend WithEvents TxtDeptCode As System.Windows.Forms.TextBox
    Friend WithEvents btnFindDept As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryDept As System.Windows.Forms.TextBox
    Friend WithEvents TxtJobCode As System.Windows.Forms.TextBox
    Friend WithEvents btnFindJob As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryJob As System.Windows.Forms.TextBox
    Friend WithEvents LblDept As System.Windows.Forms.Label
    Friend WithEvents LblJob As System.Windows.Forms.Label
    Friend WithEvents chkShowCOGS As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowIncome As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowExpense As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowProfit As System.Windows.Forms.CheckBox
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents chkShowLvl8 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl7 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl6 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl5 As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
End Class
