<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterAsset
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.DtpSetFinish = New System.Windows.Forms.DateTimePicker
        Me.Label23 = New System.Windows.Forms.Label
        Me.cboDepreciateFrequency = New System.Windows.Forms.ComboBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.DtpDepOn = New System.Windows.Forms.DateTimePicker
        Me.Label14 = New System.Windows.Forms.Label
        Me.dtpStart = New System.Windows.Forms.DateTimePicker
        Me.TxtBookValue = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtSalvageValue = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtAcquiredValue = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.LblLife1 = New System.Windows.Forms.Label
        Me.TxtDepLife = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.cboDepreciationMethod = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.ChkDepThisAsset = New System.Windows.Forms.CheckBox
        Me.ChkAcquisitionFinish = New System.Windows.Forms.CheckBox
        Me.cboAcquisitionType = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TxtPrimaryAkun_Gain_Loss = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAkun_Disposal_Income = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAkun_Accum_Dep = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAkun_Dep_Expense = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAkun_Asset = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAkun_OnProcess = New System.Windows.Forms.TextBox
        Me.TxtAccDesc_DisposalIncome = New System.Windows.Forms.TextBox
        Me.TxtAkunDisposalIncome = New System.Windows.Forms.TextBox
        Me.TxtAkunOnProcess = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.BtnFindAccDisposal = New System.Windows.Forms.Button
        Me.TextBox18 = New System.Windows.Forms.TextBox
        Me.TxtAccDesc_accum_dept = New System.Windows.Forms.TextBox
        Me.TxtAkunAccumDep = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.btnFindAccAccum = New System.Windows.Forms.Button
        Me.TextBox15 = New System.Windows.Forms.TextBox
        Me.TxtAccDesc_dept_Expense = New System.Windows.Forms.TextBox
        Me.TxtAkunDepExpense = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.BtnFindAccDepAsset = New System.Windows.Forms.Button
        Me.TextBox12 = New System.Windows.Forms.TextBox
        Me.TxtAccDesc_GainLoss = New System.Windows.Forms.TextBox
        Me.TxtAkunGainLoss = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.BtnFindAccGainLoss = New System.Windows.Forms.Button
        Me.TextBox9 = New System.Windows.Forms.TextBox
        Me.TxtAccDesc_asset = New System.Windows.Forms.TextBox
        Me.TxtAkunAsset = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.BtnFindAccAsset = New System.Windows.Forms.Button
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.TxtAccDesc_OnProgress = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.btnFindAcc_OnProcess = New System.Windows.Forms.Button
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtPrimaryJurnal_SetFinish = New System.Windows.Forms.TextBox
        Me.TxtStatusEnum = New System.Windows.Forms.TextBox
        Me.TxtStatus = New System.Windows.Forms.TextBox
        Me.BtnFindGroupCode = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtAssetCode = New System.Windows.Forms.TextBox
        Me.txtGroupCode = New System.Windows.Forms.TextBox
        Me.txtPrimaryGroup = New System.Windows.Forms.TextBox
        Me.txtPrimaryAsset = New System.Windows.Forms.TextBox
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.btnFindSubBrand = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtPrimaryUnit = New System.Windows.Forms.TextBox
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnFindUnit = New System.Windows.Forms.Button
        Me.TxtQty = New System.Windows.Forms.TextBox
        Me.TxtUnit = New System.Windows.Forms.TextBox
        Me.TxtAcc_OnProcess = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tspBar1 = New System.Windows.Forms.ToolStripProgressBar
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.gboDetail.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 38)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 17)
        Me.Label1.TabIndex = 39
        Me.Label1.Text = "Remarks"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(522, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TabControl1)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(16, 7)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Size = New System.Drawing.Size(1045, 443)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(8, 130)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1029, 266)
        Me.TabControl1.TabIndex = 1
        Me.TabControl1.TabStop = False
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.SplitContainer2)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabPage1.Size = New System.Drawing.Size(1021, 237)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Acquisition && Account Data"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(4, 4)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.DtpSetFinish)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label23)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboDepreciateFrequency)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label22)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DtpDepOn)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label14)
        Me.SplitContainer2.Panel1.Controls.Add(Me.dtpStart)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtBookValue)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtSalvageValue)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label12)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtAcquiredValue)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label13)
        Me.SplitContainer2.Panel1.Controls.Add(Me.LblLife1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtDepLife)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboDepreciationMethod)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer2.Panel1.Controls.Add(Me.ChkDepThisAsset)
        Me.SplitContainer2.Panel1.Controls.Add(Me.ChkAcquisitionFinish)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboAcquisitionType)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label8)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer2.Size = New System.Drawing.Size(1013, 229)
        Me.SplitContainer2.SplitterDistance = 501
        Me.SplitContainer2.SplitterWidth = 5
        Me.SplitContainer2.TabIndex = 0
        Me.SplitContainer2.TabStop = False
        '
        'DtpSetFinish
        '
        Me.DtpSetFinish.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpSetFinish.Checked = False
        Me.DtpSetFinish.CustomFormat = "dd MMM yyyy"
        Me.DtpSetFinish.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpSetFinish.Location = New System.Drawing.Point(336, 34)
        Me.DtpSetFinish.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DtpSetFinish.Name = "DtpSetFinish"
        Me.DtpSetFinish.Size = New System.Drawing.Size(152, 22)
        Me.DtpSetFinish.TabIndex = 10
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(277, 164)
        Me.Label23.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(75, 17)
        Me.Label23.TabIndex = 66
        Me.Label23.Text = "Frequency"
        '
        'cboDepreciateFrequency
        '
        Me.cboDepreciateFrequency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDepreciateFrequency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDepreciateFrequency.FormattingEnabled = True
        Me.cboDepreciateFrequency.Items.AddRange(New Object() {"Monthly", "Yearly"})
        Me.cboDepreciateFrequency.Location = New System.Drawing.Point(380, 161)
        Me.cboDepreciateFrequency.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboDepreciateFrequency.Name = "cboDepreciateFrequency"
        Me.cboDepreciateFrequency.Size = New System.Drawing.Size(112, 24)
        Me.cboDepreciateFrequency.TabIndex = 15
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(4, 196)
        Me.Label22.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(134, 17)
        Me.Label22.TabIndex = 58
        Me.Label22.Text = "Start Depreciate On"
        '
        'DtpDepOn
        '
        Me.DtpDepOn.CustomFormat = "MMM"
        Me.DtpDepOn.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDepOn.Location = New System.Drawing.Point(380, 194)
        Me.DtpDepOn.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DtpDepOn.Name = "DtpDepOn"
        Me.DtpDepOn.ShowUpDown = True
        Me.DtpDepOn.Size = New System.Drawing.Size(109, 22)
        Me.DtpDepOn.TabIndex = 17
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(273, 196)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(97, 17)
        Me.Label14.TabIndex = 62
        Me.Label14.Text = "Depreciate on"
        '
        'dtpStart
        '
        Me.dtpStart.CustomFormat = "MMM yyyy"
        Me.dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStart.Location = New System.Drawing.Point(153, 196)
        Me.dtpStart.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.ShowUpDown = True
        Me.dtpStart.Size = New System.Drawing.Size(115, 22)
        Me.dtpStart.TabIndex = 16
        '
        'TxtBookValue
        '
        Me.TxtBookValue.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBookValue.Location = New System.Drawing.Point(153, 164)
        Me.TxtBookValue.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtBookValue.MaxLength = 50
        Me.TxtBookValue.Name = "TxtBookValue"
        Me.TxtBookValue.ReadOnly = True
        Me.TxtBookValue.Size = New System.Drawing.Size(115, 22)
        Me.TxtBookValue.TabIndex = 71
        Me.TxtBookValue.TabStop = False
        Me.TxtBookValue.Text = "0"
        Me.TxtBookValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 167)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 17)
        Me.Label3.TabIndex = 70
        Me.Label3.Text = "Book Value"
        '
        'TxtSalvageValue
        '
        Me.TxtSalvageValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSalvageValue.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSalvageValue.Location = New System.Drawing.Point(380, 132)
        Me.TxtSalvageValue.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtSalvageValue.MaxLength = 50
        Me.TxtSalvageValue.Name = "TxtSalvageValue"
        Me.TxtSalvageValue.Size = New System.Drawing.Size(110, 22)
        Me.TxtSalvageValue.TabIndex = 14
        Me.TxtSalvageValue.Text = "0"
        Me.TxtSalvageValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(277, 132)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(99, 17)
        Me.Label12.TabIndex = 51
        Me.Label12.Text = "Salvage Value"
        '
        'txtAcquiredValue
        '
        Me.txtAcquiredValue.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtAcquiredValue.Location = New System.Drawing.Point(153, 132)
        Me.txtAcquiredValue.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtAcquiredValue.MaxLength = 20
        Me.txtAcquiredValue.Name = "txtAcquiredValue"
        Me.txtAcquiredValue.Size = New System.Drawing.Size(115, 22)
        Me.txtAcquiredValue.TabIndex = 13
        Me.txtAcquiredValue.Text = "0"
        Me.txtAcquiredValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(4, 135)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(104, 17)
        Me.Label13.TabIndex = 49
        Me.Label13.Text = "Acquired Value"
        '
        'LblLife1
        '
        Me.LblLife1.AutoSize = True
        Me.LblLife1.Location = New System.Drawing.Point(441, 100)
        Me.LblLife1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblLife1.Name = "LblLife1"
        Me.LblLife1.Size = New System.Drawing.Size(54, 17)
        Me.LblLife1.TabIndex = 48
        Me.LblLife1.Text = "months"
        '
        'TxtDepLife
        '
        Me.TxtDepLife.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDepLife.Location = New System.Drawing.Point(153, 100)
        Me.TxtDepLife.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtDepLife.MaxLength = 50
        Me.TxtDepLife.Name = "TxtDepLife"
        Me.TxtDepLife.Size = New System.Drawing.Size(279, 22)
        Me.TxtDepLife.TabIndex = 12
        Me.TxtDepLife.Text = "1"
        Me.TxtDepLife.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(4, 103)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(115, 17)
        Me.Label10.TabIndex = 46
        Me.Label10.Text = "Depreciation Life"
        '
        'cboDepreciationMethod
        '
        Me.cboDepreciationMethod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDepreciationMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDepreciationMethod.FormattingEnabled = True
        Me.cboDepreciationMethod.Items.AddRange(New Object() {"Straight Line", "Sum Year Digits", "Double Declining Balance"})
        Me.cboDepreciationMethod.Location = New System.Drawing.Point(153, 66)
        Me.cboDepreciationMethod.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboDepreciationMethod.Name = "cboDepreciationMethod"
        Me.cboDepreciationMethod.Size = New System.Drawing.Size(334, 24)
        Me.cboDepreciationMethod.TabIndex = 11
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(4, 70)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(139, 17)
        Me.Label9.TabIndex = 44
        Me.Label9.Text = "Depreciation Method"
        '
        'ChkDepThisAsset
        '
        Me.ChkDepThisAsset.AutoSize = True
        Me.ChkDepThisAsset.Location = New System.Drawing.Point(4, 38)
        Me.ChkDepThisAsset.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ChkDepThisAsset.Name = "ChkDepThisAsset"
        Me.ChkDepThisAsset.Size = New System.Drawing.Size(169, 21)
        Me.ChkDepThisAsset.TabIndex = 8
        Me.ChkDepThisAsset.Text = "Depreciate This Asset"
        Me.ChkDepThisAsset.UseVisualStyleBackColor = True
        '
        'ChkAcquisitionFinish
        '
        Me.ChkAcquisitionFinish.AutoSize = True
        Me.ChkAcquisitionFinish.Location = New System.Drawing.Point(185, 38)
        Me.ChkAcquisitionFinish.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ChkAcquisitionFinish.Name = "ChkAcquisitionFinish"
        Me.ChkAcquisitionFinish.Size = New System.Drawing.Size(139, 21)
        Me.ChkAcquisitionFinish.TabIndex = 9
        Me.ChkAcquisitionFinish.Text = "Acquisition Finish"
        Me.ChkAcquisitionFinish.UseVisualStyleBackColor = True
        '
        'cboAcquisitionType
        '
        Me.cboAcquisitionType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboAcquisitionType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAcquisitionType.FormattingEnabled = True
        Me.cboAcquisitionType.Items.AddRange(New Object() {"Define Value", "Buy/Build"})
        Me.cboAcquisitionType.Location = New System.Drawing.Point(153, 5)
        Me.cboAcquisitionType.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboAcquisitionType.Name = "cboAcquisitionType"
        Me.cboAcquisitionType.Size = New System.Drawing.Size(334, 24)
        Me.cboAcquisitionType.TabIndex = 7
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(4, 9)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(109, 17)
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "Acquistion Type"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryAkun_Gain_Loss)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryAkun_Disposal_Income)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryAkun_Accum_Dep)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryAkun_Dep_Expense)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryAkun_Asset)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryAkun_OnProcess)
        Me.GroupBox1.Controls.Add(Me.TxtAccDesc_DisposalIncome)
        Me.GroupBox1.Controls.Add(Me.TxtAkunDisposalIncome)
        Me.GroupBox1.Controls.Add(Me.TxtAkunOnProcess)
        Me.GroupBox1.Controls.Add(Me.Label21)
        Me.GroupBox1.Controls.Add(Me.BtnFindAccDisposal)
        Me.GroupBox1.Controls.Add(Me.TextBox18)
        Me.GroupBox1.Controls.Add(Me.TxtAccDesc_accum_dept)
        Me.GroupBox1.Controls.Add(Me.TxtAkunAccumDep)
        Me.GroupBox1.Controls.Add(Me.Label20)
        Me.GroupBox1.Controls.Add(Me.btnFindAccAccum)
        Me.GroupBox1.Controls.Add(Me.TextBox15)
        Me.GroupBox1.Controls.Add(Me.TxtAccDesc_dept_Expense)
        Me.GroupBox1.Controls.Add(Me.TxtAkunDepExpense)
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.BtnFindAccDepAsset)
        Me.GroupBox1.Controls.Add(Me.TextBox12)
        Me.GroupBox1.Controls.Add(Me.TxtAccDesc_GainLoss)
        Me.GroupBox1.Controls.Add(Me.TxtAkunGainLoss)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.BtnFindAccGainLoss)
        Me.GroupBox1.Controls.Add(Me.TextBox9)
        Me.GroupBox1.Controls.Add(Me.TxtAccDesc_asset)
        Me.GroupBox1.Controls.Add(Me.TxtAkunAsset)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.BtnFindAccAsset)
        Me.GroupBox1.Controls.Add(Me.TextBox6)
        Me.GroupBox1.Controls.Add(Me.TxtAccDesc_OnProgress)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.btnFindAcc_OnProcess)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(507, 229)
        Me.GroupBox1.TabIndex = 42
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Account Data"
        '
        'TxtPrimaryAkun_Gain_Loss
        '
        Me.TxtPrimaryAkun_Gain_Loss.Location = New System.Drawing.Point(113, 175)
        Me.TxtPrimaryAkun_Gain_Loss.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryAkun_Gain_Loss.Name = "TxtPrimaryAkun_Gain_Loss"
        Me.TxtPrimaryAkun_Gain_Loss.Size = New System.Drawing.Size(48, 22)
        Me.TxtPrimaryAkun_Gain_Loss.TabIndex = 72
        Me.TxtPrimaryAkun_Gain_Loss.TabStop = False
        Me.TxtPrimaryAkun_Gain_Loss.Text = "-1"
        Me.TxtPrimaryAkun_Gain_Loss.Visible = False
        '
        'TxtPrimaryAkun_Disposal_Income
        '
        Me.TxtPrimaryAkun_Disposal_Income.Location = New System.Drawing.Point(113, 150)
        Me.TxtPrimaryAkun_Disposal_Income.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryAkun_Disposal_Income.Name = "TxtPrimaryAkun_Disposal_Income"
        Me.TxtPrimaryAkun_Disposal_Income.Size = New System.Drawing.Size(48, 22)
        Me.TxtPrimaryAkun_Disposal_Income.TabIndex = 71
        Me.TxtPrimaryAkun_Disposal_Income.TabStop = False
        Me.TxtPrimaryAkun_Disposal_Income.Text = "-1"
        Me.TxtPrimaryAkun_Disposal_Income.Visible = False
        '
        'TxtPrimaryAkun_Accum_Dep
        '
        Me.TxtPrimaryAkun_Accum_Dep.Location = New System.Drawing.Point(113, 118)
        Me.TxtPrimaryAkun_Accum_Dep.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryAkun_Accum_Dep.Name = "TxtPrimaryAkun_Accum_Dep"
        Me.TxtPrimaryAkun_Accum_Dep.Size = New System.Drawing.Size(48, 22)
        Me.TxtPrimaryAkun_Accum_Dep.TabIndex = 70
        Me.TxtPrimaryAkun_Accum_Dep.TabStop = False
        Me.TxtPrimaryAkun_Accum_Dep.Text = "-1"
        Me.TxtPrimaryAkun_Accum_Dep.Visible = False
        '
        'TxtPrimaryAkun_Dep_Expense
        '
        Me.TxtPrimaryAkun_Dep_Expense.Location = New System.Drawing.Point(113, 86)
        Me.TxtPrimaryAkun_Dep_Expense.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryAkun_Dep_Expense.Name = "TxtPrimaryAkun_Dep_Expense"
        Me.TxtPrimaryAkun_Dep_Expense.Size = New System.Drawing.Size(48, 22)
        Me.TxtPrimaryAkun_Dep_Expense.TabIndex = 69
        Me.TxtPrimaryAkun_Dep_Expense.TabStop = False
        Me.TxtPrimaryAkun_Dep_Expense.Text = "-1"
        Me.TxtPrimaryAkun_Dep_Expense.Visible = False
        '
        'TxtPrimaryAkun_Asset
        '
        Me.TxtPrimaryAkun_Asset.Location = New System.Drawing.Point(113, 57)
        Me.TxtPrimaryAkun_Asset.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryAkun_Asset.Name = "TxtPrimaryAkun_Asset"
        Me.TxtPrimaryAkun_Asset.Size = New System.Drawing.Size(48, 22)
        Me.TxtPrimaryAkun_Asset.TabIndex = 68
        Me.TxtPrimaryAkun_Asset.TabStop = False
        Me.TxtPrimaryAkun_Asset.Text = "-1"
        Me.TxtPrimaryAkun_Asset.Visible = False
        '
        'TxtPrimaryAkun_OnProcess
        '
        Me.TxtPrimaryAkun_OnProcess.Location = New System.Drawing.Point(113, 23)
        Me.TxtPrimaryAkun_OnProcess.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryAkun_OnProcess.Name = "TxtPrimaryAkun_OnProcess"
        Me.TxtPrimaryAkun_OnProcess.Size = New System.Drawing.Size(48, 22)
        Me.TxtPrimaryAkun_OnProcess.TabIndex = 67
        Me.TxtPrimaryAkun_OnProcess.TabStop = False
        Me.TxtPrimaryAkun_OnProcess.Text = "-1"
        Me.TxtPrimaryAkun_OnProcess.Visible = False
        '
        'TxtAccDesc_DisposalIncome
        '
        Me.TxtAccDesc_DisposalIncome.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccDesc_DisposalIncome.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAccDesc_DisposalIncome.Location = New System.Drawing.Point(284, 151)
        Me.TxtAccDesc_DisposalIncome.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAccDesc_DisposalIncome.MaxLength = 50
        Me.TxtAccDesc_DisposalIncome.Name = "TxtAccDesc_DisposalIncome"
        Me.TxtAccDesc_DisposalIncome.ReadOnly = True
        Me.TxtAccDesc_DisposalIncome.Size = New System.Drawing.Size(214, 22)
        Me.TxtAccDesc_DisposalIncome.TabIndex = 66
        Me.TxtAccDesc_DisposalIncome.TabStop = False
        '
        'TxtAkunDisposalIncome
        '
        Me.TxtAkunDisposalIncome.Location = New System.Drawing.Point(167, 150)
        Me.TxtAkunDisposalIncome.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAkunDisposalIncome.Name = "TxtAkunDisposalIncome"
        Me.TxtAkunDisposalIncome.Size = New System.Drawing.Size(65, 22)
        Me.TxtAkunDisposalIncome.TabIndex = 22
        '
        'TxtAkunOnProcess
        '
        Me.TxtAkunOnProcess.Location = New System.Drawing.Point(167, 22)
        Me.TxtAkunOnProcess.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAkunOnProcess.Name = "TxtAkunOnProcess"
        Me.TxtAkunOnProcess.Size = New System.Drawing.Size(65, 22)
        Me.TxtAkunOnProcess.TabIndex = 18
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(15, 154)
        Me.Label21.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(111, 17)
        Me.Label21.TabIndex = 62
        Me.Label21.Text = "Disposal Income"
        '
        'BtnFindAccDisposal
        '
        Me.BtnFindAccDisposal.Location = New System.Drawing.Point(241, 151)
        Me.BtnFindAccDisposal.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFindAccDisposal.Name = "BtnFindAccDisposal"
        Me.BtnFindAccDisposal.Size = New System.Drawing.Size(35, 25)
        Me.BtnFindAccDisposal.TabIndex = 64
        Me.BtnFindAccDisposal.TabStop = False
        Me.BtnFindAccDisposal.Text = "...."
        Me.BtnFindAccDisposal.UseVisualStyleBackColor = True
        '
        'TextBox18
        '
        Me.TextBox18.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox18.Location = New System.Drawing.Point(167, 150)
        Me.TextBox18.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox18.MaxLength = 50
        Me.TextBox18.Name = "TextBox18"
        Me.TextBox18.Size = New System.Drawing.Size(60, 22)
        Me.TextBox18.TabIndex = 63
        '
        'TxtAccDesc_accum_dept
        '
        Me.TxtAccDesc_accum_dept.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccDesc_accum_dept.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAccDesc_accum_dept.Location = New System.Drawing.Point(284, 119)
        Me.TxtAccDesc_accum_dept.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAccDesc_accum_dept.MaxLength = 50
        Me.TxtAccDesc_accum_dept.Name = "TxtAccDesc_accum_dept"
        Me.TxtAccDesc_accum_dept.ReadOnly = True
        Me.TxtAccDesc_accum_dept.Size = New System.Drawing.Size(214, 22)
        Me.TxtAccDesc_accum_dept.TabIndex = 61
        Me.TxtAccDesc_accum_dept.TabStop = False
        '
        'TxtAkunAccumDep
        '
        Me.TxtAkunAccumDep.Location = New System.Drawing.Point(167, 118)
        Me.TxtAkunAccumDep.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAkunAccumDep.Name = "TxtAkunAccumDep"
        Me.TxtAkunAccumDep.Size = New System.Drawing.Size(65, 22)
        Me.TxtAkunAccumDep.TabIndex = 21
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(15, 122)
        Me.Label20.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(138, 17)
        Me.Label20.TabIndex = 57
        Me.Label20.Text = "Accum. Depreciation"
        '
        'btnFindAccAccum
        '
        Me.btnFindAccAccum.Location = New System.Drawing.Point(241, 119)
        Me.btnFindAccAccum.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFindAccAccum.Name = "btnFindAccAccum"
        Me.btnFindAccAccum.Size = New System.Drawing.Size(35, 25)
        Me.btnFindAccAccum.TabIndex = 59
        Me.btnFindAccAccum.TabStop = False
        Me.btnFindAccAccum.Text = "...."
        Me.btnFindAccAccum.UseVisualStyleBackColor = True
        '
        'TextBox15
        '
        Me.TextBox15.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox15.Location = New System.Drawing.Point(167, 118)
        Me.TextBox15.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox15.MaxLength = 50
        Me.TextBox15.Name = "TextBox15"
        Me.TextBox15.Size = New System.Drawing.Size(60, 22)
        Me.TextBox15.TabIndex = 58
        '
        'TxtAccDesc_dept_Expense
        '
        Me.TxtAccDesc_dept_Expense.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccDesc_dept_Expense.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAccDesc_dept_Expense.Location = New System.Drawing.Point(284, 87)
        Me.TxtAccDesc_dept_Expense.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAccDesc_dept_Expense.MaxLength = 50
        Me.TxtAccDesc_dept_Expense.Name = "TxtAccDesc_dept_Expense"
        Me.TxtAccDesc_dept_Expense.ReadOnly = True
        Me.TxtAccDesc_dept_Expense.Size = New System.Drawing.Size(214, 22)
        Me.TxtAccDesc_dept_Expense.TabIndex = 56
        Me.TxtAccDesc_dept_Expense.TabStop = False
        '
        'TxtAkunDepExpense
        '
        Me.TxtAkunDepExpense.Location = New System.Drawing.Point(167, 86)
        Me.TxtAkunDepExpense.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAkunDepExpense.Name = "TxtAkunDepExpense"
        Me.TxtAkunDepExpense.Size = New System.Drawing.Size(65, 22)
        Me.TxtAkunDepExpense.TabIndex = 20
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(15, 90)
        Me.Label19.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(146, 17)
        Me.Label19.TabIndex = 52
        Me.Label19.Text = "Depreciation Expense"
        '
        'BtnFindAccDepAsset
        '
        Me.BtnFindAccDepAsset.Location = New System.Drawing.Point(241, 87)
        Me.BtnFindAccDepAsset.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFindAccDepAsset.Name = "BtnFindAccDepAsset"
        Me.BtnFindAccDepAsset.Size = New System.Drawing.Size(35, 25)
        Me.BtnFindAccDepAsset.TabIndex = 54
        Me.BtnFindAccDepAsset.TabStop = False
        Me.BtnFindAccDepAsset.Text = "...."
        Me.BtnFindAccDepAsset.UseVisualStyleBackColor = True
        '
        'TextBox12
        '
        Me.TextBox12.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox12.Location = New System.Drawing.Point(167, 86)
        Me.TextBox12.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox12.MaxLength = 50
        Me.TextBox12.Name = "TextBox12"
        Me.TextBox12.Size = New System.Drawing.Size(60, 22)
        Me.TextBox12.TabIndex = 53
        '
        'TxtAccDesc_GainLoss
        '
        Me.TxtAccDesc_GainLoss.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccDesc_GainLoss.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAccDesc_GainLoss.Location = New System.Drawing.Point(284, 181)
        Me.TxtAccDesc_GainLoss.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAccDesc_GainLoss.MaxLength = 50
        Me.TxtAccDesc_GainLoss.Name = "TxtAccDesc_GainLoss"
        Me.TxtAccDesc_GainLoss.ReadOnly = True
        Me.TxtAccDesc_GainLoss.Size = New System.Drawing.Size(214, 22)
        Me.TxtAccDesc_GainLoss.TabIndex = 51
        Me.TxtAccDesc_GainLoss.TabStop = False
        '
        'TxtAkunGainLoss
        '
        Me.TxtAkunGainLoss.Location = New System.Drawing.Point(167, 180)
        Me.TxtAkunGainLoss.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAkunGainLoss.Name = "TxtAkunGainLoss"
        Me.TxtAkunGainLoss.Size = New System.Drawing.Size(65, 22)
        Me.TxtAkunGainLoss.TabIndex = 23
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(15, 183)
        Me.Label18.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(130, 17)
        Me.Label18.TabIndex = 47
        Me.Label18.Text = "Gain/Loss Disposal"
        '
        'BtnFindAccGainLoss
        '
        Me.BtnFindAccGainLoss.Location = New System.Drawing.Point(241, 181)
        Me.BtnFindAccGainLoss.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFindAccGainLoss.Name = "BtnFindAccGainLoss"
        Me.BtnFindAccGainLoss.Size = New System.Drawing.Size(35, 25)
        Me.BtnFindAccGainLoss.TabIndex = 49
        Me.BtnFindAccGainLoss.TabStop = False
        Me.BtnFindAccGainLoss.Text = "...."
        Me.BtnFindAccGainLoss.UseVisualStyleBackColor = True
        '
        'TextBox9
        '
        Me.TextBox9.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox9.Location = New System.Drawing.Point(167, 180)
        Me.TextBox9.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox9.MaxLength = 50
        Me.TextBox9.Name = "TextBox9"
        Me.TextBox9.Size = New System.Drawing.Size(60, 22)
        Me.TextBox9.TabIndex = 48
        '
        'TxtAccDesc_asset
        '
        Me.TxtAccDesc_asset.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccDesc_asset.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAccDesc_asset.Location = New System.Drawing.Point(284, 55)
        Me.TxtAccDesc_asset.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAccDesc_asset.MaxLength = 50
        Me.TxtAccDesc_asset.Name = "TxtAccDesc_asset"
        Me.TxtAccDesc_asset.ReadOnly = True
        Me.TxtAccDesc_asset.Size = New System.Drawing.Size(214, 22)
        Me.TxtAccDesc_asset.TabIndex = 46
        Me.TxtAccDesc_asset.TabStop = False
        '
        'TxtAkunAsset
        '
        Me.TxtAkunAsset.Location = New System.Drawing.Point(167, 54)
        Me.TxtAkunAsset.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAkunAsset.Name = "TxtAkunAsset"
        Me.TxtAkunAsset.Size = New System.Drawing.Size(65, 22)
        Me.TxtAkunAsset.TabIndex = 19
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(15, 58)
        Me.Label17.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(43, 17)
        Me.Label17.TabIndex = 42
        Me.Label17.Text = "Asset"
        '
        'BtnFindAccAsset
        '
        Me.BtnFindAccAsset.Location = New System.Drawing.Point(241, 55)
        Me.BtnFindAccAsset.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFindAccAsset.Name = "BtnFindAccAsset"
        Me.BtnFindAccAsset.Size = New System.Drawing.Size(35, 25)
        Me.BtnFindAccAsset.TabIndex = 44
        Me.BtnFindAccAsset.TabStop = False
        Me.BtnFindAccAsset.Text = "...."
        Me.BtnFindAccAsset.UseVisualStyleBackColor = True
        '
        'TextBox6
        '
        Me.TextBox6.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox6.Location = New System.Drawing.Point(167, 54)
        Me.TextBox6.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox6.MaxLength = 50
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(60, 22)
        Me.TextBox6.TabIndex = 43
        '
        'TxtAccDesc_OnProgress
        '
        Me.TxtAccDesc_OnProgress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccDesc_OnProgress.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAccDesc_OnProgress.Location = New System.Drawing.Point(284, 23)
        Me.TxtAccDesc_OnProgress.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAccDesc_OnProgress.MaxLength = 50
        Me.TxtAccDesc_OnProgress.Name = "TxtAccDesc_OnProgress"
        Me.TxtAccDesc_OnProgress.ReadOnly = True
        Me.TxtAccDesc_OnProgress.Size = New System.Drawing.Size(214, 22)
        Me.TxtAccDesc_OnProgress.TabIndex = 41
        Me.TxtAccDesc_OnProgress.TabStop = False
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(15, 26)
        Me.Label16.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(82, 17)
        Me.Label16.TabIndex = 12
        Me.Label16.Text = "On Process"
        '
        'btnFindAcc_OnProcess
        '
        Me.btnFindAcc_OnProcess.Location = New System.Drawing.Point(241, 23)
        Me.btnFindAcc_OnProcess.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFindAcc_OnProcess.Name = "btnFindAcc_OnProcess"
        Me.btnFindAcc_OnProcess.Size = New System.Drawing.Size(35, 25)
        Me.btnFindAcc_OnProcess.TabIndex = 39
        Me.btnFindAcc_OnProcess.TabStop = False
        Me.btnFindAcc_OnProcess.Text = "...."
        Me.btnFindAcc_OnProcess.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabPage2.Size = New System.Drawing.Size(1021, 237)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "History"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(8, 23)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryJurnal_SetFinish)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtStatusEnum)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtStatus)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindGroupCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtAssetCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtGroupCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryGroup)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryAsset)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtDescription)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindSubBrand)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryUnit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindUnit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtQty)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtUnit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtAcc_OnProcess)
        Me.SplitContainer1.Size = New System.Drawing.Size(1029, 100)
        Me.SplitContainer1.SplitterDistance = 511
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'TxtPrimaryJurnal_SetFinish
        '
        Me.TxtPrimaryJurnal_SetFinish.Location = New System.Drawing.Point(347, 7)
        Me.TxtPrimaryJurnal_SetFinish.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryJurnal_SetFinish.Name = "TxtPrimaryJurnal_SetFinish"
        Me.TxtPrimaryJurnal_SetFinish.Size = New System.Drawing.Size(48, 22)
        Me.TxtPrimaryJurnal_SetFinish.TabIndex = 48
        Me.TxtPrimaryJurnal_SetFinish.TabStop = False
        Me.TxtPrimaryJurnal_SetFinish.Text = "-1"
        Me.TxtPrimaryJurnal_SetFinish.Visible = False
        '
        'TxtStatusEnum
        '
        Me.TxtStatusEnum.Location = New System.Drawing.Point(92, 39)
        Me.TxtStatusEnum.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtStatusEnum.Name = "TxtStatusEnum"
        Me.TxtStatusEnum.Size = New System.Drawing.Size(48, 22)
        Me.TxtStatusEnum.TabIndex = 46
        Me.TxtStatusEnum.TabStop = False
        Me.TxtStatusEnum.Text = "-1"
        Me.TxtStatusEnum.Visible = False
        '
        'TxtStatus
        '
        Me.TxtStatus.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtStatus.Location = New System.Drawing.Point(346, 7)
        Me.TxtStatus.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtStatus.MaxLength = 45
        Me.TxtStatus.Name = "TxtStatus"
        Me.TxtStatus.ReadOnly = True
        Me.TxtStatus.Size = New System.Drawing.Size(153, 22)
        Me.TxtStatus.TabIndex = 45
        Me.TxtStatus.TabStop = False
        '
        'BtnFindGroupCode
        '
        Me.BtnFindGroupCode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindGroupCode.Location = New System.Drawing.Point(446, 73)
        Me.BtnFindGroupCode.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFindGroupCode.Name = "BtnFindGroupCode"
        Me.BtnFindGroupCode.Size = New System.Drawing.Size(52, 25)
        Me.BtnFindGroupCode.TabIndex = 44
        Me.BtnFindGroupCode.TabStop = False
        Me.BtnFindGroupCode.Text = "...."
        Me.BtnFindGroupCode.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 11)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 17)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Asset Code"
        '
        'txtAssetCode
        '
        Me.txtAssetCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAssetCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtAssetCode.Location = New System.Drawing.Point(163, 7)
        Me.txtAssetCode.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtAssetCode.MaxLength = 150
        Me.txtAssetCode.Name = "txtAssetCode"
        Me.txtAssetCode.Size = New System.Drawing.Size(174, 22)
        Me.txtAssetCode.TabIndex = 0
        '
        'txtGroupCode
        '
        Me.txtGroupCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtGroupCode.Location = New System.Drawing.Point(163, 71)
        Me.txtGroupCode.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtGroupCode.MaxLength = 50
        Me.txtGroupCode.Name = "txtGroupCode"
        Me.txtGroupCode.Size = New System.Drawing.Size(274, 22)
        Me.txtGroupCode.TabIndex = 2
        '
        'txtPrimaryGroup
        '
        Me.txtPrimaryGroup.Location = New System.Drawing.Point(93, 73)
        Me.txtPrimaryGroup.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimaryGroup.Multiline = True
        Me.txtPrimaryGroup.Name = "txtPrimaryGroup"
        Me.txtPrimaryGroup.Size = New System.Drawing.Size(48, 24)
        Me.txtPrimaryGroup.TabIndex = 37
        Me.txtPrimaryGroup.TabStop = False
        Me.txtPrimaryGroup.Text = "-1"
        Me.txtPrimaryGroup.Visible = False
        '
        'txtPrimaryAsset
        '
        Me.txtPrimaryAsset.Location = New System.Drawing.Point(93, 7)
        Me.txtPrimaryAsset.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimaryAsset.Name = "txtPrimaryAsset"
        Me.txtPrimaryAsset.Size = New System.Drawing.Size(48, 22)
        Me.txtPrimaryAsset.TabIndex = 36
        Me.txtPrimaryAsset.TabStop = False
        Me.txtPrimaryAsset.Text = "-1"
        Me.txtPrimaryAsset.Visible = False
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(163, 39)
        Me.txtDescription.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtDescription.MaxLength = 150
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(335, 22)
        Me.txtDescription.TabIndex = 1
        '
        'btnFindSubBrand
        '
        Me.btnFindSubBrand.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindSubBrand.Location = New System.Drawing.Point(447, 71)
        Me.btnFindSubBrand.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFindSubBrand.Name = "btnFindSubBrand"
        Me.btnFindSubBrand.Size = New System.Drawing.Size(52, 25)
        Me.btnFindSubBrand.TabIndex = 3
        Me.btnFindSubBrand.TabStop = False
        Me.btnFindSubBrand.Text = "...."
        Me.btnFindSubBrand.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(4, 43)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(79, 17)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Description"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 75)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(85, 17)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Group Code"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(299, 9)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(33, 17)
        Me.Label7.TabIndex = 47
        Me.Label7.Text = "Unit"
        '
        'TxtPrimaryUnit
        '
        Me.TxtPrimaryUnit.Location = New System.Drawing.Point(397, 7)
        Me.TxtPrimaryUnit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryUnit.Multiline = True
        Me.TxtPrimaryUnit.Name = "TxtPrimaryUnit"
        Me.TxtPrimaryUnit.Size = New System.Drawing.Size(48, 24)
        Me.TxtPrimaryUnit.TabIndex = 45
        Me.TxtPrimaryUnit.TabStop = False
        Me.TxtPrimaryUnit.Text = "-1"
        Me.TxtPrimaryUnit.Visible = False
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.Location = New System.Drawing.Point(157, 34)
        Me.txtRemarks.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtRemarks.MaxLength = 150
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(349, 22)
        Me.txtRemarks.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 7)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(30, 17)
        Me.Label6.TabIndex = 44
        Me.Label6.Text = "Qty"
        '
        'btnFindUnit
        '
        Me.btnFindUnit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindUnit.Location = New System.Drawing.Point(456, 4)
        Me.btnFindUnit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFindUnit.Name = "btnFindUnit"
        Me.btnFindUnit.Size = New System.Drawing.Size(52, 25)
        Me.btnFindUnit.TabIndex = 43
        Me.btnFindUnit.TabStop = False
        Me.btnFindUnit.Text = "...."
        Me.btnFindUnit.UseVisualStyleBackColor = True
        '
        'TxtQty
        '
        Me.TxtQty.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtQty.Location = New System.Drawing.Point(157, 4)
        Me.TxtQty.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtQty.MaxLength = 7
        Me.TxtQty.Name = "TxtQty"
        Me.TxtQty.Size = New System.Drawing.Size(132, 22)
        Me.TxtQty.TabIndex = 3
        Me.TxtQty.Text = "0.00"
        Me.TxtQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtUnit
        '
        Me.TxtUnit.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUnit.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtUnit.Location = New System.Drawing.Point(341, 5)
        Me.TxtUnit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtUnit.MaxLength = 75
        Me.TxtUnit.Name = "TxtUnit"
        Me.TxtUnit.Size = New System.Drawing.Size(105, 22)
        Me.TxtUnit.TabIndex = 4
        '
        'TxtAcc_OnProcess
        '
        Me.TxtAcc_OnProcess.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAcc_OnProcess.Location = New System.Drawing.Point(49, 71)
        Me.TxtAcc_OnProcess.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtAcc_OnProcess.MaxLength = 50
        Me.TxtAcc_OnProcess.Name = "TxtAcc_OnProcess"
        Me.TxtAcc_OnProcess.Size = New System.Drawing.Size(60, 22)
        Me.TxtAcc_OnProcess.TabIndex = 38
        Me.TxtAcc_OnProcess.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(531, 404)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(117, 32)
        Me.btnCancel.TabIndex = 25
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(397, 404)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(117, 32)
        Me.btnOK.TabIndex = 24
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 493)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip2.Size = New System.Drawing.Size(1064, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 24
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(522, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(522, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(522, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate, Me.tspBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 518)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(1064, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 23
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tspBar1
        '
        Me.tspBar1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tspBar1.Name = "tspBar1"
        Me.tspBar1.Size = New System.Drawing.Size(133, 20)
        Me.tspBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.tspBar1.Visible = False
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(480, 453)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(117, 32)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(352, 453)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(117, 32)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(224, 453)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(117, 32)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(736, 453)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(117, 32)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(608, 453)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(117, 32)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'FrmMasterAsset
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(1064, 543)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimumSize = New System.Drawing.Size(661, 353)
        Me.Name = "FrmMasterAsset"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Asset Data"
        Me.gboDetail.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryGroup As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryAsset As System.Windows.Forms.TextBox
    Friend WithEvents btnFindSubBrand As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtAssetCode As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tspBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtQty As System.Windows.Forms.TextBox
    Friend WithEvents TxtUnit As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryUnit As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnFindUnit As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents cboAcquisitionType As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents cboDepreciationMethod As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ChkDepThisAsset As System.Windows.Forms.CheckBox
    Friend WithEvents ChkAcquisitionFinish As System.Windows.Forms.CheckBox
    Friend WithEvents LblLife1 As System.Windows.Forms.Label
    Friend WithEvents TxtDepLife As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtSalvageValue As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtAcquiredValue As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cboDepreciateFrequency As System.Windows.Forms.ComboBox
    Friend WithEvents DtpDepOn As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents dtpStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtAccDesc_DisposalIncome As System.Windows.Forms.TextBox
    Friend WithEvents TxtAkunDisposalIncome As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents BtnFindAccDisposal As System.Windows.Forms.Button
    Friend WithEvents TextBox18 As System.Windows.Forms.TextBox
    Friend WithEvents TxtAccDesc_accum_dept As System.Windows.Forms.TextBox
    Friend WithEvents TxtAkunAccumDep As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents btnFindAccAccum As System.Windows.Forms.Button
    Friend WithEvents TextBox15 As System.Windows.Forms.TextBox
    Friend WithEvents TxtAccDesc_dept_Expense As System.Windows.Forms.TextBox
    Friend WithEvents TxtAkunDepExpense As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents BtnFindAccDepAsset As System.Windows.Forms.Button
    Friend WithEvents TextBox12 As System.Windows.Forms.TextBox
    Friend WithEvents TxtAccDesc_GainLoss As System.Windows.Forms.TextBox
    Friend WithEvents TxtAkunGainLoss As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents BtnFindAccGainLoss As System.Windows.Forms.Button
    Friend WithEvents TextBox9 As System.Windows.Forms.TextBox
    Friend WithEvents TxtAccDesc_asset As System.Windows.Forms.TextBox
    Friend WithEvents TxtAkunAsset As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents BtnFindAccAsset As System.Windows.Forms.Button
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TxtAccDesc_OnProgress As System.Windows.Forms.TextBox
    Friend WithEvents TxtAkunOnProcess As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents btnFindAcc_OnProcess As System.Windows.Forms.Button
    Friend WithEvents TxtAcc_OnProcess As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents TxtStatus As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindGroupCode As System.Windows.Forms.Button
    Friend WithEvents TxtStatusEnum As System.Windows.Forms.TextBox
    Friend WithEvents TxtBookValue As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryJurnal_SetFinish As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAkun_Gain_Loss As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAkun_Disposal_Income As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAkun_Accum_Dep As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAkun_Dep_Expense As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAkun_Asset As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAkun_OnProcess As System.Windows.Forms.TextBox
    Friend WithEvents DtpSetFinish As System.Windows.Forms.DateTimePicker
End Class
