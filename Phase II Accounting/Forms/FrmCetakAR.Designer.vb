<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCetakAR
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
        Me.btnFindCurrency = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.txtPrimaryArAccount = New System.Windows.Forms.TextBox
        Me.txtPrimaryPaymentAccount = New System.Windows.Forms.TextBox
        Me.txtPrimaryCustomerGroup = New System.Windows.Forms.TextBox
        Me.txtPrimaryCustomerCode = New System.Windows.Forms.TextBox
        Me.txtCustomerGroupDesc = New System.Windows.Forms.TextBox
        Me.btnFindCustomerGroup = New System.Windows.Forms.Button
        Me.txtCustomerGroupCode = New System.Windows.Forms.TextBox
        Me.txtCurrencyDesc = New System.Windows.Forms.TextBox
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.txtArAccountDesc = New System.Windows.Forms.TextBox
        Me.btnFindArAccount = New System.Windows.Forms.Button
        Me.txtArAccountCode = New System.Windows.Forms.TextBox
        Me.txtPaymentAccountDesc = New System.Windows.Forms.TextBox
        Me.btnFindPaymentAccount = New System.Windows.Forms.Button
        Me.txtPaymentAccountCode = New System.Windows.Forms.TextBox
        Me.txtCustomerDesc = New System.Windows.Forms.TextBox
        Me.btnFindCustomerCode = New System.Windows.Forms.Button
        Me.txtCustomerCode = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.rboDetail = New System.Windows.Forms.RadioButton
        Me.rboPaymentJournal = New System.Windows.Forms.RadioButton
        Me.rboPaymentAccount = New System.Windows.Forms.RadioButton
        Me.rboCurrency = New System.Windows.Forms.RadioButton
        Me.rboArAccount = New System.Windows.Forms.RadioButton
        Me.rboCustGroup = New System.Windows.Forms.RadioButton
        Me.rboCustCode = New System.Windows.Forms.RadioButton
        Me.btnClose = New System.Windows.Forms.Button
        Me.crvAR = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.btnShow = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnFindCurrency
        '
        Me.btnFindCurrency.Location = New System.Drawing.Point(229, 110)
        Me.btnFindCurrency.Name = "btnFindCurrency"
        Me.btnFindCurrency.Size = New System.Drawing.Size(31, 23)
        Me.btnFindCurrency.TabIndex = 9
        Me.btnFindCurrency.Text = "..."
        Me.btnFindCurrency.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 9)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(929, 158)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.txtPrimaryCurrency)
        Me.GroupBox3.Controls.Add(Me.txtPrimaryArAccount)
        Me.GroupBox3.Controls.Add(Me.txtPrimaryPaymentAccount)
        Me.GroupBox3.Controls.Add(Me.txtPrimaryCustomerGroup)
        Me.GroupBox3.Controls.Add(Me.txtPrimaryCustomerCode)
        Me.GroupBox3.Controls.Add(Me.txtCustomerGroupDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindCustomerGroup)
        Me.GroupBox3.Controls.Add(Me.txtCustomerGroupCode)
        Me.GroupBox3.Controls.Add(Me.txtCurrencyDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindCurrency)
        Me.GroupBox3.Controls.Add(Me.txtCurrencyCode)
        Me.GroupBox3.Controls.Add(Me.txtArAccountDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindArAccount)
        Me.GroupBox3.Controls.Add(Me.txtArAccountCode)
        Me.GroupBox3.Controls.Add(Me.txtPaymentAccountDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindPaymentAccount)
        Me.GroupBox3.Controls.Add(Me.txtPaymentAccountCode)
        Me.GroupBox3.Controls.Add(Me.txtCustomerDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindCustomerCode)
        Me.GroupBox3.Controls.Add(Me.txtCustomerCode)
        Me.GroupBox3.Location = New System.Drawing.Point(375, 10)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(548, 141)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Filter"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(15, 42)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(83, 13)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Customer Group"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(15, 115)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Currency"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(15, 91)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "AR Account"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(15, 66)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(91, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Payment Account"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(15, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(79, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Customer Code"
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(94, 112)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimaryCurrency.TabIndex = 24
        Me.txtPrimaryCurrency.TabStop = False
        Me.txtPrimaryCurrency.Text = "-1"
        Me.txtPrimaryCurrency.Visible = False
        '
        'txtPrimaryArAccount
        '
        Me.txtPrimaryArAccount.Location = New System.Drawing.Point(94, 88)
        Me.txtPrimaryArAccount.Name = "txtPrimaryArAccount"
        Me.txtPrimaryArAccount.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimaryArAccount.TabIndex = 23
        Me.txtPrimaryArAccount.TabStop = False
        Me.txtPrimaryArAccount.Text = "-1"
        Me.txtPrimaryArAccount.Visible = False
        '
        'txtPrimaryPaymentAccount
        '
        Me.txtPrimaryPaymentAccount.Location = New System.Drawing.Point(94, 63)
        Me.txtPrimaryPaymentAccount.Name = "txtPrimaryPaymentAccount"
        Me.txtPrimaryPaymentAccount.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimaryPaymentAccount.TabIndex = 22
        Me.txtPrimaryPaymentAccount.TabStop = False
        Me.txtPrimaryPaymentAccount.Text = "-1"
        Me.txtPrimaryPaymentAccount.Visible = False
        '
        'txtPrimaryCustomerGroup
        '
        Me.txtPrimaryCustomerGroup.Location = New System.Drawing.Point(94, 39)
        Me.txtPrimaryCustomerGroup.Name = "txtPrimaryCustomerGroup"
        Me.txtPrimaryCustomerGroup.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimaryCustomerGroup.TabIndex = 21
        Me.txtPrimaryCustomerGroup.TabStop = False
        Me.txtPrimaryCustomerGroup.Text = "-1"
        Me.txtPrimaryCustomerGroup.Visible = False
        '
        'txtPrimaryCustomerCode
        '
        Me.txtPrimaryCustomerCode.Location = New System.Drawing.Point(94, 16)
        Me.txtPrimaryCustomerCode.Name = "txtPrimaryCustomerCode"
        Me.txtPrimaryCustomerCode.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimaryCustomerCode.TabIndex = 20
        Me.txtPrimaryCustomerCode.TabStop = False
        Me.txtPrimaryCustomerCode.Text = "-1"
        Me.txtPrimaryCustomerCode.Visible = False
        '
        'txtCustomerGroupDesc
        '
        Me.txtCustomerGroupDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCustomerGroupDesc.Location = New System.Drawing.Point(266, 39)
        Me.txtCustomerGroupDesc.Name = "txtCustomerGroupDesc"
        Me.txtCustomerGroupDesc.ReadOnly = True
        Me.txtCustomerGroupDesc.Size = New System.Drawing.Size(277, 20)
        Me.txtCustomerGroupDesc.TabIndex = 19
        Me.txtCustomerGroupDesc.TabStop = False
        '
        'btnFindCustomerGroup
        '
        Me.btnFindCustomerGroup.Location = New System.Drawing.Point(229, 37)
        Me.btnFindCustomerGroup.Name = "btnFindCustomerGroup"
        Me.btnFindCustomerGroup.Size = New System.Drawing.Size(31, 23)
        Me.btnFindCustomerGroup.TabIndex = 3
        Me.btnFindCustomerGroup.Text = "..."
        Me.btnFindCustomerGroup.UseVisualStyleBackColor = True
        '
        'txtCustomerGroupCode
        '
        Me.txtCustomerGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerGroupCode.Location = New System.Drawing.Point(124, 39)
        Me.txtCustomerGroupCode.Name = "txtCustomerGroupCode"
        Me.txtCustomerGroupCode.Size = New System.Drawing.Size(100, 20)
        Me.txtCustomerGroupCode.TabIndex = 2
        '
        'txtCurrencyDesc
        '
        Me.txtCurrencyDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCurrencyDesc.Location = New System.Drawing.Point(266, 112)
        Me.txtCurrencyDesc.Name = "txtCurrencyDesc"
        Me.txtCurrencyDesc.ReadOnly = True
        Me.txtCurrencyDesc.Size = New System.Drawing.Size(277, 20)
        Me.txtCurrencyDesc.TabIndex = 15
        Me.txtCurrencyDesc.TabStop = False
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCurrencyCode.Location = New System.Drawing.Point(124, 112)
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.Size = New System.Drawing.Size(100, 20)
        Me.txtCurrencyCode.TabIndex = 8
        '
        'txtArAccountDesc
        '
        Me.txtArAccountDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtArAccountDesc.Location = New System.Drawing.Point(266, 88)
        Me.txtArAccountDesc.Name = "txtArAccountDesc"
        Me.txtArAccountDesc.ReadOnly = True
        Me.txtArAccountDesc.Size = New System.Drawing.Size(277, 20)
        Me.txtArAccountDesc.TabIndex = 11
        Me.txtArAccountDesc.TabStop = False
        '
        'btnFindArAccount
        '
        Me.btnFindArAccount.Location = New System.Drawing.Point(229, 86)
        Me.btnFindArAccount.Name = "btnFindArAccount"
        Me.btnFindArAccount.Size = New System.Drawing.Size(31, 23)
        Me.btnFindArAccount.TabIndex = 7
        Me.btnFindArAccount.Text = "..."
        Me.btnFindArAccount.UseVisualStyleBackColor = True
        '
        'txtArAccountCode
        '
        Me.txtArAccountCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtArAccountCode.Location = New System.Drawing.Point(124, 88)
        Me.txtArAccountCode.Name = "txtArAccountCode"
        Me.txtArAccountCode.Size = New System.Drawing.Size(100, 20)
        Me.txtArAccountCode.TabIndex = 6
        '
        'txtPaymentAccountDesc
        '
        Me.txtPaymentAccountDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPaymentAccountDesc.Location = New System.Drawing.Point(266, 63)
        Me.txtPaymentAccountDesc.Name = "txtPaymentAccountDesc"
        Me.txtPaymentAccountDesc.ReadOnly = True
        Me.txtPaymentAccountDesc.Size = New System.Drawing.Size(277, 20)
        Me.txtPaymentAccountDesc.TabIndex = 7
        Me.txtPaymentAccountDesc.TabStop = False
        '
        'btnFindPaymentAccount
        '
        Me.btnFindPaymentAccount.Location = New System.Drawing.Point(229, 61)
        Me.btnFindPaymentAccount.Name = "btnFindPaymentAccount"
        Me.btnFindPaymentAccount.Size = New System.Drawing.Size(31, 23)
        Me.btnFindPaymentAccount.TabIndex = 5
        Me.btnFindPaymentAccount.Text = "..."
        Me.btnFindPaymentAccount.UseVisualStyleBackColor = True
        '
        'txtPaymentAccountCode
        '
        Me.txtPaymentAccountCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPaymentAccountCode.Location = New System.Drawing.Point(124, 63)
        Me.txtPaymentAccountCode.Name = "txtPaymentAccountCode"
        Me.txtPaymentAccountCode.Size = New System.Drawing.Size(100, 20)
        Me.txtPaymentAccountCode.TabIndex = 4
        '
        'txtCustomerDesc
        '
        Me.txtCustomerDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCustomerDesc.Location = New System.Drawing.Point(266, 16)
        Me.txtCustomerDesc.Name = "txtCustomerDesc"
        Me.txtCustomerDesc.ReadOnly = True
        Me.txtCustomerDesc.Size = New System.Drawing.Size(277, 20)
        Me.txtCustomerDesc.TabIndex = 3
        Me.txtCustomerDesc.TabStop = False
        '
        'btnFindCustomerCode
        '
        Me.btnFindCustomerCode.Location = New System.Drawing.Point(229, 14)
        Me.btnFindCustomerCode.Name = "btnFindCustomerCode"
        Me.btnFindCustomerCode.Size = New System.Drawing.Size(31, 23)
        Me.btnFindCustomerCode.TabIndex = 1
        Me.btnFindCustomerCode.Text = "..."
        Me.btnFindCustomerCode.UseVisualStyleBackColor = True
        '
        'txtCustomerCode
        '
        Me.txtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerCode.Location = New System.Drawing.Point(124, 16)
        Me.txtCustomerCode.Name = "txtCustomerCode"
        Me.txtCustomerCode.Size = New System.Drawing.Size(100, 20)
        Me.txtCustomerCode.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rboDetail)
        Me.GroupBox2.Controls.Add(Me.btnShow)
        Me.GroupBox2.Controls.Add(Me.rboPaymentJournal)
        Me.GroupBox2.Controls.Add(Me.rboPaymentAccount)
        Me.GroupBox2.Controls.Add(Me.dtpTo)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.rboCurrency)
        Me.GroupBox2.Controls.Add(Me.dtpFrom)
        Me.GroupBox2.Controls.Add(Me.rboArAccount)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.rboCustGroup)
        Me.GroupBox2.Controls.Add(Me.rboCustCode)
        Me.GroupBox2.Location = New System.Drawing.Point(9, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(360, 148)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Group By"
        '
        'rboDetail
        '
        Me.rboDetail.AutoSize = True
        Me.rboDetail.Location = New System.Drawing.Point(140, 114)
        Me.rboDetail.Name = "rboDetail"
        Me.rboDetail.Size = New System.Drawing.Size(127, 17)
        Me.rboDetail.TabIndex = 6
        Me.rboDetail.Text = "Customer Code Detail"
        Me.rboDetail.UseVisualStyleBackColor = True
        '
        'rboPaymentJournal
        '
        Me.rboPaymentJournal.AutoSize = True
        Me.rboPaymentJournal.Location = New System.Drawing.Point(257, 68)
        Me.rboPaymentJournal.Name = "rboPaymentJournal"
        Me.rboPaymentJournal.Size = New System.Drawing.Size(103, 17)
        Me.rboPaymentJournal.TabIndex = 5
        Me.rboPaymentJournal.Text = "Payment Journal"
        Me.rboPaymentJournal.UseVisualStyleBackColor = True
        Me.rboPaymentJournal.Visible = False
        '
        'rboPaymentAccount
        '
        Me.rboPaymentAccount.AutoSize = True
        Me.rboPaymentAccount.Location = New System.Drawing.Point(12, 114)
        Me.rboPaymentAccount.Name = "rboPaymentAccount"
        Me.rboPaymentAccount.Size = New System.Drawing.Size(109, 17)
        Me.rboPaymentAccount.TabIndex = 2
        Me.rboPaymentAccount.Text = "Payment Account"
        Me.rboPaymentAccount.UseVisualStyleBackColor = True
        '
        'rboCurrency
        '
        Me.rboCurrency.AutoSize = True
        Me.rboCurrency.Location = New System.Drawing.Point(140, 91)
        Me.rboCurrency.Name = "rboCurrency"
        Me.rboCurrency.Size = New System.Drawing.Size(67, 17)
        Me.rboCurrency.TabIndex = 4
        Me.rboCurrency.Text = "Currency"
        Me.rboCurrency.UseVisualStyleBackColor = True
        '
        'rboArAccount
        '
        Me.rboArAccount.AutoSize = True
        Me.rboArAccount.Location = New System.Drawing.Point(140, 68)
        Me.rboArAccount.Name = "rboArAccount"
        Me.rboArAccount.Size = New System.Drawing.Size(83, 17)
        Me.rboArAccount.TabIndex = 3
        Me.rboArAccount.Text = "AR Account"
        Me.rboArAccount.UseVisualStyleBackColor = True
        '
        'rboCustGroup
        '
        Me.rboCustGroup.AutoSize = True
        Me.rboCustGroup.Location = New System.Drawing.Point(12, 91)
        Me.rboCustGroup.Name = "rboCustGroup"
        Me.rboCustGroup.Size = New System.Drawing.Size(101, 17)
        Me.rboCustGroup.TabIndex = 1
        Me.rboCustGroup.Text = "Customer Group"
        Me.rboCustGroup.UseVisualStyleBackColor = True
        '
        'rboCustCode
        '
        Me.rboCustCode.AutoSize = True
        Me.rboCustCode.Checked = True
        Me.rboCustCode.Location = New System.Drawing.Point(12, 68)
        Me.rboCustCode.Name = "rboCustCode"
        Me.rboCustCode.Size = New System.Drawing.Size(97, 17)
        Me.rboCustCode.TabIndex = 0
        Me.rboCustCode.TabStop = True
        Me.rboCustCode.Text = "Customer Code"
        Me.rboCustCode.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(801, 239)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 5
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'crvAR
        '
        Me.crvAR.ActiveViewIndex = -1
        Me.crvAR.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvAR.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvAR.DisplayGroupTree = False
        Me.crvAR.Location = New System.Drawing.Point(12, 173)
        Me.crvAR.Name = "crvAR"
        Me.crvAR.SelectionFormula = ""
        Me.crvAR.ShowCloseButton = False
        Me.crvAR.ShowGroupTreeButton = False
        Me.crvAR.ShowRefreshButton = False
        Me.crvAR.Size = New System.Drawing.Size(930, 333)
        Me.crvAR.TabIndex = 3
        Me.crvAR.TabStop = False
        Me.crvAR.ViewTimeSelectionFormula = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(45, 18)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(108, 20)
        Me.dtpFrom.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(157, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To"
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(183, 18)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(108, 20)
        Me.dtpTo.TabIndex = 3
        '
        'btnShow
        '
        Me.btnShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me.btnShow.Location = New System.Drawing.Point(297, 16)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(47, 19)
        Me.btnShow.TabIndex = 4
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'FrmCetakAR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(951, 514)
        Me.Controls.Add(Me.crvAR)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmCetakAR"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print AR"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnFindCurrency As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtCurrencyDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtArAccountDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindArAccount As System.Windows.Forms.Button
    Friend WithEvents txtArAccountCode As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPaymentAccountDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindPaymentAccount As System.Windows.Forms.Button
    Friend WithEvents txtPaymentAccountCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtCustomerDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCustomerCode As System.Windows.Forms.Button
    Friend WithEvents txtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rboCurrency As System.Windows.Forms.RadioButton
    Friend WithEvents rboArAccount As System.Windows.Forms.RadioButton
    Friend WithEvents rboCustGroup As System.Windows.Forms.RadioButton
    Friend WithEvents rboCustCode As System.Windows.Forms.RadioButton
    Friend WithEvents crvAR As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents rboPaymentAccount As System.Windows.Forms.RadioButton
    Friend WithEvents txtCustomerGroupDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCustomerGroup As System.Windows.Forms.Button
    Friend WithEvents txtCustomerGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryArAccount As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryPaymentAccount As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCustomerGroup As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents rboPaymentJournal As System.Windows.Forms.RadioButton
    Friend WithEvents rboDetail As System.Windows.Forms.RadioButton
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
