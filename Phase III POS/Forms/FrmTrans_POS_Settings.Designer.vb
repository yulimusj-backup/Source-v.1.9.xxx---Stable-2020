<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_POS_Settings
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
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.chk_Modify_Allow = New System.Windows.Forms.CheckBox
        Me.chk_Modify_Autho = New System.Windows.Forms.CheckBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.chk_CancelWhole_Allow = New System.Windows.Forms.CheckBox
        Me.chk_CancelWhole_Autho = New System.Windows.Forms.CheckBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.chk_CancelItem_Allow = New System.Windows.Forms.CheckBox
        Me.chk_CancelItem_Autho = New System.Windows.Forms.CheckBox
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.chk_ModifyUnitPrice_Allow = New System.Windows.Forms.CheckBox
        Me.chk_ModifyUnitPrice_Autho = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.chk_ModifyDiscount_Allow = New System.Windows.Forms.CheckBox
        Me.chk_ModifyDiscount_Autho = New System.Windows.Forms.CheckBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.chk_Reprint_Allow = New System.Windows.Forms.CheckBox
        Me.chk_Reprint_Autho = New System.Windows.Forms.CheckBox
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.dgvAutorization = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserKode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Autho_AllowDeleteItem = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Col_Autho_WholeTransDelete = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Col_Autho_AllowModify = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Col_Autho_ModifyUnitPrice = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Col_Autho_ModifyDiscount = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Col_Autho_AllowReprint = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.gboRounding = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtRounding_PrimaryAccount = New System.Windows.Forms.TextBox
        Me.btnRounding_Find = New System.Windows.Forms.Button
        Me.TxtRounding_AccountCode = New System.Windows.Forms.TextBox
        Me.cboRoundingOn = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboRoundingValue = New System.Windows.Forms.ComboBox
        Me.cboRoundingTipe = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.chkAllowRounding = New System.Windows.Forms.CheckBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.gboTransfer = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.cboBatch = New System.Windows.Forms.ComboBox
        Me.cboExpiry = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.rbo_Transf_Automatic = New System.Windows.Forms.RadioButton
        Me.rbo_Transf_Stop = New System.Windows.Forms.RadioButton
        Me.Label11 = New System.Windows.Forms.Label
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.rboNo = New System.Windows.Forms.RadioButton
        Me.rboYes = New System.Windows.Forms.RadioButton
        Me.Label12 = New System.Windows.Forms.Label
        Me.chkAllowExpiry = New System.Windows.Forms.CheckBox
        Me.chkAllowLocation = New System.Windows.Forms.CheckBox
        Me.chkAllowBatchNumber = New System.Windows.Forms.CheckBox
        Me.chkShowCompanyName = New System.Windows.Forms.CheckBox
        Me.TxtCompanyName = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.PictureBox_Logo = New System.Windows.Forms.PictureBox
        Me.chkShowLogo = New System.Windows.Forms.CheckBox
        Me.chk_Allow_CreditCard = New System.Windows.Forms.CheckBox
        Me.chk_AlowCashWithdrawal = New System.Windows.Forms.CheckBox
        Me.chk_Allow_DebitCard = New System.Windows.Forms.CheckBox
        Me.chkUseTaxRate = New System.Windows.Forms.CheckBox
        Me.TxtTaxRate = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtCompanyLogoFiles = New System.Windows.Forms.TextBox
        Me.btnFindLogo = New System.Windows.Forms.Button
        Me.btnFindCurrency = New System.Windows.Forms.Button
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.TxtPrimaryPayment = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.BtnFindPaymentOption = New System.Windows.Forms.Button
        Me.TxtPaymentOption = New System.Windows.Forms.TextBox
        Me.chk_Allow_VouvherPayment = New System.Windows.Forms.CheckBox
        Me.BtnClear = New System.Windows.Forms.Button
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.Label13 = New System.Windows.Forms.Label
        Me.TxtDefPaymentDisc = New System.Windows.Forms.TextBox
        Me.ChkUsePaymentDiscount = New System.Windows.Forms.CheckBox
        Me.TabControl2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        CType(Me.dgvAutorization, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        Me.gboRounding.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.gboTransfer.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox_Logo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(604, 421)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 15
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(523, 421)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 14
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage3)
        Me.TabControl2.Controls.Add(Me.TabPage4)
        Me.TabControl2.Controls.Add(Me.TabPage1)
        Me.TabControl2.Controls.Add(Me.TabPage2)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Location = New System.Drawing.Point(12, 186)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(667, 229)
        Me.TabControl2.TabIndex = 15
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.SplitContainer1)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(659, 203)
        Me.TabPage3.TabIndex = 0
        Me.TabPage3.Text = "Detail Settings"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 3)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox4)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox3)
        Me.SplitContainer1.Size = New System.Drawing.Size(653, 197)
        Me.SplitContainer1.SplitterDistance = 323
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'GroupBox6
        '
        Me.GroupBox6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox6.Controls.Add(Me.chk_Modify_Allow)
        Me.GroupBox6.Controls.Add(Me.chk_Modify_Autho)
        Me.GroupBox6.Location = New System.Drawing.Point(3, 105)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(317, 45)
        Me.GroupBox6.TabIndex = 2
        Me.GroupBox6.TabStop = False
        '
        'chk_Modify_Allow
        '
        Me.chk_Modify_Allow.AutoSize = True
        Me.chk_Modify_Allow.BackColor = System.Drawing.Color.White
        Me.chk_Modify_Allow.Location = New System.Drawing.Point(6, 0)
        Me.chk_Modify_Allow.Name = "chk_Modify_Allow"
        Me.chk_Modify_Allow.Size = New System.Drawing.Size(185, 17)
        Me.chk_Modify_Allow.TabIndex = 0
        Me.chk_Modify_Allow.Text = "Allow User To Modify Transaction"
        Me.chk_Modify_Allow.UseVisualStyleBackColor = False
        '
        'chk_Modify_Autho
        '
        Me.chk_Modify_Autho.AutoSize = True
        Me.chk_Modify_Autho.BackColor = System.Drawing.Color.White
        Me.chk_Modify_Autho.Location = New System.Drawing.Point(6, 23)
        Me.chk_Modify_Autho.Name = "chk_Modify_Autho"
        Me.chk_Modify_Autho.Size = New System.Drawing.Size(109, 17)
        Me.chk_Modify_Autho.TabIndex = 1
        Me.chk_Modify_Autho.Text = "Use Authorization"
        Me.chk_Modify_Autho.UseVisualStyleBackColor = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox5.Controls.Add(Me.chk_CancelWhole_Allow)
        Me.GroupBox5.Controls.Add(Me.chk_CancelWhole_Autho)
        Me.GroupBox5.Location = New System.Drawing.Point(3, 54)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(317, 45)
        Me.GroupBox5.TabIndex = 1
        Me.GroupBox5.TabStop = False
        '
        'chk_CancelWhole_Allow
        '
        Me.chk_CancelWhole_Allow.AutoSize = True
        Me.chk_CancelWhole_Allow.BackColor = System.Drawing.Color.White
        Me.chk_CancelWhole_Allow.Location = New System.Drawing.Point(6, 0)
        Me.chk_CancelWhole_Allow.Name = "chk_CancelWhole_Allow"
        Me.chk_CancelWhole_Allow.Size = New System.Drawing.Size(221, 17)
        Me.chk_CancelWhole_Allow.TabIndex = 0
        Me.chk_CancelWhole_Allow.Text = "Allow User To Cancel Whole Transaction"
        Me.chk_CancelWhole_Allow.UseVisualStyleBackColor = False
        '
        'chk_CancelWhole_Autho
        '
        Me.chk_CancelWhole_Autho.AutoSize = True
        Me.chk_CancelWhole_Autho.BackColor = System.Drawing.Color.White
        Me.chk_CancelWhole_Autho.Location = New System.Drawing.Point(6, 23)
        Me.chk_CancelWhole_Autho.Name = "chk_CancelWhole_Autho"
        Me.chk_CancelWhole_Autho.Size = New System.Drawing.Size(109, 17)
        Me.chk_CancelWhole_Autho.TabIndex = 1
        Me.chk_CancelWhole_Autho.Text = "Use Authorization"
        Me.chk_CancelWhole_Autho.UseVisualStyleBackColor = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.chk_CancelItem_Allow)
        Me.GroupBox4.Controls.Add(Me.chk_CancelItem_Autho)
        Me.GroupBox4.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(317, 45)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        '
        'chk_CancelItem_Allow
        '
        Me.chk_CancelItem_Allow.AutoSize = True
        Me.chk_CancelItem_Allow.BackColor = System.Drawing.Color.White
        Me.chk_CancelItem_Allow.Location = New System.Drawing.Point(6, 0)
        Me.chk_CancelItem_Allow.Name = "chk_CancelItem_Allow"
        Me.chk_CancelItem_Allow.Size = New System.Drawing.Size(151, 17)
        Me.chk_CancelItem_Allow.TabIndex = 0
        Me.chk_CancelItem_Allow.Text = "Allow User To Cancel Item"
        Me.chk_CancelItem_Allow.UseVisualStyleBackColor = False
        '
        'chk_CancelItem_Autho
        '
        Me.chk_CancelItem_Autho.AutoSize = True
        Me.chk_CancelItem_Autho.BackColor = System.Drawing.Color.White
        Me.chk_CancelItem_Autho.Location = New System.Drawing.Point(6, 23)
        Me.chk_CancelItem_Autho.Name = "chk_CancelItem_Autho"
        Me.chk_CancelItem_Autho.Size = New System.Drawing.Size(109, 17)
        Me.chk_CancelItem_Autho.TabIndex = 1
        Me.chk_CancelItem_Autho.Text = "Use Authorization"
        Me.chk_CancelItem_Autho.UseVisualStyleBackColor = False
        '
        'GroupBox7
        '
        Me.GroupBox7.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox7.Controls.Add(Me.chk_ModifyUnitPrice_Allow)
        Me.GroupBox7.Controls.Add(Me.chk_ModifyUnitPrice_Autho)
        Me.GroupBox7.Location = New System.Drawing.Point(7, 3)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(316, 45)
        Me.GroupBox7.TabIndex = 0
        Me.GroupBox7.TabStop = False
        '
        'chk_ModifyUnitPrice_Allow
        '
        Me.chk_ModifyUnitPrice_Allow.AutoSize = True
        Me.chk_ModifyUnitPrice_Allow.BackColor = System.Drawing.Color.White
        Me.chk_ModifyUnitPrice_Allow.Location = New System.Drawing.Point(6, 0)
        Me.chk_ModifyUnitPrice_Allow.Name = "chk_ModifyUnitPrice_Allow"
        Me.chk_ModifyUnitPrice_Allow.Size = New System.Drawing.Size(175, 17)
        Me.chk_ModifyUnitPrice_Allow.TabIndex = 0
        Me.chk_ModifyUnitPrice_Allow.Text = "Allow User To Modify Unit Price"
        Me.chk_ModifyUnitPrice_Allow.UseVisualStyleBackColor = False
        '
        'chk_ModifyUnitPrice_Autho
        '
        Me.chk_ModifyUnitPrice_Autho.AutoSize = True
        Me.chk_ModifyUnitPrice_Autho.BackColor = System.Drawing.Color.White
        Me.chk_ModifyUnitPrice_Autho.Location = New System.Drawing.Point(6, 23)
        Me.chk_ModifyUnitPrice_Autho.Name = "chk_ModifyUnitPrice_Autho"
        Me.chk_ModifyUnitPrice_Autho.Size = New System.Drawing.Size(109, 17)
        Me.chk_ModifyUnitPrice_Autho.TabIndex = 1
        Me.chk_ModifyUnitPrice_Autho.Text = "Use Authorization"
        Me.chk_ModifyUnitPrice_Autho.UseVisualStyleBackColor = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.chk_ModifyDiscount_Allow)
        Me.GroupBox2.Controls.Add(Me.chk_ModifyDiscount_Autho)
        Me.GroupBox2.Location = New System.Drawing.Point(4, 54)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(319, 45)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'chk_ModifyDiscount_Allow
        '
        Me.chk_ModifyDiscount_Allow.AutoSize = True
        Me.chk_ModifyDiscount_Allow.BackColor = System.Drawing.Color.White
        Me.chk_ModifyDiscount_Allow.Location = New System.Drawing.Point(6, 0)
        Me.chk_ModifyDiscount_Allow.Name = "chk_ModifyDiscount_Allow"
        Me.chk_ModifyDiscount_Allow.Size = New System.Drawing.Size(197, 17)
        Me.chk_ModifyDiscount_Allow.TabIndex = 0
        Me.chk_ModifyDiscount_Allow.Text = "Allow User To Modify Discount Rate"
        Me.chk_ModifyDiscount_Allow.UseVisualStyleBackColor = False
        '
        'chk_ModifyDiscount_Autho
        '
        Me.chk_ModifyDiscount_Autho.AutoSize = True
        Me.chk_ModifyDiscount_Autho.BackColor = System.Drawing.Color.White
        Me.chk_ModifyDiscount_Autho.Location = New System.Drawing.Point(6, 23)
        Me.chk_ModifyDiscount_Autho.Name = "chk_ModifyDiscount_Autho"
        Me.chk_ModifyDiscount_Autho.Size = New System.Drawing.Size(109, 17)
        Me.chk_ModifyDiscount_Autho.TabIndex = 1
        Me.chk_ModifyDiscount_Autho.Text = "Use Authorization"
        Me.chk_ModifyDiscount_Autho.UseVisualStyleBackColor = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.chk_Reprint_Allow)
        Me.GroupBox3.Controls.Add(Me.chk_Reprint_Autho)
        Me.GroupBox3.Location = New System.Drawing.Point(4, 105)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(319, 45)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'chk_Reprint_Allow
        '
        Me.chk_Reprint_Allow.AutoSize = True
        Me.chk_Reprint_Allow.BackColor = System.Drawing.Color.White
        Me.chk_Reprint_Allow.Location = New System.Drawing.Point(6, 0)
        Me.chk_Reprint_Allow.Name = "chk_Reprint_Allow"
        Me.chk_Reprint_Allow.Size = New System.Drawing.Size(172, 17)
        Me.chk_Reprint_Allow.TabIndex = 0
        Me.chk_Reprint_Allow.Text = "Allow User To Re-print Receipt"
        Me.chk_Reprint_Allow.UseVisualStyleBackColor = False
        '
        'chk_Reprint_Autho
        '
        Me.chk_Reprint_Autho.AutoSize = True
        Me.chk_Reprint_Autho.BackColor = System.Drawing.Color.White
        Me.chk_Reprint_Autho.Location = New System.Drawing.Point(6, 23)
        Me.chk_Reprint_Autho.Name = "chk_Reprint_Autho"
        Me.chk_Reprint_Autho.Size = New System.Drawing.Size(109, 17)
        Me.chk_Reprint_Autho.TabIndex = 1
        Me.chk_Reprint_Autho.Text = "Use Authorization"
        Me.chk_Reprint_Autho.UseVisualStyleBackColor = False
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.dgvAutorization)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(659, 203)
        Me.TabPage4.TabIndex = 1
        Me.TabPage4.Text = "Authorization"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'dgvAutorization
        '
        Me.dgvAutorization.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAutorization.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvAutorization.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColUserKode, Me.ColUserName, Me.Col_Autho_AllowDeleteItem, Me.Col_Autho_WholeTransDelete, Me.Col_Autho_AllowModify, Me.Col_Autho_ModifyUnitPrice, Me.Col_Autho_ModifyDiscount, Me.Col_Autho_AllowReprint, Me.ColRemarks})
        Me.dgvAutorization.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAutorization.Location = New System.Drawing.Point(3, 3)
        Me.dgvAutorization.Name = "dgvAutorization"
        Me.dgvAutorization.Size = New System.Drawing.Size(653, 197)
        Me.dgvAutorization.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 46
        '
        'ColUserKode
        '
        Me.ColUserKode.HeaderText = "User Code"
        Me.ColUserKode.Name = "ColUserKode"
        Me.ColUserKode.Width = 82
        '
        'ColUserName
        '
        Me.ColUserName.HeaderText = "User Name"
        Me.ColUserName.Name = "ColUserName"
        Me.ColUserName.Width = 85
        '
        'Col_Autho_AllowDeleteItem
        '
        Me.Col_Autho_AllowDeleteItem.HeaderText = "Cancel Item"
        Me.Col_Autho_AllowDeleteItem.Name = "Col_Autho_AllowDeleteItem"
        Me.Col_Autho_AllowDeleteItem.Width = 69
        '
        'Col_Autho_WholeTransDelete
        '
        Me.Col_Autho_WholeTransDelete.HeaderText = "Cancel Whole Transaction Deletion"
        Me.Col_Autho_WholeTransDelete.Name = "Col_Autho_WholeTransDelete"
        Me.Col_Autho_WholeTransDelete.Width = 181
        '
        'Col_Autho_AllowModify
        '
        Me.Col_Autho_AllowModify.HeaderText = "Modify Transaction"
        Me.Col_Autho_AllowModify.Name = "Col_Autho_AllowModify"
        Me.Col_Autho_AllowModify.Width = 103
        '
        'Col_Autho_ModifyUnitPrice
        '
        Me.Col_Autho_ModifyUnitPrice.HeaderText = "Modify Unit Price"
        Me.Col_Autho_ModifyUnitPrice.Name = "Col_Autho_ModifyUnitPrice"
        Me.Col_Autho_ModifyUnitPrice.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Col_Autho_ModifyUnitPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Col_Autho_ModifyUnitPrice.Width = 112
        '
        'Col_Autho_ModifyDiscount
        '
        Me.Col_Autho_ModifyDiscount.HeaderText = "Modify Discount Rate"
        Me.Col_Autho_ModifyDiscount.Name = "Col_Autho_ModifyDiscount"
        Me.Col_Autho_ModifyDiscount.Width = 115
        '
        'Col_Autho_AllowReprint
        '
        Me.Col_Autho_AllowReprint.HeaderText = "Reprint Receipt"
        Me.Col_Autho_AllowReprint.Name = "Col_Autho_AllowReprint"
        Me.Col_Autho_AllowReprint.Width = 87
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.MaxInputLength = 250
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.gboRounding)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(659, 203)
        Me.TabPage1.TabIndex = 2
        Me.TabPage1.Text = "Rounding"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'gboRounding
        '
        Me.gboRounding.Controls.Add(Me.Label5)
        Me.gboRounding.Controls.Add(Me.TxtRounding_PrimaryAccount)
        Me.gboRounding.Controls.Add(Me.btnRounding_Find)
        Me.gboRounding.Controls.Add(Me.TxtRounding_AccountCode)
        Me.gboRounding.Controls.Add(Me.cboRoundingOn)
        Me.gboRounding.Controls.Add(Me.Label4)
        Me.gboRounding.Controls.Add(Me.cboRoundingValue)
        Me.gboRounding.Controls.Add(Me.cboRoundingTipe)
        Me.gboRounding.Controls.Add(Me.Label2)
        Me.gboRounding.Controls.Add(Me.chkAllowRounding)
        Me.gboRounding.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboRounding.Location = New System.Drawing.Point(3, 3)
        Me.gboRounding.Name = "gboRounding"
        Me.gboRounding.Size = New System.Drawing.Size(653, 197)
        Me.gboRounding.TabIndex = 0
        Me.gboRounding.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 47)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(75, 13)
        Me.Label5.TabIndex = 47
        Me.Label5.Text = "Account Code"
        '
        'TxtRounding_PrimaryAccount
        '
        Me.TxtRounding_PrimaryAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRounding_PrimaryAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRounding_PrimaryAccount.Location = New System.Drawing.Point(112, 44)
        Me.TxtRounding_PrimaryAccount.MaxLength = 150
        Me.TxtRounding_PrimaryAccount.Name = "TxtRounding_PrimaryAccount"
        Me.TxtRounding_PrimaryAccount.Size = New System.Drawing.Size(29, 20)
        Me.TxtRounding_PrimaryAccount.TabIndex = 48
        Me.TxtRounding_PrimaryAccount.TabStop = False
        Me.TxtRounding_PrimaryAccount.Text = "-1"
        Me.TxtRounding_PrimaryAccount.Visible = False
        '
        'btnRounding_Find
        '
        Me.btnRounding_Find.Enabled = False
        Me.btnRounding_Find.Location = New System.Drawing.Point(351, 44)
        Me.btnRounding_Find.Name = "btnRounding_Find"
        Me.btnRounding_Find.Size = New System.Drawing.Size(48, 23)
        Me.btnRounding_Find.TabIndex = 46
        Me.btnRounding_Find.TabStop = False
        Me.btnRounding_Find.Text = "Find"
        Me.btnRounding_Find.UseVisualStyleBackColor = True
        '
        'TxtRounding_AccountCode
        '
        Me.TxtRounding_AccountCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRounding_AccountCode.Enabled = False
        Me.TxtRounding_AccountCode.Location = New System.Drawing.Point(112, 44)
        Me.TxtRounding_AccountCode.MaxLength = 150
        Me.TxtRounding_AccountCode.Name = "TxtRounding_AccountCode"
        Me.TxtRounding_AccountCode.Size = New System.Drawing.Size(233, 20)
        Me.TxtRounding_AccountCode.TabIndex = 3
        '
        'cboRoundingOn
        '
        Me.cboRoundingOn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingOn.Enabled = False
        Me.cboRoundingOn.FormattingEnabled = True
        Me.cboRoundingOn.Items.AddRange(New Object() {"Before Tax Calculation", "After Tax Calculation"})
        Me.cboRoundingOn.Location = New System.Drawing.Point(112, 44)
        Me.cboRoundingOn.Name = "cboRoundingOn"
        Me.cboRoundingOn.Size = New System.Drawing.Size(287, 21)
        Me.cboRoundingOn.TabIndex = 2
        Me.cboRoundingOn.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 47)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 13)
        Me.Label4.TabIndex = 43
        Me.Label4.Text = "Rounding On"
        Me.Label4.Visible = False
        '
        'cboRoundingValue
        '
        Me.cboRoundingValue.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingValue.Enabled = False
        Me.cboRoundingValue.FormattingEnabled = True
        Me.cboRoundingValue.Items.AddRange(New Object() {"1000", "500", "100", "50", "10", "5", "1", "0.5", "0.1", "0.05", "0.01", "0.005", "0.001"})
        Me.cboRoundingValue.Location = New System.Drawing.Point(250, 17)
        Me.cboRoundingValue.Name = "cboRoundingValue"
        Me.cboRoundingValue.Size = New System.Drawing.Size(149, 21)
        Me.cboRoundingValue.TabIndex = 1
        '
        'cboRoundingTipe
        '
        Me.cboRoundingTipe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingTipe.Enabled = False
        Me.cboRoundingTipe.FormattingEnabled = True
        Me.cboRoundingTipe.Items.AddRange(New Object() {"Up", "Down", "Middle Point"})
        Me.cboRoundingTipe.Location = New System.Drawing.Point(112, 17)
        Me.cboRoundingTipe.Name = "cboRoundingTipe"
        Me.cboRoundingTipe.Size = New System.Drawing.Size(132, 21)
        Me.cboRoundingTipe.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 13)
        Me.Label2.TabIndex = 39
        Me.Label2.Text = "Rounding Tipe"
        '
        'chkAllowRounding
        '
        Me.chkAllowRounding.AutoSize = True
        Me.chkAllowRounding.BackColor = System.Drawing.Color.White
        Me.chkAllowRounding.Location = New System.Drawing.Point(6, 0)
        Me.chkAllowRounding.Name = "chkAllowRounding"
        Me.chkAllowRounding.Size = New System.Drawing.Size(100, 17)
        Me.chkAllowRounding.TabIndex = 38
        Me.chkAllowRounding.Text = "Allow Rounding"
        Me.chkAllowRounding.UseVisualStyleBackColor = False
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.gboTransfer)
        Me.TabPage2.Controls.Add(Me.rbo_Transf_Automatic)
        Me.TabPage2.Controls.Add(Me.rbo_Transf_Stop)
        Me.TabPage2.Controls.Add(Me.Label11)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(659, 203)
        Me.TabPage2.TabIndex = 3
        Me.TabPage2.Text = "POS Data Transfer"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'gboTransfer
        '
        Me.gboTransfer.Controls.Add(Me.Label7)
        Me.gboTransfer.Controls.Add(Me.cboBatch)
        Me.gboTransfer.Controls.Add(Me.cboExpiry)
        Me.gboTransfer.Controls.Add(Me.Label8)
        Me.gboTransfer.Controls.Add(Me.Label9)
        Me.gboTransfer.Controls.Add(Me.Label10)
        Me.gboTransfer.Enabled = False
        Me.gboTransfer.Location = New System.Drawing.Point(12, 72)
        Me.gboTransfer.Name = "gboTransfer"
        Me.gboTransfer.Size = New System.Drawing.Size(439, 118)
        Me.gboTransfer.TabIndex = 2
        Me.gboTransfer.TabStop = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 94)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(233, 13)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "4. Create in adjustment to continue data transfer"
        '
        'cboBatch
        '
        Me.cboBatch.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBatch.FormattingEnabled = True
        Me.cboBatch.Items.AddRange(New Object() {"ascending order", "descending order"})
        Me.cboBatch.Location = New System.Drawing.Point(157, 65)
        Me.cboBatch.Name = "cboBatch"
        Me.cboBatch.Size = New System.Drawing.Size(121, 21)
        Me.cboBatch.TabIndex = 1
        '
        'cboExpiry
        '
        Me.cboExpiry.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboExpiry.FormattingEnabled = True
        Me.cboExpiry.Items.AddRange(New Object() {"ascending order", "descending order"})
        Me.cboExpiry.Location = New System.Drawing.Point(157, 38)
        Me.cboExpiry.Name = "cboExpiry"
        Me.cboExpiry.Size = New System.Drawing.Size(121, 21)
        Me.cboExpiry.TabIndex = 0
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 41)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(134, 13)
        Me.Label8.TabIndex = 2
        Me.Label8.Text = "2. Choose expiry date with "
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 68)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(145, 13)
        Me.Label9.TabIndex = 1
        Me.Label9.Text = "3. Choose batch number with"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(236, 13)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "1. Choose location based on branch master data"
        '
        'rbo_Transf_Automatic
        '
        Me.rbo_Transf_Automatic.AutoSize = True
        Me.rbo_Transf_Automatic.Location = New System.Drawing.Point(15, 49)
        Me.rbo_Transf_Automatic.Name = "rbo_Transf_Automatic"
        Me.rbo_Transf_Automatic.Size = New System.Drawing.Size(169, 17)
        Me.rbo_Transf_Automatic.TabIndex = 1
        Me.rbo_Transf_Automatic.Text = "Automatically resolve problems"
        Me.rbo_Transf_Automatic.UseVisualStyleBackColor = True
        '
        'rbo_Transf_Stop
        '
        Me.rbo_Transf_Stop.AutoSize = True
        Me.rbo_Transf_Stop.Checked = True
        Me.rbo_Transf_Stop.Location = New System.Drawing.Point(15, 26)
        Me.rbo_Transf_Stop.Name = "rbo_Transf_Stop"
        Me.rbo_Transf_Stop.Size = New System.Drawing.Size(342, 17)
        Me.rbo_Transf_Stop.TabIndex = 0
        Me.rbo_Transf_Stop.TabStop = True
        Me.rbo_Transf_Stop.Text = "Stop transfer, show me what's wrong, and let me do data correction"
        Me.rbo_Transf_Stop.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(12, 10)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(273, 13)
        Me.Label11.TabIndex = 9
        Me.Label11.Text = "When there no stock on inventory to apply data transfer:"
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.rboNo)
        Me.TabPage5.Controls.Add(Me.rboYes)
        Me.TabPage5.Controls.Add(Me.Label12)
        Me.TabPage5.Controls.Add(Me.chkAllowExpiry)
        Me.TabPage5.Controls.Add(Me.chkAllowLocation)
        Me.TabPage5.Controls.Add(Me.chkAllowBatchNumber)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(659, 203)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Location, Batch Number, And Expiry Setting"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'rboNo
        '
        Me.rboNo.AutoSize = True
        Me.rboNo.Checked = True
        Me.rboNo.Location = New System.Drawing.Point(10, 121)
        Me.rboNo.Name = "rboNo"
        Me.rboNo.Size = New System.Drawing.Size(39, 17)
        Me.rboNo.TabIndex = 4
        Me.rboNo.TabStop = True
        Me.rboNo.Text = "No"
        Me.rboNo.UseVisualStyleBackColor = True
        '
        'rboYes
        '
        Me.rboYes.AutoSize = True
        Me.rboYes.Location = New System.Drawing.Point(10, 98)
        Me.rboYes.Name = "rboYes"
        Me.rboYes.Size = New System.Drawing.Size(43, 17)
        Me.rboYes.TabIndex = 3
        Me.rboYes.Text = "Yes"
        Me.rboYes.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(10, 82)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(310, 13)
        Me.Label12.TabIndex = 69
        Me.Label12.Text = "Input location, batch number, or expiry date data on cashier form"
        '
        'chkAllowExpiry
        '
        Me.chkAllowExpiry.AutoSize = True
        Me.chkAllowExpiry.BackColor = System.Drawing.Color.Transparent
        Me.chkAllowExpiry.Enabled = False
        Me.chkAllowExpiry.Location = New System.Drawing.Point(13, 62)
        Me.chkAllowExpiry.Name = "chkAllowExpiry"
        Me.chkAllowExpiry.Size = New System.Drawing.Size(144, 17)
        Me.chkAllowExpiry.TabIndex = 2
        Me.chkAllowExpiry.Text = "Allow Expiry Date Control"
        Me.chkAllowExpiry.UseVisualStyleBackColor = False
        '
        'chkAllowLocation
        '
        Me.chkAllowLocation.AutoSize = True
        Me.chkAllowLocation.BackColor = System.Drawing.Color.Transparent
        Me.chkAllowLocation.Location = New System.Drawing.Point(13, 16)
        Me.chkAllowLocation.Name = "chkAllowLocation"
        Me.chkAllowLocation.Size = New System.Drawing.Size(131, 17)
        Me.chkAllowLocation.TabIndex = 0
        Me.chkAllowLocation.Text = "Allow Location Control"
        Me.chkAllowLocation.UseVisualStyleBackColor = False
        '
        'chkAllowBatchNumber
        '
        Me.chkAllowBatchNumber.AutoSize = True
        Me.chkAllowBatchNumber.BackColor = System.Drawing.Color.Transparent
        Me.chkAllowBatchNumber.Location = New System.Drawing.Point(13, 39)
        Me.chkAllowBatchNumber.Name = "chkAllowBatchNumber"
        Me.chkAllowBatchNumber.Size = New System.Drawing.Size(158, 17)
        Me.chkAllowBatchNumber.TabIndex = 1
        Me.chkAllowBatchNumber.Text = "Allow Batch Number Control"
        Me.chkAllowBatchNumber.UseVisualStyleBackColor = False
        '
        'chkShowCompanyName
        '
        Me.chkShowCompanyName.AutoSize = True
        Me.chkShowCompanyName.BackColor = System.Drawing.SystemColors.Control
        Me.chkShowCompanyName.Checked = True
        Me.chkShowCompanyName.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowCompanyName.Location = New System.Drawing.Point(169, 14)
        Me.chkShowCompanyName.Name = "chkShowCompanyName"
        Me.chkShowCompanyName.Size = New System.Drawing.Size(131, 17)
        Me.chkShowCompanyName.TabIndex = 0
        Me.chkShowCompanyName.Text = "Show Company Name"
        Me.chkShowCompanyName.UseVisualStyleBackColor = False
        '
        'TxtCompanyName
        '
        Me.TxtCompanyName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCompanyName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCompanyName.Location = New System.Drawing.Point(322, 12)
        Me.TxtCompanyName.MaxLength = 150
        Me.TxtCompanyName.Name = "TxtCompanyName"
        Me.TxtCompanyName.Size = New System.Drawing.Size(357, 20)
        Me.TxtCompanyName.TabIndex = 2
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.PictureBox_Logo)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(151, 128)
        Me.GroupBox1.TabIndex = 26
        Me.GroupBox1.TabStop = False
        '
        'PictureBox_Logo
        '
        Me.PictureBox_Logo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox_Logo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox_Logo.Location = New System.Drawing.Point(3, 16)
        Me.PictureBox_Logo.Name = "PictureBox_Logo"
        Me.PictureBox_Logo.Size = New System.Drawing.Size(145, 109)
        Me.PictureBox_Logo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox_Logo.TabIndex = 1
        Me.PictureBox_Logo.TabStop = False
        '
        'chkShowLogo
        '
        Me.chkShowLogo.AutoSize = True
        Me.chkShowLogo.BackColor = System.Drawing.SystemColors.Control
        Me.chkShowLogo.Checked = True
        Me.chkShowLogo.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLogo.Location = New System.Drawing.Point(169, 37)
        Me.chkShowLogo.Name = "chkShowLogo"
        Me.chkShowLogo.Size = New System.Drawing.Size(127, 17)
        Me.chkShowLogo.TabIndex = 1
        Me.chkShowLogo.Text = "Show Company Logo"
        Me.chkShowLogo.UseVisualStyleBackColor = False
        '
        'chk_Allow_CreditCard
        '
        Me.chk_Allow_CreditCard.AutoSize = True
        Me.chk_Allow_CreditCard.BackColor = System.Drawing.SystemColors.Control
        Me.chk_Allow_CreditCard.Location = New System.Drawing.Point(169, 139)
        Me.chk_Allow_CreditCard.Name = "chk_Allow_CreditCard"
        Me.chk_Allow_CreditCard.Size = New System.Drawing.Size(150, 17)
        Me.chk_Allow_CreditCard.TabIndex = 8
        Me.chk_Allow_CreditCard.Text = "Allow Credit Card Payment"
        Me.chk_Allow_CreditCard.UseVisualStyleBackColor = False
        '
        'chk_AlowCashWithdrawal
        '
        Me.chk_AlowCashWithdrawal.AutoSize = True
        Me.chk_AlowCashWithdrawal.BackColor = System.Drawing.SystemColors.Control
        Me.chk_AlowCashWithdrawal.Location = New System.Drawing.Point(322, 139)
        Me.chk_AlowCashWithdrawal.Name = "chk_AlowCashWithdrawal"
        Me.chk_AlowCashWithdrawal.Size = New System.Drawing.Size(134, 17)
        Me.chk_AlowCashWithdrawal.TabIndex = 10
        Me.chk_AlowCashWithdrawal.Text = "Allow Cash Withdrawal"
        Me.chk_AlowCashWithdrawal.UseVisualStyleBackColor = False
        '
        'chk_Allow_DebitCard
        '
        Me.chk_Allow_DebitCard.AutoSize = True
        Me.chk_Allow_DebitCard.BackColor = System.Drawing.SystemColors.Control
        Me.chk_Allow_DebitCard.Location = New System.Drawing.Point(169, 162)
        Me.chk_Allow_DebitCard.Name = "chk_Allow_DebitCard"
        Me.chk_Allow_DebitCard.Size = New System.Drawing.Size(148, 17)
        Me.chk_Allow_DebitCard.TabIndex = 9
        Me.chk_Allow_DebitCard.Text = "Allow Debit Card Payment"
        Me.chk_Allow_DebitCard.UseVisualStyleBackColor = False
        '
        'chkUseTaxRate
        '
        Me.chkUseTaxRate.AutoSize = True
        Me.chkUseTaxRate.BackColor = System.Drawing.SystemColors.Control
        Me.chkUseTaxRate.Location = New System.Drawing.Point(462, 139)
        Me.chkUseTaxRate.Name = "chkUseTaxRate"
        Me.chkUseTaxRate.Size = New System.Drawing.Size(92, 17)
        Me.chkUseTaxRate.TabIndex = 12
        Me.chkUseTaxRate.Text = "Use Tax Rate"
        Me.chkUseTaxRate.UseVisualStyleBackColor = False
        '
        'TxtTaxRate
        '
        Me.TxtTaxRate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTaxRate.Location = New System.Drawing.Point(554, 137)
        Me.TxtTaxRate.MaxLength = 4
        Me.TxtTaxRate.Name = "TxtTaxRate"
        Me.TxtTaxRate.Size = New System.Drawing.Size(44, 20)
        Me.TxtTaxRate.TabIndex = 13
        Me.TxtTaxRate.Text = "0"
        Me.TxtTaxRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(601, 141)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(15, 13)
        Me.Label1.TabIndex = 35
        Me.Label1.Text = "%"
        '
        'TxtCompanyLogoFiles
        '
        Me.TxtCompanyLogoFiles.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCompanyLogoFiles.Location = New System.Drawing.Point(322, 35)
        Me.TxtCompanyLogoFiles.MaxLength = 150
        Me.TxtCompanyLogoFiles.Name = "TxtCompanyLogoFiles"
        Me.TxtCompanyLogoFiles.ReadOnly = True
        Me.TxtCompanyLogoFiles.Size = New System.Drawing.Size(249, 20)
        Me.TxtCompanyLogoFiles.TabIndex = 36
        Me.TxtCompanyLogoFiles.TabStop = False
        '
        'btnFindLogo
        '
        Me.btnFindLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindLogo.Location = New System.Drawing.Point(577, 35)
        Me.btnFindLogo.Name = "btnFindLogo"
        Me.btnFindLogo.Size = New System.Drawing.Size(48, 23)
        Me.btnFindLogo.TabIndex = 3
        Me.btnFindLogo.Text = "Find"
        Me.btnFindLogo.UseVisualStyleBackColor = True
        '
        'btnFindCurrency
        '
        Me.btnFindCurrency.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCurrency.Location = New System.Drawing.Point(631, 59)
        Me.btnFindCurrency.Name = "btnFindCurrency"
        Me.btnFindCurrency.Size = New System.Drawing.Size(48, 23)
        Me.btnFindCurrency.TabIndex = 39
        Me.btnFindCurrency.TabStop = False
        Me.btnFindCurrency.Text = "Find"
        Me.btnFindCurrency.UseVisualStyleBackColor = True
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCurrencyCode.Location = New System.Drawing.Point(322, 61)
        Me.txtCurrencyCode.MaxLength = 150
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.Size = New System.Drawing.Size(303, 20)
        Me.txtCurrencyCode.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(169, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 13)
        Me.Label3.TabIndex = 50
        Me.Label3.Text = "Default Currency"
        '
        'TxtPrimaryCurrency
        '
        Me.TxtPrimaryCurrency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryCurrency.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryCurrency.Location = New System.Drawing.Point(281, 64)
        Me.TxtPrimaryCurrency.MaxLength = 150
        Me.TxtPrimaryCurrency.Name = "TxtPrimaryCurrency"
        Me.TxtPrimaryCurrency.Size = New System.Drawing.Size(36, 20)
        Me.TxtPrimaryCurrency.TabIndex = 51
        Me.TxtPrimaryCurrency.TabStop = False
        Me.TxtPrimaryCurrency.Text = "-1"
        Me.TxtPrimaryCurrency.Visible = False
        '
        'TxtPrimaryPayment
        '
        Me.TxtPrimaryPayment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryPayment.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryPayment.Location = New System.Drawing.Point(281, 90)
        Me.TxtPrimaryPayment.MaxLength = 150
        Me.TxtPrimaryPayment.Name = "TxtPrimaryPayment"
        Me.TxtPrimaryPayment.Size = New System.Drawing.Size(36, 20)
        Me.TxtPrimaryPayment.TabIndex = 57
        Me.TxtPrimaryPayment.TabStop = False
        Me.TxtPrimaryPayment.Text = "-1"
        Me.TxtPrimaryPayment.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(169, 90)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(119, 13)
        Me.Label6.TabIndex = 56
        Me.Label6.Text = "Default Payment Option"
        '
        'BtnFindPaymentOption
        '
        Me.BtnFindPaymentOption.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindPaymentOption.Location = New System.Drawing.Point(631, 85)
        Me.BtnFindPaymentOption.Name = "BtnFindPaymentOption"
        Me.BtnFindPaymentOption.Size = New System.Drawing.Size(48, 23)
        Me.BtnFindPaymentOption.TabIndex = 55
        Me.BtnFindPaymentOption.TabStop = False
        Me.BtnFindPaymentOption.Text = "Find"
        Me.BtnFindPaymentOption.UseVisualStyleBackColor = True
        '
        'TxtPaymentOption
        '
        Me.TxtPaymentOption.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPaymentOption.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPaymentOption.Location = New System.Drawing.Point(322, 87)
        Me.TxtPaymentOption.MaxLength = 150
        Me.TxtPaymentOption.Name = "TxtPaymentOption"
        Me.TxtPaymentOption.Size = New System.Drawing.Size(303, 20)
        Me.TxtPaymentOption.TabIndex = 6
        '
        'chk_Allow_VouvherPayment
        '
        Me.chk_Allow_VouvherPayment.AutoSize = True
        Me.chk_Allow_VouvherPayment.BackColor = System.Drawing.SystemColors.Control
        Me.chk_Allow_VouvherPayment.Location = New System.Drawing.Point(322, 162)
        Me.chk_Allow_VouvherPayment.Name = "chk_Allow_VouvherPayment"
        Me.chk_Allow_VouvherPayment.Size = New System.Drawing.Size(138, 17)
        Me.chk_Allow_VouvherPayment.TabIndex = 11
        Me.chk_Allow_VouvherPayment.Text = "Allow Voucher Payment"
        Me.chk_Allow_VouvherPayment.UseVisualStyleBackColor = False
        '
        'BtnClear
        '
        Me.BtnClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClear.Location = New System.Drawing.Point(630, 35)
        Me.BtnClear.Name = "BtnClear"
        Me.BtnClear.Size = New System.Drawing.Size(49, 23)
        Me.BtnClear.TabIndex = 4
        Me.BtnClear.Text = "Clear"
        Me.BtnClear.UseVisualStyleBackColor = True
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimary.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimary.Location = New System.Drawing.Point(322, 14)
        Me.TxtPrimary.MaxLength = 150
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(36, 20)
        Me.TxtPrimary.TabIndex = 60
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(631, 136)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(27, 20)
        Me.TxtReadOnly.TabIndex = 61
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'BackgroundWorker2
        '
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Filter = """JPEG Image(*.jpg)|*.jpg|GIF Image (*.gif)|*.gif|PNG Image (*.png)|*.png"""
        Me.OpenFileDialog1.Title = "Browse Image"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(169, 116)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(130, 13)
        Me.Label13.TabIndex = 64
        Me.Label13.Text = "Default Payment Discount"
        '
        'TxtDefPaymentDisc
        '
        Me.TxtDefPaymentDisc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDefPaymentDisc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDefPaymentDisc.Location = New System.Drawing.Point(322, 113)
        Me.TxtDefPaymentDisc.MaxLength = 150
        Me.TxtDefPaymentDisc.Name = "TxtDefPaymentDisc"
        Me.TxtDefPaymentDisc.Size = New System.Drawing.Size(303, 20)
        Me.TxtDefPaymentDisc.TabIndex = 7
        Me.TxtDefPaymentDisc.Text = "0"
        Me.TxtDefPaymentDisc.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ChkUsePaymentDiscount
        '
        Me.ChkUsePaymentDiscount.AutoSize = True
        Me.ChkUsePaymentDiscount.BackColor = System.Drawing.SystemColors.Control
        Me.ChkUsePaymentDiscount.Location = New System.Drawing.Point(462, 163)
        Me.ChkUsePaymentDiscount.Name = "ChkUsePaymentDiscount"
        Me.ChkUsePaymentDiscount.Size = New System.Drawing.Size(134, 17)
        Me.ChkUsePaymentDiscount.TabIndex = 14
        Me.ChkUsePaymentDiscount.Text = "Use Payment Discount"
        Me.ChkUsePaymentDiscount.UseVisualStyleBackColor = False
        '
        'FrmTrans_POS_Settings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(691, 451)
        Me.Controls.Add(Me.ChkUsePaymentDiscount)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.TxtDefPaymentDisc)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.TxtPrimary)
        Me.Controls.Add(Me.BtnClear)
        Me.Controls.Add(Me.chk_Allow_VouvherPayment)
        Me.Controls.Add(Me.TxtPrimaryPayment)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.BtnFindPaymentOption)
        Me.Controls.Add(Me.TxtPaymentOption)
        Me.Controls.Add(Me.TxtPrimaryCurrency)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnFindCurrency)
        Me.Controls.Add(Me.txtCurrencyCode)
        Me.Controls.Add(Me.btnFindLogo)
        Me.Controls.Add(Me.chkShowLogo)
        Me.Controls.Add(Me.TxtCompanyLogoFiles)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtTaxRate)
        Me.Controls.Add(Me.chkUseTaxRate)
        Me.Controls.Add(Me.chk_Allow_DebitCard)
        Me.Controls.Add(Me.chk_AlowCashWithdrawal)
        Me.Controls.Add(Me.chk_Allow_CreditCard)
        Me.Controls.Add(Me.TabControl2)
        Me.Controls.Add(Me.chkShowCompanyName)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.TxtCompanyName)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(637, 462)
        Me.Name = "FrmTrans_POS_Settings"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "General Settings"
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.dgvAutorization, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        Me.gboRounding.ResumeLayout(False)
        Me.gboRounding.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.gboTransfer.ResumeLayout(False)
        Me.gboTransfer.PerformLayout()
        Me.TabPage5.ResumeLayout(False)
        Me.TabPage5.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.PictureBox_Logo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents chk_Modify_Allow As System.Windows.Forms.CheckBox
    Friend WithEvents chk_Modify_Autho As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents chk_CancelWhole_Allow As System.Windows.Forms.CheckBox
    Friend WithEvents chk_CancelWhole_Autho As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents chk_CancelItem_Allow As System.Windows.Forms.CheckBox
    Friend WithEvents chk_CancelItem_Autho As System.Windows.Forms.CheckBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents dgvAutorization As System.Windows.Forms.DataGridView
    Friend WithEvents chkShowCompanyName As System.Windows.Forms.CheckBox
    Friend WithEvents TxtCompanyName As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkShowLogo As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox_Logo As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents chk_Reprint_Allow As System.Windows.Forms.CheckBox
    Friend WithEvents chk_Reprint_Autho As System.Windows.Forms.CheckBox
    Friend WithEvents chk_Allow_CreditCard As System.Windows.Forms.CheckBox
    Friend WithEvents chk_AlowCashWithdrawal As System.Windows.Forms.CheckBox
    Friend WithEvents chk_Allow_DebitCard As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseTaxRate As System.Windows.Forms.CheckBox
    Friend WithEvents TxtTaxRate As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents chk_ModifyUnitPrice_Allow As System.Windows.Forms.CheckBox
    Friend WithEvents chk_ModifyUnitPrice_Autho As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents chk_ModifyDiscount_Allow As System.Windows.Forms.CheckBox
    Friend WithEvents chk_ModifyDiscount_Autho As System.Windows.Forms.CheckBox
    Friend WithEvents TxtCompanyLogoFiles As System.Windows.Forms.TextBox
    Friend WithEvents btnFindLogo As System.Windows.Forms.Button
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents gboRounding As System.Windows.Forms.GroupBox
    Friend WithEvents chkAllowRounding As System.Windows.Forms.CheckBox
    Friend WithEvents cboRoundingTipe As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboRoundingOn As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboRoundingValue As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtRounding_PrimaryAccount As System.Windows.Forms.TextBox
    Friend WithEvents btnRounding_Find As System.Windows.Forms.Button
    Friend WithEvents TxtRounding_AccountCode As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCurrency As System.Windows.Forms.Button
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryPayment As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents BtnFindPaymentOption As System.Windows.Forms.Button
    Friend WithEvents TxtPaymentOption As System.Windows.Forms.TextBox
    Friend WithEvents chk_Allow_VouvherPayment As System.Windows.Forms.CheckBox
    Friend WithEvents BtnClear As System.Windows.Forms.Button
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents gboTransfer As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cboBatch As System.Windows.Forms.ComboBox
    Friend WithEvents cboExpiry As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents rbo_Transf_Automatic As System.Windows.Forms.RadioButton
    Friend WithEvents rbo_Transf_Stop As System.Windows.Forms.RadioButton
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents rboNo As System.Windows.Forms.RadioButton
    Friend WithEvents rboYes As System.Windows.Forms.RadioButton
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents chkAllowExpiry As System.Windows.Forms.CheckBox
    Friend WithEvents chkAllowLocation As System.Windows.Forms.CheckBox
    Friend WithEvents chkAllowBatchNumber As System.Windows.Forms.CheckBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserKode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Autho_AllowDeleteItem As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Col_Autho_WholeTransDelete As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Col_Autho_AllowModify As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Col_Autho_ModifyUnitPrice As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Col_Autho_ModifyDiscount As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Col_Autho_AllowReprint As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TxtDefPaymentDisc As System.Windows.Forms.TextBox
    Friend WithEvents ChkUsePaymentDiscount As System.Windows.Forms.CheckBox
End Class
