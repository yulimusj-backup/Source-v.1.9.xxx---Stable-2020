<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTools_MonthlyClosing
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
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.T1_Btnfilter = New System.Windows.Forms.Button
        Me.T1_PrimaryAkunInv = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.T1_dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.T1_dgv = New System.Windows.Forms.DataGridView
        Me.ColT1_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT1_AOPNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT1_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT1_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT1_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.T1_TxtAccDesc = New System.Windows.Forms.TextBox
        Me.T1_btnFindAcc = New System.Windows.Forms.Button
        Me.T1_TxtAccCode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.T2_dgv = New System.Windows.Forms.DataGridView
        Me.ColT2_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT2_AOPNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT2_AccCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT2_AccDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT2_Debet = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT2_Debet_Baru = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.T2_BtnFind = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.T2_dtp = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.T3_LblSelisih = New System.Windows.Forms.Label
        Me.T3_dgv = New System.Windows.Forms.DataGridView
        Me.ColT3_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT3_BundlingNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT3_In = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT3_Out = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT3_Selisih = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColT3_AllocatedAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.T3_btnFilter = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.T3_dtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.Label7 = New System.Windows.Forms.Label
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.T4_BtnFilter = New System.Windows.Forms.Button
        Me.T4_TxtPrimaryClosing = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.T1_dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.T2_dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.T3_dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(797, 521)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.T1_Btnfilter)
        Me.TabPage1.Controls.Add(Me.T1_PrimaryAkunInv)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.T1_dtpFrom)
        Me.TabPage1.Controls.Add(Me.T1_dgv)
        Me.TabPage1.Controls.Add(Me.T1_TxtAccDesc)
        Me.TabPage1.Controls.Add(Me.T1_btnFindAcc)
        Me.TabPage1.Controls.Add(Me.T1_TxtAccCode)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(789, 492)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "TabPage1"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'T1_Btnfilter
        '
        Me.T1_Btnfilter.Location = New System.Drawing.Point(341, 50)
        Me.T1_Btnfilter.Name = "T1_Btnfilter"
        Me.T1_Btnfilter.Size = New System.Drawing.Size(75, 23)
        Me.T1_Btnfilter.TabIndex = 9
        Me.T1_Btnfilter.Text = "Find"
        Me.T1_Btnfilter.UseVisualStyleBackColor = True
        '
        'T1_PrimaryAkunInv
        '
        Me.T1_PrimaryAkunInv.Location = New System.Drawing.Point(85, 23)
        Me.T1_PrimaryAkunInv.Name = "T1_PrimaryAkunInv"
        Me.T1_PrimaryAkunInv.Size = New System.Drawing.Size(44, 22)
        Me.T1_PrimaryAkunInv.TabIndex = 8
        Me.T1_PrimaryAkunInv.TabStop = False
        Me.T1_PrimaryAkunInv.Text = "-1"
        Me.T1_PrimaryAkunInv.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 54)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Period"
        '
        'T1_dtpFrom
        '
        Me.T1_dtpFrom.CustomFormat = "MMM yyyy"
        Me.T1_dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.T1_dtpFrom.Location = New System.Drawing.Point(135, 51)
        Me.T1_dtpFrom.Name = "T1_dtpFrom"
        Me.T1_dtpFrom.ShowUpDown = True
        Me.T1_dtpFrom.Size = New System.Drawing.Size(200, 22)
        Me.T1_dtpFrom.TabIndex = 6
        '
        'T1_dgv
        '
        Me.T1_dgv.AllowUserToAddRows = False
        Me.T1_dgv.AllowUserToDeleteRows = False
        Me.T1_dgv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.T1_dgv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.T1_dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.T1_dgv.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColT1_No, Me.ColT1_AOPNo, Me.ColT1_CodeNo, Me.ColT1_Desc, Me.ColT1_Amount})
        Me.T1_dgv.Location = New System.Drawing.Point(11, 79)
        Me.T1_dgv.Name = "T1_dgv"
        Me.T1_dgv.ReadOnly = True
        Me.T1_dgv.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.T1_dgv.RowTemplate.Height = 24
        Me.T1_dgv.Size = New System.Drawing.Size(770, 405)
        Me.T1_dgv.TabIndex = 5
        '
        'ColT1_No
        '
        Me.ColT1_No.HeaderText = "No"
        Me.ColT1_No.Name = "ColT1_No"
        Me.ColT1_No.ReadOnly = True
        Me.ColT1_No.Width = 51
        '
        'ColT1_AOPNo
        '
        Me.ColT1_AOPNo.HeaderText = "AOP No"
        Me.ColT1_AOPNo.Name = "ColT1_AOPNo"
        Me.ColT1_AOPNo.ReadOnly = True
        Me.ColT1_AOPNo.Width = 84
        '
        'ColT1_CodeNo
        '
        Me.ColT1_CodeNo.HeaderText = "Code No"
        Me.ColT1_CodeNo.Name = "ColT1_CodeNo"
        Me.ColT1_CodeNo.ReadOnly = True
        Me.ColT1_CodeNo.Width = 88
        '
        'ColT1_Desc
        '
        Me.ColT1_Desc.HeaderText = "Description"
        Me.ColT1_Desc.Name = "ColT1_Desc"
        Me.ColT1_Desc.ReadOnly = True
        Me.ColT1_Desc.Width = 104
        '
        'ColT1_Amount
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColT1_Amount.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColT1_Amount.HeaderText = "Amount"
        Me.ColT1_Amount.Name = "ColT1_Amount"
        Me.ColT1_Amount.ReadOnly = True
        Me.ColT1_Amount.Width = 81
        '
        'T1_TxtAccDesc
        '
        Me.T1_TxtAccDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.T1_TxtAccDesc.Location = New System.Drawing.Point(337, 22)
        Me.T1_TxtAccDesc.Name = "T1_TxtAccDesc"
        Me.T1_TxtAccDesc.ReadOnly = True
        Me.T1_TxtAccDesc.Size = New System.Drawing.Size(444, 22)
        Me.T1_TxtAccDesc.TabIndex = 4
        '
        'T1_btnFindAcc
        '
        Me.T1_btnFindAcc.Location = New System.Drawing.Point(256, 22)
        Me.T1_btnFindAcc.Name = "T1_btnFindAcc"
        Me.T1_btnFindAcc.Size = New System.Drawing.Size(75, 23)
        Me.T1_btnFindAcc.TabIndex = 3
        Me.T1_btnFindAcc.Text = "Find"
        Me.T1_btnFindAcc.UseVisualStyleBackColor = True
        '
        'T1_TxtAccCode
        '
        Me.T1_TxtAccCode.Location = New System.Drawing.Point(135, 23)
        Me.T1_TxtAccCode.Name = "T1_TxtAccCode"
        Me.T1_TxtAccCode.Size = New System.Drawing.Size(115, 22)
        Me.T1_TxtAccCode.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(121, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Inventory Account"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(818, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Tools ini digunakan untuk mencari data product/service yang nyasar sebagai invent" & _
            "ory dan kemudian di jurnal ke akun inventory"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.T2_dgv)
        Me.TabPage2.Controls.Add(Me.T2_BtnFind)
        Me.TabPage2.Controls.Add(Me.Label5)
        Me.TabPage2.Controls.Add(Me.T2_dtp)
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(789, 492)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "TabPage2"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'T2_dgv
        '
        Me.T2_dgv.AllowUserToAddRows = False
        Me.T2_dgv.AllowUserToDeleteRows = False
        Me.T2_dgv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.T2_dgv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.T2_dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.T2_dgv.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColT2_No, Me.ColT2_AOPNo, Me.ColT2_AccCode, Me.ColT2_AccDesc, Me.ColT2_Debet, Me.ColT2_Debet_Baru})
        Me.T2_dgv.Location = New System.Drawing.Point(6, 51)
        Me.T2_dgv.Name = "T2_dgv"
        Me.T2_dgv.ReadOnly = True
        Me.T2_dgv.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.T2_dgv.RowTemplate.Height = 24
        Me.T2_dgv.Size = New System.Drawing.Size(775, 433)
        Me.T2_dgv.TabIndex = 13
        '
        'ColT2_No
        '
        Me.ColT2_No.HeaderText = "No"
        Me.ColT2_No.Name = "ColT2_No"
        Me.ColT2_No.ReadOnly = True
        Me.ColT2_No.Width = 51
        '
        'ColT2_AOPNo
        '
        Me.ColT2_AOPNo.HeaderText = "AOP No"
        Me.ColT2_AOPNo.Name = "ColT2_AOPNo"
        Me.ColT2_AOPNo.ReadOnly = True
        Me.ColT2_AOPNo.Width = 84
        '
        'ColT2_AccCode
        '
        Me.ColT2_AccCode.HeaderText = "Acc Code"
        Me.ColT2_AccCode.Name = "ColT2_AccCode"
        Me.ColT2_AccCode.ReadOnly = True
        Me.ColT2_AccCode.Width = 93
        '
        'ColT2_AccDesc
        '
        Me.ColT2_AccDesc.HeaderText = "Description"
        Me.ColT2_AccDesc.Name = "ColT2_AccDesc"
        Me.ColT2_AccDesc.ReadOnly = True
        Me.ColT2_AccDesc.Width = 104
        '
        'ColT2_Debet
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColT2_Debet.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColT2_Debet.HeaderText = "Debet"
        Me.ColT2_Debet.Name = "ColT2_Debet"
        Me.ColT2_Debet.ReadOnly = True
        Me.ColT2_Debet.Width = 71
        '
        'ColT2_Debet_Baru
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColT2_Debet_Baru.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColT2_Debet_Baru.HeaderText = "Debet - new"
        Me.ColT2_Debet_Baru.Name = "ColT2_Debet_Baru"
        Me.ColT2_Debet_Baru.ReadOnly = True
        Me.ColT2_Debet_Baru.Width = 109
        '
        'T2_BtnFind
        '
        Me.T2_BtnFind.Location = New System.Drawing.Point(339, 22)
        Me.T2_BtnFind.Name = "T2_BtnFind"
        Me.T2_BtnFind.Size = New System.Drawing.Size(75, 23)
        Me.T2_BtnFind.TabIndex = 12
        Me.T2_BtnFind.Text = "Find"
        Me.T2_BtnFind.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 26)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 17)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Period"
        '
        'T2_dtp
        '
        Me.T2_dtp.CustomFormat = "MMM yyyy"
        Me.T2_dtp.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.T2_dtp.Location = New System.Drawing.Point(133, 23)
        Me.T2_dtp.Name = "T2_dtp"
        Me.T2_dtp.ShowUpDown = True
        Me.T2_dtp.Size = New System.Drawing.Size(200, 22)
        Me.T2_dtp.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 3)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(501, 17)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Tools ini digunakan untuk mencari aop yang memiliki selisih jurnal inv/expense"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.T3_LblSelisih)
        Me.TabPage3.Controls.Add(Me.T3_dgv)
        Me.TabPage3.Controls.Add(Me.T3_btnFilter)
        Me.TabPage3.Controls.Add(Me.Label6)
        Me.TabPage3.Controls.Add(Me.T3_dtpPeriod)
        Me.TabPage3.Controls.Add(Me.Label7)
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(789, 492)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "TabPage3"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'T3_LblSelisih
        '
        Me.T3_LblSelisih.AutoSize = True
        Me.T3_LblSelisih.Location = New System.Drawing.Point(421, 28)
        Me.T3_LblSelisih.Name = "T3_LblSelisih"
        Me.T3_LblSelisih.Size = New System.Drawing.Size(85, 17)
        Me.T3_LblSelisih.TabIndex = 19
        Me.T3_LblSelisih.Text = "Total Selisih"
        '
        'T3_dgv
        '
        Me.T3_dgv.AllowUserToAddRows = False
        Me.T3_dgv.AllowUserToDeleteRows = False
        Me.T3_dgv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.T3_dgv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.T3_dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.T3_dgv.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColT3_No, Me.ColT3_BundlingNo, Me.ColT3_In, Me.ColT3_Out, Me.ColT3_Selisih, Me.ColT3_AllocatedAmount})
        Me.T3_dgv.Location = New System.Drawing.Point(7, 54)
        Me.T3_dgv.Name = "T3_dgv"
        Me.T3_dgv.ReadOnly = True
        Me.T3_dgv.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.T3_dgv.RowTemplate.Height = 24
        Me.T3_dgv.Size = New System.Drawing.Size(775, 433)
        Me.T3_dgv.TabIndex = 18
        '
        'ColT3_No
        '
        Me.ColT3_No.HeaderText = "No"
        Me.ColT3_No.Name = "ColT3_No"
        Me.ColT3_No.ReadOnly = True
        Me.ColT3_No.Width = 51
        '
        'ColT3_BundlingNo
        '
        Me.ColT3_BundlingNo.HeaderText = "Bundling No"
        Me.ColT3_BundlingNo.Name = "ColT3_BundlingNo"
        Me.ColT3_BundlingNo.ReadOnly = True
        Me.ColT3_BundlingNo.Width = 110
        '
        'ColT3_In
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColT3_In.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColT3_In.HeaderText = "In"
        Me.ColT3_In.Name = "ColT3_In"
        Me.ColT3_In.ReadOnly = True
        Me.ColT3_In.Width = 44
        '
        'ColT3_Out
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColT3_Out.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColT3_Out.HeaderText = "Out"
        Me.ColT3_Out.Name = "ColT3_Out"
        Me.ColT3_Out.ReadOnly = True
        Me.ColT3_Out.Width = 56
        '
        'ColT3_Selisih
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColT3_Selisih.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColT3_Selisih.HeaderText = "Selisih"
        Me.ColT3_Selisih.Name = "ColT3_Selisih"
        Me.ColT3_Selisih.ReadOnly = True
        Me.ColT3_Selisih.Width = 74
        '
        'ColT3_AllocatedAmount
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColT3_AllocatedAmount.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColT3_AllocatedAmount.HeaderText = "Allocated Amount"
        Me.ColT3_AllocatedAmount.Name = "ColT3_AllocatedAmount"
        Me.ColT3_AllocatedAmount.ReadOnly = True
        Me.ColT3_AllocatedAmount.Width = 143
        '
        'T3_btnFilter
        '
        Me.T3_btnFilter.Location = New System.Drawing.Point(340, 25)
        Me.T3_btnFilter.Name = "T3_btnFilter"
        Me.T3_btnFilter.Size = New System.Drawing.Size(75, 23)
        Me.T3_btnFilter.TabIndex = 17
        Me.T3_btnFilter.Text = "Filter"
        Me.T3_btnFilter.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(7, 29)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 17)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Period"
        '
        'T3_dtpPeriod
        '
        Me.T3_dtpPeriod.CustomFormat = "MMM yyyy"
        Me.T3_dtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.T3_dtpPeriod.Location = New System.Drawing.Point(134, 26)
        Me.T3_dtpPeriod.Name = "T3_dtpPeriod"
        Me.T3_dtpPeriod.ShowUpDown = True
        Me.T3_dtpPeriod.Size = New System.Drawing.Size(200, 22)
        Me.T3_dtpPeriod.TabIndex = 15
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 6)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(689, 17)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Tools ini digunakan untuk mencari item detail yang amount bundlingnya tidak seimb" & _
            "ang pada in dan out-nya"
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.Label9)
        Me.TabPage4.Controls.Add(Me.T4_BtnFilter)
        Me.TabPage4.Controls.Add(Me.T4_TxtPrimaryClosing)
        Me.TabPage4.Controls.Add(Me.Label8)
        Me.TabPage4.Location = New System.Drawing.Point(4, 25)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(789, 492)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "TabPage4"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'T4_BtnFilter
        '
        Me.T4_BtnFilter.Location = New System.Drawing.Point(256, 22)
        Me.T4_BtnFilter.Name = "T4_BtnFilter"
        Me.T4_BtnFilter.Size = New System.Drawing.Size(75, 23)
        Me.T4_BtnFilter.TabIndex = 6
        Me.T4_BtnFilter.Text = "Filter"
        Me.T4_BtnFilter.UseVisualStyleBackColor = True
        '
        'T4_TxtPrimaryClosing
        '
        Me.T4_TxtPrimaryClosing.Location = New System.Drawing.Point(135, 23)
        Me.T4_TxtPrimaryClosing.Name = "T4_TxtPrimaryClosing"
        Me.T4_TxtPrimaryClosing.Size = New System.Drawing.Size(115, 22)
        Me.T4_TxtPrimaryClosing.TabIndex = 5
        Me.T4_TxtPrimaryClosing.Text = "-1"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 25)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(106, 17)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "Primary Closing"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 3)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(375, 17)
        Me.Label9.TabIndex = 7
        Me.Label9.Text = "Tools ini digunakan untuk sinkron data jurnal ke item akun"
        '
        'FrmTools_MonthlyClosing
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(797, 521)
        Me.Controls.Add(Me.TabControl1)
        Me.MinimumSize = New System.Drawing.Size(805, 558)
        Me.Name = "FrmTools_MonthlyClosing"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SMESS Tools"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.T1_dgv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.T2_dgv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.T3_dgv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents T1_TxtAccDesc As System.Windows.Forms.TextBox
    Friend WithEvents T1_btnFindAcc As System.Windows.Forms.Button
    Friend WithEvents T1_TxtAccCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents T1_dgv As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents T1_dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents ColT1_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT1_AOPNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT1_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT1_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT1_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents T1_PrimaryAkunInv As System.Windows.Forms.TextBox
    Friend WithEvents T1_Btnfilter As System.Windows.Forms.Button
    Friend WithEvents T2_BtnFind As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents T2_dtp As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents T2_dgv As System.Windows.Forms.DataGridView
    Friend WithEvents ColT2_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT2_AOPNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT2_AccCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT2_AccDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT2_Debet As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT2_Debet_Baru As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents T3_dgv As System.Windows.Forms.DataGridView
    Friend WithEvents T3_btnFilter As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents T3_dtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents T3_LblSelisih As System.Windows.Forms.Label
    Friend WithEvents ColT3_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT3_BundlingNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT3_In As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT3_Out As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT3_Selisih As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColT3_AllocatedAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents T4_BtnFilter As System.Windows.Forms.Button
    Friend WithEvents T4_TxtPrimaryClosing As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
End Class
