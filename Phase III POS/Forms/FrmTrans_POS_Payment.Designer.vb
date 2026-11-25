<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_POS_Payment
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.LblKeterangan = New System.Windows.Forms.Label
        Me.spc = New System.Windows.Forms.SplitContainer
        Me.Label21 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtApprovalCode = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtAccountNumber = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtChange = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtTotalPayment = New System.Windows.Forms.TextBox
        Me.TxtPrimaryPayment = New System.Windows.Forms.TextBox
        Me.cboPaymentTipe = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtTotalPayment2 = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtChange2 = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.Col_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPaymentCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPaymentType = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColPaymentBankCharge = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPaymentAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccountNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColApprovalCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPayment = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.TxtBankCharge = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtRounding = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtGrandTotal = New System.Windows.Forms.TextBox
        Me.TxtTaxPercentage = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.TxtTaxAmount = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.TxtSubTotal = New System.Windows.Forms.TextBox
        Me.TxtDiscString = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.TxtDisc = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtTotalSales = New System.Windows.Forms.TextBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtWithdrawal = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.spc.Panel1.SuspendLayout()
        Me.spc.Panel2.SuspendLayout()
        Me.spc.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LblKeterangan
        '
        Me.LblKeterangan.AutoSize = True
        Me.LblKeterangan.Location = New System.Drawing.Point(12, 9)
        Me.LblKeterangan.Name = "LblKeterangan"
        Me.LblKeterangan.Size = New System.Drawing.Size(208, 13)
        Me.LblKeterangan.TabIndex = 0
        Me.LblKeterangan.Text = "F1 - Single Payment, F2 - Multiple Payment"
        '
        'spc
        '
        Me.spc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.spc.Location = New System.Drawing.Point(12, 207)
        Me.spc.Name = "spc"
        '
        'spc.Panel1
        '
        Me.spc.Panel1.Controls.Add(Me.Label21)
        Me.spc.Panel1.Controls.Add(Me.TxtRemarks)
        Me.spc.Panel1.Controls.Add(Me.Label8)
        Me.spc.Panel1.Controls.Add(Me.TxtApprovalCode)
        Me.spc.Panel1.Controls.Add(Me.Label7)
        Me.spc.Panel1.Controls.Add(Me.TxtAccountNumber)
        Me.spc.Panel1.Controls.Add(Me.Label6)
        Me.spc.Panel1.Controls.Add(Me.TxtChange)
        Me.spc.Panel1.Controls.Add(Me.Label5)
        Me.spc.Panel1.Controls.Add(Me.TxtTotalPayment)
        Me.spc.Panel1.Controls.Add(Me.TxtPrimaryPayment)
        Me.spc.Panel1.Controls.Add(Me.cboPaymentTipe)
        Me.spc.Panel1.Controls.Add(Me.Label2)
        '
        'spc.Panel2
        '
        Me.spc.Panel2.Controls.Add(Me.SplitContainer1)
        Me.spc.Size = New System.Drawing.Size(554, 201)
        Me.spc.SplitterDistance = 275
        Me.spc.TabIndex = 1
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(3, 91)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(49, 13)
        Me.Label21.TabIndex = 30
        Me.Label21.Text = "Remarks"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(103, 88)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Multiline = True
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(169, 53)
        Me.TxtRemarks.TabIndex = 3
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 176)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(77, 13)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Approval Code"
        '
        'TxtApprovalCode
        '
        Me.TxtApprovalCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtApprovalCode.Location = New System.Drawing.Point(103, 173)
        Me.TxtApprovalCode.MaxLength = 250
        Me.TxtApprovalCode.Name = "TxtApprovalCode"
        Me.TxtApprovalCode.Size = New System.Drawing.Size(169, 20)
        Me.TxtApprovalCode.TabIndex = 5
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 150)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(87, 13)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "Account Number"
        '
        'TxtAccountNumber
        '
        Me.TxtAccountNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccountNumber.Location = New System.Drawing.Point(103, 147)
        Me.TxtAccountNumber.MaxLength = 250
        Me.TxtAccountNumber.Name = "TxtAccountNumber"
        Me.TxtAccountNumber.Size = New System.Drawing.Size(169, 20)
        Me.TxtAccountNumber.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 65)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(44, 13)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Change"
        '
        'TxtChange
        '
        Me.TxtChange.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtChange.Location = New System.Drawing.Point(103, 62)
        Me.TxtChange.Name = "TxtChange"
        Me.TxtChange.ReadOnly = True
        Me.TxtChange.Size = New System.Drawing.Size(169, 20)
        Me.TxtChange.TabIndex = 2
        Me.TxtChange.TabStop = False
        Me.TxtChange.Text = "0"
        Me.TxtChange.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 39)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(75, 13)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Total Payment"
        '
        'TxtTotalPayment
        '
        Me.TxtTotalPayment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalPayment.Location = New System.Drawing.Point(103, 36)
        Me.TxtTotalPayment.MaxLength = 12
        Me.TxtTotalPayment.Name = "TxtTotalPayment"
        Me.TxtTotalPayment.Size = New System.Drawing.Size(169, 20)
        Me.TxtTotalPayment.TabIndex = 1
        Me.TxtTotalPayment.Text = "0"
        Me.TxtTotalPayment.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtPrimaryPayment
        '
        Me.TxtPrimaryPayment.Location = New System.Drawing.Point(66, 9)
        Me.TxtPrimaryPayment.Name = "TxtPrimaryPayment"
        Me.TxtPrimaryPayment.Size = New System.Drawing.Size(31, 20)
        Me.TxtPrimaryPayment.TabIndex = 2
        Me.TxtPrimaryPayment.TabStop = False
        Me.TxtPrimaryPayment.Text = "-1"
        Me.TxtPrimaryPayment.Visible = False
        '
        'cboPaymentTipe
        '
        Me.cboPaymentTipe.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboPaymentTipe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPaymentTipe.FormattingEnabled = True
        Me.cboPaymentTipe.Location = New System.Drawing.Point(103, 9)
        Me.cboPaymentTipe.Name = "cboPaymentTipe"
        Me.cboPaymentTipe.Size = New System.Drawing.Size(169, 21)
        Me.cboPaymentTipe.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(75, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Payment Type"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtWithdrawal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label15)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtTotalPayment2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtChange2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label9)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvItem)
        Me.SplitContainer1.Size = New System.Drawing.Size(275, 201)
        Me.SplitContainer1.SplitterDistance = 86
        Me.SplitContainer1.TabIndex = 27
        Me.SplitContainer1.TabStop = False
        '
        'TxtTotalPayment2
        '
        Me.TxtTotalPayment2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalPayment2.Location = New System.Drawing.Point(103, 9)
        Me.TxtTotalPayment2.Name = "TxtTotalPayment2"
        Me.TxtTotalPayment2.ReadOnly = True
        Me.TxtTotalPayment2.Size = New System.Drawing.Size(169, 20)
        Me.TxtTotalPayment2.TabIndex = 0
        Me.TxtTotalPayment2.TabStop = False
        Me.TxtTotalPayment2.Text = "0"
        Me.TxtTotalPayment2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 12)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(75, 13)
        Me.Label11.TabIndex = 12
        Me.Label11.Text = "Total Payment"
        '
        'TxtChange2
        '
        Me.TxtChange2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtChange2.Location = New System.Drawing.Point(103, 35)
        Me.TxtChange2.Name = "TxtChange2"
        Me.TxtChange2.ReadOnly = True
        Me.TxtChange2.Size = New System.Drawing.Size(169, 20)
        Me.TxtChange2.TabIndex = 1
        Me.TxtChange2.TabStop = False
        Me.TxtChange2.Text = "0"
        Me.TxtChange2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 38)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(44, 13)
        Me.Label9.TabIndex = 16
        Me.Label9.Text = "Change"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col_No, Me.ColPaymentCode, Me.ColPaymentType, Me.ColPaymentBankCharge, Me.ColPaymentAmount, Me.ColAccountNumber, Me.ColApprovalCode, Me.ColRemarks, Me.ColPrimaryPayment})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(0, 0)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.Size = New System.Drawing.Size(275, 111)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 0
        '
        'Col_No
        '
        Me.Col_No.HeaderText = "No"
        Me.Col_No.Name = "Col_No"
        Me.Col_No.Width = 46
        '
        'ColPaymentCode
        '
        Me.ColPaymentCode.HeaderText = "Payment Code"
        Me.ColPaymentCode.Name = "ColPaymentCode"
        Me.ColPaymentCode.Width = 101
        '
        'ColPaymentType
        '
        Me.ColPaymentType.HeaderText = "Payment Type"
        Me.ColPaymentType.Items.AddRange(New Object() {"Credit_Card_Payment", "Debit_Card_Payment", "Cash_Withdrawal", "Voucher_Payment", "Cash"})
        Me.ColPaymentType.Name = "ColPaymentType"
        Me.ColPaymentType.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPaymentType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'ColPaymentBankCharge
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPaymentBankCharge.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColPaymentBankCharge.HeaderText = "Bank Charge"
        Me.ColPaymentBankCharge.Name = "ColPaymentBankCharge"
        Me.ColPaymentBankCharge.Width = 94
        '
        'ColPaymentAmount
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPaymentAmount.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColPaymentAmount.HeaderText = "Payment Amount"
        Me.ColPaymentAmount.MaxInputLength = 12
        Me.ColPaymentAmount.Name = "ColPaymentAmount"
        Me.ColPaymentAmount.Width = 112
        '
        'ColAccountNumber
        '
        Me.ColAccountNumber.HeaderText = "Account Number"
        Me.ColAccountNumber.MaxInputLength = 150
        Me.ColAccountNumber.Name = "ColAccountNumber"
        Me.ColAccountNumber.Width = 112
        '
        'ColApprovalCode
        '
        Me.ColApprovalCode.HeaderText = "Approval Code"
        Me.ColApprovalCode.MaxInputLength = 150
        Me.ColApprovalCode.Name = "ColApprovalCode"
        Me.ColApprovalCode.Width = 102
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.MaxInputLength = 150
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'ColPrimaryPayment
        '
        Me.ColPrimaryPayment.HeaderText = "PrimaryPayment"
        Me.ColPrimaryPayment.Name = "ColPrimaryPayment"
        Me.ColPrimaryPayment.Visible = False
        Me.ColPrimaryPayment.Width = 107
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 35.0!)
        Me.Label3.Location = New System.Drawing.Point(12, 414)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(270, 54)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Grand Total"
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(491, 414)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(491, 443)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'TxtBankCharge
        '
        Me.TxtBankCharge.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBankCharge.Location = New System.Drawing.Point(115, 155)
        Me.TxtBankCharge.Name = "TxtBankCharge"
        Me.TxtBankCharge.ReadOnly = True
        Me.TxtBankCharge.Size = New System.Drawing.Size(448, 20)
        Me.TxtBankCharge.TabIndex = 64
        Me.TxtBankCharge.TabStop = False
        Me.TxtBankCharge.Text = "0"
        Me.TxtBankCharge.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(14, 158)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(69, 13)
        Me.Label22.TabIndex = 63
        Me.Label22.Text = "Bank Charge"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 184)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 13)
        Me.Label1.TabIndex = 62
        Me.Label1.Text = "Grand Total"
        '
        'TxtRounding
        '
        Me.TxtRounding.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRounding.Location = New System.Drawing.Point(116, 129)
        Me.TxtRounding.Name = "TxtRounding"
        Me.TxtRounding.ReadOnly = True
        Me.TxtRounding.Size = New System.Drawing.Size(447, 20)
        Me.TxtRounding.TabIndex = 61
        Me.TxtRounding.TabStop = False
        Me.TxtRounding.Text = "0"
        Me.TxtRounding.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(15, 132)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(53, 13)
        Me.Label16.TabIndex = 60
        Me.Label16.Text = "Rounding"
        '
        'TxtGrandTotal
        '
        Me.TxtGrandTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtGrandTotal.Location = New System.Drawing.Point(115, 181)
        Me.TxtGrandTotal.Name = "TxtGrandTotal"
        Me.TxtGrandTotal.ReadOnly = True
        Me.TxtGrandTotal.Size = New System.Drawing.Size(448, 20)
        Me.TxtGrandTotal.TabIndex = 59
        Me.TxtGrandTotal.TabStop = False
        Me.TxtGrandTotal.Text = "0"
        Me.TxtGrandTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtTaxPercentage
        '
        Me.TxtTaxPercentage.Location = New System.Drawing.Point(115, 103)
        Me.TxtTaxPercentage.Name = "TxtTaxPercentage"
        Me.TxtTaxPercentage.ReadOnly = True
        Me.TxtTaxPercentage.Size = New System.Drawing.Size(68, 20)
        Me.TxtTaxPercentage.TabIndex = 58
        Me.TxtTaxPercentage.TabStop = False
        Me.TxtTaxPercentage.Text = "0"
        Me.TxtTaxPercentage.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(15, 106)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(25, 13)
        Me.Label14.TabIndex = 57
        Me.Label14.Text = "Tax"
        '
        'TxtTaxAmount
        '
        Me.TxtTaxAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTaxAmount.Location = New System.Drawing.Point(210, 103)
        Me.TxtTaxAmount.Name = "TxtTaxAmount"
        Me.TxtTaxAmount.ReadOnly = True
        Me.TxtTaxAmount.Size = New System.Drawing.Size(353, 20)
        Me.TxtTaxAmount.TabIndex = 56
        Me.TxtTaxAmount.TabStop = False
        Me.TxtTaxAmount.Text = "0"
        Me.TxtTaxAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(15, 80)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(53, 13)
        Me.Label13.TabIndex = 55
        Me.Label13.Text = "Sub Total"
        '
        'TxtSubTotal
        '
        Me.TxtSubTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSubTotal.Location = New System.Drawing.Point(115, 77)
        Me.TxtSubTotal.Name = "TxtSubTotal"
        Me.TxtSubTotal.ReadOnly = True
        Me.TxtSubTotal.Size = New System.Drawing.Size(448, 20)
        Me.TxtSubTotal.TabIndex = 54
        Me.TxtSubTotal.TabStop = False
        Me.TxtSubTotal.Text = "0"
        Me.TxtSubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtDiscString
        '
        Me.TxtDiscString.Location = New System.Drawing.Point(115, 51)
        Me.TxtDiscString.Name = "TxtDiscString"
        Me.TxtDiscString.Size = New System.Drawing.Size(89, 20)
        Me.TxtDiscString.TabIndex = 53
        Me.TxtDiscString.Text = "0%"
        Me.TxtDiscString.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(15, 54)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(49, 13)
        Me.Label12.TabIndex = 52
        Me.Label12.Text = "Discount"
        '
        'TxtDisc
        '
        Me.TxtDisc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDisc.Location = New System.Drawing.Point(210, 51)
        Me.TxtDisc.Name = "TxtDisc"
        Me.TxtDisc.ReadOnly = True
        Me.TxtDisc.Size = New System.Drawing.Size(353, 20)
        Me.TxtDisc.TabIndex = 51
        Me.TxtDisc.TabStop = False
        Me.TxtDisc.Text = "0"
        Me.TxtDisc.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(15, 28)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 13)
        Me.Label4.TabIndex = 50
        Me.Label4.Text = "Total Sales"
        '
        'TxtTotalSales
        '
        Me.TxtTotalSales.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalSales.Location = New System.Drawing.Point(115, 25)
        Me.TxtTotalSales.Name = "TxtTotalSales"
        Me.TxtTotalSales.ReadOnly = True
        Me.TxtTotalSales.Size = New System.Drawing.Size(448, 20)
        Me.TxtTotalSales.TabIndex = 49
        Me.TxtTotalSales.TabStop = False
        Me.TxtTotalSales.Text = "0"
        Me.TxtTotalSales.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'BackgroundWorker1
        '
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(189, 106)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(15, 13)
        Me.Label10.TabIndex = 65
        Me.Label10.Text = "%"
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(82, 181)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(27, 20)
        Me.TxtReadOnly.TabIndex = 66
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'TxtWithdrawal
        '
        Me.TxtWithdrawal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtWithdrawal.Location = New System.Drawing.Point(103, 61)
        Me.TxtWithdrawal.Name = "TxtWithdrawal"
        Me.TxtWithdrawal.ReadOnly = True
        Me.TxtWithdrawal.Size = New System.Drawing.Size(169, 20)
        Me.TxtWithdrawal.TabIndex = 17
        Me.TxtWithdrawal.TabStop = False
        Me.TxtWithdrawal.Text = "0"
        Me.TxtWithdrawal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(3, 64)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(60, 13)
        Me.Label15.TabIndex = 18
        Me.Label15.Text = "Withdrawal"
        '
        'FrmTrans_POS_Payment
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(569, 474)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.TxtBankCharge)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtRounding)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.TxtGrandTotal)
        Me.Controls.Add(Me.TxtTaxPercentage)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.TxtTaxAmount)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.TxtSubTotal)
        Me.Controls.Add(Me.TxtDiscString)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.TxtDisc)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtTotalSales)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.spc)
        Me.Controls.Add(Me.LblKeterangan)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(585, 457)
        Me.Name = "FrmTrans_POS_Payment"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Payment Details"
        Me.spc.Panel1.ResumeLayout(False)
        Me.spc.Panel1.PerformLayout()
        Me.spc.Panel2.ResumeLayout(False)
        Me.spc.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LblKeterangan As System.Windows.Forms.Label
    Friend WithEvents spc As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtPrimaryPayment As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtChange As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalPayment As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtApprovalCode As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtAccountNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtChange2 As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalPayment2 As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents TxtBankCharge As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtRounding As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtGrandTotal As System.Windows.Forms.TextBox
    Friend WithEvents TxtTaxPercentage As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TxtTaxAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TxtSubTotal As System.Windows.Forms.TextBox
    Friend WithEvents TxtDiscString As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtDisc As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalSales As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Col_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPaymentCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPaymentType As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColPaymentBankCharge As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPaymentAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccountNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColApprovalCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPayment As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents cboPaymentTipe As System.Windows.Forms.ComboBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents TxtWithdrawal As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
End Class
