<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanARPaid
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
        Me.txtPrimaryCust = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.txtCurrencyDesc = New System.Windows.Forms.TextBox
        Me.btnFindCurrency = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.TxtPONumber = New System.Windows.Forms.TextBox
        Me.LblCustPONo = New System.Windows.Forms.Label
        Me.TxtPrimaryPO = New System.Windows.Forms.TextBox
        Me.TxtSuppCode = New System.Windows.Forms.TextBox
        Me.btnFindPO = New System.Windows.Forms.Button
        Me.LblPONo = New System.Windows.Forms.Label
        Me.TxtSONumber = New System.Windows.Forms.TextBox
        Me.TxtPrimarySO = New System.Windows.Forms.TextBox
        Me.TxtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.TxtCustomerPONumber = New System.Windows.Forms.TextBox
        Me.btnFindSO = New System.Windows.Forms.Button
        Me.LblSONo = New System.Windows.Forms.Label
        Me.TxtSupplierName = New System.Windows.Forms.TextBox
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.btnFindSupplier = New System.Windows.Forms.Button
        Me.LblSuppNo = New System.Windows.Forms.Label
        Me.TxtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.btnFindCodeNo = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.txtArAccountCode = New System.Windows.Forms.TextBox
        Me.txtCustCode = New System.Windows.Forms.TextBox
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.txtPrimaryArAccount = New System.Windows.Forms.TextBox
        Me.txtArAccountDesc = New System.Windows.Forms.TextBox
        Me.btnFindArAccount = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtCustDesc = New System.Windows.Forms.TextBox
        Me.btnFindCust = New System.Windows.Forms.Button
        Me.LblCustCode = New System.Windows.Forms.Label
        Me.btnShow = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.CboPayment = New System.Windows.Forms.ComboBox
        Me.DtpAsOf = New System.Windows.Forms.DateTimePicker
        Me.Label7 = New System.Windows.Forms.Label
        Me.crvAP = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.chkUseSupplieInvoiceNo = New System.Windows.Forms.CheckBox
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtPrimaryCust
        '
        Me.txtPrimaryCust.Location = New System.Drawing.Point(53, 20)
        Me.txtPrimaryCust.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimaryCust.Name = "txtPrimaryCust"
        Me.txtPrimaryCust.Size = New System.Drawing.Size(36, 22)
        Me.txtPrimaryCust.TabIndex = 20
        Me.txtPrimaryCust.TabStop = False
        Me.txtPrimaryCust.Text = "-1"
        Me.txtPrimaryCust.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(8, 82)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(65, 17)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Currency"
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(116, 162)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 28)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'txtCurrencyDesc
        '
        Me.txtCurrencyDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCurrencyDesc.Location = New System.Drawing.Point(301, 80)
        Me.txtCurrencyDesc.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCurrencyDesc.Name = "txtCurrencyDesc"
        Me.txtCurrencyDesc.ReadOnly = True
        Me.txtCurrencyDesc.Size = New System.Drawing.Size(384, 22)
        Me.txtCurrencyDesc.TabIndex = 15
        Me.txtCurrencyDesc.TabStop = False
        '
        'btnFindCurrency
        '
        Me.btnFindCurrency.Location = New System.Drawing.Point(233, 78)
        Me.btnFindCurrency.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindCurrency.Name = "btnFindCurrency"
        Me.btnFindCurrency.Size = New System.Drawing.Size(60, 28)
        Me.btnFindCurrency.TabIndex = 9
        Me.btnFindCurrency.TabStop = False
        Me.btnFindCurrency.Text = "Find"
        Me.btnFindCurrency.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.TxtPONumber)
        Me.GroupBox3.Controls.Add(Me.LblCustPONo)
        Me.GroupBox3.Controls.Add(Me.TxtPrimaryPO)
        Me.GroupBox3.Controls.Add(Me.TxtSuppCode)
        Me.GroupBox3.Controls.Add(Me.btnFindPO)
        Me.GroupBox3.Controls.Add(Me.LblPONo)
        Me.GroupBox3.Controls.Add(Me.TxtSONumber)
        Me.GroupBox3.Controls.Add(Me.TxtPrimarySO)
        Me.GroupBox3.Controls.Add(Me.TxtPrimarySupplier)
        Me.GroupBox3.Controls.Add(Me.TxtCustomerPONumber)
        Me.GroupBox3.Controls.Add(Me.btnFindSO)
        Me.GroupBox3.Controls.Add(Me.LblSONo)
        Me.GroupBox3.Controls.Add(Me.TxtSupplierName)
        Me.GroupBox3.Controls.Add(Me.TxtCodeNo)
        Me.GroupBox3.Controls.Add(Me.btnFindSupplier)
        Me.GroupBox3.Controls.Add(Me.LblSuppNo)
        Me.GroupBox3.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.GroupBox3.Controls.Add(Me.TxtDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindCodeNo)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.Controls.Add(Me.txtCurrencyCode)
        Me.GroupBox3.Controls.Add(Me.txtArAccountCode)
        Me.GroupBox3.Controls.Add(Me.txtCustCode)
        Me.GroupBox3.Controls.Add(Me.txtPrimaryCurrency)
        Me.GroupBox3.Controls.Add(Me.txtPrimaryArAccount)
        Me.GroupBox3.Controls.Add(Me.txtPrimaryCust)
        Me.GroupBox3.Controls.Add(Me.txtCurrencyDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindCurrency)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.txtArAccountDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindArAccount)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.txtCustDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindCust)
        Me.GroupBox3.Controls.Add(Me.LblCustCode)
        Me.GroupBox3.Location = New System.Drawing.Point(265, 12)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox3.Size = New System.Drawing.Size(693, 167)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Filter"
        '
        'TxtPONumber
        '
        Me.TxtPONumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPONumber.Location = New System.Drawing.Point(99, 139)
        Me.TxtPONumber.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPONumber.Name = "TxtPONumber"
        Me.TxtPONumber.Size = New System.Drawing.Size(127, 22)
        Me.TxtPONumber.TabIndex = 2
        '
        'LblCustPONo
        '
        Me.LblCustPONo.AutoSize = True
        Me.LblCustPONo.Location = New System.Drawing.Point(301, 143)
        Me.LblCustPONo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblCustPONo.Name = "LblCustPONo"
        Me.LblCustPONo.Size = New System.Drawing.Size(118, 17)
        Me.LblCustPONo.TabIndex = 35
        Me.LblCustPONo.Text = "Cust. PO Number"
        '
        'TxtPrimaryPO
        '
        Me.TxtPrimaryPO.Location = New System.Drawing.Point(55, 139)
        Me.TxtPrimaryPO.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryPO.Name = "TxtPrimaryPO"
        Me.TxtPrimaryPO.Size = New System.Drawing.Size(36, 22)
        Me.TxtPrimaryPO.TabIndex = 39
        Me.TxtPrimaryPO.TabStop = False
        Me.TxtPrimaryPO.Text = "-1"
        Me.TxtPrimaryPO.Visible = False
        '
        'TxtSuppCode
        '
        Me.TxtSuppCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSuppCode.Location = New System.Drawing.Point(100, 21)
        Me.TxtSuppCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSuppCode.Name = "TxtSuppCode"
        Me.TxtSuppCode.Size = New System.Drawing.Size(127, 22)
        Me.TxtSuppCode.TabIndex = 1
        '
        'btnFindPO
        '
        Me.btnFindPO.Location = New System.Drawing.Point(233, 137)
        Me.btnFindPO.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindPO.Name = "btnFindPO"
        Me.btnFindPO.Size = New System.Drawing.Size(60, 28)
        Me.btnFindPO.TabIndex = 37
        Me.btnFindPO.TabStop = False
        Me.btnFindPO.Text = "Find"
        Me.btnFindPO.UseVisualStyleBackColor = True
        '
        'LblPONo
        '
        Me.LblPONo.AutoSize = True
        Me.LblPONo.Location = New System.Drawing.Point(8, 143)
        Me.LblPONo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblPONo.Name = "LblPONo"
        Me.LblPONo.Size = New System.Drawing.Size(50, 17)
        Me.LblPONo.TabIndex = 38
        Me.LblPONo.Text = "PO No"
        '
        'TxtSONumber
        '
        Me.TxtSONumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSONumber.Location = New System.Drawing.Point(99, 140)
        Me.TxtSONumber.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSONumber.Name = "TxtSONumber"
        Me.TxtSONumber.Size = New System.Drawing.Size(127, 22)
        Me.TxtSONumber.TabIndex = 4
        '
        'TxtPrimarySO
        '
        Me.TxtPrimarySO.Location = New System.Drawing.Point(53, 140)
        Me.TxtPrimarySO.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimarySO.Name = "TxtPrimarySO"
        Me.TxtPrimarySO.Size = New System.Drawing.Size(36, 22)
        Me.TxtPrimarySO.TabIndex = 34
        Me.TxtPrimarySO.TabStop = False
        Me.TxtPrimarySO.Text = "-1"
        Me.TxtPrimarySO.Visible = False
        '
        'TxtPrimarySupplier
        '
        Me.TxtPrimarySupplier.Location = New System.Drawing.Point(56, 21)
        Me.TxtPrimarySupplier.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimarySupplier.Name = "TxtPrimarySupplier"
        Me.TxtPrimarySupplier.Size = New System.Drawing.Size(36, 22)
        Me.TxtPrimarySupplier.TabIndex = 25
        Me.TxtPrimarySupplier.TabStop = False
        Me.TxtPrimarySupplier.Text = "-1"
        Me.TxtPrimarySupplier.Visible = False
        '
        'TxtCustomerPONumber
        '
        Me.TxtCustomerPONumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerPONumber.Location = New System.Drawing.Point(427, 140)
        Me.TxtCustomerPONumber.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCustomerPONumber.Name = "TxtCustomerPONumber"
        Me.TxtCustomerPONumber.ReadOnly = True
        Me.TxtCustomerPONumber.Size = New System.Drawing.Size(257, 22)
        Me.TxtCustomerPONumber.TabIndex = 33
        Me.TxtCustomerPONumber.TabStop = False
        '
        'btnFindSO
        '
        Me.btnFindSO.Location = New System.Drawing.Point(233, 137)
        Me.btnFindSO.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindSO.Name = "btnFindSO"
        Me.btnFindSO.Size = New System.Drawing.Size(60, 28)
        Me.btnFindSO.TabIndex = 31
        Me.btnFindSO.TabStop = False
        Me.btnFindSO.Text = "Find"
        Me.btnFindSO.UseVisualStyleBackColor = True
        '
        'LblSONo
        '
        Me.LblSONo.AutoSize = True
        Me.LblSONo.Location = New System.Drawing.Point(8, 143)
        Me.LblSONo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblSONo.Name = "LblSONo"
        Me.LblSONo.Size = New System.Drawing.Size(50, 17)
        Me.LblSONo.TabIndex = 32
        Me.LblSONo.Text = "SO No"
        '
        'TxtSupplierName
        '
        Me.TxtSupplierName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSupplierName.Location = New System.Drawing.Point(303, 21)
        Me.TxtSupplierName.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSupplierName.Name = "TxtSupplierName"
        Me.TxtSupplierName.ReadOnly = True
        Me.TxtSupplierName.Size = New System.Drawing.Size(384, 22)
        Me.TxtSupplierName.TabIndex = 24
        Me.TxtSupplierName.TabStop = False
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNo.Location = New System.Drawing.Point(99, 110)
        Me.TxtCodeNo.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.Size = New System.Drawing.Size(127, 22)
        Me.TxtCodeNo.TabIndex = 3
        '
        'btnFindSupplier
        '
        Me.btnFindSupplier.Location = New System.Drawing.Point(235, 17)
        Me.btnFindSupplier.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindSupplier.Name = "btnFindSupplier"
        Me.btnFindSupplier.Size = New System.Drawing.Size(60, 28)
        Me.btnFindSupplier.TabIndex = 23
        Me.btnFindSupplier.TabStop = False
        Me.btnFindSupplier.Text = "Find"
        Me.btnFindSupplier.UseVisualStyleBackColor = True
        '
        'LblSuppNo
        '
        Me.LblSuppNo.AutoSize = True
        Me.LblSuppNo.Location = New System.Drawing.Point(9, 23)
        Me.LblSuppNo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblSuppNo.Name = "LblSuppNo"
        Me.LblSuppNo.Size = New System.Drawing.Size(60, 17)
        Me.LblSuppNo.TabIndex = 21
        Me.LblSuppNo.Text = "Supplier"
        '
        'TxtPrimaryCodeNo
        '
        Me.TxtPrimaryCodeNo.Location = New System.Drawing.Point(53, 110)
        Me.TxtPrimaryCodeNo.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryCodeNo.Name = "TxtPrimaryCodeNo"
        Me.TxtPrimaryCodeNo.Size = New System.Drawing.Size(36, 22)
        Me.TxtPrimaryCodeNo.TabIndex = 29
        Me.TxtPrimaryCodeNo.TabStop = False
        Me.TxtPrimaryCodeNo.Text = "-1"
        Me.TxtPrimaryCodeNo.Visible = False
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(301, 110)
        Me.TxtDesc.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.ReadOnly = True
        Me.TxtDesc.Size = New System.Drawing.Size(384, 22)
        Me.TxtDesc.TabIndex = 28
        Me.TxtDesc.TabStop = False
        '
        'btnFindCodeNo
        '
        Me.btnFindCodeNo.Location = New System.Drawing.Point(233, 107)
        Me.btnFindCodeNo.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindCodeNo.Name = "btnFindCodeNo"
        Me.btnFindCodeNo.Size = New System.Drawing.Size(60, 28)
        Me.btnFindCodeNo.TabIndex = 26
        Me.btnFindCodeNo.TabStop = False
        Me.btnFindCodeNo.Text = "Find"
        Me.btnFindCodeNo.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 113)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(63, 17)
        Me.Label9.TabIndex = 27
        Me.Label9.Text = "Code No"
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCurrencyCode.Location = New System.Drawing.Point(99, 80)
        Me.txtCurrencyCode.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.Size = New System.Drawing.Size(127, 22)
        Me.txtCurrencyCode.TabIndex = 2
        '
        'txtArAccountCode
        '
        Me.txtArAccountCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtArAccountCode.Location = New System.Drawing.Point(99, 50)
        Me.txtArAccountCode.Margin = New System.Windows.Forms.Padding(4)
        Me.txtArAccountCode.Name = "txtArAccountCode"
        Me.txtArAccountCode.Size = New System.Drawing.Size(127, 22)
        Me.txtArAccountCode.TabIndex = 1
        '
        'txtCustCode
        '
        Me.txtCustCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustCode.Location = New System.Drawing.Point(99, 20)
        Me.txtCustCode.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCustCode.Name = "txtCustCode"
        Me.txtCustCode.Size = New System.Drawing.Size(127, 22)
        Me.txtCustCode.TabIndex = 0
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(53, 80)
        Me.txtPrimaryCurrency.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(36, 22)
        Me.txtPrimaryCurrency.TabIndex = 24
        Me.txtPrimaryCurrency.TabStop = False
        Me.txtPrimaryCurrency.Text = "-1"
        Me.txtPrimaryCurrency.Visible = False
        '
        'txtPrimaryArAccount
        '
        Me.txtPrimaryArAccount.Location = New System.Drawing.Point(53, 50)
        Me.txtPrimaryArAccount.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimaryArAccount.Name = "txtPrimaryArAccount"
        Me.txtPrimaryArAccount.Size = New System.Drawing.Size(36, 22)
        Me.txtPrimaryArAccount.TabIndex = 23
        Me.txtPrimaryArAccount.TabStop = False
        Me.txtPrimaryArAccount.Text = "-1"
        Me.txtPrimaryArAccount.Visible = False
        '
        'txtArAccountDesc
        '
        Me.txtArAccountDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtArAccountDesc.Location = New System.Drawing.Point(301, 50)
        Me.txtArAccountDesc.Margin = New System.Windows.Forms.Padding(4)
        Me.txtArAccountDesc.Name = "txtArAccountDesc"
        Me.txtArAccountDesc.ReadOnly = True
        Me.txtArAccountDesc.Size = New System.Drawing.Size(384, 22)
        Me.txtArAccountDesc.TabIndex = 11
        Me.txtArAccountDesc.TabStop = False
        '
        'btnFindArAccount
        '
        Me.btnFindArAccount.Location = New System.Drawing.Point(233, 48)
        Me.btnFindArAccount.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindArAccount.Name = "btnFindArAccount"
        Me.btnFindArAccount.Size = New System.Drawing.Size(60, 28)
        Me.btnFindArAccount.TabIndex = 7
        Me.btnFindArAccount.TabStop = False
        Me.btnFindArAccount.Text = "Find"
        Me.btnFindArAccount.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 54)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(82, 17)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "AR Account"
        '
        'txtCustDesc
        '
        Me.txtCustDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCustDesc.Location = New System.Drawing.Point(301, 20)
        Me.txtCustDesc.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCustDesc.Name = "txtCustDesc"
        Me.txtCustDesc.ReadOnly = True
        Me.txtCustDesc.Size = New System.Drawing.Size(384, 22)
        Me.txtCustDesc.TabIndex = 3
        Me.txtCustDesc.TabStop = False
        '
        'btnFindCust
        '
        Me.btnFindCust.Location = New System.Drawing.Point(233, 17)
        Me.btnFindCust.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindCust.Name = "btnFindCust"
        Me.btnFindCust.Size = New System.Drawing.Size(60, 28)
        Me.btnFindCust.TabIndex = 1
        Me.btnFindCust.TabStop = False
        Me.btnFindCust.Text = "Find"
        Me.btnFindCust.UseVisualStyleBackColor = True
        '
        'LblCustCode
        '
        Me.LblCustCode.AutoSize = True
        Me.LblCustCode.Location = New System.Drawing.Point(8, 23)
        Me.LblCustCode.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblCustCode.Name = "LblCustCode"
        Me.LblCustCode.Size = New System.Drawing.Size(68, 17)
        Me.LblCustCode.TabIndex = 0
        Me.LblCustCode.Text = "Customer"
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(8, 162)
        Me.btnShow.Margin = New System.Windows.Forms.Padding(4)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(100, 28)
        Me.btnShow.TabIndex = 4
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.chkUseSupplieInvoiceNo)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.CboPayment)
        Me.GroupBox1.Controls.Add(Me.DtpAsOf)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.btnShow)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 11)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(965, 212)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.dtpFrom)
        Me.GroupBox2.Controls.Add(Me.dtpTo)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 71)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Size = New System.Drawing.Size(249, 84)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Document Date"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(7, 62)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(25, 17)
        Me.Label8.TabIndex = 10
        Me.Label8.Text = "To"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 31)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "From"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(75, 28)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(168, 22)
        Me.dtpFrom.TabIndex = 0
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(75, 58)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(168, 22)
        Me.dtpTo.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 44)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Payment"
        '
        'CboPayment
        '
        Me.CboPayment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPayment.FormattingEnabled = True
        Me.CboPayment.Location = New System.Drawing.Point(83, 41)
        Me.CboPayment.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CboPayment.Name = "CboPayment"
        Me.CboPayment.Size = New System.Drawing.Size(175, 24)
        Me.CboPayment.TabIndex = 1
        '
        'DtpAsOf
        '
        Me.DtpAsOf.CustomFormat = "dd MMM yyyy"
        Me.DtpAsOf.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpAsOf.Location = New System.Drawing.Point(83, 12)
        Me.DtpAsOf.Margin = New System.Windows.Forms.Padding(4)
        Me.DtpAsOf.Name = "DtpAsOf"
        Me.DtpAsOf.Size = New System.Drawing.Size(175, 22)
        Me.DtpAsOf.TabIndex = 0
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 15)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(43, 17)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "As Of"
        '
        'crvAP
        '
        Me.crvAP.ActiveViewIndex = -1
        Me.crvAP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvAP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvAP.DisplayGroupTree = False
        Me.crvAP.EnableDrillDown = False
        Me.crvAP.Location = New System.Drawing.Point(13, 231)
        Me.crvAP.Margin = New System.Windows.Forms.Padding(4)
        Me.crvAP.Name = "crvAP"
        Me.crvAP.SelectionFormula = ""
        Me.crvAP.ShowCloseButton = False
        Me.crvAP.ShowGroupTreeButton = False
        Me.crvAP.ShowRefreshButton = False
        Me.crvAP.Size = New System.Drawing.Size(966, 391)
        Me.crvAP.TabIndex = 3
        Me.crvAP.TabStop = False
        Me.crvAP.ViewTimeSelectionFormula = ""
        '
        'chkUseSupplieInvoiceNo
        '
        Me.chkUseSupplieInvoiceNo.AutoSize = True
        Me.chkUseSupplieInvoiceNo.Location = New System.Drawing.Point(265, 186)
        Me.chkUseSupplieInvoiceNo.Name = "chkUseSupplieInvoiceNo"
        Me.chkUseSupplieInvoiceNo.Size = New System.Drawing.Size(181, 21)
        Me.chkUseSupplieInvoiceNo.TabIndex = 10
        Me.chkUseSupplieInvoiceNo.Text = "Use Supplier Invoice No"
        Me.chkUseSupplieInvoiceNo.UseVisualStyleBackColor = True
        '
        'FrmLaporanARPaid
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(995, 634)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.crvAP)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimumSize = New System.Drawing.Size(1003, 671)
        Me.Name = "FrmLaporanARPaid"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Account Receivable Paid Report"
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtPrimaryCust As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents txtCurrencyDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCurrency As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryArAccount As System.Windows.Forms.TextBox
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents txtArAccountDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindArAccount As System.Windows.Forms.Button
    Friend WithEvents txtArAccountCode As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtCustDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCust As System.Windows.Forms.Button
    Friend WithEvents txtCustCode As System.Windows.Forms.TextBox
    Friend WithEvents LblCustCode As System.Windows.Forms.Label
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents crvAP As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents DtpAsOf As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CboPayment As System.Windows.Forms.ComboBox
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents LblCustPONo As System.Windows.Forms.Label
    Friend WithEvents TxtSONumber As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimarySO As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerPONumber As System.Windows.Forms.TextBox
    Friend WithEvents btnFindSO As System.Windows.Forms.Button
    Friend WithEvents LblSONo As System.Windows.Forms.Label
    Friend WithEvents TxtSuppCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents TxtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents btnFindSupplier As System.Windows.Forms.Button
    Friend WithEvents LblSuppNo As System.Windows.Forms.Label
    Friend WithEvents TxtPONumber As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryPO As System.Windows.Forms.TextBox
    Friend WithEvents btnFindPO As System.Windows.Forms.Button
    Friend WithEvents LblPONo As System.Windows.Forms.Label
    Friend WithEvents chkUseSupplieInvoiceNo As System.Windows.Forms.CheckBox
End Class
