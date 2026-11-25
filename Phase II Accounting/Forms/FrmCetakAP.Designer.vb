<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCetakAP
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
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.txtPrimaryPaymentAccount = New System.Windows.Forms.TextBox
        Me.rboArAccount = New System.Windows.Forms.RadioButton
        Me.txtPrimaryArAccount = New System.Windows.Forms.TextBox
        Me.rboCurrency = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.rboDetail = New System.Windows.Forms.RadioButton
        Me.rboPaymentJournal = New System.Windows.Forms.RadioButton
        Me.rboPaymentAccount = New System.Windows.Forms.RadioButton
        Me.rboSuppCode = New System.Windows.Forms.RadioButton
        Me.txtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.btnShow = New System.Windows.Forms.Button
        Me.txtCurrencyDesc = New System.Windows.Forms.TextBox
        Me.btnFindCurrency = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.txtArAccountDesc = New System.Windows.Forms.TextBox
        Me.txtArAccountCode = New System.Windows.Forms.TextBox
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.btnFindArAccount = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtPaymentAccountDesc = New System.Windows.Forms.TextBox
        Me.btnFindPaymentAccount = New System.Windows.Forms.Button
        Me.txtPaymentAccountCode = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtSupplierDesc = New System.Windows.Forms.TextBox
        Me.btnFindSupplier = New System.Windows.Forms.Button
        Me.txtSupplierCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkSupplierInvoice = New System.Windows.Forms.CheckBox
        Me.crvAP = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From"
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(94, 90)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimaryCurrency.TabIndex = 24
        Me.txtPrimaryCurrency.TabStop = False
        Me.txtPrimaryCurrency.Text = "-1"
        Me.txtPrimaryCurrency.Visible = False
        '
        'txtPrimaryPaymentAccount
        '
        Me.txtPrimaryPaymentAccount.Location = New System.Drawing.Point(94, 41)
        Me.txtPrimaryPaymentAccount.Name = "txtPrimaryPaymentAccount"
        Me.txtPrimaryPaymentAccount.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimaryPaymentAccount.TabIndex = 22
        Me.txtPrimaryPaymentAccount.TabStop = False
        Me.txtPrimaryPaymentAccount.Text = "-1"
        Me.txtPrimaryPaymentAccount.Visible = False
        '
        'rboArAccount
        '
        Me.rboArAccount.AutoSize = True
        Me.rboArAccount.Location = New System.Drawing.Point(4, 66)
        Me.rboArAccount.Name = "rboArAccount"
        Me.rboArAccount.Size = New System.Drawing.Size(83, 17)
        Me.rboArAccount.TabIndex = 2
        Me.rboArAccount.TabStop = True
        Me.rboArAccount.Text = "AR Account"
        Me.rboArAccount.UseVisualStyleBackColor = True
        '
        'txtPrimaryArAccount
        '
        Me.txtPrimaryArAccount.Location = New System.Drawing.Point(94, 66)
        Me.txtPrimaryArAccount.Name = "txtPrimaryArAccount"
        Me.txtPrimaryArAccount.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimaryArAccount.TabIndex = 23
        Me.txtPrimaryArAccount.TabStop = False
        Me.txtPrimaryArAccount.Text = "-1"
        Me.txtPrimaryArAccount.Visible = False
        '
        'rboCurrency
        '
        Me.rboCurrency.AutoSize = True
        Me.rboCurrency.Location = New System.Drawing.Point(121, 19)
        Me.rboCurrency.Name = "rboCurrency"
        Me.rboCurrency.Size = New System.Drawing.Size(67, 17)
        Me.rboCurrency.TabIndex = 3
        Me.rboCurrency.TabStop = True
        Me.rboCurrency.Text = "Currency"
        Me.rboCurrency.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rboDetail)
        Me.GroupBox2.Controls.Add(Me.rboPaymentJournal)
        Me.GroupBox2.Controls.Add(Me.rboPaymentAccount)
        Me.GroupBox2.Controls.Add(Me.rboCurrency)
        Me.GroupBox2.Controls.Add(Me.rboArAccount)
        Me.GroupBox2.Controls.Add(Me.rboSuppCode)
        Me.GroupBox2.Location = New System.Drawing.Point(0, 87)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(249, 94)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Group By"
        '
        'rboDetail
        '
        Me.rboDetail.AutoSize = True
        Me.rboDetail.Location = New System.Drawing.Point(121, 43)
        Me.rboDetail.Name = "rboDetail"
        Me.rboDetail.Size = New System.Drawing.Size(121, 17)
        Me.rboDetail.TabIndex = 4
        Me.rboDetail.Text = "Supplier Code Detail"
        Me.rboDetail.UseVisualStyleBackColor = True
        '
        'rboPaymentJournal
        '
        Me.rboPaymentJournal.AutoSize = True
        Me.rboPaymentJournal.Location = New System.Drawing.Point(121, 66)
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
        Me.rboPaymentAccount.Location = New System.Drawing.Point(4, 42)
        Me.rboPaymentAccount.Name = "rboPaymentAccount"
        Me.rboPaymentAccount.Size = New System.Drawing.Size(109, 17)
        Me.rboPaymentAccount.TabIndex = 1
        Me.rboPaymentAccount.TabStop = True
        Me.rboPaymentAccount.Text = "Payment Account"
        Me.rboPaymentAccount.UseVisualStyleBackColor = True
        '
        'rboSuppCode
        '
        Me.rboSuppCode.AutoSize = True
        Me.rboSuppCode.Checked = True
        Me.rboSuppCode.Location = New System.Drawing.Point(4, 19)
        Me.rboSuppCode.Name = "rboSuppCode"
        Me.rboSuppCode.Size = New System.Drawing.Size(91, 17)
        Me.rboSuppCode.TabIndex = 0
        Me.rboSuppCode.TabStop = True
        Me.rboSuppCode.Text = "Supplier Code"
        Me.rboSuppCode.UseVisualStyleBackColor = True
        '
        'txtPrimarySupplier
        '
        Me.txtPrimarySupplier.Location = New System.Drawing.Point(94, 16)
        Me.txtPrimarySupplier.Name = "txtPrimarySupplier"
        Me.txtPrimarySupplier.Size = New System.Drawing.Size(28, 20)
        Me.txtPrimarySupplier.TabIndex = 20
        Me.txtPrimarySupplier.TabStop = False
        Me.txtPrimarySupplier.Text = "-1"
        Me.txtPrimarySupplier.Visible = False
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(42, 36)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(207, 20)
        Me.dtpTo.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(42, 10)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(207, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(825, 14)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(75, 23)
        Me.btnShow.TabIndex = 4
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'txtCurrencyDesc
        '
        Me.txtCurrencyDesc.Location = New System.Drawing.Point(315, 90)
        Me.txtCurrencyDesc.Name = "txtCurrencyDesc"
        Me.txtCurrencyDesc.ReadOnly = True
        Me.txtCurrencyDesc.Size = New System.Drawing.Size(246, 20)
        Me.txtCurrencyDesc.TabIndex = 15
        Me.txtCurrencyDesc.TabStop = False
        '
        'btnFindCurrency
        '
        Me.btnFindCurrency.Location = New System.Drawing.Point(250, 88)
        Me.btnFindCurrency.Name = "btnFindCurrency"
        Me.btnFindCurrency.Size = New System.Drawing.Size(59, 23)
        Me.btnFindCurrency.TabIndex = 9
        Me.btnFindCurrency.Text = "..."
        Me.btnFindCurrency.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(15, 93)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Currency"
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(825, 43)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'txtArAccountDesc
        '
        Me.txtArAccountDesc.Location = New System.Drawing.Point(315, 66)
        Me.txtArAccountDesc.Name = "txtArAccountDesc"
        Me.txtArAccountDesc.ReadOnly = True
        Me.txtArAccountDesc.Size = New System.Drawing.Size(246, 20)
        Me.txtArAccountDesc.TabIndex = 11
        Me.txtArAccountDesc.TabStop = False
        '
        'txtArAccountCode
        '
        Me.txtArAccountCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtArAccountCode.Location = New System.Drawing.Point(124, 66)
        Me.txtArAccountCode.Name = "txtArAccountCode"
        Me.txtArAccountCode.Size = New System.Drawing.Size(120, 20)
        Me.txtArAccountCode.TabIndex = 6
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCurrencyCode.Location = New System.Drawing.Point(124, 90)
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.Size = New System.Drawing.Size(120, 20)
        Me.txtCurrencyCode.TabIndex = 8
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtPrimaryCurrency)
        Me.GroupBox3.Controls.Add(Me.txtPrimaryArAccount)
        Me.GroupBox3.Controls.Add(Me.txtPrimaryPaymentAccount)
        Me.GroupBox3.Controls.Add(Me.txtPrimarySupplier)
        Me.GroupBox3.Controls.Add(Me.txtCurrencyDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindCurrency)
        Me.GroupBox3.Controls.Add(Me.txtCurrencyCode)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.txtArAccountDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindArAccount)
        Me.GroupBox3.Controls.Add(Me.txtArAccountCode)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.txtPaymentAccountDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindPaymentAccount)
        Me.GroupBox3.Controls.Add(Me.txtPaymentAccountCode)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.txtSupplierDesc)
        Me.GroupBox3.Controls.Add(Me.btnFindSupplier)
        Me.GroupBox3.Controls.Add(Me.txtSupplierCode)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Location = New System.Drawing.Point(254, 10)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(566, 141)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Filter"
        '
        'btnFindArAccount
        '
        Me.btnFindArAccount.Location = New System.Drawing.Point(250, 64)
        Me.btnFindArAccount.Name = "btnFindArAccount"
        Me.btnFindArAccount.Size = New System.Drawing.Size(59, 23)
        Me.btnFindArAccount.TabIndex = 7
        Me.btnFindArAccount.Text = "..."
        Me.btnFindArAccount.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(15, 69)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "AR Account"
        '
        'txtPaymentAccountDesc
        '
        Me.txtPaymentAccountDesc.Location = New System.Drawing.Point(315, 41)
        Me.txtPaymentAccountDesc.Name = "txtPaymentAccountDesc"
        Me.txtPaymentAccountDesc.ReadOnly = True
        Me.txtPaymentAccountDesc.Size = New System.Drawing.Size(246, 20)
        Me.txtPaymentAccountDesc.TabIndex = 7
        Me.txtPaymentAccountDesc.TabStop = False
        '
        'btnFindPaymentAccount
        '
        Me.btnFindPaymentAccount.Location = New System.Drawing.Point(250, 39)
        Me.btnFindPaymentAccount.Name = "btnFindPaymentAccount"
        Me.btnFindPaymentAccount.Size = New System.Drawing.Size(59, 23)
        Me.btnFindPaymentAccount.TabIndex = 5
        Me.btnFindPaymentAccount.Text = "..."
        Me.btnFindPaymentAccount.UseVisualStyleBackColor = True
        '
        'txtPaymentAccountCode
        '
        Me.txtPaymentAccountCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPaymentAccountCode.Location = New System.Drawing.Point(124, 41)
        Me.txtPaymentAccountCode.Name = "txtPaymentAccountCode"
        Me.txtPaymentAccountCode.Size = New System.Drawing.Size(120, 20)
        Me.txtPaymentAccountCode.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(15, 44)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(91, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Payment Account"
        '
        'txtSupplierDesc
        '
        Me.txtSupplierDesc.Location = New System.Drawing.Point(315, 16)
        Me.txtSupplierDesc.Name = "txtSupplierDesc"
        Me.txtSupplierDesc.ReadOnly = True
        Me.txtSupplierDesc.Size = New System.Drawing.Size(246, 20)
        Me.txtSupplierDesc.TabIndex = 3
        Me.txtSupplierDesc.TabStop = False
        '
        'btnFindSupplier
        '
        Me.btnFindSupplier.Location = New System.Drawing.Point(250, 14)
        Me.btnFindSupplier.Name = "btnFindSupplier"
        Me.btnFindSupplier.Size = New System.Drawing.Size(59, 23)
        Me.btnFindSupplier.TabIndex = 1
        Me.btnFindSupplier.Text = "..."
        Me.btnFindSupplier.UseVisualStyleBackColor = True
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierCode.Location = New System.Drawing.Point(124, 16)
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.Size = New System.Drawing.Size(120, 20)
        Me.txtSupplierCode.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(15, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Supplier Code"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.chkSupplierInvoice)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.btnShow)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.dtpTo)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.dtpFrom)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, -1)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(933, 187)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'chkSupplierInvoice
        '
        Me.chkSupplierInvoice.AutoSize = True
        Me.chkSupplierInvoice.Location = New System.Drawing.Point(42, 63)
        Me.chkSupplierInvoice.Name = "chkSupplierInvoice"
        Me.chkSupplierInvoice.Size = New System.Drawing.Size(132, 17)
        Me.chkSupplierInvoice.TabIndex = 6
        Me.chkSupplierInvoice.Text = "Show Supplier Invoice"
        Me.chkSupplierInvoice.UseVisualStyleBackColor = True
        '
        'crvAP
        '
        Me.crvAP.ActiveViewIndex = -1
        Me.crvAP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvAP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvAP.DisplayGroupTree = False
        Me.crvAP.Location = New System.Drawing.Point(12, 197)
        Me.crvAP.Name = "crvAP"
        Me.crvAP.SelectionFormula = ""
        Me.crvAP.ShowCloseButton = False
        Me.crvAP.ShowGroupTreeButton = False
        Me.crvAP.ShowRefreshButton = False
        Me.crvAP.Size = New System.Drawing.Size(933, 383)
        Me.crvAP.TabIndex = 5
        Me.crvAP.TabStop = False
        Me.crvAP.ViewTimeSelectionFormula = ""
        '
        'FrmCetakAP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(957, 592)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.crvAP)
        Me.MinimumSize = New System.Drawing.Size(973, 630)
        Me.Name = "FrmCetakAP"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print AP"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryPaymentAccount As System.Windows.Forms.TextBox
    Friend WithEvents rboArAccount As System.Windows.Forms.RadioButton
    Friend WithEvents txtPrimaryArAccount As System.Windows.Forms.TextBox
    Friend WithEvents rboCurrency As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rboPaymentAccount As System.Windows.Forms.RadioButton
    Friend WithEvents rboSuppCode As System.Windows.Forms.RadioButton
    Friend WithEvents txtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents txtCurrencyDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCurrency As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents txtArAccountDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtArAccountCode As System.Windows.Forms.TextBox
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnFindArAccount As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPaymentAccountDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindPaymentAccount As System.Windows.Forms.Button
    Friend WithEvents txtPaymentAccountCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSupplierDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindSupplier As System.Windows.Forms.Button
    Friend WithEvents txtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents crvAP As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents chkSupplierInvoice As System.Windows.Forms.CheckBox
    Friend WithEvents rboPaymentJournal As System.Windows.Forms.RadioButton
    Friend WithEvents rboDetail As System.Windows.Forms.RadioButton
End Class
