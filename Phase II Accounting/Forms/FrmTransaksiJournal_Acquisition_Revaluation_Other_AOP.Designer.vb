<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiJournal_Acquisition_Revaluation_Other_AOP
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
        Me.btnFindSupplier = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.txtSupplierCode = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.cboCodeNo = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtPrimaryPO = New System.Windows.Forms.TextBox
        Me.cboPurchaseOrder = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtPrimaryAOP = New System.Windows.Forms.TextBox
        Me.CboAOPNumber = New System.Windows.Forms.ComboBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.TxtQtyperBatch = New System.Windows.Forms.TextBox
        Me.lblUnits = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtAtPrice = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtTotal = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtExpenseAccount = New System.Windows.Forms.TextBox
        Me.TxtPrimaryExpenseAccount = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnFindSupplier
        '
        Me.btnFindSupplier.Location = New System.Drawing.Point(324, 19)
        Me.btnFindSupplier.Name = "btnFindSupplier"
        Me.btnFindSupplier.Size = New System.Drawing.Size(32, 20)
        Me.btnFindSupplier.TabIndex = 135
        Me.btnFindSupplier.TabStop = False
        Me.btnFindSupplier.Text = "...."
        Me.btnFindSupplier.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 136
        Me.Label1.Text = "Supplier Code"
        '
        'txtPrimarySupplier
        '
        Me.txtPrimarySupplier.Location = New System.Drawing.Point(96, 20)
        Me.txtPrimarySupplier.Name = "txtPrimarySupplier"
        Me.txtPrimarySupplier.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimarySupplier.TabIndex = 137
        Me.txtPrimarySupplier.TabStop = False
        Me.txtPrimarySupplier.Text = "-1"
        Me.txtPrimarySupplier.Visible = False
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierCode.Location = New System.Drawing.Point(127, 18)
        Me.txtSupplierCode.MaxLength = 50
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.Size = New System.Drawing.Size(191, 20)
        Me.txtSupplierCode.TabIndex = 134
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryExpenseAccount)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.TxtExpenseAccount)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.TxtTotal)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TxtAtPrice)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.TxtQtyperBatch)
        Me.GroupBox1.Controls.Add(Me.lblUnits)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryAOP)
        Me.GroupBox1.Controls.Add(Me.CboAOPNumber)
        Me.GroupBox1.Controls.Add(Me.txtPrimaryCodeNo)
        Me.GroupBox1.Controls.Add(Me.cboCodeNo)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.txtPrimaryPO)
        Me.GroupBox1.Controls.Add(Me.cboPurchaseOrder)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtPrimarySupplier)
        Me.GroupBox1.Controls.Add(Me.btnFindSupplier)
        Me.GroupBox1.Controls.Add(Me.txtSupplierCode)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(370, 240)
        Me.GroupBox1.TabIndex = 138
        Me.GroupBox1.TabStop = False
        '
        'txtPrimaryCodeNo
        '
        Me.txtPrimaryCodeNo.Location = New System.Drawing.Point(151, 74)
        Me.txtPrimaryCodeNo.Name = "txtPrimaryCodeNo"
        Me.txtPrimaryCodeNo.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryCodeNo.TabIndex = 194
        Me.txtPrimaryCodeNo.TabStop = False
        Me.txtPrimaryCodeNo.Text = "-1"
        Me.txtPrimaryCodeNo.Visible = False
        '
        'cboCodeNo
        '
        Me.cboCodeNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCodeNo.FormattingEnabled = True
        Me.cboCodeNo.Location = New System.Drawing.Point(127, 72)
        Me.cboCodeNo.Name = "cboCodeNo"
        Me.cboCodeNo.Size = New System.Drawing.Size(229, 21)
        Me.cboCodeNo.TabIndex = 190
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(17, 80)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(49, 13)
        Me.Label10.TabIndex = 191
        Me.Label10.Text = "Code No"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(17, 103)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(98, 13)
        Me.Label19.TabIndex = 192
        Me.Label19.Text = "Purchase Order No"
        '
        'txtPrimaryPO
        '
        Me.txtPrimaryPO.Location = New System.Drawing.Point(150, 99)
        Me.txtPrimaryPO.Name = "txtPrimaryPO"
        Me.txtPrimaryPO.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryPO.TabIndex = 193
        Me.txtPrimaryPO.TabStop = False
        Me.txtPrimaryPO.Text = "-1"
        Me.txtPrimaryPO.Visible = False
        '
        'cboPurchaseOrder
        '
        Me.cboPurchaseOrder.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPurchaseOrder.FormattingEnabled = True
        Me.cboPurchaseOrder.Location = New System.Drawing.Point(127, 100)
        Me.cboPurchaseOrder.Name = "cboPurchaseOrder"
        Me.cboPurchaseOrder.Size = New System.Drawing.Size(229, 21)
        Me.cboPurchaseOrder.TabIndex = 189
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(17, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 13)
        Me.Label2.TabIndex = 196
        Me.Label2.Text = "AOP Number"
        '
        'TxtPrimaryAOP
        '
        Me.TxtPrimaryAOP.Location = New System.Drawing.Point(150, 45)
        Me.TxtPrimaryAOP.Name = "TxtPrimaryAOP"
        Me.TxtPrimaryAOP.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryAOP.TabIndex = 197
        Me.TxtPrimaryAOP.TabStop = False
        Me.TxtPrimaryAOP.Text = "-1"
        Me.TxtPrimaryAOP.Visible = False
        '
        'CboAOPNumber
        '
        Me.CboAOPNumber.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboAOPNumber.FormattingEnabled = True
        Me.CboAOPNumber.Location = New System.Drawing.Point(127, 45)
        Me.CboAOPNumber.Name = "CboAOPNumber"
        Me.CboAOPNumber.Size = New System.Drawing.Size(229, 21)
        Me.CboAOPNumber.TabIndex = 195
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(17, 127)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(23, 13)
        Me.Label13.TabIndex = 200
        Me.Label13.Text = "Qty"
        '
        'TxtQtyperBatch
        '
        Me.TxtQtyperBatch.Location = New System.Drawing.Point(127, 127)
        Me.TxtQtyperBatch.MaxLength = 50
        Me.TxtQtyperBatch.Name = "TxtQtyperBatch"
        Me.TxtQtyperBatch.Size = New System.Drawing.Size(192, 20)
        Me.TxtQtyperBatch.TabIndex = 198
        Me.TxtQtyperBatch.Text = "0"
        Me.TxtQtyperBatch.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblUnits
        '
        Me.lblUnits.AutoSize = True
        Me.lblUnits.Location = New System.Drawing.Point(326, 130)
        Me.lblUnits.Name = "lblUnits"
        Me.lblUnits.Size = New System.Drawing.Size(31, 13)
        Me.lblUnits.TabIndex = 199
        Me.lblUnits.Text = "Units"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(17, 153)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(44, 13)
        Me.Label3.TabIndex = 202
        Me.Label3.Text = "At Price"
        '
        'TxtAtPrice
        '
        Me.TxtAtPrice.Location = New System.Drawing.Point(127, 153)
        Me.TxtAtPrice.MaxLength = 50
        Me.TxtAtPrice.Name = "TxtAtPrice"
        Me.TxtAtPrice.ReadOnly = True
        Me.TxtAtPrice.Size = New System.Drawing.Size(230, 20)
        Me.TxtAtPrice.TabIndex = 201
        Me.TxtAtPrice.Text = "0"
        Me.TxtAtPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(17, 179)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(31, 13)
        Me.Label4.TabIndex = 204
        Me.Label4.Text = "Total"
        '
        'TxtTotal
        '
        Me.TxtTotal.Location = New System.Drawing.Point(127, 179)
        Me.TxtTotal.MaxLength = 50
        Me.TxtTotal.Name = "TxtTotal"
        Me.TxtTotal.ReadOnly = True
        Me.TxtTotal.Size = New System.Drawing.Size(230, 20)
        Me.TxtTotal.TabIndex = 203
        Me.TxtTotal.Text = "0"
        Me.TxtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(16, 205)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(91, 13)
        Me.Label5.TabIndex = 206
        Me.Label5.Text = "Expense Account"
        '
        'TxtExpenseAccount
        '
        Me.TxtExpenseAccount.Location = New System.Drawing.Point(126, 205)
        Me.TxtExpenseAccount.MaxLength = 50
        Me.TxtExpenseAccount.Name = "TxtExpenseAccount"
        Me.TxtExpenseAccount.ReadOnly = True
        Me.TxtExpenseAccount.Size = New System.Drawing.Size(230, 20)
        Me.TxtExpenseAccount.TabIndex = 205
        '
        'TxtPrimaryExpenseAccount
        '
        Me.TxtPrimaryExpenseAccount.Location = New System.Drawing.Point(126, 205)
        Me.TxtPrimaryExpenseAccount.Name = "TxtPrimaryExpenseAccount"
        Me.TxtPrimaryExpenseAccount.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryExpenseAccount.TabIndex = 207
        Me.TxtPrimaryExpenseAccount.TabStop = False
        Me.TxtPrimaryExpenseAccount.Text = "-1"
        Me.TxtPrimaryExpenseAccount.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(186, 260)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 140
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(70, 260)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 139
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmTransaksiJournal_Acquisition_Revaluation_Other_AOP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(388, 298)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiJournal_Acquisition_Revaluation_Other_AOP"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Input_AOP"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnFindSupplier As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents txtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryAOP As System.Windows.Forms.TextBox
    Friend WithEvents CboAOPNumber As System.Windows.Forms.ComboBox
    Friend WithEvents txtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents cboCodeNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryPO As System.Windows.Forms.TextBox
    Friend WithEvents cboPurchaseOrder As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtAtPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TxtQtyperBatch As System.Windows.Forms.TextBox
    Friend WithEvents lblUnits As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtExpenseAccount As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryExpenseAccount As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
End Class
