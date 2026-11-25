<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmSettingAP_AR
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
        Me.btnFindAR = New System.Windows.Forms.Button
        Me.TxtAkunAR = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAkunAR = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAkunAP = New System.Windows.Forms.TextBox
        Me.TxtAkunAP = New System.Windows.Forms.TextBox
        Me.btnFindAkunAP = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.TxtCustomeDOAddress = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.btnFindCustomDOAddress = New System.Windows.Forms.Button
        Me.chkUseCustomDO = New System.Windows.Forms.CheckBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.TxtCustomTaxInvoiceAddress = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnFindCustomTaxInvoice = New System.Windows.Forms.Button
        Me.chkUseCustomTaxInvoice = New System.Windows.Forms.CheckBox
        Me.TxtCustomInvoiceAddress = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnFindCustomInvoiceReport = New System.Windows.Forms.Button
        Me.chkUseCustomInvoice = New System.Windows.Forms.CheckBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtTaxInvoice = New System.Windows.Forms.TextBox
        Me.chkTaxInvoice = New System.Windows.Forms.CheckBox
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtTaxAOP = New System.Windows.Forms.TextBox
        Me.ChkTaxAOP = New System.Windows.Forms.CheckBox
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.TxtSignedByName = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtSignedByPosition = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TabControl1.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Account Receivable"
        '
        'btnFindAR
        '
        Me.btnFindAR.Location = New System.Drawing.Point(381, 4)
        Me.btnFindAR.Name = "btnFindAR"
        Me.btnFindAR.Size = New System.Drawing.Size(54, 23)
        Me.btnFindAR.TabIndex = 2
        Me.btnFindAR.TabStop = False
        Me.btnFindAR.Text = "Find"
        Me.btnFindAR.UseVisualStyleBackColor = True
        '
        'TxtAkunAR
        '
        Me.TxtAkunAR.Location = New System.Drawing.Point(117, 6)
        Me.TxtAkunAR.Name = "TxtAkunAR"
        Me.TxtAkunAR.Size = New System.Drawing.Size(258, 20)
        Me.TxtAkunAR.TabIndex = 0
        '
        'TxtPrimaryAkunAR
        '
        Me.TxtPrimaryAkunAR.Location = New System.Drawing.Point(337, 6)
        Me.TxtPrimaryAkunAR.Name = "TxtPrimaryAkunAR"
        Me.TxtPrimaryAkunAR.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryAkunAR.TabIndex = 3
        Me.TxtPrimaryAkunAR.TabStop = False
        Me.TxtPrimaryAkunAR.Text = "-1"
        Me.TxtPrimaryAkunAR.Visible = False
        '
        'TxtPrimaryAkunAP
        '
        Me.TxtPrimaryAkunAP.Location = New System.Drawing.Point(337, 30)
        Me.TxtPrimaryAkunAP.Name = "TxtPrimaryAkunAP"
        Me.TxtPrimaryAkunAP.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryAkunAP.TabIndex = 7
        Me.TxtPrimaryAkunAP.TabStop = False
        Me.TxtPrimaryAkunAP.Text = "-1"
        Me.TxtPrimaryAkunAP.Visible = False
        '
        'TxtAkunAP
        '
        Me.TxtAkunAP.Location = New System.Drawing.Point(117, 32)
        Me.TxtAkunAP.Name = "TxtAkunAP"
        Me.TxtAkunAP.Size = New System.Drawing.Size(258, 20)
        Me.TxtAkunAP.TabIndex = 1
        '
        'btnFindAkunAP
        '
        Me.btnFindAkunAP.Location = New System.Drawing.Point(381, 30)
        Me.btnFindAkunAP.Name = "btnFindAkunAP"
        Me.btnFindAkunAP.Size = New System.Drawing.Size(54, 23)
        Me.btnFindAkunAP.TabIndex = 3
        Me.btnFindAkunAP.TabStop = False
        Me.btnFindAkunAP.Text = "Find"
        Me.btnFindAkunAP.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Account Payable"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(386, 359)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(305, 358)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(449, 341)
        Me.TabControl1.TabIndex = 8
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.TxtCustomeDOAddress)
        Me.TabPage4.Controls.Add(Me.Label7)
        Me.TabPage4.Controls.Add(Me.btnFindCustomDOAddress)
        Me.TabPage4.Controls.Add(Me.chkUseCustomDO)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(441, 315)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Delivery Order Setting"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'TxtCustomeDOAddress
        '
        Me.TxtCustomeDOAddress.Location = New System.Drawing.Point(93, 29)
        Me.TxtCustomeDOAddress.Name = "TxtCustomeDOAddress"
        Me.TxtCustomeDOAddress.Size = New System.Drawing.Size(261, 20)
        Me.TxtCustomeDOAddress.TabIndex = 11
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 13)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Report Address"
        '
        'btnFindCustomDOAddress
        '
        Me.btnFindCustomDOAddress.Location = New System.Drawing.Point(360, 29)
        Me.btnFindCustomDOAddress.Name = "btnFindCustomDOAddress"
        Me.btnFindCustomDOAddress.Size = New System.Drawing.Size(75, 20)
        Me.btnFindCustomDOAddress.TabIndex = 12
        Me.btnFindCustomDOAddress.Text = "Find"
        Me.btnFindCustomDOAddress.UseVisualStyleBackColor = True
        '
        'chkUseCustomDO
        '
        Me.chkUseCustomDO.AutoSize = True
        Me.chkUseCustomDO.Location = New System.Drawing.Point(6, 6)
        Me.chkUseCustomDO.Name = "chkUseCustomDO"
        Me.chkUseCustomDO.Size = New System.Drawing.Size(153, 17)
        Me.chkUseCustomDO.TabIndex = 10
        Me.chkUseCustomDO.Text = "Use Custom Delivery Order"
        Me.chkUseCustomDO.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.TxtSignedByPosition)
        Me.TabPage2.Controls.Add(Me.Label9)
        Me.TabPage2.Controls.Add(Me.TxtSignedByName)
        Me.TabPage2.Controls.Add(Me.Label8)
        Me.TabPage2.Controls.Add(Me.TxtCustomTaxInvoiceAddress)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.btnFindCustomTaxInvoice)
        Me.TabPage2.Controls.Add(Me.chkUseCustomTaxInvoice)
        Me.TabPage2.Controls.Add(Me.TxtCustomInvoiceAddress)
        Me.TabPage2.Controls.Add(Me.Label5)
        Me.TabPage2.Controls.Add(Me.btnFindCustomInvoiceReport)
        Me.TabPage2.Controls.Add(Me.chkUseCustomInvoice)
        Me.TabPage2.Controls.Add(Me.Label3)
        Me.TabPage2.Controls.Add(Me.TxtTaxInvoice)
        Me.TabPage2.Controls.Add(Me.chkTaxInvoice)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(441, 315)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Invoice Setting"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TxtCustomTaxInvoiceAddress
        '
        Me.TxtCustomTaxInvoiceAddress.Location = New System.Drawing.Point(107, 103)
        Me.TxtCustomTaxInvoiceAddress.Name = "TxtCustomTaxInvoiceAddress"
        Me.TxtCustomTaxInvoiceAddress.Size = New System.Drawing.Size(247, 20)
        Me.TxtCustomTaxInvoiceAddress.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 106)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 13)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "Report Address"
        '
        'btnFindCustomTaxInvoice
        '
        Me.btnFindCustomTaxInvoice.Location = New System.Drawing.Point(360, 103)
        Me.btnFindCustomTaxInvoice.Name = "btnFindCustomTaxInvoice"
        Me.btnFindCustomTaxInvoice.Size = New System.Drawing.Size(75, 20)
        Me.btnFindCustomTaxInvoice.TabIndex = 7
        Me.btnFindCustomTaxInvoice.Text = "Find"
        Me.btnFindCustomTaxInvoice.UseVisualStyleBackColor = True
        '
        'chkUseCustomTaxInvoice
        '
        Me.chkUseCustomTaxInvoice.AutoSize = True
        Me.chkUseCustomTaxInvoice.Location = New System.Drawing.Point(6, 80)
        Me.chkUseCustomTaxInvoice.Name = "chkUseCustomTaxInvoice"
        Me.chkUseCustomTaxInvoice.Size = New System.Drawing.Size(142, 17)
        Me.chkUseCustomTaxInvoice.TabIndex = 5
        Me.chkUseCustomTaxInvoice.Text = "Use Custom Tax Invoice"
        Me.chkUseCustomTaxInvoice.UseVisualStyleBackColor = True
        '
        'TxtCustomInvoiceAddress
        '
        Me.TxtCustomInvoiceAddress.Location = New System.Drawing.Point(107, 54)
        Me.TxtCustomInvoiceAddress.Name = "TxtCustomInvoiceAddress"
        Me.TxtCustomInvoiceAddress.Size = New System.Drawing.Size(247, 20)
        Me.TxtCustomInvoiceAddress.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 57)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 13)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Report Address"
        '
        'btnFindCustomInvoiceReport
        '
        Me.btnFindCustomInvoiceReport.Location = New System.Drawing.Point(360, 54)
        Me.btnFindCustomInvoiceReport.Name = "btnFindCustomInvoiceReport"
        Me.btnFindCustomInvoiceReport.Size = New System.Drawing.Size(75, 20)
        Me.btnFindCustomInvoiceReport.TabIndex = 4
        Me.btnFindCustomInvoiceReport.Text = "Find"
        Me.btnFindCustomInvoiceReport.UseVisualStyleBackColor = True
        '
        'chkUseCustomInvoice
        '
        Me.chkUseCustomInvoice.AutoSize = True
        Me.chkUseCustomInvoice.Location = New System.Drawing.Point(6, 31)
        Me.chkUseCustomInvoice.Name = "chkUseCustomInvoice"
        Me.chkUseCustomInvoice.Size = New System.Drawing.Size(121, 17)
        Me.chkUseCustomInvoice.TabIndex = 2
        Me.chkUseCustomInvoice.Text = "Use Custom Invoice"
        Me.chkUseCustomInvoice.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(181, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(15, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "%"
        '
        'TxtTaxInvoice
        '
        Me.TxtTaxInvoice.Location = New System.Drawing.Point(107, 6)
        Me.TxtTaxInvoice.Name = "TxtTaxInvoice"
        Me.TxtTaxInvoice.ReadOnly = True
        Me.TxtTaxInvoice.Size = New System.Drawing.Size(68, 20)
        Me.TxtTaxInvoice.TabIndex = 1
        Me.TxtTaxInvoice.TabStop = False
        Me.TxtTaxInvoice.Text = "0"
        Me.TxtTaxInvoice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'chkTaxInvoice
        '
        Me.chkTaxInvoice.AutoSize = True
        Me.chkTaxInvoice.Location = New System.Drawing.Point(6, 8)
        Me.chkTaxInvoice.Name = "chkTaxInvoice"
        Me.chkTaxInvoice.Size = New System.Drawing.Size(66, 17)
        Me.chkTaxInvoice.TabIndex = 0
        Me.chkTaxInvoice.Text = "Use Tax"
        Me.chkTaxInvoice.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.Label4)
        Me.TabPage3.Controls.Add(Me.TxtTaxAOP)
        Me.TabPage3.Controls.Add(Me.ChkTaxAOP)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(441, 315)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "AOP Setting"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(181, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(15, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "%"
        '
        'TxtTaxAOP
        '
        Me.TxtTaxAOP.Location = New System.Drawing.Point(93, 6)
        Me.TxtTaxAOP.Name = "TxtTaxAOP"
        Me.TxtTaxAOP.ReadOnly = True
        Me.TxtTaxAOP.Size = New System.Drawing.Size(82, 20)
        Me.TxtTaxAOP.TabIndex = 4
        Me.TxtTaxAOP.TabStop = False
        Me.TxtTaxAOP.Text = "0"
        Me.TxtTaxAOP.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ChkTaxAOP
        '
        Me.ChkTaxAOP.AutoSize = True
        Me.ChkTaxAOP.Location = New System.Drawing.Point(6, 8)
        Me.ChkTaxAOP.Name = "ChkTaxAOP"
        Me.ChkTaxAOP.Size = New System.Drawing.Size(66, 17)
        Me.ChkTaxAOP.TabIndex = 3
        Me.ChkTaxAOP.Text = "Use Tax"
        Me.ChkTaxAOP.UseVisualStyleBackColor = True
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.TxtAkunAR)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.TxtPrimaryAkunAP)
        Me.TabPage1.Controls.Add(Me.btnFindAR)
        Me.TabPage1.Controls.Add(Me.TxtAkunAP)
        Me.TabPage1.Controls.Add(Me.TxtPrimaryAkunAR)
        Me.TabPage1.Controls.Add(Me.btnFindAkunAP)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(441, 315)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Journal Setting"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.Filter = "Report file|*.rpt|All file|*.*"
        '
        'TxtSignedByName
        '
        Me.TxtSignedByName.Location = New System.Drawing.Point(107, 129)
        Me.TxtSignedByName.Name = "TxtSignedByName"
        Me.TxtSignedByName.Size = New System.Drawing.Size(328, 20)
        Me.TxtSignedByName.TabIndex = 10
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 132)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(86, 13)
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "Signed By Name"
        '
        'TxtSignedByPosition
        '
        Me.TxtSignedByPosition.Location = New System.Drawing.Point(107, 155)
        Me.TxtSignedByPosition.Name = "TxtSignedByPosition"
        Me.TxtSignedByPosition.Size = New System.Drawing.Size(328, 20)
        Me.TxtSignedByPosition.TabIndex = 12
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 158)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(95, 13)
        Me.Label9.TabIndex = 13
        Me.Label9.Text = "Signed By Position"
        '
        'FrmSettingAP_AR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(473, 393)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "FrmSettingAP_AR"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Journal Setting"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnFindAR As System.Windows.Forms.Button
    Friend WithEvents TxtAkunAR As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAkunAR As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAkunAP As System.Windows.Forms.TextBox
    Friend WithEvents TxtAkunAP As System.Windows.Forms.TextBox
    Friend WithEvents btnFindAkunAP As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtTaxInvoice As System.Windows.Forms.TextBox
    Friend WithEvents chkTaxInvoice As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtTaxAOP As System.Windows.Forms.TextBox
    Friend WithEvents ChkTaxAOP As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseCustomInvoice As System.Windows.Forms.CheckBox
    Friend WithEvents TxtCustomTaxInvoiceAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnFindCustomTaxInvoice As System.Windows.Forms.Button
    Friend WithEvents chkUseCustomTaxInvoice As System.Windows.Forms.CheckBox
    Friend WithEvents TxtCustomInvoiceAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnFindCustomInvoiceReport As System.Windows.Forms.Button
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents TxtCustomeDOAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnFindCustomDOAddress As System.Windows.Forms.Button
    Friend WithEvents chkUseCustomDO As System.Windows.Forms.CheckBox
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TxtSignedByPosition As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtSignedByName As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label

End Class
