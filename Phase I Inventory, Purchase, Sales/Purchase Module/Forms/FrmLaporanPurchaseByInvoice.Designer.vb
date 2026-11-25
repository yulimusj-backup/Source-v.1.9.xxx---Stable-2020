<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPurchaseByInvoice
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
        Me.BtnPrint = New System.Windows.Forms.Button
        Me.ChkUsingSupplierNumber = New System.Windows.Forms.CheckBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnClose = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.ChkShowRatio = New System.Windows.Forms.CheckBox
        Me.TxtAccountDesc = New System.Windows.Forms.TextBox
        Me.TxtSuppliserDesc = New System.Windows.Forms.TextBox
        Me.ChkAllProductPO = New System.Windows.Forms.CheckBox
        Me.BtnFindSupp = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtPrimaryAcc = New System.Windows.Forms.TextBox
        Me.TxtKodeSupplier = New System.Windows.Forms.TextBox
        Me.TxtAccCode = New System.Windows.Forms.TextBox
        Me.TxtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.BtnFindAcc = New System.Windows.Forms.Button
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CboFormat = New System.Windows.Forms.ComboBox
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtCodeNoDesc = New System.Windows.Forms.TextBox
        Me.BtnFindCodeNo = New System.Windows.Forms.Button
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtCodeNo = New System.Windows.Forms.TextBox
        Me.txtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.Crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BtnPrint
        '
        Me.BtnPrint.Location = New System.Drawing.Point(120, 116)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(53, 23)
        Me.BtnPrint.TabIndex = 10
        Me.BtnPrint.Text = "Print"
        Me.BtnPrint.UseVisualStyleBackColor = True
        '
        'ChkUsingSupplierNumber
        '
        Me.ChkUsingSupplierNumber.AutoSize = True
        Me.ChkUsingSupplierNumber.Checked = True
        Me.ChkUsingSupplierNumber.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkUsingSupplierNumber.Location = New System.Drawing.Point(6, 11)
        Me.ChkUsingSupplierNumber.Name = "ChkUsingSupplierNumber"
        Me.ChkUsingSupplierNumber.Size = New System.Drawing.Size(164, 17)
        Me.ChkUsingSupplierNumber.TabIndex = 0
        Me.ChkUsingSupplierNumber.Text = "Use Supplier Invoice Number"
        Me.ChkUsingSupplierNumber.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.BtnClose)
        Me.GroupBox1.Controls.Add(Me.BtnPrint)
        Me.GroupBox1.Controls.Add(Me.TabControl1)
        Me.GroupBox1.Controls.Add(Me.ChkUsingSupplierNumber)
        Me.GroupBox1.Controls.Add(Me.CboFormat)
        Me.GroupBox1.Controls.Add(Me.DtpTo)
        Me.GroupBox1.Controls.Add(Me.DtpFrom)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(712, 149)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(59, 116)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(53, 23)
        Me.BtnClose.TabIndex = 11
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(203, 11)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(503, 128)
        Me.TabControl1.TabIndex = 13
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.ChkShowRatio)
        Me.TabPage1.Controls.Add(Me.TxtAccountDesc)
        Me.TabPage1.Controls.Add(Me.TxtSuppliserDesc)
        Me.TabPage1.Controls.Add(Me.ChkAllProductPO)
        Me.TabPage1.Controls.Add(Me.BtnFindSupp)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.TxtPrimaryAcc)
        Me.TabPage1.Controls.Add(Me.TxtKodeSupplier)
        Me.TabPage1.Controls.Add(Me.TxtAccCode)
        Me.TabPage1.Controls.Add(Me.TxtPrimarySupplier)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.BtnFindAcc)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Size = New System.Drawing.Size(495, 102)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Additional Filter"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'ChkShowRatio
        '
        Me.ChkShowRatio.AutoSize = True
        Me.ChkShowRatio.Location = New System.Drawing.Point(182, 67)
        Me.ChkShowRatio.Name = "ChkShowRatio"
        Me.ChkShowRatio.Size = New System.Drawing.Size(142, 17)
        Me.ChkShowRatio.TabIndex = 20
        Me.ChkShowRatio.Text = "Show Ratio (If Available)"
        Me.ChkShowRatio.UseVisualStyleBackColor = True
        '
        'TxtAccountDesc
        '
        Me.TxtAccountDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccountDesc.Location = New System.Drawing.Point(259, 31)
        Me.TxtAccountDesc.Name = "TxtAccountDesc"
        Me.TxtAccountDesc.ReadOnly = True
        Me.TxtAccountDesc.Size = New System.Drawing.Size(231, 20)
        Me.TxtAccountDesc.TabIndex = 19
        Me.TxtAccountDesc.TabStop = False
        '
        'TxtSuppliserDesc
        '
        Me.TxtSuppliserDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSuppliserDesc.Location = New System.Drawing.Point(259, 6)
        Me.TxtSuppliserDesc.Name = "TxtSuppliserDesc"
        Me.TxtSuppliserDesc.ReadOnly = True
        Me.TxtSuppliserDesc.Size = New System.Drawing.Size(231, 20)
        Me.TxtSuppliserDesc.TabIndex = 19
        Me.TxtSuppliserDesc.TabStop = False
        '
        'ChkAllProductPO
        '
        Me.ChkAllProductPO.AutoSize = True
        Me.ChkAllProductPO.Location = New System.Drawing.Point(6, 67)
        Me.ChkAllProductPO.Name = "ChkAllProductPO"
        Me.ChkAllProductPO.Size = New System.Drawing.Size(171, 17)
        Me.ChkAllProductPO.TabIndex = 9
        Me.ChkAllProductPO.Text = "Always Display All Invoice Item"
        Me.ChkAllProductPO.UseVisualStyleBackColor = True
        '
        'BtnFindSupp
        '
        Me.BtnFindSupp.Location = New System.Drawing.Point(216, 6)
        Me.BtnFindSupp.Name = "BtnFindSupp"
        Me.BtnFindSupp.Size = New System.Drawing.Size(37, 20)
        Me.BtnFindSupp.TabIndex = 6
        Me.BtnFindSupp.TabStop = False
        Me.BtnFindSupp.Text = "Find"
        Me.BtnFindSupp.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 10)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Supplier Code"
        '
        'TxtPrimaryAcc
        '
        Me.TxtPrimaryAcc.Location = New System.Drawing.Point(98, 31)
        Me.TxtPrimaryAcc.Name = "TxtPrimaryAcc"
        Me.TxtPrimaryAcc.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryAcc.TabIndex = 17
        Me.TxtPrimaryAcc.TabStop = False
        Me.TxtPrimaryAcc.Text = "-1"
        Me.TxtPrimaryAcc.Visible = False
        '
        'TxtKodeSupplier
        '
        Me.TxtKodeSupplier.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtKodeSupplier.Location = New System.Drawing.Point(98, 6)
        Me.TxtKodeSupplier.Name = "TxtKodeSupplier"
        Me.TxtKodeSupplier.Size = New System.Drawing.Size(113, 20)
        Me.TxtKodeSupplier.TabIndex = 5
        '
        'TxtAccCode
        '
        Me.TxtAccCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAccCode.Location = New System.Drawing.Point(98, 31)
        Me.TxtAccCode.Name = "TxtAccCode"
        Me.TxtAccCode.Size = New System.Drawing.Size(113, 20)
        Me.TxtAccCode.TabIndex = 7
        '
        'TxtPrimarySupplier
        '
        Me.TxtPrimarySupplier.Location = New System.Drawing.Point(98, 7)
        Me.TxtPrimarySupplier.Name = "TxtPrimarySupplier"
        Me.TxtPrimarySupplier.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimarySupplier.TabIndex = 13
        Me.TxtPrimarySupplier.TabStop = False
        Me.TxtPrimarySupplier.Text = "-1"
        Me.TxtPrimarySupplier.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 32)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 13)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Payable Account"
        '
        'BtnFindAcc
        '
        Me.BtnFindAcc.Location = New System.Drawing.Point(216, 31)
        Me.BtnFindAcc.Name = "BtnFindAcc"
        Me.BtnFindAcc.Size = New System.Drawing.Size(37, 20)
        Me.BtnFindAcc.TabIndex = 8
        Me.BtnFindAcc.TabStop = False
        Me.BtnFindAcc.Text = "Find"
        Me.BtnFindAcc.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvItemProductGroup)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Size = New System.Drawing.Size(495, 102)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Product Filter"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(2, 2)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidth = 24
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(491, 98)
        Me.dgvItemProductGroup.TabIndex = 8
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.ReadOnly = True
        Me.ColProductGroup_No.Width = 46
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.Width = 39
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.HeaderText = "Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.Width = 57
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.ReadOnly = True
        Me.ColProductGroup_Desc.Width = 85
        '
        'ColProductGroup_Remarks
        '
        Me.ColProductGroup_Remarks.HeaderText = "Remarks"
        Me.ColProductGroup_Remarks.Name = "ColProductGroup_Remarks"
        Me.ColProductGroup_Remarks.ReadOnly = True
        Me.ColProductGroup_Remarks.Visible = False
        Me.ColProductGroup_Remarks.Width = 89
        '
        'ColProductGroup_PrimaryAcuan
        '
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.ReadOnly = True
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 81
        '
        'CboFormat
        '
        Me.CboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFormat.FormattingEnabled = True
        Me.CboFormat.Items.AddRange(New Object() {"Rekap", "Detail"})
        Me.CboFormat.Location = New System.Drawing.Point(48, 30)
        Me.CboFormat.Name = "CboFormat"
        Me.CboFormat.Size = New System.Drawing.Size(126, 21)
        Me.CboFormat.TabIndex = 1
        '
        'DtpTo
        '
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(48, 84)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(126, 20)
        Me.DtpTo.TabIndex = 3
        '
        'DtpFrom
        '
        Me.DtpFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(48, 58)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(126, 20)
        Me.DtpFrom.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 34)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(39, 13)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "Format"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 84)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "To"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "From"
        '
        'TxtCodeNoDesc
        '
        Me.TxtCodeNoDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeNoDesc.Location = New System.Drawing.Point(464, 261)
        Me.TxtCodeNoDesc.Name = "TxtCodeNoDesc"
        Me.TxtCodeNoDesc.ReadOnly = True
        Me.TxtCodeNoDesc.Size = New System.Drawing.Size(236, 20)
        Me.TxtCodeNoDesc.TabIndex = 19
        Me.TxtCodeNoDesc.TabStop = False
        '
        'BtnFindCodeNo
        '
        Me.BtnFindCodeNo.Location = New System.Drawing.Point(421, 261)
        Me.BtnFindCodeNo.Name = "BtnFindCodeNo"
        Me.BtnFindCodeNo.Size = New System.Drawing.Size(37, 20)
        Me.BtnFindCodeNo.TabIndex = 4
        Me.BtnFindCodeNo.TabStop = False
        Me.BtnFindCodeNo.Text = "Find"
        Me.BtnFindCodeNo.UseVisualStyleBackColor = True
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(208, 264)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(49, 13)
        Me.Label19.TabIndex = 6
        Me.Label19.Text = "Code No"
        '
        'txtCodeNo
        '
        Me.txtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCodeNo.Location = New System.Drawing.Point(303, 261)
        Me.txtCodeNo.Name = "txtCodeNo"
        Me.txtCodeNo.Size = New System.Drawing.Size(113, 20)
        Me.txtCodeNo.TabIndex = 3
        '
        'txtPrimaryCodeNo
        '
        Me.txtPrimaryCodeNo.Location = New System.Drawing.Point(325, 236)
        Me.txtPrimaryCodeNo.Name = "txtPrimaryCodeNo"
        Me.txtPrimaryCodeNo.Size = New System.Drawing.Size(30, 20)
        Me.txtPrimaryCodeNo.TabIndex = 7
        Me.txtPrimaryCodeNo.TabStop = False
        Me.txtPrimaryCodeNo.Text = "-1"
        Me.txtPrimaryCodeNo.Visible = False
        '
        'Crv
        '
        Me.Crv.ActiveViewIndex = -1
        Me.Crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Crv.DisplayGroupTree = False
        Me.Crv.EnableDrillDown = False
        Me.Crv.Location = New System.Drawing.Point(12, 155)
        Me.Crv.Name = "Crv"
        Me.Crv.SelectionFormula = ""
        Me.Crv.ShowCloseButton = False
        Me.Crv.ShowGroupTreeButton = False
        Me.Crv.ShowRefreshButton = False
        Me.Crv.Size = New System.Drawing.Size(714, 344)
        Me.Crv.TabIndex = 12
        Me.Crv.TabStop = False
        Me.Crv.ViewTimeSelectionFormula = ""
        '
        'BackgroundWorker3
        '
        '
        'FrmLaporanPurchaseByInvoice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(740, 514)
        Me.Controls.Add(Me.Crv)
        Me.Controls.Add(Me.TxtCodeNoDesc)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BtnFindCodeNo)
        Me.Controls.Add(Me.txtPrimaryCodeNo)
        Me.Controls.Add(Me.txtCodeNo)
        Me.Controls.Add(Me.Label19)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(756, 552)
        Me.Name = "FrmLaporanPurchaseByInvoice"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Report By Invoice"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
    Friend WithEvents ChkUsingSupplierNumber As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryAcc As System.Windows.Forms.TextBox
    Friend WithEvents TxtAccCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BtnFindAcc As System.Windows.Forms.Button
    Friend WithEvents TxtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents TxtKodeSupplier As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BtnFindSupp As System.Windows.Forms.Button
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents ChkAllProductPO As System.Windows.Forms.CheckBox
    Friend WithEvents TxtAccountDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtSuppliserDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtCodeNoDesc As System.Windows.Forms.TextBox
    Friend WithEvents ChkShowRatio As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
End Class
