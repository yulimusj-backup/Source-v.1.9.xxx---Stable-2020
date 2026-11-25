<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanLogUser
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.txtUser = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnFindUser = New System.Windows.Forms.Button
        Me.txtPrimaryUser = New System.Windows.Forms.TextBox
        Me.CboStatusPO = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BgwUser = New System.ComponentModel.BackgroundWorker
        Me.txtReadonly = New System.Windows.Forms.TextBox
        Me.txtPrimarySalaryFormula = New System.Windows.Forms.TextBox
        Me.btnFindSalaryFormula = New System.Windows.Forms.Button
        Me.txtSalaryFormulaCode = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(669, 42)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(50, 23)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(669, 15)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(50, 23)
        Me.btnFilter.TabIndex = 9
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(221, 18)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(107, 20)
        Me.dtpTo.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(194, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "To"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(83, 18)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(105, 20)
        Me.dtpFrom.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Activity Date"
        '
        'crv
        '
        Me.crv.ActiveViewIndex = -1
        Me.crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv.DisplayGroupTree = False
        Me.crv.EnableDrillDown = False
        Me.crv.Location = New System.Drawing.Point(0, 84)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(738, 430)
        Me.crv.TabIndex = 0
        Me.crv.TabStop = False
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'txtUser
        '
        Me.txtUser.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtUser.Location = New System.Drawing.Point(430, 18)
        Me.txtUser.Name = "txtUser"
        Me.txtUser.Size = New System.Drawing.Size(133, 20)
        Me.txtUser.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(391, 22)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(29, 13)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "User"
        '
        'btnFindUser
        '
        Me.btnFindUser.Location = New System.Drawing.Point(569, 19)
        Me.btnFindUser.Name = "btnFindUser"
        Me.btnFindUser.Size = New System.Drawing.Size(31, 18)
        Me.btnFindUser.TabIndex = 6
        Me.btnFindUser.TabStop = False
        Me.btnFindUser.Text = "..."
        Me.btnFindUser.UseVisualStyleBackColor = True
        '
        'txtPrimaryUser
        '
        Me.txtPrimaryUser.Location = New System.Drawing.Point(421, 18)
        Me.txtPrimaryUser.Name = "txtPrimaryUser"
        Me.txtPrimaryUser.Size = New System.Drawing.Size(34, 20)
        Me.txtPrimaryUser.TabIndex = 14
        Me.txtPrimaryUser.Text = "-1"
        Me.txtPrimaryUser.Visible = False
        '
        'CboStatusPO
        '
        Me.CboStatusPO.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboStatusPO.FormattingEnabled = True
        Me.CboStatusPO.Items.AddRange(New Object() {"All Activity", "Create", "Modify", "Delete"})
        Me.CboStatusPO.Location = New System.Drawing.Point(83, 45)
        Me.CboStatusPO.Name = "CboStatusPO"
        Me.CboStatusPO.Size = New System.Drawing.Size(245, 21)
        Me.CboStatusPO.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 49)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(68, 13)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Activity Type"
        '
        'txtReadonly
        '
        Me.txtReadonly.Location = New System.Drawing.Point(594, 18)
        Me.txtReadonly.Name = "txtReadonly"
        Me.txtReadonly.ReadOnly = True
        Me.txtReadonly.Size = New System.Drawing.Size(34, 20)
        Me.txtReadonly.TabIndex = 17
        Me.txtReadonly.Text = " "
        Me.txtReadonly.Visible = False
        '
        'txtPrimarySalaryFormula
        '
        Me.txtPrimarySalaryFormula.Location = New System.Drawing.Point(421, 45)
        Me.txtPrimarySalaryFormula.Name = "txtPrimarySalaryFormula"
        Me.txtPrimarySalaryFormula.Size = New System.Drawing.Size(34, 20)
        Me.txtPrimarySalaryFormula.TabIndex = 15
        Me.txtPrimarySalaryFormula.Text = "-1"
        Me.txtPrimarySalaryFormula.Visible = False
        '
        'btnFindSalaryFormula
        '
        Me.btnFindSalaryFormula.Location = New System.Drawing.Point(569, 46)
        Me.btnFindSalaryFormula.Name = "btnFindSalaryFormula"
        Me.btnFindSalaryFormula.Size = New System.Drawing.Size(31, 18)
        Me.btnFindSalaryFormula.TabIndex = 8
        Me.btnFindSalaryFormula.TabStop = False
        Me.btnFindSalaryFormula.Text = "..."
        Me.btnFindSalaryFormula.UseVisualStyleBackColor = True
        '
        'txtSalaryFormulaCode
        '
        Me.txtSalaryFormulaCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSalaryFormulaCode.Location = New System.Drawing.Point(430, 45)
        Me.txtSalaryFormulaCode.Name = "txtSalaryFormulaCode"
        Me.txtSalaryFormulaCode.Size = New System.Drawing.Size(133, 20)
        Me.txtSalaryFormulaCode.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(314, 49)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 13)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Salary Formula Code"
        '
        'FrmLaporanLogUser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(740, 514)
        Me.Controls.Add(Me.txtPrimarySalaryFormula)
        Me.Controls.Add(Me.btnFindSalaryFormula)
        Me.Controls.Add(Me.txtSalaryFormulaCode)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtReadonly)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.CboStatusPO)
        Me.Controls.Add(Me.txtPrimaryUser)
        Me.Controls.Add(Me.btnFindUser)
        Me.Controls.Add(Me.txtUser)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.crv)
        Me.MinimumSize = New System.Drawing.Size(756, 552)
        Me.Name = "FrmLaporanLogUser"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "User activity report on salary formula"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents txtUser As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnFindUser As System.Windows.Forms.Button
    Friend WithEvents txtPrimaryUser As System.Windows.Forms.TextBox
    Friend WithEvents CboStatusPO As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwUser As System.ComponentModel.BackgroundWorker
    Friend WithEvents txtReadonly As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimarySalaryFormula As System.Windows.Forms.TextBox
    Friend WithEvents btnFindSalaryFormula As System.Windows.Forms.Button
    Friend WithEvents txtSalaryFormulaCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
End Class
