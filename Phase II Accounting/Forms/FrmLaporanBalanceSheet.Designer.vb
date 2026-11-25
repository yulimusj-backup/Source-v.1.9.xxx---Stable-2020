<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanBalanceSheet
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
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.cboFormatLaporan = New System.Windows.Forms.ComboBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.dtpFrom_I = New System.Windows.Forms.DateTimePicker
        Me.LblPeriod = New System.Windows.Forms.Label
        Me.dtpFrom_II = New System.Windows.Forms.DateTimePicker
        Me.LblPeriodII = New System.Windows.Forms.Label
        Me.chkUseTotal = New System.Windows.Forms.CheckBox
        Me.chkShowLvl4 = New System.Windows.Forms.CheckBox
        Me.chkShowLvl3 = New System.Windows.Forms.CheckBox
        Me.chkShowLvl2 = New System.Windows.Forms.CheckBox
        Me.chkHideZero = New System.Windows.Forms.CheckBox
        Me.chkHideZeroInactiveAccount = New System.Windows.Forms.CheckBox
        Me.chkShowAccountCode = New System.Windows.Forms.CheckBox
        Me.TxtPrimaryJob = New System.Windows.Forms.TextBox
        Me.btnFindJob = New System.Windows.Forms.Button
        Me.TxtJobCode = New System.Windows.Forms.TextBox
        Me.LblJob = New System.Windows.Forms.Label
        Me.TxtPrimaryDept = New System.Windows.Forms.TextBox
        Me.btnFindDept = New System.Windows.Forms.Button
        Me.TxtDeptCode = New System.Windows.Forms.TextBox
        Me.LblDept = New System.Windows.Forms.Label
        Me.chkShowLvl8 = New System.Windows.Forms.CheckBox
        Me.chkShowLvl7 = New System.Windows.Forms.CheckBox
        Me.chkShowLvl6 = New System.Windows.Forms.CheckBox
        Me.chkShowLvl5 = New System.Windows.Forms.CheckBox
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.Location = New System.Drawing.Point(355, 12)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(517, 538)
        Me.crv1.TabIndex = 11
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'cboFormatLaporan
        '
        Me.cboFormatLaporan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormatLaporan.FormattingEnabled = True
        Me.cboFormatLaporan.Items.AddRange(New Object() {"Balance Sheet By Period", "Balance Sheet 2 Period Comparison", "Balance Sheet Vs Budget Comparison", "Balance Sheet 6 Months Period", "Balance Sheet 6 Months Period w/ Budget", "Balance Sheet 1 Year Period", "Balance Sheet 1 Year Period w/ Budget"})
        Me.cboFormatLaporan.Location = New System.Drawing.Point(8, 19)
        Me.cboFormatLaporan.Name = "cboFormatLaporan"
        Me.cboFormatLaporan.Size = New System.Drawing.Size(304, 21)
        Me.cboFormatLaporan.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(231, 282)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 23)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(149, 282)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(76, 23)
        Me.btnShow.TabIndex = 8
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'dtpFrom_I
        '
        Me.dtpFrom_I.CustomFormat = "MMM yyyy"
        Me.dtpFrom_I.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom_I.Location = New System.Drawing.Point(70, 49)
        Me.dtpFrom_I.Name = "dtpFrom_I"
        Me.dtpFrom_I.ShowUpDown = True
        Me.dtpFrom_I.Size = New System.Drawing.Size(90, 20)
        Me.dtpFrom_I.TabIndex = 1
        '
        'LblPeriod
        '
        Me.LblPeriod.AutoSize = True
        Me.LblPeriod.Location = New System.Drawing.Point(6, 54)
        Me.LblPeriod.Name = "LblPeriod"
        Me.LblPeriod.Size = New System.Drawing.Size(43, 13)
        Me.LblPeriod.TabIndex = 41
        Me.LblPeriod.Text = "Period I"
        '
        'dtpFrom_II
        '
        Me.dtpFrom_II.CustomFormat = "MMM yyyy"
        Me.dtpFrom_II.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom_II.Location = New System.Drawing.Point(222, 48)
        Me.dtpFrom_II.Name = "dtpFrom_II"
        Me.dtpFrom_II.ShowUpDown = True
        Me.dtpFrom_II.Size = New System.Drawing.Size(90, 20)
        Me.dtpFrom_II.TabIndex = 2
        '
        'LblPeriodII
        '
        Me.LblPeriodII.AutoSize = True
        Me.LblPeriodII.Location = New System.Drawing.Point(170, 54)
        Me.LblPeriodII.Name = "LblPeriodII"
        Me.LblPeriodII.Size = New System.Drawing.Size(46, 13)
        Me.LblPeriodII.TabIndex = 43
        Me.LblPeriodII.Text = "Period II"
        '
        'chkUseTotal
        '
        Me.chkUseTotal.AutoSize = True
        Me.chkUseTotal.Location = New System.Drawing.Point(9, 81)
        Me.chkUseTotal.Name = "chkUseTotal"
        Me.chkUseTotal.Size = New System.Drawing.Size(72, 17)
        Me.chkUseTotal.TabIndex = 3
        Me.chkUseTotal.Text = "Use Total"
        Me.chkUseTotal.UseVisualStyleBackColor = True
        '
        'chkShowLvl4
        '
        Me.chkShowLvl4.AutoSize = True
        Me.chkShowLvl4.Checked = True
        Me.chkShowLvl4.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl4.Location = New System.Drawing.Point(9, 231)
        Me.chkShowLvl4.Name = "chkShowLvl4"
        Me.chkShowLvl4.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl4.TabIndex = 6
        Me.chkShowLvl4.Text = "Show level 4 accounts"
        Me.chkShowLvl4.UseVisualStyleBackColor = True
        '
        'chkShowLvl3
        '
        Me.chkShowLvl3.AutoSize = True
        Me.chkShowLvl3.Checked = True
        Me.chkShowLvl3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl3.Location = New System.Drawing.Point(9, 209)
        Me.chkShowLvl3.Name = "chkShowLvl3"
        Me.chkShowLvl3.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl3.TabIndex = 5
        Me.chkShowLvl3.Text = "Show level 3 accounts"
        Me.chkShowLvl3.UseVisualStyleBackColor = True
        '
        'chkShowLvl2
        '
        Me.chkShowLvl2.AutoSize = True
        Me.chkShowLvl2.Checked = True
        Me.chkShowLvl2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl2.Location = New System.Drawing.Point(9, 186)
        Me.chkShowLvl2.Name = "chkShowLvl2"
        Me.chkShowLvl2.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl2.TabIndex = 4
        Me.chkShowLvl2.Text = "Show level 2 accounts"
        Me.chkShowLvl2.UseVisualStyleBackColor = True
        '
        'chkHideZero
        '
        Me.chkHideZero.AutoSize = True
        Me.chkHideZero.Location = New System.Drawing.Point(9, 126)
        Me.chkHideZero.Name = "chkHideZero"
        Me.chkHideZero.Size = New System.Drawing.Size(151, 17)
        Me.chkHideZero.TabIndex = 5
        Me.chkHideZero.Text = "Hide zero amount account"
        Me.chkHideZero.UseVisualStyleBackColor = True
        '
        'chkHideZeroInactiveAccount
        '
        Me.chkHideZeroInactiveAccount.AutoSize = True
        Me.chkHideZeroInactiveAccount.Location = New System.Drawing.Point(9, 149)
        Me.chkHideZeroInactiveAccount.Name = "chkHideZeroInactiveAccount"
        Me.chkHideZeroInactiveAccount.Size = New System.Drawing.Size(153, 17)
        Me.chkHideZeroInactiveAccount.TabIndex = 6
        Me.chkHideZeroInactiveAccount.Text = "Hide zero inactive account"
        Me.chkHideZeroInactiveAccount.UseVisualStyleBackColor = True
        '
        'chkShowAccountCode
        '
        Me.chkShowAccountCode.AutoSize = True
        Me.chkShowAccountCode.Checked = True
        Me.chkShowAccountCode.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowAccountCode.Location = New System.Drawing.Point(9, 103)
        Me.chkShowAccountCode.Name = "chkShowAccountCode"
        Me.chkShowAccountCode.Size = New System.Drawing.Size(122, 17)
        Me.chkShowAccountCode.TabIndex = 4
        Me.chkShowAccountCode.Text = "Show account code"
        Me.chkShowAccountCode.UseVisualStyleBackColor = True
        '
        'TxtPrimaryJob
        '
        Me.TxtPrimaryJob.Location = New System.Drawing.Point(500, 86)
        Me.TxtPrimaryJob.Name = "TxtPrimaryJob"
        Me.TxtPrimaryJob.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryJob.TabIndex = 67
        Me.TxtPrimaryJob.Text = "-1"
        Me.TxtPrimaryJob.Visible = False
        '
        'btnFindJob
        '
        Me.btnFindJob.Location = New System.Drawing.Point(652, 86)
        Me.btnFindJob.Name = "btnFindJob"
        Me.btnFindJob.Size = New System.Drawing.Size(37, 21)
        Me.btnFindJob.TabIndex = 13
        Me.btnFindJob.TabStop = False
        Me.btnFindJob.Text = "Find"
        Me.btnFindJob.UseVisualStyleBackColor = True
        '
        'TxtJobCode
        '
        Me.TxtJobCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtJobCode.Location = New System.Drawing.Point(536, 87)
        Me.TxtJobCode.Name = "TxtJobCode"
        Me.TxtJobCode.Size = New System.Drawing.Size(110, 20)
        Me.TxtJobCode.TabIndex = 12
        '
        'LblJob
        '
        Me.LblJob.AutoSize = True
        Me.LblJob.Location = New System.Drawing.Point(452, 90)
        Me.LblJob.Name = "LblJob"
        Me.LblJob.Size = New System.Drawing.Size(24, 13)
        Me.LblJob.TabIndex = 66
        Me.LblJob.Text = "Job"
        '
        'TxtPrimaryDept
        '
        Me.TxtPrimaryDept.Location = New System.Drawing.Point(500, 63)
        Me.TxtPrimaryDept.Name = "TxtPrimaryDept"
        Me.TxtPrimaryDept.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryDept.TabIndex = 65
        Me.TxtPrimaryDept.Text = "-1"
        Me.TxtPrimaryDept.Visible = False
        '
        'btnFindDept
        '
        Me.btnFindDept.Location = New System.Drawing.Point(652, 63)
        Me.btnFindDept.Name = "btnFindDept"
        Me.btnFindDept.Size = New System.Drawing.Size(37, 21)
        Me.btnFindDept.TabIndex = 11
        Me.btnFindDept.TabStop = False
        Me.btnFindDept.Text = "Find"
        Me.btnFindDept.UseVisualStyleBackColor = True
        '
        'TxtDeptCode
        '
        Me.TxtDeptCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDeptCode.Location = New System.Drawing.Point(536, 64)
        Me.TxtDeptCode.Name = "TxtDeptCode"
        Me.TxtDeptCode.Size = New System.Drawing.Size(110, 20)
        Me.TxtDeptCode.TabIndex = 10
        '
        'LblDept
        '
        Me.LblDept.AutoSize = True
        Me.LblDept.Location = New System.Drawing.Point(452, 67)
        Me.LblDept.Name = "LblDept"
        Me.LblDept.Size = New System.Drawing.Size(68, 13)
        Me.LblDept.TabIndex = 64
        Me.LblDept.Text = "Departement"
        '
        'chkShowLvl8
        '
        Me.chkShowLvl8.AutoSize = True
        Me.chkShowLvl8.Checked = True
        Me.chkShowLvl8.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl8.Location = New System.Drawing.Point(173, 231)
        Me.chkShowLvl8.Name = "chkShowLvl8"
        Me.chkShowLvl8.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl8.TabIndex = 71
        Me.chkShowLvl8.Text = "Show level 8 accounts"
        Me.chkShowLvl8.UseVisualStyleBackColor = True
        '
        'chkShowLvl7
        '
        Me.chkShowLvl7.AutoSize = True
        Me.chkShowLvl7.Checked = True
        Me.chkShowLvl7.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl7.Location = New System.Drawing.Point(173, 208)
        Me.chkShowLvl7.Name = "chkShowLvl7"
        Me.chkShowLvl7.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl7.TabIndex = 70
        Me.chkShowLvl7.Text = "Show level 7 accounts"
        Me.chkShowLvl7.UseVisualStyleBackColor = True
        '
        'chkShowLvl6
        '
        Me.chkShowLvl6.AutoSize = True
        Me.chkShowLvl6.Checked = True
        Me.chkShowLvl6.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl6.Location = New System.Drawing.Point(173, 185)
        Me.chkShowLvl6.Name = "chkShowLvl6"
        Me.chkShowLvl6.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl6.TabIndex = 69
        Me.chkShowLvl6.Text = "Show level 6 accounts"
        Me.chkShowLvl6.UseVisualStyleBackColor = True
        '
        'chkShowLvl5
        '
        Me.chkShowLvl5.AutoSize = True
        Me.chkShowLvl5.Checked = True
        Me.chkShowLvl5.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl5.Location = New System.Drawing.Point(9, 254)
        Me.chkShowLvl5.Name = "chkShowLvl5"
        Me.chkShowLvl5.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl5.TabIndex = 68
        Me.chkShowLvl5.Text = "Show level 5 accounts"
        Me.chkShowLvl5.UseVisualStyleBackColor = True
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(11, 287)
        Me.LinkLabel1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(70, 13)
        Me.LinkLabel1.TabIndex = 7
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Detail Setting"
        Me.LinkLabel1.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cboFormatLaporan)
        Me.GroupBox1.Controls.Add(Me.chkShowLvl8)
        Me.GroupBox1.Controls.Add(Me.chkShowLvl7)
        Me.GroupBox1.Controls.Add(Me.LinkLabel1)
        Me.GroupBox1.Controls.Add(Me.chkShowLvl6)
        Me.GroupBox1.Controls.Add(Me.dtpFrom_II)
        Me.GroupBox1.Controls.Add(Me.chkShowLvl5)
        Me.GroupBox1.Controls.Add(Me.LblPeriod)
        Me.GroupBox1.Controls.Add(Me.dtpFrom_I)
        Me.GroupBox1.Controls.Add(Me.LblPeriodII)
        Me.GroupBox1.Controls.Add(Me.chkUseTotal)
        Me.GroupBox1.Controls.Add(Me.chkHideZeroInactiveAccount)
        Me.GroupBox1.Controls.Add(Me.chkHideZero)
        Me.GroupBox1.Controls.Add(Me.chkShowAccountCode)
        Me.GroupBox1.Controls.Add(Me.btnShow)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.chkShowLvl4)
        Me.GroupBox1.Controls.Add(Me.chkShowLvl2)
        Me.GroupBox1.Controls.Add(Me.chkShowLvl3)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(320, 316)
        Me.GroupBox1.TabIndex = 73
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter By"
        '
        'FrmLaporanBalanceSheet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(884, 562)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TxtPrimaryJob)
        Me.Controls.Add(Me.btnFindJob)
        Me.Controls.Add(Me.TxtJobCode)
        Me.Controls.Add(Me.LblJob)
        Me.Controls.Add(Me.TxtPrimaryDept)
        Me.Controls.Add(Me.btnFindDept)
        Me.Controls.Add(Me.TxtDeptCode)
        Me.Controls.Add(Me.LblDept)
        Me.Controls.Add(Me.crv1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 600)
        Me.Name = "FrmLaporanBalanceSheet"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Balance Sheet"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents cboFormatLaporan As System.Windows.Forms.ComboBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents dtpFrom_I As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblPeriod As System.Windows.Forms.Label
    Friend WithEvents dtpFrom_II As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblPeriodII As System.Windows.Forms.Label
    Friend WithEvents chkUseTotal As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl4 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl3 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkHideZero As System.Windows.Forms.CheckBox
    Friend WithEvents chkHideZeroInactiveAccount As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowAccountCode As System.Windows.Forms.CheckBox
    Friend WithEvents TxtPrimaryJob As System.Windows.Forms.TextBox
    Friend WithEvents btnFindJob As System.Windows.Forms.Button
    Friend WithEvents TxtJobCode As System.Windows.Forms.TextBox
    Friend WithEvents LblJob As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryDept As System.Windows.Forms.TextBox
    Friend WithEvents btnFindDept As System.Windows.Forms.Button
    Friend WithEvents TxtDeptCode As System.Windows.Forms.TextBox
    Friend WithEvents LblDept As System.Windows.Forms.Label
    Friend WithEvents chkShowLvl8 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl7 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl6 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl5 As System.Windows.Forms.CheckBox
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
End Class
