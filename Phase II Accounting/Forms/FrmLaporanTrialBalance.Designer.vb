<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanTrialBalance
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.cboFormatLaporan = New System.Windows.Forms.ComboBox
        Me.cboJenis = New System.Windows.Forms.ComboBox
        Me.dtpFrom_I = New System.Windows.Forms.DateTimePicker
        Me.LblPeriod = New System.Windows.Forms.Label
        Me.dtpTo_I = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.gboPeriod_I = New System.Windows.Forms.GroupBox
        Me.gboPeriod_II = New System.Windows.Forms.GroupBox
        Me.dtpFrom_II = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpTo_II = New System.Windows.Forms.DateTimePicker
        Me.LblPeriodII = New System.Windows.Forms.Label
        Me.chkHideZeroInactiveAccount = New System.Windows.Forms.CheckBox
        Me.chkHideZero = New System.Windows.Forms.CheckBox
        Me.chkShowLvl2 = New System.Windows.Forms.CheckBox
        Me.chkShowLvl3 = New System.Windows.Forms.CheckBox
        Me.chkShowLvl4 = New System.Windows.Forms.CheckBox
        Me.chkUseTotal = New System.Windows.Forms.CheckBox
        Me.chkShowAccountCode = New System.Windows.Forms.CheckBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.LblJob = New System.Windows.Forms.Label
        Me.LblDept = New System.Windows.Forms.Label
        Me.TxtPrimaryJob = New System.Windows.Forms.TextBox
        Me.btnFindJob = New System.Windows.Forms.Button
        Me.TxtJobCode = New System.Windows.Forms.TextBox
        Me.TxtPrimaryDept = New System.Windows.Forms.TextBox
        Me.btnFindDept = New System.Windows.Forms.Button
        Me.TxtDeptCode = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.gboPeriod_I.SuspendLayout()
        Me.gboPeriod_II.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
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
        Me.crv1.EnableDrillDown = False
        Me.crv1.Location = New System.Drawing.Point(285, 12)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(587, 488)
        Me.crv1.TabIndex = 12
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(185, 68)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 23)
        Me.btnClose.TabIndex = 11
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(185, 44)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(76, 23)
        Me.btnShow.TabIndex = 10
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'cboFormatLaporan
        '
        Me.cboFormatLaporan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormatLaporan.FormattingEnabled = True
        Me.cboFormatLaporan.Items.AddRange(New Object() {"Trial Balance By Period", "Trial Balance 2 Period Comparison", "Trial Balance 2 Period Comparison Landscape", "Trial Balance Vs Budget Comparison"})
        Me.cboFormatLaporan.Location = New System.Drawing.Point(6, 19)
        Me.cboFormatLaporan.Name = "cboFormatLaporan"
        Me.cboFormatLaporan.Size = New System.Drawing.Size(255, 21)
        Me.cboFormatLaporan.TabIndex = 0
        '
        'cboJenis
        '
        Me.cboJenis.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboJenis.FormattingEnabled = True
        Me.cboJenis.Items.AddRange(New Object() {"Trial Balance", "Adjusted Trial Balance"})
        Me.cboJenis.Location = New System.Drawing.Point(6, 19)
        Me.cboJenis.Name = "cboJenis"
        Me.cboJenis.Size = New System.Drawing.Size(234, 21)
        Me.cboJenis.TabIndex = 0
        Me.cboJenis.TabStop = False
        Me.cboJenis.Visible = False
        '
        'dtpFrom_I
        '
        Me.dtpFrom_I.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom_I.CustomFormat = "MMM yyyy"
        Me.dtpFrom_I.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom_I.Location = New System.Drawing.Point(42, 19)
        Me.dtpFrom_I.Name = "dtpFrom_I"
        Me.dtpFrom_I.ShowUpDown = True
        Me.dtpFrom_I.Size = New System.Drawing.Size(111, 20)
        Me.dtpFrom_I.TabIndex = 0
        '
        'LblPeriod
        '
        Me.LblPeriod.AutoSize = True
        Me.LblPeriod.Location = New System.Drawing.Point(6, 24)
        Me.LblPeriod.Name = "LblPeriod"
        Me.LblPeriod.Size = New System.Drawing.Size(30, 13)
        Me.LblPeriod.TabIndex = 30
        Me.LblPeriod.Text = "From"
        '
        'dtpTo_I
        '
        Me.dtpTo_I.CustomFormat = "MMM yyyy"
        Me.dtpTo_I.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo_I.Location = New System.Drawing.Point(55, 46)
        Me.dtpTo_I.Name = "dtpTo_I"
        Me.dtpTo_I.ShowUpDown = True
        Me.dtpTo_I.Size = New System.Drawing.Size(120, 20)
        Me.dtpTo_I.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 52)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 32
        Me.Label1.Text = "Period I"
        '
        'gboPeriod_I
        '
        Me.gboPeriod_I.Controls.Add(Me.dtpFrom_I)
        Me.gboPeriod_I.Controls.Add(Me.LblPeriod)
        Me.gboPeriod_I.Controls.Add(Me.gboPeriod_II)
        Me.gboPeriod_I.Enabled = False
        Me.gboPeriod_I.Location = New System.Drawing.Point(605, 131)
        Me.gboPeriod_I.Name = "gboPeriod_I"
        Me.gboPeriod_I.Size = New System.Drawing.Size(161, 76)
        Me.gboPeriod_I.TabIndex = 8
        Me.gboPeriod_I.TabStop = False
        Me.gboPeriod_I.Text = "Period I"
        Me.gboPeriod_I.Visible = False
        '
        'gboPeriod_II
        '
        Me.gboPeriod_II.Controls.Add(Me.dtpFrom_II)
        Me.gboPeriod_II.Controls.Add(Me.Label2)
        Me.gboPeriod_II.Location = New System.Drawing.Point(77, 45)
        Me.gboPeriod_II.Name = "gboPeriod_II"
        Me.gboPeriod_II.Size = New System.Drawing.Size(161, 76)
        Me.gboPeriod_II.TabIndex = 9
        Me.gboPeriod_II.TabStop = False
        Me.gboPeriod_II.Text = "Period II"
        '
        'dtpFrom_II
        '
        Me.dtpFrom_II.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom_II.CustomFormat = "MMM yyyy"
        Me.dtpFrom_II.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom_II.Location = New System.Drawing.Point(42, 19)
        Me.dtpFrom_II.Name = "dtpFrom_II"
        Me.dtpFrom_II.ShowUpDown = True
        Me.dtpFrom_II.Size = New System.Drawing.Size(111, 20)
        Me.dtpFrom_II.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "From"
        '
        'dtpTo_II
        '
        Me.dtpTo_II.CustomFormat = "MMM yyyy"
        Me.dtpTo_II.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo_II.Location = New System.Drawing.Point(55, 73)
        Me.dtpTo_II.Name = "dtpTo_II"
        Me.dtpTo_II.ShowUpDown = True
        Me.dtpTo_II.Size = New System.Drawing.Size(120, 20)
        Me.dtpTo_II.TabIndex = 2
        '
        'LblPeriodII
        '
        Me.LblPeriodII.AutoSize = True
        Me.LblPeriodII.Location = New System.Drawing.Point(6, 78)
        Me.LblPeriodII.Name = "LblPeriodII"
        Me.LblPeriodII.Size = New System.Drawing.Size(46, 13)
        Me.LblPeriodII.TabIndex = 32
        Me.LblPeriodII.Text = "Period II"
        '
        'chkHideZeroInactiveAccount
        '
        Me.chkHideZeroInactiveAccount.AutoSize = True
        Me.chkHideZeroInactiveAccount.Location = New System.Drawing.Point(9, 127)
        Me.chkHideZeroInactiveAccount.Name = "chkHideZeroInactiveAccount"
        Me.chkHideZeroInactiveAccount.Size = New System.Drawing.Size(153, 17)
        Me.chkHideZeroInactiveAccount.TabIndex = 4
        Me.chkHideZeroInactiveAccount.Text = "Hide zero inactive account"
        Me.chkHideZeroInactiveAccount.UseVisualStyleBackColor = True
        '
        'chkHideZero
        '
        Me.chkHideZero.AutoSize = True
        Me.chkHideZero.Location = New System.Drawing.Point(9, 150)
        Me.chkHideZero.Name = "chkHideZero"
        Me.chkHideZero.Size = New System.Drawing.Size(151, 17)
        Me.chkHideZero.TabIndex = 5
        Me.chkHideZero.Text = "Hide zero amount account"
        Me.chkHideZero.UseVisualStyleBackColor = True
        '
        'chkShowLvl2
        '
        Me.chkShowLvl2.AutoSize = True
        Me.chkShowLvl2.Checked = True
        Me.chkShowLvl2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl2.Location = New System.Drawing.Point(9, 192)
        Me.chkShowLvl2.Name = "chkShowLvl2"
        Me.chkShowLvl2.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl2.TabIndex = 7
        Me.chkShowLvl2.Text = "Show level 2 accounts"
        Me.chkShowLvl2.UseVisualStyleBackColor = True
        '
        'chkShowLvl3
        '
        Me.chkShowLvl3.AutoSize = True
        Me.chkShowLvl3.Checked = True
        Me.chkShowLvl3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl3.Location = New System.Drawing.Point(9, 215)
        Me.chkShowLvl3.Name = "chkShowLvl3"
        Me.chkShowLvl3.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl3.TabIndex = 8
        Me.chkShowLvl3.Text = "Show level 3 accounts"
        Me.chkShowLvl3.UseVisualStyleBackColor = True
        '
        'chkShowLvl4
        '
        Me.chkShowLvl4.AutoSize = True
        Me.chkShowLvl4.Checked = True
        Me.chkShowLvl4.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLvl4.Location = New System.Drawing.Point(9, 238)
        Me.chkShowLvl4.Name = "chkShowLvl4"
        Me.chkShowLvl4.Size = New System.Drawing.Size(134, 17)
        Me.chkShowLvl4.TabIndex = 9
        Me.chkShowLvl4.Text = "Show level 4 accounts"
        Me.chkShowLvl4.UseVisualStyleBackColor = True
        '
        'chkUseTotal
        '
        Me.chkUseTotal.AutoSize = True
        Me.chkUseTotal.Location = New System.Drawing.Point(9, 106)
        Me.chkUseTotal.Name = "chkUseTotal"
        Me.chkUseTotal.Size = New System.Drawing.Size(72, 17)
        Me.chkUseTotal.TabIndex = 3
        Me.chkUseTotal.Text = "Use Total"
        Me.chkUseTotal.UseVisualStyleBackColor = True
        '
        'chkShowAccountCode
        '
        Me.chkShowAccountCode.AutoSize = True
        Me.chkShowAccountCode.Checked = True
        Me.chkShowAccountCode.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowAccountCode.Location = New System.Drawing.Point(9, 170)
        Me.chkShowAccountCode.Name = "chkShowAccountCode"
        Me.chkShowAccountCode.Size = New System.Drawing.Size(124, 17)
        Me.chkShowAccountCode.TabIndex = 6
        Me.chkShowAccountCode.Text = "Show Account Code"
        Me.chkShowAccountCode.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'LblJob
        '
        Me.LblJob.AutoSize = True
        Me.LblJob.Location = New System.Drawing.Point(441, 80)
        Me.LblJob.Name = "LblJob"
        Me.LblJob.Size = New System.Drawing.Size(24, 13)
        Me.LblJob.TabIndex = 66
        Me.LblJob.Text = "Job"
        '
        'LblDept
        '
        Me.LblDept.AutoSize = True
        Me.LblDept.Location = New System.Drawing.Point(441, 57)
        Me.LblDept.Name = "LblDept"
        Me.LblDept.Size = New System.Drawing.Size(68, 13)
        Me.LblDept.TabIndex = 64
        Me.LblDept.Text = "Departement"
        '
        'TxtPrimaryJob
        '
        Me.TxtPrimaryJob.Location = New System.Drawing.Point(489, 76)
        Me.TxtPrimaryJob.Name = "TxtPrimaryJob"
        Me.TxtPrimaryJob.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryJob.TabIndex = 67
        Me.TxtPrimaryJob.Text = "-1"
        Me.TxtPrimaryJob.Visible = False
        '
        'btnFindJob
        '
        Me.btnFindJob.Location = New System.Drawing.Point(641, 79)
        Me.btnFindJob.Name = "btnFindJob"
        Me.btnFindJob.Size = New System.Drawing.Size(37, 21)
        Me.btnFindJob.TabIndex = 63
        Me.btnFindJob.TabStop = False
        Me.btnFindJob.Text = "Find"
        Me.btnFindJob.UseVisualStyleBackColor = True
        '
        'TxtJobCode
        '
        Me.TxtJobCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtJobCode.Location = New System.Drawing.Point(515, 77)
        Me.TxtJobCode.Name = "TxtJobCode"
        Me.TxtJobCode.Size = New System.Drawing.Size(120, 20)
        Me.TxtJobCode.TabIndex = 62
        '
        'TxtPrimaryDept
        '
        Me.TxtPrimaryDept.Location = New System.Drawing.Point(489, 53)
        Me.TxtPrimaryDept.Name = "TxtPrimaryDept"
        Me.TxtPrimaryDept.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryDept.TabIndex = 65
        Me.TxtPrimaryDept.Text = "-1"
        Me.TxtPrimaryDept.Visible = False
        '
        'btnFindDept
        '
        Me.btnFindDept.Location = New System.Drawing.Point(641, 56)
        Me.btnFindDept.Name = "btnFindDept"
        Me.btnFindDept.Size = New System.Drawing.Size(37, 21)
        Me.btnFindDept.TabIndex = 61
        Me.btnFindDept.TabStop = False
        Me.btnFindDept.Text = "Find"
        Me.btnFindDept.UseVisualStyleBackColor = True
        '
        'TxtDeptCode
        '
        Me.TxtDeptCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDeptCode.Location = New System.Drawing.Point(515, 54)
        Me.TxtDeptCode.Name = "TxtDeptCode"
        Me.TxtDeptCode.Size = New System.Drawing.Size(120, 20)
        Me.TxtDeptCode.TabIndex = 60
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cboFormatLaporan)
        Me.GroupBox2.Controls.Add(Me.cboJenis)
        Me.GroupBox2.Controls.Add(Me.LblPeriodII)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.dtpTo_I)
        Me.GroupBox2.Controls.Add(Me.btnClose)
        Me.GroupBox2.Controls.Add(Me.btnShow)
        Me.GroupBox2.Controls.Add(Me.dtpTo_II)
        Me.GroupBox2.Controls.Add(Me.chkUseTotal)
        Me.GroupBox2.Controls.Add(Me.chkShowAccountCode)
        Me.GroupBox2.Controls.Add(Me.chkShowLvl4)
        Me.GroupBox2.Controls.Add(Me.chkHideZeroInactiveAccount)
        Me.GroupBox2.Controls.Add(Me.chkShowLvl3)
        Me.GroupBox2.Controls.Add(Me.chkHideZero)
        Me.GroupBox2.Controls.Add(Me.chkShowLvl2)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(267, 277)
        Me.GroupBox2.TabIndex = 71
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Filter By"
        '
        'FrmLaporanTrialBalance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(884, 512)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.LblJob)
        Me.Controls.Add(Me.LblDept)
        Me.Controls.Add(Me.TxtPrimaryJob)
        Me.Controls.Add(Me.btnFindJob)
        Me.Controls.Add(Me.TxtJobCode)
        Me.Controls.Add(Me.TxtPrimaryDept)
        Me.Controls.Add(Me.btnFindDept)
        Me.Controls.Add(Me.TxtDeptCode)
        Me.Controls.Add(Me.gboPeriod_I)
        Me.Controls.Add(Me.crv1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 550)
        Me.Name = "FrmLaporanTrialBalance"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Trial Balance"
        Me.gboPeriod_I.ResumeLayout(False)
        Me.gboPeriod_I.PerformLayout()
        Me.gboPeriod_II.ResumeLayout(False)
        Me.gboPeriod_II.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents cboFormatLaporan As System.Windows.Forms.ComboBox
    Friend WithEvents cboJenis As System.Windows.Forms.ComboBox
    Friend WithEvents dtpFrom_I As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblPeriod As System.Windows.Forms.Label
    Friend WithEvents dtpTo_I As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents gboPeriod_I As System.Windows.Forms.GroupBox
    Friend WithEvents gboPeriod_II As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFrom_II As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpTo_II As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LblPeriodII As System.Windows.Forms.Label
    Friend WithEvents chkHideZeroInactiveAccount As System.Windows.Forms.CheckBox
    Friend WithEvents chkHideZero As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl3 As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLvl4 As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseTotal As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowAccountCode As System.Windows.Forms.CheckBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents LblJob As System.Windows.Forms.Label
    Friend WithEvents LblDept As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryJob As System.Windows.Forms.TextBox
    Friend WithEvents btnFindJob As System.Windows.Forms.Button
    Friend WithEvents TxtJobCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryDept As System.Windows.Forms.TextBox
    Friend WithEvents btnFindDept As System.Windows.Forms.Button
    Friend WithEvents TxtDeptCode As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
End Class
