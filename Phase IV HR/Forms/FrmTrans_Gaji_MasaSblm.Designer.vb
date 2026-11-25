<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_MasaSblm
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
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.btnEdit = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnFind = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TxtPrimaryEmployee = New System.Windows.Forms.TextBox
        Me.TxtEmployer = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtIncomeTaxPaid = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtNettIncome = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.DtpWorkPeriodStart = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.DtpWorkPeriodTo = New System.Windows.Forms.DateTimePicker
        Me.btnFindEmployee = New System.Windows.Forms.Button
        Me.TxtEmployeeCode = New System.Windows.Forms.TextBox
        Me.TxtEmployeeName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtA1FormNo = New System.Windows.Forms.TextBox
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(264, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 272)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(544, 22)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 30
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(264, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(264, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(274, 232)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 4
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 294)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(544, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 29
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(264, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(445, 232)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "A1 Form No"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(280, 190)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(164, 190)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 26)
        Me.btnOK.TabIndex = 12
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(78, 43)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(43, 20)
        Me.TxtPrimary.TabIndex = 4
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Visible = False
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(127, 165)
        Me.TxtRemarks.MaxLength = 100
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(392, 20)
        Me.TxtRemarks.TabIndex = 11
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(188, 232)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 3
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 165)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 13)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Remarks"
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(17, 232)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 1
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryEmployee)
        Me.gboDetail.Controls.Add(Me.TxtEmployer)
        Me.gboDetail.Controls.Add(Me.Label8)
        Me.gboDetail.Controls.Add(Me.TxtIncomeTaxPaid)
        Me.gboDetail.Controls.Add(Me.Label7)
        Me.gboDetail.Controls.Add(Me.TxtNettIncome)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnFindEmployee)
        Me.gboDetail.Controls.Add(Me.TxtEmployeeCode)
        Me.gboDetail.Controls.Add(Me.TxtEmployeeName)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.dtpDate)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtPrimary)
        Me.gboDetail.Controls.Add(Me.TxtRemarks)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.TxtA1FormNo)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Location = New System.Drawing.Point(9, 4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(523, 222)
        Me.gboDetail.TabIndex = 100
        Me.gboDetail.TabStop = False
        '
        'TxtPrimaryEmployee
        '
        Me.TxtPrimaryEmployee.Location = New System.Drawing.Point(78, 19)
        Me.TxtPrimaryEmployee.Name = "TxtPrimaryEmployee"
        Me.TxtPrimaryEmployee.Size = New System.Drawing.Size(43, 20)
        Me.TxtPrimaryEmployee.TabIndex = 0
        Me.TxtPrimaryEmployee.TabStop = False
        Me.TxtPrimaryEmployee.Visible = False
        '
        'TxtEmployer
        '
        Me.TxtEmployer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmployer.Location = New System.Drawing.Point(126, 91)
        Me.TxtEmployer.MaxLength = 100
        Me.TxtEmployer.Name = "TxtEmployer"
        Me.TxtEmployer.Size = New System.Drawing.Size(392, 20)
        Me.TxtEmployer.TabIndex = 8
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(5, 91)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(50, 13)
        Me.Label8.TabIndex = 45
        Me.Label8.Text = "Employer"
        '
        'TxtIncomeTaxPaid
        '
        Me.TxtIncomeTaxPaid.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtIncomeTaxPaid.Location = New System.Drawing.Point(127, 141)
        Me.TxtIncomeTaxPaid.MaxLength = 100
        Me.TxtIncomeTaxPaid.Name = "TxtIncomeTaxPaid"
        Me.TxtIncomeTaxPaid.Size = New System.Drawing.Size(392, 20)
        Me.TxtIncomeTaxPaid.TabIndex = 10
        Me.TxtIncomeTaxPaid.Text = "0"
        Me.TxtIncomeTaxPaid.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 141)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(87, 13)
        Me.Label7.TabIndex = 43
        Me.Label7.Text = "Income Tax Paid"
        '
        'TxtNettIncome
        '
        Me.TxtNettIncome.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNettIncome.Location = New System.Drawing.Point(126, 116)
        Me.TxtNettIncome.MaxLength = 100
        Me.TxtNettIncome.Name = "TxtNettIncome"
        Me.TxtNettIncome.Size = New System.Drawing.Size(392, 20)
        Me.TxtNettIncome.TabIndex = 9
        Me.TxtNettIncome.Text = "0"
        Me.TxtNettIncome.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(5, 116)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(65, 13)
        Me.Label6.TabIndex = 41
        Me.Label6.Text = "Nett Income"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(5, 69)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(66, 13)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "Work Period"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(127, 67)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpWorkPeriodStart)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.DtpWorkPeriodTo)
        Me.SplitContainer1.Size = New System.Drawing.Size(390, 19)
        Me.SplitContainer1.SplitterDistance = 192
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 5
        Me.SplitContainer1.TabStop = False
        '
        'DtpWorkPeriodStart
        '
        Me.DtpWorkPeriodStart.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpWorkPeriodStart.CustomFormat = "MMM yyyy"
        Me.DtpWorkPeriodStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpWorkPeriodStart.Location = New System.Drawing.Point(0, 0)
        Me.DtpWorkPeriodStart.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpWorkPeriodStart.Name = "DtpWorkPeriodStart"
        Me.DtpWorkPeriodStart.ShowUpDown = True
        Me.DtpWorkPeriodStart.Size = New System.Drawing.Size(190, 20)
        Me.DtpWorkPeriodStart.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 2)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(20, 13)
        Me.Label5.TabIndex = 41
        Me.Label5.Text = "To"
        '
        'DtpWorkPeriodTo
        '
        Me.DtpWorkPeriodTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpWorkPeriodTo.CustomFormat = "MMM yyyy"
        Me.DtpWorkPeriodTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpWorkPeriodTo.Location = New System.Drawing.Point(28, 0)
        Me.DtpWorkPeriodTo.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpWorkPeriodTo.Name = "DtpWorkPeriodTo"
        Me.DtpWorkPeriodTo.ShowUpDown = True
        Me.DtpWorkPeriodTo.Size = New System.Drawing.Size(167, 20)
        Me.DtpWorkPeriodTo.TabIndex = 7
        '
        'btnFindEmployee
        '
        Me.btnFindEmployee.Location = New System.Drawing.Point(211, 18)
        Me.btnFindEmployee.Name = "btnFindEmployee"
        Me.btnFindEmployee.Size = New System.Drawing.Size(50, 20)
        Me.btnFindEmployee.TabIndex = 1
        Me.btnFindEmployee.TabStop = False
        Me.btnFindEmployee.Text = "Find"
        Me.btnFindEmployee.UseVisualStyleBackColor = True
        '
        'TxtEmployeeCode
        '
        Me.TxtEmployeeCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtEmployeeCode.Location = New System.Drawing.Point(127, 19)
        Me.TxtEmployeeCode.MaxLength = 50
        Me.TxtEmployeeCode.Name = "TxtEmployeeCode"
        Me.TxtEmployeeCode.Size = New System.Drawing.Size(79, 20)
        Me.TxtEmployeeCode.TabIndex = 0
        '
        'TxtEmployeeName
        '
        Me.TxtEmployeeName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmployeeName.Location = New System.Drawing.Point(267, 19)
        Me.TxtEmployeeName.MaxLength = 100
        Me.TxtEmployeeName.Name = "TxtEmployeeName"
        Me.TxtEmployeeName.ReadOnly = True
        Me.TxtEmployeeName.Size = New System.Drawing.Size(251, 20)
        Me.TxtEmployeeName.TabIndex = 2
        Me.TxtEmployeeName.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(5, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 13)
        Me.Label3.TabIndex = 35
        Me.Label3.Text = "Employee"
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(368, 43)
        Me.dtpDate.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(151, 20)
        Me.dtpDate.TabIndex = 4
        '
        'TxtA1FormNo
        '
        Me.TxtA1FormNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtA1FormNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtA1FormNo.Location = New System.Drawing.Point(127, 43)
        Me.TxtA1FormNo.MaxLength = 50
        Me.TxtA1FormNo.Name = "TxtA1FormNo"
        Me.TxtA1FormNo.Size = New System.Drawing.Size(236, 20)
        Me.TxtA1FormNo.TabIndex = 3
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(103, 232)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 2
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(359, 232)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 6
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'FrmTrans_Gaji_MasaSblm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(544, 316)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnNew)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(560, 354)
        Me.Name = "FrmTrans_Gaji_MasaSblm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Previous Period Income"
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents DtpWorkPeriodStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnFindEmployee As System.Windows.Forms.Button
    Friend WithEvents TxtEmployeeCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtA1FormNo As System.Windows.Forms.TextBox
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents TxtIncomeTaxPaid As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtNettIncome As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents DtpWorkPeriodTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtEmployer As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryEmployee As System.Windows.Forms.TextBox
End Class
