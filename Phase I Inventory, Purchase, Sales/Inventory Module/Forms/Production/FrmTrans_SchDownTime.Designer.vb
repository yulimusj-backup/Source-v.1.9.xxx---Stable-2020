<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SchDownTime
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
        Me.components = New System.ComponentModel.Container
        Me.txtReason = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtDocumentNo = New System.Windows.Forms.TextBox
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.txtReadonly = New System.Windows.Forms.TextBox
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.colWU_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_StartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_StartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_EndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_EndTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LlbWorkUnit = New System.Windows.Forms.LinkLabel
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ApplyToAllWokUnitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ApplyToWorkUnitGroupToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtReason
        '
        Me.txtReason.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReason.BackColor = System.Drawing.SystemColors.Window
        Me.txtReason.Location = New System.Drawing.Point(132, 30)
        Me.txtReason.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtReason.MaxLength = 250
        Me.txtReason.Multiline = True
        Me.txtReason.Name = "txtReason"
        Me.txtReason.Size = New System.Drawing.Size(334, 66)
        Me.txtReason.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(4, 30)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(57, 17)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Reason"
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(132, 0)
        Me.txtPrimary.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(37, 22)
        Me.txtPrimary.TabIndex = 0
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(332, 0)
        Me.dtpDate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(133, 22)
        Me.dtpDate.TabIndex = 1
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(8, 12)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtReason)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtDocumentNo)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer1.Size = New System.Drawing.Size(944, 100)
        Me.SplitContainer1.SplitterDistance = 470
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 99999999
        Me.SplitContainer1.TabStop = False
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(284, 5)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 17)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Date"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 5)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 17)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Document No"
        '
        'txtDocumentNo
        '
        Me.txtDocumentNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDocumentNo.BackColor = System.Drawing.SystemColors.Control
        Me.txtDocumentNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDocumentNo.Location = New System.Drawing.Point(132, 0)
        Me.txtDocumentNo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtDocumentNo.MaxLength = 250
        Me.txtDocumentNo.Name = "txtDocumentNo"
        Me.txtDocumentNo.ReadOnly = True
        Me.txtDocumentNo.Size = New System.Drawing.Size(143, 22)
        Me.txtDocumentNo.TabIndex = 0
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.Location = New System.Drawing.Point(77, 0)
        Me.txtRemarks.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtRemarks.MaxLength = 250
        Me.txtRemarks.Multiline = True
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(380, 96)
        Me.txtRemarks.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 5)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 17)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Remarks"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(245, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'txtReadonly
        '
        Me.txtReadonly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReadonly.Location = New System.Drawing.Point(932, 546)
        Me.txtReadonly.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtReadonly.Name = "txtReadonly"
        Me.txtReadonly.ReadOnly = True
        Me.txtReadonly.Size = New System.Drawing.Size(37, 22)
        Me.txtReadonly.TabIndex = 30
        Me.txtReadonly.Visible = False
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(245, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(245, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(245, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(484, 489)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(135, 32)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(341, 489)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(135, 32)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 583)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(1003, 25)
        Me.StatusStrip1.TabIndex = 31
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(673, 535)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(107, 32)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(443, 535)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(107, 32)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(213, 535)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(107, 32)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(328, 535)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(107, 32)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(557, 535)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(107, 32)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.LlbWorkUnit)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(21, 1)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Size = New System.Drawing.Size(960, 528)
        Me.gboDetail.TabIndex = 23
        Me.gboDetail.TabStop = False
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colWU_No, Me.colWU_Code, Me.colWU_Desc, Me.colWU_StartDate, Me.colWU_StartTime, Me.colWU_EndDate, Me.colWU_EndTime, Me.colWU_Remarks, Me.colWU_Primary})
        Me.dgvItem.Location = New System.Drawing.Point(7, 116)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(948, 366)
        Me.dgvItem.TabIndex = 0
        '
        'colWU_No
        '
        Me.colWU_No.Frozen = True
        Me.colWU_No.HeaderText = "No"
        Me.colWU_No.Name = "colWU_No"
        Me.colWU_No.ReadOnly = True
        Me.colWU_No.Width = 51
        '
        'colWU_Code
        '
        Me.colWU_Code.Frozen = True
        Me.colWU_Code.HeaderText = "Work Unit"
        Me.colWU_Code.MaxInputLength = 250
        Me.colWU_Code.Name = "colWU_Code"
        Me.colWU_Code.Width = 95
        '
        'colWU_Desc
        '
        Me.colWU_Desc.HeaderText = "Desc"
        Me.colWU_Desc.Name = "colWU_Desc"
        Me.colWU_Desc.ReadOnly = True
        Me.colWU_Desc.Width = 65
        '
        'colWU_StartDate
        '
        Me.colWU_StartDate.HeaderText = "Start Date"
        Me.colWU_StartDate.Name = "colWU_StartDate"
        Me.colWU_StartDate.Width = 97
        '
        'colWU_StartTime
        '
        Me.colWU_StartTime.HeaderText = "Start Time"
        Me.colWU_StartTime.Name = "colWU_StartTime"
        Me.colWU_StartTime.Width = 98
        '
        'colWU_EndDate
        '
        Me.colWU_EndDate.HeaderText = "End Date"
        Me.colWU_EndDate.Name = "colWU_EndDate"
        Me.colWU_EndDate.Width = 92
        '
        'colWU_EndTime
        '
        Me.colWU_EndTime.HeaderText = "End Time"
        Me.colWU_EndTime.Name = "colWU_EndTime"
        Me.colWU_EndTime.Width = 93
        '
        'colWU_Remarks
        '
        Me.colWU_Remarks.HeaderText = "Remarks"
        Me.colWU_Remarks.Name = "colWU_Remarks"
        Me.colWU_Remarks.Width = 89
        '
        'colWU_Primary
        '
        Me.colWU_Primary.HeaderText = "Primary Work Unit"
        Me.colWU_Primary.Name = "colWU_Primary"
        Me.colWU_Primary.ReadOnly = True
        Me.colWU_Primary.Visible = False
        Me.colWU_Primary.Width = 117
        '
        'LlbWorkUnit
        '
        Me.LlbWorkUnit.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LlbWorkUnit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbWorkUnit.AutoSize = True
        Me.LlbWorkUnit.Location = New System.Drawing.Point(889, 485)
        Me.LlbWorkUnit.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LlbWorkUnit.Name = "LlbWorkUnit"
        Me.LlbWorkUnit.Size = New System.Drawing.Size(63, 17)
        Me.LlbWorkUnit.TabIndex = 33
        Me.LlbWorkUnit.TabStop = True
        Me.LlbWorkUnit.Text = "Advance"
        Me.LlbWorkUnit.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'BackgroundWorker1
        '
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ApplyToAllWokUnitToolStripMenuItem, Me.ApplyToWorkUnitGroupToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(251, 52)
        '
        'ApplyToAllWokUnitToolStripMenuItem
        '
        Me.ApplyToAllWokUnitToolStripMenuItem.Name = "ApplyToAllWokUnitToolStripMenuItem"
        Me.ApplyToAllWokUnitToolStripMenuItem.Size = New System.Drawing.Size(250, 24)
        Me.ApplyToAllWokUnitToolStripMenuItem.Text = "Apply to All Work Unit"
        '
        'ApplyToWorkUnitGroupToolStripMenuItem
        '
        Me.ApplyToWorkUnitGroupToolStripMenuItem.Name = "ApplyToWorkUnitGroupToolStripMenuItem"
        Me.ApplyToWorkUnitGroupToolStripMenuItem.Size = New System.Drawing.Size(250, 24)
        Me.ApplyToWorkUnitGroupToolStripMenuItem.Text = "Apply to Work Unit Group"
        '
        'BackgroundWorker2
        '
        '
        'FrmTrans_SchDownTime
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1003, 608)
        Me.Controls.Add(Me.txtReadonly)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.gboDetail)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimumSize = New System.Drawing.Size(1011, 645)
        Me.Name = "FrmTrans_SchDownTime"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Scheduled Down Time"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtReason As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents txtDocumentNo As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents txtReadonly As System.Windows.Forms.TextBox
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents LlbWorkUnit As System.Windows.Forms.LinkLabel
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ApplyToAllWokUnitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ApplyToWorkUnitGroupToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents colWU_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_StartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_StartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_EndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_EndTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
