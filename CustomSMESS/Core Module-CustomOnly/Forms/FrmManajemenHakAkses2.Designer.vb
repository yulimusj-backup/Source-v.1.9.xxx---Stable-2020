<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmManajemenHakAkses2
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
        Me.btnCancel = New System.Windows.Forms.Button
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TxtReadonly = New System.Windows.Forms.TextBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.trvAccess = New System.Windows.Forms.TreeView
        Me.TxtAccesDesc = New System.Windows.Forms.TextBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtPrimaryGroup = New System.Windows.Forms.TextBox
        Me.TxtDescription = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtGroupCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.StatusStrip2.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(472, 329)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 402)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(939, 22)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 32
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(462, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(462, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(603, 370)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(462, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TxtReadonly)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryGroup)
        Me.gboDetail.Controls.Add(Me.TxtDescription)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.TxtGroupCode)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Location = New System.Drawing.Point(7, 3)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(925, 361)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'TxtReadonly
        '
        Me.TxtReadonly.Location = New System.Drawing.Point(89, 48)
        Me.TxtReadonly.Name = "TxtReadonly"
        Me.TxtReadonly.ReadOnly = True
        Me.TxtReadonly.Size = New System.Drawing.Size(43, 20)
        Me.TxtReadonly.TabIndex = 29
        Me.TxtReadonly.TabStop = False
        Me.TxtReadonly.Visible = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer1.Location = New System.Drawing.Point(9, 110)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.trvAccess)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtAccesDesc)
        Me.SplitContainer1.Size = New System.Drawing.Size(901, 213)
        Me.SplitContainer1.SplitterDistance = 330
        Me.SplitContainer1.TabIndex = 2
        Me.SplitContainer1.TabStop = False
        '
        'trvAccess
        '
        Me.trvAccess.CheckBoxes = True
        Me.trvAccess.Dock = System.Windows.Forms.DockStyle.Fill
        Me.trvAccess.Location = New System.Drawing.Point(0, 0)
        Me.trvAccess.Name = "trvAccess"
        Me.trvAccess.Size = New System.Drawing.Size(328, 211)
        Me.trvAccess.TabIndex = 0
        '
        'TxtAccesDesc
        '
        Me.TxtAccesDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtAccesDesc.Location = New System.Drawing.Point(0, 0)
        Me.TxtAccesDesc.Multiline = True
        Me.TxtAccesDesc.Name = "TxtAccesDesc"
        Me.TxtAccesDesc.ReadOnly = True
        Me.TxtAccesDesc.Size = New System.Drawing.Size(565, 211)
        Me.TxtAccesDesc.TabIndex = 0
        Me.TxtAccesDesc.TabStop = False
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(372, 329)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 26)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtPrimaryGroup
        '
        Me.TxtPrimaryGroup.Location = New System.Drawing.Point(89, 19)
        Me.TxtPrimaryGroup.Name = "TxtPrimaryGroup"
        Me.TxtPrimaryGroup.Size = New System.Drawing.Size(43, 20)
        Me.TxtPrimaryGroup.TabIndex = 27
        Me.TxtPrimaryGroup.TabStop = False
        Me.TxtPrimaryGroup.Visible = False
        '
        'TxtDescription
        '
        Me.TxtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescription.Location = New System.Drawing.Point(138, 45)
        Me.TxtDescription.MaxLength = 150
        Me.TxtDescription.Multiline = True
        Me.TxtDescription.Name = "TxtDescription"
        Me.TxtDescription.Size = New System.Drawing.Size(772, 59)
        Me.TxtDescription.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "Description"
        '
        'TxtGroupCode
        '
        Me.TxtGroupCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtGroupCode.Location = New System.Drawing.Point(138, 19)
        Me.TxtGroupCode.MaxLength = 50
        Me.TxtGroupCode.Name = "TxtGroupCode"
        Me.TxtGroupCode.Size = New System.Drawing.Size(772, 20)
        Me.TxtGroupCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Group Code"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 424)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(939, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 31
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(462, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(431, 370)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(259, 370)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(517, 370)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(345, 370)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'FrmManajemenHakAkses2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(939, 446)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.MinimumSize = New System.Drawing.Size(467, 244)
        Me.Name = "FrmManajemenHakAkses2"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Access Management-User Group"
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryGroup As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents trvAccess As System.Windows.Forms.TreeView
    Friend WithEvents TxtAccesDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtReadonly As System.Windows.Forms.TextBox
End Class
