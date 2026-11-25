<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMast_PayOutDropCash
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
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnEdit = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.btnFind = New System.Windows.Forms.Button
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.btnFindAccount = New System.Windows.Forms.Button
        Me.TxtPrimaryAccountCost = New System.Windows.Forms.TextBox
        Me.TxtAccountCostCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.CboTipe = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.LblAccount = New System.Windows.Forms.Label
        Me.btnDelete = New System.Windows.Forms.Button
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(111, 109)
        Me.TxtRemarks.MaxLength = 150
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(433, 20)
        Me.TxtRemarks.TabIndex = 4
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(288, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(240, 177)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(110, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(15, 109)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(49, 13)
        Me.Label7.TabIndex = 35
        Me.Label7.Text = "Remarks"
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(110, 30)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimary.TabIndex = 23
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(8, 177)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(110, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(288, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(455, 135)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(89, 20)
        Me.TxtReadOnly.TabIndex = 6
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(111, 56)
        Me.TxtDesc.MaxLength = 150
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(433, 20)
        Me.TxtDesc.TabIndex = 1
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 224)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(591, 22)
        Me.StatusStrip2.TabIndex = 46
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(472, 177)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(110, 26)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(124, 177)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(110, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(288, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 246)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(591, 22)
        Me.StatusStrip1.TabIndex = 45
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(288, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.btnFindAccount)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryAccountCost)
        Me.gboDetail.Controls.Add(Me.TxtAccountCostCode)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.CboTipe)
        Me.gboDetail.Controls.Add(Me.TxtReadOnly)
        Me.gboDetail.Controls.Add(Me.TxtRemarks)
        Me.gboDetail.Controls.Add(Me.Label7)
        Me.gboDetail.Controls.Add(Me.TxtDesc)
        Me.gboDetail.Controls.Add(Me.TxtPrimary)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtCode)
        Me.gboDetail.Controls.Add(Me.LblAccount)
        Me.gboDetail.Location = New System.Drawing.Point(17, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(559, 171)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'btnFindAccount
        '
        Me.btnFindAccount.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindAccount.Location = New System.Drawing.Point(508, 82)
        Me.btnFindAccount.Name = "btnFindAccount"
        Me.btnFindAccount.Size = New System.Drawing.Size(36, 20)
        Me.btnFindAccount.TabIndex = 4
        Me.btnFindAccount.TabStop = False
        Me.btnFindAccount.Text = "Find"
        Me.btnFindAccount.UseVisualStyleBackColor = True
        '
        'TxtPrimaryAccountCost
        '
        Me.TxtPrimaryAccountCost.Location = New System.Drawing.Point(374, 83)
        Me.TxtPrimaryAccountCost.Name = "TxtPrimaryAccountCost"
        Me.TxtPrimaryAccountCost.Size = New System.Drawing.Size(51, 20)
        Me.TxtPrimaryAccountCost.TabIndex = 40
        Me.TxtPrimaryAccountCost.TabStop = False
        Me.TxtPrimaryAccountCost.Text = "-1"
        Me.TxtPrimaryAccountCost.Visible = False
        '
        'TxtAccountCostCode
        '
        Me.TxtAccountCostCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccountCostCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAccountCostCode.Location = New System.Drawing.Point(374, 82)
        Me.TxtAccountCostCode.MaxLength = 150
        Me.TxtAccountCostCode.Name = "TxtAccountCostCode"
        Me.TxtAccountCostCode.Size = New System.Drawing.Size(128, 20)
        Me.TxtAccountCostCode.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(15, 85)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(31, 13)
        Me.Label3.TabIndex = 37
        Me.Label3.Text = "Type"
        '
        'CboTipe
        '
        Me.CboTipe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboTipe.FormattingEnabled = True
        Me.CboTipe.Items.AddRange(New Object() {"Payout", "Drop Cash"})
        Me.CboTipe.Location = New System.Drawing.Point(111, 82)
        Me.CboTipe.Name = "CboTipe"
        Me.CboTipe.Size = New System.Drawing.Size(165, 21)
        Me.CboTipe.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Description"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "Code"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(282, 139)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(110, 26)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(166, 139)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(110, 26)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtCode
        '
        Me.TxtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCode.Location = New System.Drawing.Point(111, 30)
        Me.TxtCode.MaxLength = 150
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(433, 20)
        Me.TxtCode.TabIndex = 0
        '
        'LblAccount
        '
        Me.LblAccount.AutoSize = True
        Me.LblAccount.Location = New System.Drawing.Point(282, 85)
        Me.LblAccount.Name = "LblAccount"
        Me.LblAccount.Size = New System.Drawing.Size(71, 13)
        Me.LblAccount.TabIndex = 39
        Me.LblAccount.Text = "Cost Account"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(356, 177)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(110, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'FrmMast_PayOutDropCash
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(591, 268)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnDelete)
        Me.MinimizeBox = False
        Me.Name = "FrmMast_PayOutDropCash"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Payout Or Drop Cash"
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CboTipe As System.Windows.Forms.ComboBox
    Friend WithEvents btnFindAccount As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryAccountCost As System.Windows.Forms.TextBox
    Friend WithEvents LblAccount As System.Windows.Forms.Label
    Friend WithEvents TxtAccountCostCode As System.Windows.Forms.TextBox
End Class
