<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterSalesPersonGroup
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
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnFind = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtPrimarySalesPersonGroup = New System.Windows.Forms.TextBox
        Me.TxtSalesPersonGroupName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtSalesPersonGroupCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.gboDetail.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(319, 84)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(107, 32)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(483, 132)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(107, 32)
        Me.btnExit.TabIndex = 35
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(253, 132)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(107, 32)
        Me.btnEdit.TabIndex = 33
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(302, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(24, 132)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(107, 32)
        Me.btnFind.TabIndex = 31
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtPrimarySalesPersonGroup)
        Me.gboDetail.Controls.Add(Me.TxtSalesPersonGroupName)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.TxtSalesPersonGroupCode)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Location = New System.Drawing.Point(12, 2)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Size = New System.Drawing.Size(589, 123)
        Me.gboDetail.TabIndex = 36
        Me.gboDetail.TabStop = False
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(164, 84)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(107, 32)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtPrimarySalesPersonGroup
        '
        Me.TxtPrimarySalesPersonGroup.Location = New System.Drawing.Point(104, 23)
        Me.TxtPrimarySalesPersonGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimarySalesPersonGroup.Name = "TxtPrimarySalesPersonGroup"
        Me.TxtPrimarySalesPersonGroup.Size = New System.Drawing.Size(56, 22)
        Me.TxtPrimarySalesPersonGroup.TabIndex = 32
        Me.TxtPrimarySalesPersonGroup.TabStop = False
        Me.TxtPrimarySalesPersonGroup.Visible = False
        '
        'TxtSalesPersonGroupName
        '
        Me.TxtSalesPersonGroupName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSalesPersonGroupName.Location = New System.Drawing.Point(169, 55)
        Me.TxtSalesPersonGroupName.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSalesPersonGroupName.MaxLength = 250
        Me.TxtSalesPersonGroupName.Name = "TxtSalesPersonGroupName"
        Me.TxtSalesPersonGroupName.Size = New System.Drawing.Size(399, 22)
        Me.TxtSalesPersonGroupName.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 55)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 17)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Name"
        '
        'TxtSalesPersonGroupCode
        '
        Me.TxtSalesPersonGroupCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSalesPersonGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSalesPersonGroupCode.Location = New System.Drawing.Point(169, 23)
        Me.TxtSalesPersonGroupCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSalesPersonGroupCode.MaxLength = 250
        Me.TxtSalesPersonGroupCode.Name = "TxtSalesPersonGroupCode"
        Me.TxtSalesPersonGroupCode.Size = New System.Drawing.Size(399, 22)
        Me.TxtSalesPersonGroupCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 23)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 17)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Code"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(302, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 181)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(625, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 37
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(302, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 206)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip2.Size = New System.Drawing.Size(625, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 38
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(302, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(368, 132)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(107, 32)
        Me.btnDelete.TabIndex = 34
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(139, 132)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(107, 32)
        Me.btnNew.TabIndex = 32
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'FrmMasterSalesPersonGroup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(625, 231)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.MinimumSize = New System.Drawing.Size(633, 268)
        Me.Name = "FrmMasterSalesPersonGroup"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales Person Group"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtPrimarySalesPersonGroup As System.Windows.Forms.TextBox
    Friend WithEvents TxtSalesPersonGroupName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtSalesPersonGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
End Class
