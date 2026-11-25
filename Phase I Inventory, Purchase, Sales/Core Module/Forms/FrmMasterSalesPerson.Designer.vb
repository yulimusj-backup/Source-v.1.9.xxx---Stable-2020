<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterSalesPerson
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
        Me.TxtPrimarySalesPerson = New System.Windows.Forms.TextBox
        Me.TxtSalesPersonName = New System.Windows.Forms.TextBox
        Me.btnEdit = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnFind = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TxtPrimarySalesPersonGroup = New System.Windows.Forms.TextBox
        Me.btnFindSalesPersonGroup = New System.Windows.Forms.Button
        Me.TxtSalesPersonGroup = New System.Windows.Forms.TextBox
        Me.TxtSalesPersonGroupDesc = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtSalesPersonCode = New System.Windows.Forms.TextBox
        Me.btnNew = New System.Windows.Forms.Button
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(302, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 211)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip2.Size = New System.Drawing.Size(625, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 30
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(302, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
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
        Me.btnDelete.Location = New System.Drawing.Point(368, 165)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(107, 32)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 236)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(625, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 29
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(302, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(483, 165)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(107, 32)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
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
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(319, 114)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(107, 32)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(164, 114)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(107, 32)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtPrimarySalesPerson
        '
        Me.TxtPrimarySalesPerson.Location = New System.Drawing.Point(104, 23)
        Me.TxtPrimarySalesPerson.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimarySalesPerson.Name = "TxtPrimarySalesPerson"
        Me.TxtPrimarySalesPerson.Size = New System.Drawing.Size(56, 22)
        Me.TxtPrimarySalesPerson.TabIndex = 32
        Me.TxtPrimarySalesPerson.TabStop = False
        Me.TxtPrimarySalesPerson.Visible = False
        '
        'TxtSalesPersonName
        '
        Me.TxtSalesPersonName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSalesPersonName.Location = New System.Drawing.Point(146, 55)
        Me.TxtSalesPersonName.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSalesPersonName.MaxLength = 250
        Me.TxtSalesPersonName.Name = "TxtSalesPersonName"
        Me.TxtSalesPersonName.Size = New System.Drawing.Size(422, 22)
        Me.TxtSalesPersonName.TabIndex = 1
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(253, 165)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(107, 32)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
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
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(24, 165)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(107, 32)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TxtPrimarySalesPersonGroup)
        Me.gboDetail.Controls.Add(Me.btnFindSalesPersonGroup)
        Me.gboDetail.Controls.Add(Me.TxtSalesPersonGroup)
        Me.gboDetail.Controls.Add(Me.TxtSalesPersonGroupDesc)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtPrimarySalesPerson)
        Me.gboDetail.Controls.Add(Me.TxtSalesPersonName)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.TxtSalesPersonCode)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Location = New System.Drawing.Point(12, 5)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Size = New System.Drawing.Size(589, 153)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'TxtPrimarySalesPersonGroup
        '
        Me.TxtPrimarySalesPersonGroup.Location = New System.Drawing.Point(82, 85)
        Me.TxtPrimarySalesPersonGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimarySalesPersonGroup.Name = "TxtPrimarySalesPersonGroup"
        Me.TxtPrimarySalesPersonGroup.Size = New System.Drawing.Size(56, 22)
        Me.TxtPrimarySalesPersonGroup.TabIndex = 35
        Me.TxtPrimarySalesPersonGroup.TabStop = False
        Me.TxtPrimarySalesPersonGroup.Visible = False
        '
        'btnFindSalesPersonGroup
        '
        Me.btnFindSalesPersonGroup.Location = New System.Drawing.Point(255, 85)
        Me.btnFindSalesPersonGroup.Name = "btnFindSalesPersonGroup"
        Me.btnFindSalesPersonGroup.Size = New System.Drawing.Size(57, 23)
        Me.btnFindSalesPersonGroup.TabIndex = 3
        Me.btnFindSalesPersonGroup.TabStop = False
        Me.btnFindSalesPersonGroup.Text = "Find"
        Me.btnFindSalesPersonGroup.UseVisualStyleBackColor = True
        '
        'TxtSalesPersonGroup
        '
        Me.TxtSalesPersonGroup.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSalesPersonGroup.Location = New System.Drawing.Point(146, 85)
        Me.TxtSalesPersonGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSalesPersonGroup.MaxLength = 250
        Me.TxtSalesPersonGroup.Name = "TxtSalesPersonGroup"
        Me.TxtSalesPersonGroup.Size = New System.Drawing.Size(102, 22)
        Me.TxtSalesPersonGroup.TabIndex = 2
        '
        'TxtSalesPersonGroupDesc
        '
        Me.TxtSalesPersonGroupDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSalesPersonGroupDesc.Location = New System.Drawing.Point(319, 85)
        Me.TxtSalesPersonGroupDesc.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSalesPersonGroupDesc.MaxLength = 250
        Me.TxtSalesPersonGroupDesc.Name = "TxtSalesPersonGroupDesc"
        Me.TxtSalesPersonGroupDesc.ReadOnly = True
        Me.TxtSalesPersonGroupDesc.Size = New System.Drawing.Size(249, 22)
        Me.TxtSalesPersonGroupDesc.TabIndex = 4
        Me.TxtSalesPersonGroupDesc.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 85)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 17)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Group"
        '
        'TxtSalesPersonCode
        '
        Me.TxtSalesPersonCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSalesPersonCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSalesPersonCode.Location = New System.Drawing.Point(146, 23)
        Me.TxtSalesPersonCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSalesPersonCode.MaxLength = 250
        Me.TxtSalesPersonCode.Name = "TxtSalesPersonCode"
        Me.TxtSalesPersonCode.Size = New System.Drawing.Size(422, 22)
        Me.TxtSalesPersonCode.TabIndex = 0
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(139, 165)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(107, 32)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'FrmMasterSalesPerson
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(625, 261)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnNew)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimumSize = New System.Drawing.Size(633, 298)
        Me.Name = "FrmMasterSalesPerson"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales Person Data"
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
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
    Friend WithEvents TxtPrimarySalesPerson As System.Windows.Forms.TextBox
    Friend WithEvents TxtSalesPersonName As System.Windows.Forms.TextBox
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TxtSalesPersonCode As System.Windows.Forms.TextBox
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFindSalesPersonGroup As System.Windows.Forms.Button
    Friend WithEvents TxtSalesPersonGroup As System.Windows.Forms.TextBox
    Friend WithEvents TxtSalesPersonGroupDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimarySalesPersonGroup As System.Windows.Forms.TextBox
End Class
