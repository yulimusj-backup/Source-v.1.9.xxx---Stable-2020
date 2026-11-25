<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterUnit
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
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TxtTaxUnitDesc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryTaxUnit = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.BtnFindTaxUnit = New System.Windows.Forms.Button
        Me.TxtTaxUnit = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtPrimaryUnit = New System.Windows.Forms.TextBox
        Me.TxtUnitName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtUnitCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TxtTaxUnitDesc)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryTaxUnit)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.BtnFindTaxUnit)
        Me.gboDetail.Controls.Add(Me.TxtTaxUnit)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryUnit)
        Me.gboDetail.Controls.Add(Me.TxtUnitName)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.TxtUnitCode)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Location = New System.Drawing.Point(12, 10)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Size = New System.Drawing.Size(586, 157)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'TxtTaxUnitDesc
        '
        Me.TxtTaxUnitDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTaxUnitDesc.Location = New System.Drawing.Point(388, 85)
        Me.TxtTaxUnitDesc.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtTaxUnitDesc.MaxLength = 250
        Me.TxtTaxUnitDesc.Name = "TxtTaxUnitDesc"
        Me.TxtTaxUnitDesc.ReadOnly = True
        Me.TxtTaxUnitDesc.Size = New System.Drawing.Size(177, 22)
        Me.TxtTaxUnitDesc.TabIndex = 37
        Me.TxtTaxUnitDesc.TabStop = False
        '
        'TxtPrimaryTaxUnit
        '
        Me.TxtPrimaryTaxUnit.Location = New System.Drawing.Point(100, 123)
        Me.TxtPrimaryTaxUnit.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryTaxUnit.Name = "TxtPrimaryTaxUnit"
        Me.TxtPrimaryTaxUnit.Size = New System.Drawing.Size(56, 22)
        Me.TxtPrimaryTaxUnit.TabIndex = 36
        Me.TxtPrimaryTaxUnit.TabStop = False
        Me.TxtPrimaryTaxUnit.Text = "-1"
        Me.TxtPrimaryTaxUnit.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 88)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(151, 17)
        Me.Label3.TabIndex = 35
        Me.Label3.Text = "Tax Measurement Ref."
        '
        'BtnFindTaxUnit
        '
        Me.BtnFindTaxUnit.Location = New System.Drawing.Point(328, 83)
        Me.BtnFindTaxUnit.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnFindTaxUnit.Name = "BtnFindTaxUnit"
        Me.BtnFindTaxUnit.Size = New System.Drawing.Size(52, 26)
        Me.BtnFindTaxUnit.TabIndex = 3
        Me.BtnFindTaxUnit.TabStop = False
        Me.BtnFindTaxUnit.Text = "..."
        Me.BtnFindTaxUnit.UseVisualStyleBackColor = True
        '
        'TxtTaxUnit
        '
        Me.TxtTaxUnit.Location = New System.Drawing.Point(170, 85)
        Me.TxtTaxUnit.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtTaxUnit.MaxLength = 250
        Me.TxtTaxUnit.Name = "TxtTaxUnit"
        Me.TxtTaxUnit.Size = New System.Drawing.Size(150, 22)
        Me.TxtTaxUnit.TabIndex = 2
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(309, 118)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(107, 32)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(170, 118)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(107, 32)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtPrimaryUnit
        '
        Me.TxtPrimaryUnit.Location = New System.Drawing.Point(104, 23)
        Me.TxtPrimaryUnit.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryUnit.Name = "TxtPrimaryUnit"
        Me.TxtPrimaryUnit.Size = New System.Drawing.Size(56, 22)
        Me.TxtPrimaryUnit.TabIndex = 32
        Me.TxtPrimaryUnit.TabStop = False
        Me.TxtPrimaryUnit.Visible = False
        '
        'TxtUnitName
        '
        Me.TxtUnitName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUnitName.Location = New System.Drawing.Point(170, 55)
        Me.TxtUnitName.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtUnitName.MaxLength = 100
        Me.TxtUnitName.Name = "TxtUnitName"
        Me.TxtUnitName.Size = New System.Drawing.Size(395, 22)
        Me.TxtUnitName.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 55)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(74, 17)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Unit Name"
        '
        'TxtUnitCode
        '
        Me.TxtUnitCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUnitCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtUnitCode.Location = New System.Drawing.Point(170, 23)
        Me.TxtUnitCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtUnitCode.MaxLength = 50
        Me.TxtUnitCode.Name = "TxtUnitCode"
        Me.TxtUnitCode.Size = New System.Drawing.Size(395, 22)
        Me.TxtUnitCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 23)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 17)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Unit Code"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(482, 174)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(107, 32)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(367, 174)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(107, 32)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(252, 174)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(107, 32)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(23, 174)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(107, 32)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(138, 174)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(107, 32)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 214)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip2.Size = New System.Drawing.Size(612, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 22
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(296, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(296, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 239)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(612, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 21
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(296, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(296, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'FrmMasterUnit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(612, 264)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.gboDetail)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(630, 264)
        Me.Name = "FrmMasterUnit"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Unit of Measurement Data"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TxtPrimaryUnit As System.Windows.Forms.TextBox
    Friend WithEvents TxtUnitName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtUnitCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BtnFindTaxUnit As System.Windows.Forms.Button
    Friend WithEvents TxtTaxUnit As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryTaxUnit As System.Windows.Forms.TextBox
    Friend WithEvents TxtTaxUnitDesc As System.Windows.Forms.TextBox
End Class
