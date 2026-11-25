<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterTaxProductReference
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
        Me.BtnExit = New System.Windows.Forms.Button
        Me.BtnDelete = New System.Windows.Forms.Button
        Me.BtnEdit = New System.Windows.Forms.Button
        Me.BtnNew = New System.Windows.Forms.Button
        Me.BtnFind = New System.Windows.Forms.Button
        Me.GboDetail = New System.Windows.Forms.GroupBox
        Me.CboType = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.BtnOk = New System.Windows.Forms.Button
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.TslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.TslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.TslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.TslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.GboDetail.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnExit
        '
        Me.BtnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnExit.Location = New System.Drawing.Point(467, 236)
        Me.BtnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(99, 32)
        Me.BtnExit.TabIndex = 4
        Me.BtnExit.Text = "E&xit"
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BtnDelete
        '
        Me.BtnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnDelete.Location = New System.Drawing.Point(360, 236)
        Me.BtnDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(99, 32)
        Me.BtnDelete.TabIndex = 3
        Me.BtnDelete.Text = "&Delete"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'BtnEdit
        '
        Me.BtnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnEdit.Location = New System.Drawing.Point(253, 236)
        Me.BtnEdit.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(99, 32)
        Me.BtnEdit.TabIndex = 2
        Me.BtnEdit.Text = "&Edit"
        Me.BtnEdit.UseVisualStyleBackColor = True
        '
        'BtnNew
        '
        Me.BtnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnNew.Location = New System.Drawing.Point(146, 236)
        Me.BtnNew.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnNew.Name = "BtnNew"
        Me.BtnNew.Size = New System.Drawing.Size(99, 32)
        Me.BtnNew.TabIndex = 1
        Me.BtnNew.Text = "&New"
        Me.BtnNew.UseVisualStyleBackColor = True
        '
        'BtnFind
        '
        Me.BtnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnFind.Location = New System.Drawing.Point(39, 236)
        Me.BtnFind.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(99, 32)
        Me.BtnFind.TabIndex = 0
        Me.BtnFind.Text = "&Find"
        Me.BtnFind.UseVisualStyleBackColor = True
        '
        'GboDetail
        '
        Me.GboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GboDetail.Controls.Add(Me.CboType)
        Me.GboDetail.Controls.Add(Me.Label4)
        Me.GboDetail.Controls.Add(Me.TxtRemarks)
        Me.GboDetail.Controls.Add(Me.TxtReadOnly)
        Me.GboDetail.Controls.Add(Me.BtnCancel)
        Me.GboDetail.Controls.Add(Me.BtnOk)
        Me.GboDetail.Controls.Add(Me.TxtPrimary)
        Me.GboDetail.Controls.Add(Me.TxtDesc)
        Me.GboDetail.Controls.Add(Me.TxtCode)
        Me.GboDetail.Controls.Add(Me.Label3)
        Me.GboDetail.Controls.Add(Me.Label2)
        Me.GboDetail.Controls.Add(Me.Label1)
        Me.GboDetail.Location = New System.Drawing.Point(16, 5)
        Me.GboDetail.Margin = New System.Windows.Forms.Padding(4)
        Me.GboDetail.Name = "GboDetail"
        Me.GboDetail.Padding = New System.Windows.Forms.Padding(4)
        Me.GboDetail.Size = New System.Drawing.Size(561, 223)
        Me.GboDetail.TabIndex = 15
        Me.GboDetail.TabStop = False
        '
        'CboType
        '
        Me.CboType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboType.AutoCompleteCustomSource.AddRange(New String() {"Good", "Service"})
        Me.CboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboType.FormattingEnabled = True
        Me.CboType.Items.AddRange(New Object() {"Goods", "Services"})
        Me.CboType.Location = New System.Drawing.Point(112, 16)
        Me.CboType.Name = "CboType"
        Me.CboType.Size = New System.Drawing.Size(430, 24)
        Me.CboType.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 19)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 17)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Type"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(112, 111)
        Me.TxtRemarks.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Multiline = True
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(430, 53)
        Me.TxtRemarks.TabIndex = 4
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(177, 223)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(56, 22)
        Me.TxtReadOnly.TabIndex = 10
        Me.TxtReadOnly.Visible = False
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnCancel.Location = New System.Drawing.Point(293, 180)
        Me.BtnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(117, 32)
        Me.BtnCancel.TabIndex = 6
        Me.BtnCancel.Text = "&Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'BtnOk
        '
        Me.BtnOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnOk.Location = New System.Drawing.Point(157, 180)
        Me.BtnOk.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnOk.Name = "BtnOk"
        Me.BtnOk.Size = New System.Drawing.Size(117, 32)
        Me.BtnOk.TabIndex = 5
        Me.BtnOk.Text = "&OK"
        Me.BtnOk.UseVisualStyleBackColor = True
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(112, 47)
        Me.TxtPrimary.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(56, 22)
        Me.TxtPrimary.TabIndex = 7
        Me.TxtPrimary.Visible = False
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(112, 79)
        Me.TxtDesc.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtDesc.MaxLength = 1000
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(430, 22)
        Me.TxtDesc.TabIndex = 3
        '
        'TxtCode
        '
        Me.TxtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCode.Location = New System.Drawing.Point(112, 47)
        Me.TxtCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCode.MaxLength = 250
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(430, 22)
        Me.TxtCode.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 115)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 17)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Remarks"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 83)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(79, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Description"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 51)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Code"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TslCreatedBy, Me.TslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 286)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip2.Size = New System.Drawing.Size(604, 25)
        Me.StatusStrip2.TabIndex = 22
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'TslCreatedBy
        '
        Me.TslCreatedBy.Name = "TslCreatedBy"
        Me.TslCreatedBy.Size = New System.Drawing.Size(292, 20)
        Me.TslCreatedBy.Spring = True
        Me.TslCreatedBy.Text = "Created By:"
        '
        'TslCreatedDate
        '
        Me.TslCreatedDate.Name = "TslCreatedDate"
        Me.TslCreatedDate.Size = New System.Drawing.Size(292, 20)
        Me.TslCreatedDate.Spring = True
        Me.TslCreatedDate.Text = "Created Date:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TslModifiedBy, Me.TslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 311)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(604, 25)
        Me.StatusStrip1.TabIndex = 21
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'TslModifiedBy
        '
        Me.TslModifiedBy.Name = "TslModifiedBy"
        Me.TslModifiedBy.Size = New System.Drawing.Size(292, 20)
        Me.TslModifiedBy.Spring = True
        Me.TslModifiedBy.Text = "Modified By:"
        '
        'TslModifiedDate
        '
        Me.TslModifiedDate.Name = "TslModifiedDate"
        Me.TslModifiedDate.Size = New System.Drawing.Size(292, 20)
        Me.TslModifiedDate.Spring = True
        Me.TslModifiedDate.Text = "Modified Date:"
        '
        'FrmMasterTaxProductReference
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnExit
        Me.ClientSize = New System.Drawing.Size(604, 336)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.BtnDelete)
        Me.Controls.Add(Me.BtnEdit)
        Me.Controls.Add(Me.BtnNew)
        Me.Controls.Add(Me.BtnFind)
        Me.Controls.Add(Me.GboDetail)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(622, 383)
        Me.Name = "FrmMasterTaxProductReference"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Tax Product Reference"
        Me.GboDetail.ResumeLayout(False)
        Me.GboDetail.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents BtnNew As System.Windows.Forms.Button
    Friend WithEvents BtnFind As System.Windows.Forms.Button
    Friend WithEvents GboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents BtnOk As System.Windows.Forms.Button
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents TslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents TslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents CboType As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
End Class
