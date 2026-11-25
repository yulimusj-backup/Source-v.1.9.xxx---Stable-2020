<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterTaxFacilityStampReference
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
        Me.BtnExit.Location = New System.Drawing.Point(350, 214)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(74, 26)
        Me.BtnExit.TabIndex = 4
        Me.BtnExit.Text = "E&xit"
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BtnDelete
        '
        Me.BtnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnDelete.Location = New System.Drawing.Point(270, 214)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(74, 26)
        Me.BtnDelete.TabIndex = 3
        Me.BtnDelete.Text = "&Delete"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'BtnEdit
        '
        Me.BtnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnEdit.Location = New System.Drawing.Point(190, 214)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(74, 26)
        Me.BtnEdit.TabIndex = 2
        Me.BtnEdit.Text = "&Edit"
        Me.BtnEdit.UseVisualStyleBackColor = True
        '
        'BtnNew
        '
        Me.BtnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnNew.Location = New System.Drawing.Point(110, 214)
        Me.BtnNew.Name = "BtnNew"
        Me.BtnNew.Size = New System.Drawing.Size(74, 26)
        Me.BtnNew.TabIndex = 1
        Me.BtnNew.Text = "&New"
        Me.BtnNew.UseVisualStyleBackColor = True
        '
        'BtnFind
        '
        Me.BtnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnFind.Location = New System.Drawing.Point(29, 214)
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(74, 26)
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
        Me.GboDetail.Location = New System.Drawing.Point(12, 4)
        Me.GboDetail.Name = "GboDetail"
        Me.GboDetail.Size = New System.Drawing.Size(421, 194)
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
        Me.CboType.Items.AddRange(New Object() {"07", "08"})
        Me.CboType.Location = New System.Drawing.Point(106, 18)
        Me.CboType.Margin = New System.Windows.Forms.Padding(2)
        Me.CboType.Name = "CboType"
        Me.CboType.Size = New System.Drawing.Size(304, 21)
        Me.CboType.TabIndex = 12
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(91, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Transaction Code"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(106, 96)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Multiline = True
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(305, 61)
        Me.TxtRemarks.TabIndex = 3
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(133, 193)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(43, 20)
        Me.TxtReadOnly.TabIndex = 10
        Me.TxtReadOnly.Visible = False
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnCancel.Location = New System.Drawing.Point(218, 162)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(88, 26)
        Me.BtnCancel.TabIndex = 5
        Me.BtnCancel.Text = "&Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'BtnOk
        '
        Me.BtnOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnOk.Location = New System.Drawing.Point(116, 162)
        Me.BtnOk.Name = "BtnOk"
        Me.BtnOk.Size = New System.Drawing.Size(88, 26)
        Me.BtnOk.TabIndex = 4
        Me.BtnOk.Text = "&OK"
        Me.BtnOk.UseVisualStyleBackColor = True
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(106, 44)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(43, 20)
        Me.TxtPrimary.TabIndex = 0
        Me.TxtPrimary.Visible = False
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(106, 70)
        Me.TxtDesc.MaxLength = 250
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(305, 20)
        Me.TxtDesc.TabIndex = 2
        '
        'TxtCode
        '
        Me.TxtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCode.Location = New System.Drawing.Point(106, 44)
        Me.TxtCode.MaxLength = 250
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(305, 20)
        Me.TxtCode.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Remarks"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Description"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 47)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Code"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TslCreatedBy, Me.TslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 267)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(460, 22)
        Me.StatusStrip2.TabIndex = 22
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'TslCreatedBy
        '
        Me.TslCreatedBy.Name = "TslCreatedBy"
        Me.TslCreatedBy.Size = New System.Drawing.Size(222, 17)
        Me.TslCreatedBy.Spring = True
        Me.TslCreatedBy.Text = "Created By:"
        '
        'TslCreatedDate
        '
        Me.TslCreatedDate.Name = "TslCreatedDate"
        Me.TslCreatedDate.Size = New System.Drawing.Size(222, 17)
        Me.TslCreatedDate.Spring = True
        Me.TslCreatedDate.Text = "Created Date:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TslModifiedBy, Me.TslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 289)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(460, 22)
        Me.StatusStrip1.TabIndex = 21
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'TslModifiedBy
        '
        Me.TslModifiedBy.Name = "TslModifiedBy"
        Me.TslModifiedBy.Size = New System.Drawing.Size(222, 17)
        Me.TslModifiedBy.Spring = True
        Me.TslModifiedBy.Text = "Modified By:"
        '
        'TslModifiedDate
        '
        Me.TslModifiedDate.Name = "TslModifiedDate"
        Me.TslModifiedDate.Size = New System.Drawing.Size(222, 17)
        Me.TslModifiedDate.Spring = True
        Me.TslModifiedDate.Text = "Modified Date:"
        '
        'FrmMasterTaxFacilityStampReference
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnExit
        Me.ClientSize = New System.Drawing.Size(460, 311)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.BtnDelete)
        Me.Controls.Add(Me.BtnEdit)
        Me.Controls.Add(Me.BtnNew)
        Me.Controls.Add(Me.BtnFind)
        Me.Controls.Add(Me.GboDetail)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(476, 350)
        Me.Name = "FrmMasterTaxFacilityStampReference"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Tax Facility Stamp Reference"
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
