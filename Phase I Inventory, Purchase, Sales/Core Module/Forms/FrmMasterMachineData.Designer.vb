<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterMachineData
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
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtMachineGroupDesc = New System.Windows.Forms.TextBox
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.BtnFindMachineGroup = New System.Windows.Forms.Button
        Me.TxtPrimaryMachineGroup = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtMachineGroupCode = New System.Windows.Forms.TextBox
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.BtnOk = New System.Windows.Forms.Button
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.ChkActive = New System.Windows.Forms.CheckBox
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.TxtMachineDataCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
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
        Me.BtnExit.Location = New System.Drawing.Point(414, 278)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(88, 26)
        Me.BtnExit.TabIndex = 20
        Me.BtnExit.Text = "E&xit"
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BtnDelete
        '
        Me.BtnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnDelete.Location = New System.Drawing.Point(320, 278)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(88, 26)
        Me.BtnDelete.TabIndex = 19
        Me.BtnDelete.Text = "&Delete"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'BtnEdit
        '
        Me.BtnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnEdit.Location = New System.Drawing.Point(226, 278)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(88, 26)
        Me.BtnEdit.TabIndex = 18
        Me.BtnEdit.Text = "&Edit"
        Me.BtnEdit.UseVisualStyleBackColor = True
        '
        'BtnNew
        '
        Me.BtnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnNew.Location = New System.Drawing.Point(132, 278)
        Me.BtnNew.Name = "BtnNew"
        Me.BtnNew.Size = New System.Drawing.Size(88, 26)
        Me.BtnNew.TabIndex = 17
        Me.BtnNew.Text = "&New"
        Me.BtnNew.UseVisualStyleBackColor = True
        '
        'BtnFind
        '
        Me.BtnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnFind.Location = New System.Drawing.Point(38, 278)
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(88, 26)
        Me.BtnFind.TabIndex = 16
        Me.BtnFind.Text = "&Find"
        Me.BtnFind.UseVisualStyleBackColor = True
        '
        'GboDetail
        '
        Me.GboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GboDetail.Controls.Add(Me.Label5)
        Me.GboDetail.Controls.Add(Me.TxtMachineGroupDesc)
        Me.GboDetail.Controls.Add(Me.TxtRemarks)
        Me.GboDetail.Controls.Add(Me.BtnFindMachineGroup)
        Me.GboDetail.Controls.Add(Me.TxtPrimaryMachineGroup)
        Me.GboDetail.Controls.Add(Me.Label4)
        Me.GboDetail.Controls.Add(Me.TxtMachineGroupCode)
        Me.GboDetail.Controls.Add(Me.BtnCancel)
        Me.GboDetail.Controls.Add(Me.BtnOk)
        Me.GboDetail.Controls.Add(Me.TxtPrimary)
        Me.GboDetail.Controls.Add(Me.ChkActive)
        Me.GboDetail.Controls.Add(Me.TxtDesc)
        Me.GboDetail.Controls.Add(Me.TxtMachineDataCode)
        Me.GboDetail.Controls.Add(Me.Label3)
        Me.GboDetail.Controls.Add(Me.Label2)
        Me.GboDetail.Controls.Add(Me.Label1)
        Me.GboDetail.Controls.Add(Me.TxtReadOnly)
        Me.GboDetail.Location = New System.Drawing.Point(12, 4)
        Me.GboDetail.Name = "GboDetail"
        Me.GboDetail.Size = New System.Drawing.Size(511, 267)
        Me.GboDetail.TabIndex = 15
        Me.GboDetail.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 126)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Remarks"
        '
        'TxtMachineGroupDesc
        '
        Me.TxtMachineGroupDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMachineGroupDesc.Location = New System.Drawing.Point(133, 97)
        Me.TxtMachineGroupDesc.Name = "TxtMachineGroupDesc"
        Me.TxtMachineGroupDesc.ReadOnly = True
        Me.TxtMachineGroupDesc.Size = New System.Drawing.Size(365, 20)
        Me.TxtMachineGroupDesc.TabIndex = 7
        Me.TxtMachineGroupDesc.TabStop = False
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(133, 123)
        Me.TxtRemarks.Multiline = True
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(365, 78)
        Me.TxtRemarks.TabIndex = 8
        '
        'BtnFindMachineGroup
        '
        Me.BtnFindMachineGroup.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindMachineGroup.Location = New System.Drawing.Point(460, 71)
        Me.BtnFindMachineGroup.Name = "BtnFindMachineGroup"
        Me.BtnFindMachineGroup.Size = New System.Drawing.Size(38, 20)
        Me.BtnFindMachineGroup.TabIndex = 6
        Me.BtnFindMachineGroup.TabStop = False
        Me.BtnFindMachineGroup.Text = "..."
        Me.BtnFindMachineGroup.UseVisualStyleBackColor = True
        '
        'TxtPrimaryMachineGroup
        '
        Me.TxtPrimaryMachineGroup.Location = New System.Drawing.Point(84, 71)
        Me.TxtPrimaryMachineGroup.Name = "TxtPrimaryMachineGroup"
        Me.TxtPrimaryMachineGroup.Size = New System.Drawing.Size(43, 20)
        Me.TxtPrimaryMachineGroup.TabIndex = 4
        Me.TxtPrimaryMachineGroup.TabStop = False
        Me.TxtPrimaryMachineGroup.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 74)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(108, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Machine Group Code"
        '
        'TxtMachineGroupCode
        '
        Me.TxtMachineGroupCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMachineGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMachineGroupCode.Location = New System.Drawing.Point(133, 71)
        Me.TxtMachineGroupCode.Name = "TxtMachineGroupCode"
        Me.TxtMachineGroupCode.Size = New System.Drawing.Size(321, 20)
        Me.TxtMachineGroupCode.TabIndex = 5
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnCancel.Location = New System.Drawing.Point(265, 232)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(88, 26)
        Me.BtnCancel.TabIndex = 11
        Me.BtnCancel.Text = "&Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'BtnOk
        '
        Me.BtnOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BtnOk.Location = New System.Drawing.Point(163, 232)
        Me.BtnOk.Name = "BtnOk"
        Me.BtnOk.Size = New System.Drawing.Size(88, 26)
        Me.BtnOk.TabIndex = 10
        Me.BtnOk.Text = "&OK"
        Me.BtnOk.UseVisualStyleBackColor = True
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(84, 19)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(43, 20)
        Me.TxtPrimary.TabIndex = 1
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Visible = False
        '
        'ChkActive
        '
        Me.ChkActive.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ChkActive.AutoSize = True
        Me.ChkActive.Location = New System.Drawing.Point(133, 207)
        Me.ChkActive.Name = "ChkActive"
        Me.ChkActive.Size = New System.Drawing.Size(56, 17)
        Me.ChkActive.TabIndex = 9
        Me.ChkActive.Text = "Active"
        Me.ChkActive.UseVisualStyleBackColor = True
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(133, 45)
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(365, 20)
        Me.TxtDesc.TabIndex = 3
        '
        'TxtMachineDataCode
        '
        Me.TxtMachineDataCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMachineDataCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMachineDataCode.Location = New System.Drawing.Point(133, 19)
        Me.TxtMachineDataCode.Name = "TxtMachineDataCode"
        Me.TxtMachineDataCode.Size = New System.Drawing.Size(365, 20)
        Me.TxtMachineDataCode.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 100)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(108, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Machine Group Desc"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Description"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(102, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Machine Data Code"
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(133, 181)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(43, 20)
        Me.TxtReadOnly.TabIndex = 8
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TslCreatedBy, Me.TslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 314)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(535, 22)
        Me.StatusStrip2.TabIndex = 22
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'TslCreatedBy
        '
        Me.TslCreatedBy.Name = "TslCreatedBy"
        Me.TslCreatedBy.Size = New System.Drawing.Size(260, 17)
        Me.TslCreatedBy.Spring = True
        Me.TslCreatedBy.Text = "Created By:"
        '
        'TslCreatedDate
        '
        Me.TslCreatedDate.Name = "TslCreatedDate"
        Me.TslCreatedDate.Size = New System.Drawing.Size(260, 17)
        Me.TslCreatedDate.Spring = True
        Me.TslCreatedDate.Text = "Created Date:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TslModifiedBy, Me.TslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 336)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(535, 22)
        Me.StatusStrip1.TabIndex = 21
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'TslModifiedBy
        '
        Me.TslModifiedBy.Name = "TslModifiedBy"
        Me.TslModifiedBy.Size = New System.Drawing.Size(260, 17)
        Me.TslModifiedBy.Spring = True
        Me.TslModifiedBy.Text = "Modified By:"
        '
        'TslModifiedDate
        '
        Me.TslModifiedDate.Name = "TslModifiedDate"
        Me.TslModifiedDate.Size = New System.Drawing.Size(260, 17)
        Me.TslModifiedDate.Spring = True
        Me.TslModifiedDate.Text = "Modified Date:"
        '
        'FrmMasterMachineData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(535, 358)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.BtnDelete)
        Me.Controls.Add(Me.BtnEdit)
        Me.Controls.Add(Me.BtnNew)
        Me.Controls.Add(Me.BtnFind)
        Me.Controls.Add(Me.GboDetail)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(551, 396)
        Me.Name = "FrmMasterMachineData"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Machine Data"
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
    Friend WithEvents ChkActive As System.Windows.Forms.CheckBox
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtMachineDataCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents TslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents TslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtPrimaryMachineGroup As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtMachineGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindMachineGroup As System.Windows.Forms.Button
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtMachineGroupDesc As System.Windows.Forms.TextBox
End Class
