<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_POS_Payment_Load
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
        Me.TxtPrimarySession = New System.Windows.Forms.TextBox
        Me.CBOSession_ID = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TxtPrimarySession
        '
        Me.TxtPrimarySession.Location = New System.Drawing.Point(71, 12)
        Me.TxtPrimarySession.Name = "TxtPrimarySession"
        Me.TxtPrimarySession.Size = New System.Drawing.Size(31, 20)
        Me.TxtPrimarySession.TabIndex = 17
        Me.TxtPrimarySession.TabStop = False
        Me.TxtPrimarySession.Text = "-1"
        Me.TxtPrimarySession.Visible = False
        '
        'CBOSession_ID
        '
        Me.CBOSession_ID.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CBOSession_ID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBOSession_ID.FormattingEnabled = True
        Me.CBOSession_ID.Location = New System.Drawing.Point(108, 12)
        Me.CBOSession_ID.Name = "CBOSession_ID"
        Me.CBOSession_ID.Size = New System.Drawing.Size(237, 21)
        Me.CBOSession_ID.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(61, 13)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Session_ID"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(270, 39)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(270, 39)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 27
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'FrmTrans_POS_Payment_Load
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(357, 71)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.TxtPrimarySession)
        Me.Controls.Add(Me.CBOSession_ID)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnCancel)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(373, 215)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(373, 109)
        Me.Name = "FrmTrans_POS_Payment_Load"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login_Session_ID"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtPrimarySession As System.Windows.Forms.TextBox
    Friend WithEvents CBOSession_ID As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
