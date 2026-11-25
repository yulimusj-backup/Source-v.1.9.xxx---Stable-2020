<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmSet_IO_Location
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnFindWUGroup = New System.Windows.Forms.Button
        Me.TxtWUGroupCode = New System.Windows.Forms.TextBox
        Me.TxtWUGroupDesc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryWUGroup = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtPrimaryInputLocation = New System.Windows.Forms.TextBox
        Me.TxtInputLocationDesc = New System.Windows.Forms.TextBox
        Me.TxtInputLocationCode = New System.Windows.Forms.TextBox
        Me.btnFindInputLocation = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtPrimaryOutputLocation = New System.Windows.Forms.TextBox
        Me.TxtOutputLocationDesc = New System.Windows.Forms.TextBox
        Me.TxtOutputLocationCode = New System.Windows.Forms.TextBox
        Me.btnFindOutputLocation = New System.Windows.Forms.Button
        Me.btnSet = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Group WU"
        '
        'btnFindWUGroup
        '
        Me.btnFindWUGroup.Location = New System.Drawing.Point(233, 12)
        Me.btnFindWUGroup.Name = "btnFindWUGroup"
        Me.btnFindWUGroup.Size = New System.Drawing.Size(75, 23)
        Me.btnFindWUGroup.TabIndex = 1
        Me.btnFindWUGroup.Text = "Find"
        Me.btnFindWUGroup.UseVisualStyleBackColor = True
        '
        'TxtWUGroupCode
        '
        Me.TxtWUGroupCode.Location = New System.Drawing.Point(127, 12)
        Me.TxtWUGroupCode.Name = "TxtWUGroupCode"
        Me.TxtWUGroupCode.Size = New System.Drawing.Size(100, 22)
        Me.TxtWUGroupCode.TabIndex = 2
        '
        'TxtWUGroupDesc
        '
        Me.TxtWUGroupDesc.Location = New System.Drawing.Point(314, 12)
        Me.TxtWUGroupDesc.Name = "TxtWUGroupDesc"
        Me.TxtWUGroupDesc.ReadOnly = True
        Me.TxtWUGroupDesc.Size = New System.Drawing.Size(220, 22)
        Me.TxtWUGroupDesc.TabIndex = 3
        '
        'TxtPrimaryWUGroup
        '
        Me.TxtPrimaryWUGroup.Location = New System.Drawing.Point(64, 12)
        Me.TxtPrimaryWUGroup.Name = "TxtPrimaryWUGroup"
        Me.TxtPrimaryWUGroup.Size = New System.Drawing.Size(42, 22)
        Me.TxtPrimaryWUGroup.TabIndex = 4
        Me.TxtPrimaryWUGroup.TabStop = False
        Me.TxtPrimaryWUGroup.Text = "-1"
        Me.TxtPrimaryWUGroup.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(97, 17)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Input Location"
        '
        'TxtPrimaryInputLocation
        '
        Me.TxtPrimaryInputLocation.Location = New System.Drawing.Point(64, 41)
        Me.TxtPrimaryInputLocation.Name = "TxtPrimaryInputLocation"
        Me.TxtPrimaryInputLocation.Size = New System.Drawing.Size(42, 22)
        Me.TxtPrimaryInputLocation.TabIndex = 9
        Me.TxtPrimaryInputLocation.TabStop = False
        Me.TxtPrimaryInputLocation.Text = "-1"
        Me.TxtPrimaryInputLocation.Visible = False
        '
        'TxtInputLocationDesc
        '
        Me.TxtInputLocationDesc.Location = New System.Drawing.Point(314, 41)
        Me.TxtInputLocationDesc.Name = "TxtInputLocationDesc"
        Me.TxtInputLocationDesc.ReadOnly = True
        Me.TxtInputLocationDesc.Size = New System.Drawing.Size(220, 22)
        Me.TxtInputLocationDesc.TabIndex = 8
        '
        'TxtInputLocationCode
        '
        Me.TxtInputLocationCode.Location = New System.Drawing.Point(127, 41)
        Me.TxtInputLocationCode.Name = "TxtInputLocationCode"
        Me.TxtInputLocationCode.Size = New System.Drawing.Size(100, 22)
        Me.TxtInputLocationCode.TabIndex = 7
        '
        'btnFindInputLocation
        '
        Me.btnFindInputLocation.Location = New System.Drawing.Point(233, 41)
        Me.btnFindInputLocation.Name = "btnFindInputLocation"
        Me.btnFindInputLocation.Size = New System.Drawing.Size(75, 23)
        Me.btnFindInputLocation.TabIndex = 6
        Me.btnFindInputLocation.Text = "Find"
        Me.btnFindInputLocation.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(109, 17)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Output Location"
        '
        'TxtPrimaryOutputLocation
        '
        Me.TxtPrimaryOutputLocation.Location = New System.Drawing.Point(64, 69)
        Me.TxtPrimaryOutputLocation.Name = "TxtPrimaryOutputLocation"
        Me.TxtPrimaryOutputLocation.Size = New System.Drawing.Size(42, 22)
        Me.TxtPrimaryOutputLocation.TabIndex = 14
        Me.TxtPrimaryOutputLocation.TabStop = False
        Me.TxtPrimaryOutputLocation.Text = "-1"
        Me.TxtPrimaryOutputLocation.Visible = False
        '
        'TxtOutputLocationDesc
        '
        Me.TxtOutputLocationDesc.Location = New System.Drawing.Point(314, 69)
        Me.TxtOutputLocationDesc.Name = "TxtOutputLocationDesc"
        Me.TxtOutputLocationDesc.ReadOnly = True
        Me.TxtOutputLocationDesc.Size = New System.Drawing.Size(220, 22)
        Me.TxtOutputLocationDesc.TabIndex = 13
        '
        'TxtOutputLocationCode
        '
        Me.TxtOutputLocationCode.Location = New System.Drawing.Point(127, 69)
        Me.TxtOutputLocationCode.Name = "TxtOutputLocationCode"
        Me.TxtOutputLocationCode.Size = New System.Drawing.Size(100, 22)
        Me.TxtOutputLocationCode.TabIndex = 12
        '
        'btnFindOutputLocation
        '
        Me.btnFindOutputLocation.Location = New System.Drawing.Point(233, 69)
        Me.btnFindOutputLocation.Name = "btnFindOutputLocation"
        Me.btnFindOutputLocation.Size = New System.Drawing.Size(75, 23)
        Me.btnFindOutputLocation.TabIndex = 11
        Me.btnFindOutputLocation.Text = "Find"
        Me.btnFindOutputLocation.UseVisualStyleBackColor = True
        '
        'btnSet
        '
        Me.btnSet.Location = New System.Drawing.Point(127, 97)
        Me.btnSet.Name = "btnSet"
        Me.btnSet.Size = New System.Drawing.Size(75, 23)
        Me.btnSet.TabIndex = 15
        Me.btnSet.Text = "Set"
        Me.btnSet.UseVisualStyleBackColor = True
        '
        'FrmSet_IO_Location
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(546, 130)
        Me.Controls.Add(Me.btnSet)
        Me.Controls.Add(Me.TxtPrimaryOutputLocation)
        Me.Controls.Add(Me.TxtOutputLocationDesc)
        Me.Controls.Add(Me.TxtOutputLocationCode)
        Me.Controls.Add(Me.btnFindOutputLocation)
        Me.Controls.Add(Me.TxtPrimaryInputLocation)
        Me.Controls.Add(Me.TxtInputLocationDesc)
        Me.Controls.Add(Me.TxtInputLocationCode)
        Me.Controls.Add(Me.btnFindInputLocation)
        Me.Controls.Add(Me.TxtPrimaryWUGroup)
        Me.Controls.Add(Me.TxtWUGroupDesc)
        Me.Controls.Add(Me.TxtWUGroupCode)
        Me.Controls.Add(Me.btnFindWUGroup)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "FrmSet_IO_Location"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FrmSet_IO_Location"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnFindWUGroup As System.Windows.Forms.Button
    Friend WithEvents TxtWUGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtWUGroupDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryWUGroup As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryInputLocation As System.Windows.Forms.TextBox
    Friend WithEvents TxtInputLocationDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtInputLocationCode As System.Windows.Forms.TextBox
    Friend WithEvents btnFindInputLocation As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryOutputLocation As System.Windows.Forms.TextBox
    Friend WithEvents TxtOutputLocationDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtOutputLocationCode As System.Windows.Forms.TextBox
    Friend WithEvents btnFindOutputLocation As System.Windows.Forms.Button
    Friend WithEvents btnSet As System.Windows.Forms.Button
End Class
