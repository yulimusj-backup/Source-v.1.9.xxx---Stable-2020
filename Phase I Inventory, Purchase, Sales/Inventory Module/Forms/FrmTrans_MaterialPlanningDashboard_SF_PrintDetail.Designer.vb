<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_SF_PrintDetail
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
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboOrderNo = New System.Windows.Forms.ComboBox
        Me.cboGroupBy = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(192, 62)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 19)
        Me.btnCancel.TabIndex = 0
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(131, 62)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 19)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 13)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Order No"
        '
        'cboOrderNo
        '
        Me.cboOrderNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboOrderNo.FormattingEnabled = True
        Me.cboOrderNo.Items.AddRange(New Object() {"Sales Order Number", "Cust. PO Number", "Both"})
        Me.cboOrderNo.Location = New System.Drawing.Point(91, 10)
        Me.cboOrderNo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboOrderNo.Name = "cboOrderNo"
        Me.cboOrderNo.Size = New System.Drawing.Size(153, 21)
        Me.cboOrderNo.TabIndex = 3
        '
        'cboGroupBy
        '
        Me.cboGroupBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGroupBy.FormattingEnabled = True
        Me.cboGroupBy.Items.AddRange(New Object() {"Customer", "Product Code"})
        Me.cboGroupBy.Location = New System.Drawing.Point(91, 34)
        Me.cboGroupBy.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboGroupBy.Name = "cboGroupBy"
        Me.cboGroupBy.Size = New System.Drawing.Size(153, 21)
        Me.cboGroupBy.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 37)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Group By"
        '
        'FrmTrans_MaterialPlanningDashboard_SF_PrintDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(257, 90)
        Me.Controls.Add(Me.cboGroupBy)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboOrderNo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(261, 127)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_SF_PrintDetail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents cboOrderNo As System.Windows.Forms.ComboBox
    Public WithEvents cboGroupBy As System.Windows.Forms.ComboBox
End Class
