<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_MP_DetailTolerance
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
        Me.TxtQty_Order = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.LblUnit = New System.Windows.Forms.Label
        Me.LblUnit1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtQty_UpperTolerance = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtPerct_UpperTolerance = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtPerct_LowerTolerance = New System.Windows.Forms.TextBox
        Me.LblUnit2 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtQty_LowerTolerance = New System.Windows.Forms.TextBox
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(309, 101)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(228, 101)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtQty_Order
        '
        Me.TxtQty_Order.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQty_Order.Location = New System.Drawing.Point(138, 12)
        Me.TxtQty_Order.Name = "TxtQty_Order"
        Me.TxtQty_Order.ReadOnly = True
        Me.TxtQty_Order.Size = New System.Drawing.Size(198, 22)
        Me.TxtQty_Order.TabIndex = 0
        Me.TxtQty_Order.TabStop = False
        Me.TxtQty_Order.Text = "0"
        Me.TxtQty_Order.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Order Qty"
        '
        'LblUnit
        '
        Me.LblUnit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnit.AutoSize = True
        Me.LblUnit.Location = New System.Drawing.Point(342, 15)
        Me.LblUnit.Name = "LblUnit"
        Me.LblUnit.Size = New System.Drawing.Size(33, 17)
        Me.LblUnit.TabIndex = 4
        Me.LblUnit.Text = "Unit"
        '
        'LblUnit1
        '
        Me.LblUnit1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnit1.AutoSize = True
        Me.LblUnit1.Location = New System.Drawing.Point(342, 43)
        Me.LblUnit1.Name = "LblUnit1"
        Me.LblUnit1.Size = New System.Drawing.Size(33, 17)
        Me.LblUnit1.TabIndex = 7
        Me.LblUnit1.Text = "Unit"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 43)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(115, 17)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Upper Tolerance"
        '
        'TxtQty_UpperTolerance
        '
        Me.TxtQty_UpperTolerance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQty_UpperTolerance.Location = New System.Drawing.Point(228, 40)
        Me.TxtQty_UpperTolerance.Name = "TxtQty_UpperTolerance"
        Me.TxtQty_UpperTolerance.Size = New System.Drawing.Size(108, 22)
        Me.TxtQty_UpperTolerance.TabIndex = 2
        Me.TxtQty_UpperTolerance.Text = "0"
        Me.TxtQty_UpperTolerance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(202, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 17)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "%"
        '
        'TxtPerct_UpperTolerance
        '
        Me.TxtPerct_UpperTolerance.Location = New System.Drawing.Point(138, 40)
        Me.TxtPerct_UpperTolerance.Name = "TxtPerct_UpperTolerance"
        Me.TxtPerct_UpperTolerance.Size = New System.Drawing.Size(58, 22)
        Me.TxtPerct_UpperTolerance.TabIndex = 1
        Me.TxtPerct_UpperTolerance.Text = "0"
        Me.TxtPerct_UpperTolerance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(202, 71)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(20, 17)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "%"
        '
        'TxtPerct_LowerTolerance
        '
        Me.TxtPerct_LowerTolerance.Location = New System.Drawing.Point(138, 68)
        Me.TxtPerct_LowerTolerance.Name = "TxtPerct_LowerTolerance"
        Me.TxtPerct_LowerTolerance.Size = New System.Drawing.Size(58, 22)
        Me.TxtPerct_LowerTolerance.TabIndex = 3
        Me.TxtPerct_LowerTolerance.Text = "0"
        Me.TxtPerct_LowerTolerance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblUnit2
        '
        Me.LblUnit2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnit2.AutoSize = True
        Me.LblUnit2.Location = New System.Drawing.Point(342, 71)
        Me.LblUnit2.Name = "LblUnit2"
        Me.LblUnit2.Size = New System.Drawing.Size(33, 17)
        Me.LblUnit2.TabIndex = 12
        Me.LblUnit2.Text = "Unit"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(12, 71)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(114, 17)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "Lower Tolerance"
        '
        'TxtQty_LowerTolerance
        '
        Me.TxtQty_LowerTolerance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQty_LowerTolerance.Location = New System.Drawing.Point(228, 68)
        Me.TxtQty_LowerTolerance.Name = "TxtQty_LowerTolerance"
        Me.TxtQty_LowerTolerance.Size = New System.Drawing.Size(108, 22)
        Me.TxtQty_LowerTolerance.TabIndex = 4
        Me.TxtQty_LowerTolerance.Text = "0"
        Me.TxtQty_LowerTolerance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(130, 102)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(58, 22)
        Me.TxtPrimary.TabIndex = 15
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtPrimary.Visible = False
        '
        'FrmTrans_MaterialPlanningDashboard_MP_DetailTolerance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(396, 136)
        Me.Controls.Add(Me.TxtPrimary)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TxtPerct_LowerTolerance)
        Me.Controls.Add(Me.LblUnit2)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.TxtQty_LowerTolerance)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtPerct_UpperTolerance)
        Me.Controls.Add(Me.LblUnit1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TxtQty_UpperTolerance)
        Me.Controls.Add(Me.LblUnit)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtQty_Order)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(404, 173)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_MP_DetailTolerance"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Tolerance"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtQty_Order As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblUnit As System.Windows.Forms.Label
    Friend WithEvents LblUnit1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtQty_UpperTolerance As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtPerct_UpperTolerance As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtPerct_LowerTolerance As System.Windows.Forms.TextBox
    Friend WithEvents LblUnit2 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtQty_LowerTolerance As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
End Class
