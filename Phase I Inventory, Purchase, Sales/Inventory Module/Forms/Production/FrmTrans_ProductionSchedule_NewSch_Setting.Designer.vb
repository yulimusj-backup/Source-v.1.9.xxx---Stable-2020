<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_NewSch_Setting
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ChkQtyTarget = New System.Windows.Forms.CheckBox
        Me.ChkDesc = New System.Windows.Forms.CheckBox
        Me.ChkCodeNo = New System.Windows.Forms.CheckBox
        Me.chkJONUmber = New System.Windows.Forms.CheckBox
        Me.chkMONumber = New System.Windows.Forms.CheckBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.ChkQtyTarget)
        Me.GroupBox1.Controls.Add(Me.ChkDesc)
        Me.GroupBox1.Controls.Add(Me.ChkCodeNo)
        Me.GroupBox1.Controls.Add(Me.chkJONUmber)
        Me.GroupBox1.Controls.Add(Me.chkMONumber)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(425, 81)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Please specify visible columns"
        '
        'ChkQtyTarget
        '
        Me.ChkQtyTarget.AutoSize = True
        Me.ChkQtyTarget.Location = New System.Drawing.Point(311, 21)
        Me.ChkQtyTarget.Name = "ChkQtyTarget"
        Me.ChkQtyTarget.Size = New System.Drawing.Size(98, 21)
        Me.ChkQtyTarget.TabIndex = 4
        Me.ChkQtyTarget.Text = "Qty Target"
        Me.ChkQtyTarget.UseVisualStyleBackColor = True
        '
        'ChkDesc
        '
        Me.ChkDesc.AutoSize = True
        Me.ChkDesc.Location = New System.Drawing.Point(154, 48)
        Me.ChkDesc.Name = "ChkDesc"
        Me.ChkDesc.Size = New System.Drawing.Size(101, 21)
        Me.ChkDesc.TabIndex = 3
        Me.ChkDesc.Text = "Description"
        Me.ChkDesc.UseVisualStyleBackColor = True
        '
        'ChkCodeNo
        '
        Me.ChkCodeNo.AutoSize = True
        Me.ChkCodeNo.Location = New System.Drawing.Point(154, 21)
        Me.ChkCodeNo.Name = "ChkCodeNo"
        Me.ChkCodeNo.Size = New System.Drawing.Size(85, 21)
        Me.ChkCodeNo.TabIndex = 2
        Me.ChkCodeNo.Text = "Code No"
        Me.ChkCodeNo.UseVisualStyleBackColor = True
        '
        'chkJONUmber
        '
        Me.chkJONUmber.AutoSize = True
        Me.chkJONUmber.Location = New System.Drawing.Point(6, 48)
        Me.chkJONUmber.Name = "chkJONUmber"
        Me.chkJONUmber.Size = New System.Drawing.Size(102, 21)
        Me.chkJONUmber.TabIndex = 1
        Me.chkJONUmber.Text = "JO Number"
        Me.chkJONUmber.UseVisualStyleBackColor = True
        '
        'chkMONumber
        '
        Me.chkMONumber.AutoSize = True
        Me.chkMONumber.Location = New System.Drawing.Point(6, 21)
        Me.chkMONumber.Name = "chkMONumber"
        Me.chkMONumber.Size = New System.Drawing.Size(106, 21)
        Me.chkMONumber.TabIndex = 0
        Me.chkMONumber.Text = "MO Number"
        Me.chkMONumber.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(365, 99)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmTrans_ProductionSchedule_NewSch_Setting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(449, 131)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(457, 168)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(457, 168)
        Me.Name = "FrmTrans_ProductionSchedule_NewSch_Setting"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Columns Setting"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkQtyTarget As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDesc As System.Windows.Forms.CheckBox
    Friend WithEvents ChkCodeNo As System.Windows.Forms.CheckBox
    Friend WithEvents chkJONUmber As System.Windows.Forms.CheckBox
    Friend WithEvents chkMONumber As System.Windows.Forms.CheckBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
End Class
