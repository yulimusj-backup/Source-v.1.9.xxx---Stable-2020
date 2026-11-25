<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm2021
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
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing = New System.Windows.Forms.Button
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing = New System.Windows.Forms.Label
        Me.Label88 = New System.Windows.Forms.Label
        Me.prg2018_UpdateItemProdQtyMonthlyClosing = New System.Windows.Forms.ProgressBar
        Me.SuspendLayout()
        '
        'btnUpdate2018_ItemProdQtyMonthlyClosing
        '
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.Location = New System.Drawing.Point(876, 8)
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.Name = "btnUpdate2018_ItemProdQtyMonthlyClosing"
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.Size = New System.Drawing.Size(83, 28)
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.TabIndex = 149
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.TabStop = False
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.Text = "Update"
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.UseVisualStyleBackColor = True
        '
        'LblProgress2018_UpdateItemProdQtyMonthlyClosing
        '
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing.AutoSize = True
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing.Location = New System.Drawing.Point(325, 37)
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing.Name = "LblProgress2018_UpdateItemProdQtyMonthlyClosing"
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing.Size = New System.Drawing.Size(81, 17)
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing.TabIndex = 148
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing.Text = "Progress...."
        '
        'Label88
        '
        Me.Label88.AutoSize = True
        Me.Label88.Location = New System.Drawing.Point(12, 17)
        Me.Label88.Name = "Label88"
        Me.Label88.Size = New System.Drawing.Size(32, 17)
        Me.Label88.TabIndex = 147
        Me.Label88.Text = "xxxx"
        '
        'prg2018_UpdateItemProdQtyMonthlyClosing
        '
        Me.prg2018_UpdateItemProdQtyMonthlyClosing.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prg2018_UpdateItemProdQtyMonthlyClosing.Location = New System.Drawing.Point(325, 11)
        Me.prg2018_UpdateItemProdQtyMonthlyClosing.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg2018_UpdateItemProdQtyMonthlyClosing.Name = "prg2018_UpdateItemProdQtyMonthlyClosing"
        Me.prg2018_UpdateItemProdQtyMonthlyClosing.Size = New System.Drawing.Size(544, 23)
        Me.prg2018_UpdateItemProdQtyMonthlyClosing.TabIndex = 146
        '
        'Frm2021
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(972, 263)
        Me.Controls.Add(Me.btnUpdate2018_ItemProdQtyMonthlyClosing)
        Me.Controls.Add(Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing)
        Me.Controls.Add(Me.Label88)
        Me.Controls.Add(Me.prg2018_UpdateItemProdQtyMonthlyClosing)
        Me.Name = "Frm2021"
        Me.Text = "Frm2021"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnUpdate2018_ItemProdQtyMonthlyClosing As System.Windows.Forms.Button
    Friend WithEvents LblProgress2018_UpdateItemProdQtyMonthlyClosing As System.Windows.Forms.Label
    Friend WithEvents Label88 As System.Windows.Forms.Label
    Friend WithEvents prg2018_UpdateItemProdQtyMonthlyClosing As System.Windows.Forms.ProgressBar
End Class
