<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonthlyClosing_Export
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
        Me.ChkRecap = New System.Windows.Forms.CheckBox
        Me.ChkRecapWithDetail = New System.Windows.Forms.CheckBox
        Me.ChkDetail = New System.Windows.Forms.CheckBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.BtnExport = New System.Windows.Forms.Button
        Me.BtnClose = New System.Windows.Forms.Button
        Me.SaveDia = New System.Windows.Forms.SaveFileDialog
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ChkRecap
        '
        Me.ChkRecap.AutoSize = True
        Me.ChkRecap.Location = New System.Drawing.Point(10, 20)
        Me.ChkRecap.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ChkRecap.Name = "ChkRecap"
        Me.ChkRecap.Size = New System.Drawing.Size(58, 17)
        Me.ChkRecap.TabIndex = 0
        Me.ChkRecap.Text = "Recap"
        Me.ChkRecap.UseVisualStyleBackColor = True
        '
        'ChkRecapWithDetail
        '
        Me.ChkRecapWithDetail.AutoSize = True
        Me.ChkRecapWithDetail.Location = New System.Drawing.Point(10, 41)
        Me.ChkRecapWithDetail.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ChkRecapWithDetail.Name = "ChkRecapWithDetail"
        Me.ChkRecapWithDetail.Size = New System.Drawing.Size(166, 17)
        Me.ChkRecapWithDetail.TabIndex = 1
        Me.ChkRecapWithDetail.Text = "Recap W/ Detail Transaction"
        Me.ChkRecapWithDetail.UseVisualStyleBackColor = True
        '
        'ChkDetail
        '
        Me.ChkDetail.AutoSize = True
        Me.ChkDetail.Location = New System.Drawing.Point(10, 63)
        Me.ChkDetail.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ChkDetail.Name = "ChkDetail"
        Me.ChkDetail.Size = New System.Drawing.Size(53, 17)
        Me.ChkDetail.TabIndex = 2
        Me.ChkDetail.Text = "Detail"
        Me.ChkDetail.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.ChkRecap)
        Me.GroupBox1.Controls.Add(Me.ChkDetail)
        Me.GroupBox1.Controls.Add(Me.ChkRecapWithDetail)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 10)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(305, 88)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Format"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 124)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Label1"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(9, 102)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(202, 19)
        Me.ProgressBar1.TabIndex = 5
        '
        'BtnExport
        '
        Me.BtnExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnExport.Location = New System.Drawing.Point(215, 102)
        Me.BtnExport.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.BtnExport.Name = "BtnExport"
        Me.BtnExport.Size = New System.Drawing.Size(47, 19)
        Me.BtnExport.TabIndex = 6
        Me.BtnExport.Text = "Export"
        Me.BtnExport.UseVisualStyleBackColor = True
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(267, 102)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(47, 19)
        Me.BtnClose.TabIndex = 7
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'FrmTransaksiMonthlyClosing_Export
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(325, 150)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.BtnExport)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(341, 188)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(341, 188)
        Me.Name = "FrmTransaksiMonthlyClosing_Export"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Save to file"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ChkRecap As System.Windows.Forms.CheckBox
    Friend WithEvents ChkRecapWithDetail As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDetail As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents BtnExport As System.Windows.Forms.Button
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents SaveDia As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
