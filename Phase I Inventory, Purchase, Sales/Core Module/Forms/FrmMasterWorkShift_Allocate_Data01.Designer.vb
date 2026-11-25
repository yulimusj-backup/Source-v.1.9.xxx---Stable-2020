<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterWorkShift_Allocate_Data01
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
        Me.btnExit = New System.Windows.Forms.Button
        Me.DGVPayment = New System.Windows.Forms.DataGridView
        CType(Me.DGVPayment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(326, 166)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'DGVPayment
        '
        Me.DGVPayment.AllowUserToAddRows = False
        Me.DGVPayment.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVPayment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVPayment.Location = New System.Drawing.Point(0, 0)
        Me.DGVPayment.Margin = New System.Windows.Forms.Padding(2)
        Me.DGVPayment.MultiSelect = False
        Me.DGVPayment.Name = "DGVPayment"
        Me.DGVPayment.RowTemplate.Height = 24
        Me.DGVPayment.Size = New System.Drawing.Size(732, 358)
        Me.DGVPayment.TabIndex = 8
        '
        'FrmMasterWorkShift_Allocate_Data01
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(732, 358)
        Me.Controls.Add(Me.DGVPayment)
        Me.Controls.Add(Me.btnExit)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(748, 396)
        Me.Name = "FrmMasterWorkShift_Allocate_Data01"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data Allocate"
        CType(Me.DGVPayment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents DGVPayment As System.Windows.Forms.DataGridView
End Class
