<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPOOtorisasi_Removed
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemoved = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimaryPO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnChanged = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.llbChange = New System.Windows.Forms.LinkLabel
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColPONumber, Me.ColRemoved, Me.ColPrimaryPO})
        Me.DataGridView1.Location = New System.Drawing.Point(9, 10)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(345, 170)
        Me.DataGridView1.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 46
        '
        'ColPONumber
        '
        Me.ColPONumber.HeaderText = "PO Number"
        Me.ColPONumber.Name = "ColPONumber"
        Me.ColPONumber.Width = 87
        '
        'ColRemoved
        '
        Me.ColRemoved.HeaderText = "Removed"
        Me.ColRemoved.Name = "ColRemoved"
        Me.ColRemoved.Width = 59
        '
        'ColPrimaryPO
        '
        Me.ColPrimaryPO.HeaderText = "Primary PO"
        Me.ColPrimaryPO.Name = "ColPrimaryPO"
        Me.ColPrimaryPO.Visible = False
        '
        'btnChanged
        '
        Me.btnChanged.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnChanged.Location = New System.Drawing.Point(9, 184)
        Me.btnChanged.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnChanged.Name = "btnChanged"
        Me.btnChanged.Size = New System.Drawing.Size(56, 19)
        Me.btnChanged.TabIndex = 1
        Me.btnChanged.Text = "Change"
        Me.btnChanged.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(298, 184)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(56, 19)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'llbChange
        '
        Me.llbChange.AutoSize = True
        Me.llbChange.Location = New System.Drawing.Point(70, 188)
        Me.llbChange.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbChange.Name = "llbChange"
        Me.llbChange.Size = New System.Drawing.Size(65, 13)
        Me.llbChange.TabIndex = 3
        Me.llbChange.TabStop = True
        Me.llbChange.Text = "Uncheck All"
        '
        'FrmTransaksiPOOtorisasi_Removed
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(363, 213)
        Me.Controls.Add(Me.llbChange)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnChanged)
        Me.Controls.Add(Me.DataGridView1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(379, 251)
        Me.Name = "FrmTransaksiPOOtorisasi_Removed"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Removed PO"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemoved As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimaryPO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnChanged As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents llbChange As System.Windows.Forms.LinkLabel
End Class
