<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmUpdate_ChkInv
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
        Me.dgvInv = New System.Windows.Forms.DataGridView
        Me.ColInvNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvInv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvInv
        '
        Me.dgvInv.AllowUserToAddRows = False
        Me.dgvInv.AllowUserToDeleteRows = False
        Me.dgvInv.AllowUserToResizeColumns = False
        Me.dgvInv.AllowUserToResizeRows = False
        Me.dgvInv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvInv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvInv.ColumnHeadersHeight = 25
        Me.dgvInv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvInv.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColInvNo})
        Me.dgvInv.Location = New System.Drawing.Point(12, 12)
        Me.dgvInv.Name = "dgvInv"
        Me.dgvInv.ReadOnly = True
        Me.dgvInv.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvInv.RowTemplate.Height = 24
        Me.dgvInv.Size = New System.Drawing.Size(755, 409)
        Me.dgvInv.TabIndex = 0
        '
        'ColInvNo
        '
        Me.ColInvNo.HeaderText = "Invoice No"
        Me.ColInvNo.Name = "ColInvNo"
        Me.ColInvNo.ReadOnly = True
        Me.ColInvNo.Width = 99
        '
        'FrmUpdate_ChkInv
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(779, 433)
        Me.Controls.Add(Me.dgvInv)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(787, 470)
        Me.Name = "FrmUpdate_ChkInv"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Incomplete Invoice"
        CType(Me.dgvInv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvInv As System.Windows.Forms.DataGridView
    Friend WithEvents ColInvNo As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
