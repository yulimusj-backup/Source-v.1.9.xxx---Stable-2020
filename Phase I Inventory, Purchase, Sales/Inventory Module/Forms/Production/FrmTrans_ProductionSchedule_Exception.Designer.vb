<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_Exception
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
        Me.dgvException = New System.Windows.Forms.DataGridView
        Me.ColException_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColException_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColException_AllowIgnore = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.dgvException, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvException
        '
        Me.dgvException.AllowUserToAddRows = False
        Me.dgvException.AllowUserToDeleteRows = False
        Me.dgvException.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvException.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvException.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvException.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColException_No, Me.ColException_Desc, Me.ColException_AllowIgnore})
        Me.dgvException.Location = New System.Drawing.Point(12, 30)
        Me.dgvException.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvException.Name = "dgvException"
        Me.dgvException.ReadOnly = True
        Me.dgvException.RowHeadersVisible = False
        Me.dgvException.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvException.RowTemplate.Height = 24
        Me.dgvException.Size = New System.Drawing.Size(668, 291)
        Me.dgvException.TabIndex = 3
        '
        'ColException_No
        '
        Me.ColException_No.HeaderText = "No"
        Me.ColException_No.Name = "ColException_No"
        Me.ColException_No.ReadOnly = True
        Me.ColException_No.Width = 51
        '
        'ColException_Desc
        '
        Me.ColException_Desc.HeaderText = "Description"
        Me.ColException_Desc.Name = "ColException_Desc"
        Me.ColException_Desc.ReadOnly = True
        Me.ColException_Desc.Width = 104
        '
        'ColException_AllowIgnore
        '
        Me.ColException_AllowIgnore.HeaderText = "Allow Ignore"
        Me.ColException_AllowIgnore.Name = "ColException_AllowIgnore"
        Me.ColException_AllowIgnore.ReadOnly = True
        Me.ColException_AllowIgnore.Visible = False
        Me.ColException_AllowIgnore.Width = 109
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(605, 328)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(524, 328)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(370, 17)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "We found the following exception, continue save anyway?"
        '
        'FrmTrans_ProductionSchedule_Exception
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(692, 363)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.dgvException)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(700, 400)
        Me.Name = "FrmTrans_ProductionSchedule_Exception"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Exception"
        CType(Me.dgvException, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvException As System.Windows.Forms.DataGridView
    Friend WithEvents ColException_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColException_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColException_AllowIgnore As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
