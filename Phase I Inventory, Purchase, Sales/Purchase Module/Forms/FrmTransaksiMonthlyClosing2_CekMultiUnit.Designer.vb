<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonthlyClosing2_CekMultiUnit
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
        Me.dgv1 = New System.Windows.Forms.DataGridView
        Me.ColCostGroup = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRefUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUpdate = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.llbUncekAll = New System.Windows.Forms.LinkLabel
        Me.llbCekAll = New System.Windows.Forms.LinkLabel
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgv1
        '
        Me.dgv1.AllowUserToAddRows = False
        Me.dgv1.AllowUserToDeleteRows = False
        Me.dgv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgv1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgv1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCostGroup, Me.ColCodeNo, Me.ColDesc, Me.ColRefUnit, Me.ColUpdate})
        Me.dgv1.Location = New System.Drawing.Point(12, 46)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgv1.RowTemplate.Height = 24
        Me.dgv1.Size = New System.Drawing.Size(723, 513)
        Me.dgv1.TabIndex = 0
        '
        'ColCostGroup
        '
        Me.ColCostGroup.HeaderText = "Cost Group"
        Me.ColCostGroup.Name = "ColCostGroup"
        Me.ColCostGroup.ReadOnly = True
        Me.ColCostGroup.Width = 105
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.Width = 88
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.Width = 104
        '
        'ColRefUnit
        '
        Me.ColRefUnit.HeaderText = "Unit"
        Me.ColRefUnit.Name = "ColRefUnit"
        Me.ColRefUnit.ReadOnly = True
        Me.ColRefUnit.Width = 58
        '
        'ColUpdate
        '
        Me.ColUpdate.HeaderText = "Update"
        Me.ColUpdate.Name = "ColUpdate"
        Me.ColUpdate.ReadOnly = True
        Me.ColUpdate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColUpdate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColUpdate.Width = 79
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(660, 565)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(663, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "The following code no, doesn't have unit of measurement required to perform overh" & _
            "ead cost calculation,"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(109, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "please fix it first."
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(9, 568)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(109, 17)
        Me.LinkLabel1.TabIndex = 4
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Update Process"
        '
        'llbUncekAll
        '
        Me.llbUncekAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncekAll.AutoSize = True
        Me.llbUncekAll.Location = New System.Drawing.Point(572, 568)
        Me.llbUncekAll.Name = "llbUncekAll"
        Me.llbUncekAll.Size = New System.Drawing.Size(82, 17)
        Me.llbUncekAll.TabIndex = 5
        Me.llbUncekAll.TabStop = True
        Me.llbUncekAll.Text = "Uncheck All"
        '
        'llbCekAll
        '
        Me.llbCekAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCekAll.AutoSize = True
        Me.llbCekAll.Location = New System.Drawing.Point(500, 568)
        Me.llbCekAll.Name = "llbCekAll"
        Me.llbCekAll.Size = New System.Drawing.Size(66, 17)
        Me.llbCekAll.TabIndex = 6
        Me.llbCekAll.TabStop = True
        Me.llbCekAll.Text = "Check All"
        '
        'FrmTransaksiMonthlyClosing2_CekMultiUnit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(747, 600)
        Me.Controls.Add(Me.llbCekAll)
        Me.Controls.Add(Me.llbUncekAll)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.dgv1)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(755, 637)
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiMonthlyClosing2_CekMultiUnit"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Information"
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents ColCostGroup As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRefUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUpdate As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents llbUncekAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCekAll As System.Windows.Forms.LinkLabel
End Class
