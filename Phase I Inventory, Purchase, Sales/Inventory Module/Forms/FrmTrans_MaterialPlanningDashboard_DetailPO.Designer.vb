<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_DetailPO
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPODate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupplier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInclude = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimaryPO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnOK = New System.Windows.Forms.Button
        Me.llbMP_CekAll = New System.Windows.Forms.LinkLabel
        Me.llbMP_UnCekAll = New System.Windows.Forms.LinkLabel
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(525, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "There are some unauthorized po. Please specify which on to include in calculation" & _
            ""
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColPONumber, Me.ColPODate, Me.ColSupplier, Me.ColInclude, Me.ColPrimaryPO})
        Me.dgvItem.Location = New System.Drawing.Point(12, 30)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(687, 480)
        Me.dgvItem.TabIndex = 1
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColPONumber
        '
        Me.ColPONumber.HeaderText = "PO Number"
        Me.ColPONumber.Name = "ColPONumber"
        Me.ColPONumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColPONumber.Width = 107
        '
        'ColPODate
        '
        Me.ColPODate.HeaderText = "PO Date"
        Me.ColPODate.Name = "ColPODate"
        Me.ColPODate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColPODate.Width = 87
        '
        'ColSupplier
        '
        Me.ColSupplier.HeaderText = "Supplier"
        Me.ColSupplier.Name = "ColSupplier"
        Me.ColSupplier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSupplier.Width = 66
        '
        'ColInclude
        '
        Me.ColInclude.HeaderText = "Include"
        Me.ColInclude.Name = "ColInclude"
        Me.ColInclude.Width = 59
        '
        'ColPrimaryPO
        '
        Me.ColPrimaryPO.HeaderText = "Primary PO"
        Me.ColPrimaryPO.Name = "ColPrimaryPO"
        Me.ColPrimaryPO.Visible = False
        Me.ColPrimaryPO.Width = 105
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(624, 514)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'llbMP_CekAll
        '
        Me.llbMP_CekAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbMP_CekAll.AutoSize = True
        Me.llbMP_CekAll.Location = New System.Drawing.Point(13, 520)
        Me.llbMP_CekAll.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbMP_CekAll.Name = "llbMP_CekAll"
        Me.llbMP_CekAll.Size = New System.Drawing.Size(66, 17)
        Me.llbMP_CekAll.TabIndex = 89
        Me.llbMP_CekAll.TabStop = True
        Me.llbMP_CekAll.Text = "Check All"
        Me.llbMP_CekAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbMP_UnCekAll
        '
        Me.llbMP_UnCekAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbMP_UnCekAll.AutoSize = True
        Me.llbMP_UnCekAll.Location = New System.Drawing.Point(87, 520)
        Me.llbMP_UnCekAll.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbMP_UnCekAll.Name = "llbMP_UnCekAll"
        Me.llbMP_UnCekAll.Size = New System.Drawing.Size(82, 17)
        Me.llbMP_UnCekAll.TabIndex = 90
        Me.llbMP_UnCekAll.TabStop = True
        Me.llbMP_UnCekAll.Text = "Uncheck All"
        Me.llbMP_UnCekAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'FrmTrans_MaterialPlanningDashboard_DetailPO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(721, 548)
        Me.ControlBox = False
        Me.Controls.Add(Me.llbMP_UnCekAll)
        Me.Controls.Add(Me.llbMP_CekAll)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(729, 549)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_DetailPO"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Unauthorized PO"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPODate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupplier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInclude As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimaryPO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents llbMP_CekAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbMP_UnCekAll As System.Windows.Forms.LinkLabel
End Class
