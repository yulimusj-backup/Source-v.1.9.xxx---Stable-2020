<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_NewSch3_Exception
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
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColAll_Process_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_WU = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_IdentifierProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_PrimaryWU = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.AllowUserToDeleteRows = False
        Me.DgvItem.AllowUserToResizeRows = False
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem.ColumnHeadersHeight = 25
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAll_Process_No, Me.ColAll_Process_Code, Me.ColAll_Process_CodeNo, Me.ColAll_Process_Desc, Me.ColAll_Process_WU, Me.ColAll_Process_PrimaryProcess, Me.ColAll_Process_PrimaryCodeNo, Me.ColAll_IdentifierProcess, Me.ColAll_Process_PrimaryWU})
        Me.DgvItem.Location = New System.Drawing.Point(13, 30)
        Me.DgvItem.Margin = New System.Windows.Forms.Padding(4)
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.ReadOnly = True
        Me.DgvItem.RowHeadersVisible = False
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.Size = New System.Drawing.Size(806, 448)
        Me.DgvItem.TabIndex = 1
        '
        'ColAll_Process_No
        '
        Me.ColAll_Process_No.HeaderText = "No"
        Me.ColAll_Process_No.Name = "ColAll_Process_No"
        Me.ColAll_Process_No.ReadOnly = True
        Me.ColAll_Process_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_No.Width = 32
        '
        'ColAll_Process_Code
        '
        Me.ColAll_Process_Code.HeaderText = "Code"
        Me.ColAll_Process_Code.Name = "ColAll_Process_Code"
        Me.ColAll_Process_Code.ReadOnly = True
        Me.ColAll_Process_Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_Code.Width = 47
        '
        'ColAll_Process_CodeNo
        '
        Me.ColAll_Process_CodeNo.HeaderText = "Code No"
        Me.ColAll_Process_CodeNo.Name = "ColAll_Process_CodeNo"
        Me.ColAll_Process_CodeNo.ReadOnly = True
        Me.ColAll_Process_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_CodeNo.Width = 69
        '
        'ColAll_Process_Desc
        '
        Me.ColAll_Process_Desc.HeaderText = "Description"
        Me.ColAll_Process_Desc.Name = "ColAll_Process_Desc"
        Me.ColAll_Process_Desc.ReadOnly = True
        Me.ColAll_Process_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_Desc.Width = 85
        '
        'ColAll_Process_WU
        '
        Me.ColAll_Process_WU.HeaderText = "Work Unit"
        Me.ColAll_Process_WU.Name = "ColAll_Process_WU"
        Me.ColAll_Process_WU.ReadOnly = True
        Me.ColAll_Process_WU.Width = 95
        '
        'ColAll_Process_PrimaryProcess
        '
        Me.ColAll_Process_PrimaryProcess.HeaderText = "Primary Process"
        Me.ColAll_Process_PrimaryProcess.Name = "ColAll_Process_PrimaryProcess"
        Me.ColAll_Process_PrimaryProcess.ReadOnly = True
        Me.ColAll_Process_PrimaryProcess.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_PrimaryProcess.Visible = False
        Me.ColAll_Process_PrimaryProcess.Width = 117
        '
        'ColAll_Process_PrimaryCodeNo
        '
        Me.ColAll_Process_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColAll_Process_PrimaryCodeNo.Name = "ColAll_Process_PrimaryCodeNo"
        Me.ColAll_Process_PrimaryCodeNo.ReadOnly = True
        Me.ColAll_Process_PrimaryCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_PrimaryCodeNo.Visible = False
        Me.ColAll_Process_PrimaryCodeNo.Width = 121
        '
        'ColAll_IdentifierProcess
        '
        Me.ColAll_IdentifierProcess.HeaderText = "Identifier"
        Me.ColAll_IdentifierProcess.Name = "ColAll_IdentifierProcess"
        Me.ColAll_IdentifierProcess.ReadOnly = True
        Me.ColAll_IdentifierProcess.Visible = False
        Me.ColAll_IdentifierProcess.Width = 87
        '
        'ColAll_Process_PrimaryWU
        '
        Me.ColAll_Process_PrimaryWU.HeaderText = "Primary WU"
        Me.ColAll_Process_PrimaryWU.Name = "ColAll_Process_PrimaryWU"
        Me.ColAll_Process_PrimaryWU.ReadOnly = True
        Me.ColAll_Process_PrimaryWU.Visible = False
        Me.ColAll_Process_PrimaryWU.Width = 108
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(446, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "The following work unit is not found in product data master, please fix:"
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(444, 113)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(106, 40)
        Me.btnClose.TabIndex = 3
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Button1"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'FrmTrans_ProductionSchedule_NewSch3_Exception
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(832, 491)
        Me.Controls.Add(Me.DgvItem)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label1)
        Me.MinimizeBox = False
        Me.Name = "FrmTrans_ProductionSchedule_NewSch3_Exception"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Exception"
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColAll_Process_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_WU As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_IdentifierProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_PrimaryWU As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class
