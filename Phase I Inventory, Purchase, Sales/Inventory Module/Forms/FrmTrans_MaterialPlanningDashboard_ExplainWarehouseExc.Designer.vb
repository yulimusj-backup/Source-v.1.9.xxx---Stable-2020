<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_ExplainWarehouseExc
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTrans_MaterialPlanningDashboard_ExplainWarehouseExc))
        Me.dgvWarehouse = New System.Windows.Forms.DataGridView
        Me.ColWarehouse_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColWarehouse_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse_ExcludeInclude = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColWarehouse_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnOK = New System.Windows.Forms.Button
        CType(Me.dgvWarehouse, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvWarehouse
        '
        Me.dgvWarehouse.AllowUserToAddRows = False
        Me.dgvWarehouse.AllowUserToDeleteRows = False
        Me.dgvWarehouse.AllowUserToResizeColumns = False
        Me.dgvWarehouse.AllowUserToResizeRows = False
        Me.dgvWarehouse.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvWarehouse.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvWarehouse.ColumnHeadersHeight = 25
        Me.dgvWarehouse.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvWarehouse.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWarehouse_No, Me.ColWarehouse_Level, Me.ColWarehouse_Code, Me.ColWarehouse_Desc, Me.ColWarehouse_ExcludeInclude, Me.ColWarehouse_Remarks, Me.ColWarehouse_PrimaryAcuan})
        Me.dgvWarehouse.Location = New System.Drawing.Point(13, 13)
        Me.dgvWarehouse.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvWarehouse.MultiSelect = False
        Me.dgvWarehouse.Name = "dgvWarehouse"
        Me.dgvWarehouse.ReadOnly = True
        Me.dgvWarehouse.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvWarehouse.RowTemplate.Height = 24
        Me.dgvWarehouse.Size = New System.Drawing.Size(666, 468)
        Me.dgvWarehouse.TabIndex = 9
        '
        'ColWarehouse_No
        '
        Me.ColWarehouse_No.Frozen = True
        Me.ColWarehouse_No.HeaderText = "No"
        Me.ColWarehouse_No.Name = "ColWarehouse_No"
        Me.ColWarehouse_No.ReadOnly = True
        Me.ColWarehouse_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWarehouse_No.Width = 32
        '
        'ColWarehouse_Level
        '
        Me.ColWarehouse_Level.Frozen = True
        Me.ColWarehouse_Level.HeaderText = "Level"
        Me.ColWarehouse_Level.Items.AddRange(New Object() {"Warehouse", "Location"})
        Me.ColWarehouse_Level.Name = "ColWarehouse_Level"
        Me.ColWarehouse_Level.ReadOnly = True
        Me.ColWarehouse_Level.Width = 48
        '
        'ColWarehouse_Code
        '
        Me.ColWarehouse_Code.Frozen = True
        Me.ColWarehouse_Code.HeaderText = "Code"
        Me.ColWarehouse_Code.Name = "ColWarehouse_Code"
        Me.ColWarehouse_Code.ReadOnly = True
        Me.ColWarehouse_Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWarehouse_Code.Width = 47
        '
        'ColWarehouse_Desc
        '
        Me.ColWarehouse_Desc.HeaderText = "Description"
        Me.ColWarehouse_Desc.Name = "ColWarehouse_Desc"
        Me.ColWarehouse_Desc.ReadOnly = True
        Me.ColWarehouse_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWarehouse_Desc.Width = 85
        '
        'ColWarehouse_ExcludeInclude
        '
        Me.ColWarehouse_ExcludeInclude.HeaderText = "Include/Exclude"
        Me.ColWarehouse_ExcludeInclude.Items.AddRange(New Object() {"Include", "Exclude"})
        Me.ColWarehouse_ExcludeInclude.Name = "ColWarehouse_ExcludeInclude"
        Me.ColWarehouse_ExcludeInclude.ReadOnly = True
        Me.ColWarehouse_ExcludeInclude.Width = 112
        '
        'ColWarehouse_Remarks
        '
        Me.ColWarehouse_Remarks.HeaderText = "Remarks"
        Me.ColWarehouse_Remarks.Name = "ColWarehouse_Remarks"
        Me.ColWarehouse_Remarks.ReadOnly = True
        Me.ColWarehouse_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWarehouse_Remarks.Width = 70
        '
        'ColWarehouse_PrimaryAcuan
        '
        Me.ColWarehouse_PrimaryAcuan.HeaderText = "Primary Acuan"
        Me.ColWarehouse_PrimaryAcuan.Name = "ColWarehouse_PrimaryAcuan"
        Me.ColWarehouse_PrimaryAcuan.ReadOnly = True
        Me.ColWarehouse_PrimaryAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWarehouse_PrimaryAcuan.Visible = False
        Me.ColWarehouse_PrimaryAcuan.Width = 106
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnOK.Location = New System.Drawing.Point(604, 488)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 10
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmTrans_MaterialPlanningDashboard_ExplainWarehouseExc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnOK
        Me.ClientSize = New System.Drawing.Size(692, 523)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.dgvWarehouse)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(700, 560)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_ExplainWarehouseExc"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Warehouse Exclusion Detail"
        CType(Me.dgvWarehouse, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvWarehouse As System.Windows.Forms.DataGridView
    Friend WithEvents ColWarehouse_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColWarehouse_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse_ExcludeInclude As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColWarehouse_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnOK As System.Windows.Forms.Button
End Class
