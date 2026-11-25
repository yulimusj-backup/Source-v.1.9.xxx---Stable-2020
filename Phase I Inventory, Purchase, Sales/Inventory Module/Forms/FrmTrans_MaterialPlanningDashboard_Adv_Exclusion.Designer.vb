<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_Adv_Exclusion
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
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.rboTipe_Exclusion = New System.Windows.Forms.RadioButton
        Me.rboTipe_Inclusion = New System.Windows.Forms.RadioButton
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkUseWarehouseExclusion = New System.Windows.Forms.CheckBox
        Me.dgvWarehouse = New System.Windows.Forms.DataGridView
        Me.ColWarehouse_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColWarehouse_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse_ExcludeInclude = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColWarehouse_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboSupplierSelectionMode = New System.Windows.Forms.ComboBox
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvWarehouse, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(660, 557)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(87, 31)
        Me.btnCancel.TabIndex = 0
        Me.btnCancel.TabStop = False
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        Me.btnCancel.Visible = False
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(660, 557)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(87, 31)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeight = 25
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(7, 34)
        Me.dgvItemProductGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(713, 438)
        Me.dgvItemProductGroup.TabIndex = 3
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.Frozen = True
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_No.Width = 32
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.Frozen = True
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.Width = 48
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.Frozen = True
        Me.ColProductGroup_CodeNo.HeaderText = "Group Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_CodeNo.Width = 91
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Desc.Width = 85
        '
        'ColProductGroup_Remarks
        '
        Me.ColProductGroup_Remarks.HeaderText = "Remarks"
        Me.ColProductGroup_Remarks.Name = "ColProductGroup_Remarks"
        Me.ColProductGroup_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Remarks.Width = 70
        '
        'ColProductGroup_PrimaryAcuan
        '
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary Acuan"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 106
        '
        'BackgroundWorker3
        '
        '
        'rboTipe_Exclusion
        '
        Me.rboTipe_Exclusion.AutoSize = True
        Me.rboTipe_Exclusion.Checked = True
        Me.rboTipe_Exclusion.Location = New System.Drawing.Point(3, 6)
        Me.rboTipe_Exclusion.Name = "rboTipe_Exclusion"
        Me.rboTipe_Exclusion.Size = New System.Drawing.Size(190, 21)
        Me.rboTipe_Exclusion.TabIndex = 4
        Me.rboTipe_Exclusion.TabStop = True
        Me.rboTipe_Exclusion.Text = "Exclude the following item"
        Me.rboTipe_Exclusion.UseVisualStyleBackColor = True
        '
        'rboTipe_Inclusion
        '
        Me.rboTipe_Inclusion.AutoSize = True
        Me.rboTipe_Inclusion.Location = New System.Drawing.Point(264, 6)
        Me.rboTipe_Inclusion.Name = "rboTipe_Inclusion"
        Me.rboTipe_Inclusion.Size = New System.Drawing.Size(230, 21)
        Me.rboTipe_Inclusion.TabIndex = 5
        Me.rboTipe_Inclusion.Text = "Only calculate the following item"
        Me.rboTipe_Inclusion.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 43)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(735, 508)
        Me.TabControl1.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.rboTipe_Inclusion)
        Me.TabPage1.Controls.Add(Me.dgvItemProductGroup)
        Me.TabPage1.Controls.Add(Me.rboTipe_Exclusion)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(727, 479)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Product Group Scope"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label1)
        Me.TabPage2.Controls.Add(Me.chkUseWarehouseExclusion)
        Me.TabPage2.Controls.Add(Me.dgvWarehouse)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(727, 479)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Warehouse/Location Scope"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 454)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(670, 17)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Note: When location and warehouse data are in conflict, location data will be use" & _
            " as calculation reference"
        '
        'chkUseWarehouseExclusion
        '
        Me.chkUseWarehouseExclusion.AutoSize = True
        Me.chkUseWarehouseExclusion.Location = New System.Drawing.Point(7, 6)
        Me.chkUseWarehouseExclusion.Name = "chkUseWarehouseExclusion"
        Me.chkUseWarehouseExclusion.Size = New System.Drawing.Size(383, 21)
        Me.chkUseWarehouseExclusion.TabIndex = 9
        Me.chkUseWarehouseExclusion.Text = "Only Calculate Qty On Hand For Selected Location Only"
        Me.chkUseWarehouseExclusion.UseVisualStyleBackColor = True
        '
        'dgvWarehouse
        '
        Me.dgvWarehouse.AllowUserToResizeColumns = False
        Me.dgvWarehouse.AllowUserToResizeRows = False
        Me.dgvWarehouse.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvWarehouse.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvWarehouse.ColumnHeadersHeight = 25
        Me.dgvWarehouse.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvWarehouse.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWarehouse_No, Me.ColWarehouse_Level, Me.ColWarehouse_Code, Me.ColWarehouse_Desc, Me.ColWarehouse_ExcludeInclude, Me.ColWarehouse_Remarks, Me.ColWarehouse_PrimaryAcuan})
        Me.dgvWarehouse.Location = New System.Drawing.Point(7, 34)
        Me.dgvWarehouse.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvWarehouse.MultiSelect = False
        Me.dgvWarehouse.Name = "dgvWarehouse"
        Me.dgvWarehouse.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvWarehouse.RowTemplate.Height = 24
        Me.dgvWarehouse.Size = New System.Drawing.Size(713, 407)
        Me.dgvWarehouse.TabIndex = 8
        '
        'ColWarehouse_No
        '
        Me.ColWarehouse_No.Frozen = True
        Me.ColWarehouse_No.HeaderText = "No"
        Me.ColWarehouse_No.Name = "ColWarehouse_No"
        Me.ColWarehouse_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWarehouse_No.Width = 32
        '
        'ColWarehouse_Level
        '
        Me.ColWarehouse_Level.Frozen = True
        Me.ColWarehouse_Level.HeaderText = "Level"
        Me.ColWarehouse_Level.Items.AddRange(New Object() {"Warehouse", "Location"})
        Me.ColWarehouse_Level.Name = "ColWarehouse_Level"
        Me.ColWarehouse_Level.Width = 48
        '
        'ColWarehouse_Code
        '
        Me.ColWarehouse_Code.Frozen = True
        Me.ColWarehouse_Code.HeaderText = "Code"
        Me.ColWarehouse_Code.Name = "ColWarehouse_Code"
        Me.ColWarehouse_Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWarehouse_Code.Width = 47
        '
        'ColWarehouse_Desc
        '
        Me.ColWarehouse_Desc.HeaderText = "Description"
        Me.ColWarehouse_Desc.Name = "ColWarehouse_Desc"
        Me.ColWarehouse_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWarehouse_Desc.Width = 85
        '
        'ColWarehouse_ExcludeInclude
        '
        Me.ColWarehouse_ExcludeInclude.HeaderText = "Include/Exclude"
        Me.ColWarehouse_ExcludeInclude.Items.AddRange(New Object() {"Include", "Exclude"})
        Me.ColWarehouse_ExcludeInclude.Name = "ColWarehouse_ExcludeInclude"
        Me.ColWarehouse_ExcludeInclude.Width = 112
        '
        'ColWarehouse_Remarks
        '
        Me.ColWarehouse_Remarks.HeaderText = "Remarks"
        Me.ColWarehouse_Remarks.Name = "ColWarehouse_Remarks"
        Me.ColWarehouse_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWarehouse_Remarks.Width = 70
        '
        'ColWarehouse_PrimaryAcuan
        '
        Me.ColWarehouse_PrimaryAcuan.HeaderText = "Primary Acuan"
        Me.ColWarehouse_PrimaryAcuan.Name = "ColWarehouse_PrimaryAcuan"
        Me.ColWarehouse_PrimaryAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWarehouse_PrimaryAcuan.Visible = False
        Me.ColWarehouse_PrimaryAcuan.Width = 106
        '
        'BackgroundWorker1
        '
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(161, 17)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Supplier Selection Mode"
        '
        'cboSupplierSelectionMode
        '
        Me.cboSupplierSelectionMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSupplierSelectionMode.FormattingEnabled = True
        Me.cboSupplierSelectionMode.Items.AddRange(New Object() {"By Product Setting", "The cheapest in the last 3 order", "Prefered supplier", "Most recent PO"})
        Me.cboSupplierSelectionMode.Location = New System.Drawing.Point(179, 12)
        Me.cboSupplierSelectionMode.Name = "cboSupplierSelectionMode"
        Me.cboSupplierSelectionMode.Size = New System.Drawing.Size(219, 24)
        Me.cboSupplierSelectionMode.TabIndex = 1
        '
        'FrmTrans_MaterialPlanningDashboard_Adv_Exclusion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(759, 600)
        Me.Controls.Add(Me.cboSupplierSelectionMode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.TabControl1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(667, 519)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_Adv_Exclusion"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Additional Options"
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.dgvWarehouse, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents rboTipe_Exclusion As System.Windows.Forms.RadioButton
    Friend WithEvents rboTipe_Inclusion As System.Windows.Forms.RadioButton
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents chkUseWarehouseExclusion As System.Windows.Forms.CheckBox
    Friend WithEvents dgvWarehouse As System.Windows.Forms.DataGridView
    Friend WithEvents ColWarehouse_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColWarehouse_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse_ExcludeInclude As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColWarehouse_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboSupplierSelectionMode As System.Windows.Forms.ComboBox
End Class
