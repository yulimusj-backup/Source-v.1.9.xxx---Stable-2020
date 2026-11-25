<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_NewSch
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColPriority = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyTarget = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyProduce = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTargetFinishDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchedule = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimaryMO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.btnUp = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.llbColumnsSettings = New System.Windows.Forms.LinkLabel
        Me.TxtFilter_PrimaryMO = New System.Windows.Forms.TextBox
        Me.TxtFilter_MONo = New System.Windows.Forms.TextBox
        Me.LblFilter_FindMO = New System.Windows.Forms.Label
        Me.btnFilter_FindMO = New System.Windows.Forms.Button
        Me.btnFilter_FindCodeNo = New System.Windows.Forms.Button
        Me.TxtFilter_Desc = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtFilter_CodeNo = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.llbMP_Filter = New System.Windows.Forms.LinkLabel
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(499, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "The following item is ready for scheduling, please make sure the data is valid: "
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColPriority, Me.ColJO, Me.ColMO, Me.ColCodeNo, Me.ColDesc, Me.ColQtyTarget, Me.ColQtyProduce, Me.ColUnit, Me.ColTargetFinishDate, Me.ColSchedule, Me.ColPrimaryMO, Me.ColIdentifierJO})
        Me.dgvItem.Location = New System.Drawing.Point(12, 89)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(968, 491)
        Me.dgvItem.TabIndex = 1
        '
        'ColPriority
        '
        Me.ColPriority.HeaderText = "Priority"
        Me.ColPriority.Name = "ColPriority"
        Me.ColPriority.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPriority.Width = 58
        '
        'ColJO
        '
        Me.ColJO.HeaderText = "JO No"
        Me.ColJO.Name = "ColJO"
        Me.ColJO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJO.Width = 54
        '
        'ColMO
        '
        Me.ColMO.HeaderText = "MO No"
        Me.ColMO.Name = "ColMO"
        Me.ColMO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMO.Width = 58
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCodeNo.Width = 69
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Width = 85
        '
        'ColQtyTarget
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyTarget.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQtyTarget.HeaderText = "Qty Target"
        Me.ColQtyTarget.Name = "ColQtyTarget"
        Me.ColQtyTarget.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyTarget.Width = 82
        '
        'ColQtyProduce
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyProduce.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyProduce.HeaderText = "Qty Produce"
        Me.ColQtyProduce.Name = "ColQtyProduce"
        Me.ColQtyProduce.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyProduce.Width = 93
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 39
        '
        'ColTargetFinishDate
        '
        DataGridViewCellStyle3.Format = "dd MMM yyyy"
        Me.ColTargetFinishDate.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColTargetFinishDate.HeaderText = "Target Finish Date"
        Me.ColTargetFinishDate.Name = "ColTargetFinishDate"
        Me.ColTargetFinishDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTargetFinishDate.Width = 131
        '
        'ColSchedule
        '
        Me.ColSchedule.HeaderText = "Schedule"
        Me.ColSchedule.Name = "ColSchedule"
        Me.ColSchedule.Width = 73
        '
        'ColPrimaryMO
        '
        Me.ColPrimaryMO.HeaderText = "Primary MO"
        Me.ColPrimaryMO.Name = "ColPrimaryMO"
        Me.ColPrimaryMO.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPrimaryMO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryMO.Visible = False
        Me.ColPrimaryMO.Width = 88
        '
        'ColIdentifierJO
        '
        Me.ColIdentifierJO.HeaderText = "Identifier JO"
        Me.ColIdentifierJO.Name = "ColIdentifierJO"
        Me.ColIdentifierJO.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColIdentifierJO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColIdentifierJO.Visible = False
        Me.ColIdentifierJO.Width = 90
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(905, 585)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 36)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(824, 585)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 36)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.Location = New System.Drawing.Point(899, 64)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(82, 17)
        Me.llbUncheckAll.TabIndex = 4
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        Me.llbUncheckAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbCheckAll
        '
        Me.llbCheckAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(827, 64)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(66, 17)
        Me.llbCheckAll.TabIndex = 5
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        Me.llbCheckAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'btnUp
        '
        Me.btnUp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUp.BackgroundImage = Global.Inventory_Module.My.Resources.Resources.arrow_up
        Me.btnUp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnUp.Location = New System.Drawing.Point(12, 585)
        Me.btnUp.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnUp.Name = "btnUp"
        Me.btnUp.Size = New System.Drawing.Size(37, 36)
        Me.btnUp.TabIndex = 7
        Me.btnUp.TabStop = False
        Me.btnUp.UseVisualStyleBackColor = True
        Me.btnUp.Visible = False
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackgroundImage = Global.Inventory_Module.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.Location = New System.Drawing.Point(55, 585)
        Me.btnDown.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(37, 36)
        Me.btnDown.TabIndex = 6
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        Me.btnDown.Visible = False
        '
        'llbColumnsSettings
        '
        Me.llbColumnsSettings.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbColumnsSettings.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbColumnsSettings.AutoSize = True
        Me.llbColumnsSettings.Location = New System.Drawing.Point(9, 594)
        Me.llbColumnsSettings.Name = "llbColumnsSettings"
        Me.llbColumnsSettings.Size = New System.Drawing.Size(110, 17)
        Me.llbColumnsSettings.TabIndex = 8
        Me.llbColumnsSettings.TabStop = True
        Me.llbColumnsSettings.Text = "Column Settings"
        Me.llbColumnsSettings.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'TxtFilter_PrimaryMO
        '
        Me.TxtFilter_PrimaryMO.Location = New System.Drawing.Point(553, 32)
        Me.TxtFilter_PrimaryMO.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtFilter_PrimaryMO.Name = "TxtFilter_PrimaryMO"
        Me.TxtFilter_PrimaryMO.Size = New System.Drawing.Size(28, 22)
        Me.TxtFilter_PrimaryMO.TabIndex = 120
        Me.TxtFilter_PrimaryMO.TabStop = False
        Me.TxtFilter_PrimaryMO.Text = "-1"
        Me.TxtFilter_PrimaryMO.Visible = False
        '
        'TxtFilter_MONo
        '
        Me.TxtFilter_MONo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtFilter_MONo.Location = New System.Drawing.Point(579, 30)
        Me.TxtFilter_MONo.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtFilter_MONo.Name = "TxtFilter_MONo"
        Me.TxtFilter_MONo.Size = New System.Drawing.Size(224, 22)
        Me.TxtFilter_MONo.TabIndex = 117
        '
        'LblFilter_FindMO
        '
        Me.LblFilter_FindMO.AutoSize = True
        Me.LblFilter_FindMO.Location = New System.Drawing.Point(452, 32)
        Me.LblFilter_FindMO.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblFilter_FindMO.Name = "LblFilter_FindMO"
        Me.LblFilter_FindMO.Size = New System.Drawing.Size(52, 17)
        Me.LblFilter_FindMO.TabIndex = 119
        Me.LblFilter_FindMO.Text = "MO No"
        '
        'btnFilter_FindMO
        '
        Me.btnFilter_FindMO.Location = New System.Drawing.Point(812, 30)
        Me.btnFilter_FindMO.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFilter_FindMO.Name = "btnFilter_FindMO"
        Me.btnFilter_FindMO.Size = New System.Drawing.Size(52, 25)
        Me.btnFilter_FindMO.TabIndex = 118
        Me.btnFilter_FindMO.TabStop = False
        Me.btnFilter_FindMO.Text = "Find"
        Me.btnFilter_FindMO.UseVisualStyleBackColor = True
        '
        'btnFilter_FindCodeNo
        '
        Me.btnFilter_FindCodeNo.Location = New System.Drawing.Point(373, 26)
        Me.btnFilter_FindCodeNo.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFilter_FindCodeNo.Name = "btnFilter_FindCodeNo"
        Me.btnFilter_FindCodeNo.Size = New System.Drawing.Size(52, 28)
        Me.btnFilter_FindCodeNo.TabIndex = 116
        Me.btnFilter_FindCodeNo.TabStop = False
        Me.btnFilter_FindCodeNo.Text = "Find"
        Me.btnFilter_FindCodeNo.UseVisualStyleBackColor = True
        '
        'TxtFilter_Desc
        '
        Me.TxtFilter_Desc.Location = New System.Drawing.Point(140, 61)
        Me.TxtFilter_Desc.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtFilter_Desc.Name = "TxtFilter_Desc"
        Me.TxtFilter_Desc.Size = New System.Drawing.Size(284, 22)
        Me.TxtFilter_Desc.TabIndex = 113
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 64)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(79, 17)
        Me.Label2.TabIndex = 115
        Me.Label2.Text = "Description"
        '
        'TxtFilter_CodeNo
        '
        Me.TxtFilter_CodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtFilter_CodeNo.Location = New System.Drawing.Point(140, 30)
        Me.TxtFilter_CodeNo.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtFilter_CodeNo.Name = "TxtFilter_CodeNo"
        Me.TxtFilter_CodeNo.Size = New System.Drawing.Size(224, 22)
        Me.TxtFilter_CodeNo.TabIndex = 112
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(13, 34)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(63, 17)
        Me.Label8.TabIndex = 114
        Me.Label8.Text = "Code No"
        '
        'llbMP_Filter
        '
        Me.llbMP_Filter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbMP_Filter.AutoSize = True
        Me.llbMP_Filter.Location = New System.Drawing.Point(783, 64)
        Me.llbMP_Filter.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbMP_Filter.Name = "llbMP_Filter"
        Me.llbMP_Filter.Size = New System.Drawing.Size(39, 17)
        Me.llbMP_Filter.TabIndex = 122
        Me.llbMP_Filter.TabStop = True
        Me.llbMP_Filter.Text = "Filter"
        Me.llbMP_Filter.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'FrmTrans_ProductionSchedule_NewSch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(995, 634)
        Me.Controls.Add(Me.TxtFilter_PrimaryMO)
        Me.Controls.Add(Me.btnFilter_FindCodeNo)
        Me.Controls.Add(Me.TxtFilter_MONo)
        Me.Controls.Add(Me.LblFilter_FindMO)
        Me.Controls.Add(Me.TxtFilter_Desc)
        Me.Controls.Add(Me.btnFilter_FindMO)
        Me.Controls.Add(Me.llbMP_Filter)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtFilter_CodeNo)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.llbColumnsSettings)
        Me.Controls.Add(Me.btnUp)
        Me.Controls.Add(Me.btnDown)
        Me.Controls.Add(Me.llbCheckAll)
        Me.Controls.Add(Me.llbUncheckAll)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1003, 671)
        Me.Name = "FrmTrans_ProductionSchedule_NewSch"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Schedule Item"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents btnUp As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents ColPriority As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyTarget As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyProduce As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTargetFinishDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchedule As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimaryMO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents llbColumnsSettings As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtFilter_MONo As System.Windows.Forms.TextBox
    Friend WithEvents LblFilter_FindMO As System.Windows.Forms.Label
    Friend WithEvents btnFilter_FindMO As System.Windows.Forms.Button
    Friend WithEvents btnFilter_FindCodeNo As System.Windows.Forms.Button
    Friend WithEvents TxtFilter_Desc As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtFilter_CodeNo As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents llbMP_Filter As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtFilter_PrimaryMO As System.Windows.Forms.TextBox
End Class
