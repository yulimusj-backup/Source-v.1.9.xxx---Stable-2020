<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_work_unitionSchedule_EditOT
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
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker
        Me.dgvOT = New System.Windows.Forms.DataGridView
        Me.ColWorkUnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkUnitDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySPL = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSPLNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDelete = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryWU = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvFilterWU = New System.Windows.Forms.DataGridView
        Me.ColFilterWU_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterWU_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColFilterWU_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterWU_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterWU_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnFindSPL = New System.Windows.Forms.Button
        Me.TxtSPLNo = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnAddBySPL = New System.Windows.Forms.Button
        Me.btnRemoveBySPL = New System.Windows.Forms.Button
        Me.btnSetOvertime = New System.Windows.Forms.Button
        Me.btnRemoveOvertime = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.lblFirmPlanDate = New System.Windows.Forms.Label
        Me.TxtPrimarySPL = New System.Windows.Forms.TextBox
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.BackgroundWorker_CloseForm = New System.ComponentModel.BackgroundWorker
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.BtnCancel = New System.Windows.Forms.Button
        CType(Me.dgvOT, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvFilterWU, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dtpStartDate
        '
        Me.dtpStartDate.CustomFormat = "dd MMM yyyy"
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStartDate.Location = New System.Drawing.Point(91, 32)
        Me.dtpStartDate.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(201, 22)
        Me.dtpStartDate.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Start"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 63)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(33, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "End"
        '
        'dtpEndDate
        '
        Me.dtpEndDate.CustomFormat = "dd MMM yyyy"
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEndDate.Location = New System.Drawing.Point(91, 60)
        Me.dtpEndDate.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(201, 22)
        Me.dtpEndDate.TabIndex = 2
        '
        'dgvOT
        '
        Me.dgvOT.AllowUserToAddRows = False
        Me.dgvOT.AllowUserToDeleteRows = False
        Me.dgvOT.AllowUserToResizeColumns = False
        Me.dgvOT.AllowUserToResizeRows = False
        Me.dgvOT.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvOT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvOT.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWorkUnitCode, Me.ColWorkUnitDesc, Me.ColStartDate, Me.ColStartTime, Me.ColEndDate, Me.ColEndTime, Me.ColRemarks, Me.ColPrimarySPL, Me.ColSPLNo, Me.ColDelete, Me.ColPrimary, Me.ColPrimaryWU})
        Me.dgvOT.Location = New System.Drawing.Point(12, 224)
        Me.dgvOT.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvOT.Name = "dgvOT"
        Me.dgvOT.RowHeadersVisible = False
        Me.dgvOT.RowTemplate.Height = 24
        Me.dgvOT.Size = New System.Drawing.Size(892, 405)
        Me.dgvOT.TabIndex = 4
        '
        'ColWorkUnitCode
        '
        Me.ColWorkUnitCode.HeaderText = "Code"
        Me.ColWorkUnitCode.Name = "ColWorkUnitCode"
        Me.ColWorkUnitCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColWorkUnitDesc
        '
        Me.ColWorkUnitDesc.HeaderText = "Description"
        Me.ColWorkUnitDesc.Name = "ColWorkUnitDesc"
        Me.ColWorkUnitDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColStartDate
        '
        Me.ColStartDate.HeaderText = "Start Date"
        Me.ColStartDate.Name = "ColStartDate"
        Me.ColStartDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColStartTime
        '
        Me.ColStartTime.HeaderText = ""
        Me.ColStartTime.Name = "ColStartTime"
        Me.ColStartTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColEndDate
        '
        Me.ColEndDate.HeaderText = "End Date"
        Me.ColEndDate.Name = "ColEndDate"
        Me.ColEndDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColEndTime
        '
        Me.ColEndTime.HeaderText = ""
        Me.ColEndTime.Name = "ColEndTime"
        Me.ColEndTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Visible = False
        '
        'ColPrimarySPL
        '
        Me.ColPrimarySPL.HeaderText = "Primary SPL"
        Me.ColPrimarySPL.Name = "ColPrimarySPL"
        Me.ColPrimarySPL.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimarySPL.Visible = False
        '
        'ColSPLNo
        '
        Me.ColSPLNo.HeaderText = "SPL No"
        Me.ColSPLNo.Name = "ColSPLNo"
        Me.ColSPLNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSPLNo.Visible = False
        '
        'ColDelete
        '
        Me.ColDelete.HeaderText = "Delete"
        Me.ColDelete.Name = "ColDelete"
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.Visible = False
        '
        'ColPrimaryWU
        '
        Me.ColPrimaryWU.HeaderText = "PrimaryWU"
        Me.ColPrimaryWU.Name = "ColPrimaryWU"
        Me.ColPrimaryWU.Visible = False
        '
        'dgvFilterWU
        '
        Me.dgvFilterWU.AllowUserToResizeColumns = False
        Me.dgvFilterWU.AllowUserToResizeRows = False
        Me.dgvFilterWU.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvFilterWU.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvFilterWU.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColFilterWU_No, Me.ColFilterWU_Level, Me.ColFilterWU_Code, Me.ColFilterWU_Desc, Me.ColFilterWU_Primary})
        Me.dgvFilterWU.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvFilterWU.Location = New System.Drawing.Point(3, 17)
        Me.dgvFilterWU.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvFilterWU.Name = "dgvFilterWU"
        Me.dgvFilterWU.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvFilterWU.RowTemplate.Height = 24
        Me.dgvFilterWU.Size = New System.Drawing.Size(601, 187)
        Me.dgvFilterWU.TabIndex = 8
        '
        'ColFilterWU_No
        '
        Me.ColFilterWU_No.HeaderText = "No"
        Me.ColFilterWU_No.Name = "ColFilterWU_No"
        Me.ColFilterWU_No.Width = 51
        '
        'ColFilterWU_Level
        '
        Me.ColFilterWU_Level.HeaderText = "Level"
        Me.ColFilterWU_Level.Items.AddRange(New Object() {"Work Unit Group", "Work Unit"})
        Me.ColFilterWU_Level.Name = "ColFilterWU_Level"
        Me.ColFilterWU_Level.Width = 48
        '
        'ColFilterWU_Code
        '
        Me.ColFilterWU_Code.HeaderText = "Code"
        Me.ColFilterWU_Code.Name = "ColFilterWU_Code"
        Me.ColFilterWU_Code.Width = 66
        '
        'ColFilterWU_Desc
        '
        Me.ColFilterWU_Desc.HeaderText = "Description"
        Me.ColFilterWU_Desc.Name = "ColFilterWU_Desc"
        Me.ColFilterWU_Desc.Width = 104
        '
        'ColFilterWU_Primary
        '
        Me.ColFilterWU_Primary.HeaderText = "Primary"
        Me.ColFilterWU_Primary.Name = "ColFilterWU_Primary"
        Me.ColFilterWU_Primary.Visible = False
        Me.ColFilterWU_Primary.Width = 81
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 91)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 17)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "SPL"
        Me.Label3.Visible = False
        '
        'btnFindSPL
        '
        Me.btnFindSPL.Location = New System.Drawing.Point(237, 89)
        Me.btnFindSPL.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnFindSPL.Name = "btnFindSPL"
        Me.btnFindSPL.Size = New System.Drawing.Size(53, 22)
        Me.btnFindSPL.TabIndex = 10
        Me.btnFindSPL.TabStop = False
        Me.btnFindSPL.Text = "Find"
        Me.btnFindSPL.UseVisualStyleBackColor = True
        Me.btnFindSPL.Visible = False
        '
        'TxtSPLNo
        '
        Me.TxtSPLNo.Location = New System.Drawing.Point(91, 89)
        Me.TxtSPLNo.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtSPLNo.Name = "TxtSPLNo"
        Me.TxtSPLNo.Size = New System.Drawing.Size(141, 22)
        Me.TxtSPLNo.TabIndex = 11
        Me.TxtSPLNo.TabStop = False
        Me.TxtSPLNo.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dgvFilterWU)
        Me.GroupBox1.Location = New System.Drawing.Point(297, 12)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(607, 206)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Work Unit"
        '
        'btnAddBySPL
        '
        Me.btnAddBySPL.Location = New System.Drawing.Point(12, 116)
        Me.btnAddBySPL.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnAddBySPL.Name = "btnAddBySPL"
        Me.btnAddBySPL.Size = New System.Drawing.Size(133, 25)
        Me.btnAddBySPL.TabIndex = 13
        Me.btnAddBySPL.TabStop = False
        Me.btnAddBySPL.Text = "Add By SPL"
        Me.btnAddBySPL.UseVisualStyleBackColor = True
        Me.btnAddBySPL.Visible = False
        '
        'btnRemoveBySPL
        '
        Me.btnRemoveBySPL.Location = New System.Drawing.Point(151, 116)
        Me.btnRemoveBySPL.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnRemoveBySPL.Name = "btnRemoveBySPL"
        Me.btnRemoveBySPL.Size = New System.Drawing.Size(133, 25)
        Me.btnRemoveBySPL.TabIndex = 15
        Me.btnRemoveBySPL.TabStop = False
        Me.btnRemoveBySPL.Text = "Remove By SPL"
        Me.btnRemoveBySPL.UseVisualStyleBackColor = True
        Me.btnRemoveBySPL.Visible = False
        '
        'btnSetOvertime
        '
        Me.btnSetOvertime.Location = New System.Drawing.Point(12, 89)
        Me.btnSetOvertime.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnSetOvertime.Name = "btnSetOvertime"
        Me.btnSetOvertime.Size = New System.Drawing.Size(133, 25)
        Me.btnSetOvertime.TabIndex = 16
        Me.btnSetOvertime.Text = "Add Overtime"
        Me.btnSetOvertime.UseVisualStyleBackColor = True
        '
        'btnRemoveOvertime
        '
        Me.btnRemoveOvertime.Location = New System.Drawing.Point(151, 89)
        Me.btnRemoveOvertime.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnRemoveOvertime.Name = "btnRemoveOvertime"
        Me.btnRemoveOvertime.Size = New System.Drawing.Size(133, 25)
        Me.btnRemoveOvertime.TabIndex = 17
        Me.btnRemoveOvertime.Text = "Remove Overtime"
        Me.btnRemoveOvertime.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(12, 119)
        Me.btnFilter.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(133, 25)
        Me.btnFilter.TabIndex = 18
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 17)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Firm Plan"
        '
        'lblFirmPlanDate
        '
        Me.lblFirmPlanDate.AutoSize = True
        Me.lblFirmPlanDate.Location = New System.Drawing.Point(87, 12)
        Me.lblFirmPlanDate.Name = "lblFirmPlanDate"
        Me.lblFirmPlanDate.Size = New System.Drawing.Size(101, 17)
        Me.lblFirmPlanDate.TabIndex = 20
        Me.lblFirmPlanDate.Text = "Firm Plan Date"
        '
        'TxtPrimarySPL
        '
        Me.TxtPrimarySPL.Location = New System.Drawing.Point(52, 89)
        Me.TxtPrimarySPL.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtPrimarySPL.Name = "TxtPrimarySPL"
        Me.TxtPrimarySPL.Size = New System.Drawing.Size(33, 22)
        Me.TxtPrimarySPL.TabIndex = 21
        Me.TxtPrimarySPL.TabStop = False
        Me.TxtPrimarySPL.Text = "-1"
        Me.TxtPrimarySPL.Visible = False
        '
        'BackgroundWorker3
        '
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(443, 318)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(33, 22)
        Me.TxtReadOnly.TabIndex = 22
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'BackgroundWorker_CloseForm
        '
        '
        'llbCheckAll
        '
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(750, 633)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(66, 17)
        Me.llbCheckAll.TabIndex = 23
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.Location = New System.Drawing.Point(822, 633)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(82, 17)
        Me.llbUncheckAll.TabIndex = 24
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        '
        'BtnCancel
        '
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(391, 316)
        Me.BtnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(133, 25)
        Me.BtnCancel.TabIndex = 25
        Me.BtnCancel.Text = "Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'FrmTrans_work_unitionSchedule_EditOT
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(916, 658)
        Me.Controls.Add(Me.llbUncheckAll)
        Me.Controls.Add(Me.llbCheckAll)
        Me.Controls.Add(Me.TxtPrimarySPL)
        Me.Controls.Add(Me.lblFirmPlanDate)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.btnRemoveOvertime)
        Me.Controls.Add(Me.btnSetOvertime)
        Me.Controls.Add(Me.btnRemoveBySPL)
        Me.Controls.Add(Me.btnAddBySPL)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TxtSPLNo)
        Me.Controls.Add(Me.btnFindSPL)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dgvOT)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpEndDate)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpStartDate)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(924, 695)
        Me.Name = "FrmTrans_work_unitionSchedule_EditOT"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Overtime Data"
        CType(Me.dgvOT, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvFilterWU, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dgvOT As System.Windows.Forms.DataGridView
    Friend WithEvents dgvFilterWU As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnFindSPL As System.Windows.Forms.Button
    Friend WithEvents TxtSPLNo As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnAddBySPL As System.Windows.Forms.Button
    Friend WithEvents btnRemoveBySPL As System.Windows.Forms.Button
    Friend WithEvents btnSetOvertime As System.Windows.Forms.Button
    Friend WithEvents btnRemoveOvertime As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblFirmPlanDate As System.Windows.Forms.Label
    Friend WithEvents TxtPrimarySPL As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColFilterWU_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterWU_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColFilterWU_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterWU_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterWU_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker_CloseForm As System.ComponentModel.BackgroundWorker
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents ColWorkUnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkUnitDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySPL As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSPLNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDelete As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryWU As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
End Class
