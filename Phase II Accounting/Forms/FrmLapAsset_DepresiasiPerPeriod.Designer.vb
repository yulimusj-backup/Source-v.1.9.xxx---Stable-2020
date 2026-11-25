<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLapAsset_DepresiasiPerPeriod
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
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.chkShowDisposedAsset = New System.Windows.Forms.CheckBox
        Me.ChkShowNonDepreciateAsset = New System.Windows.Forms.CheckBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColPrimaryGroup = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGroup = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAsset = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAssetCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAssetDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAqcValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAcqDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccumatedDep = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDepreciation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRevaluation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDisposal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemainingValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColShow = New System.Windows.Forms.DataGridViewButtonColumn
        Me.TabControl1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(67, 12)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.ShowUpDown = True
        Me.dtpFrom.Size = New System.Drawing.Size(121, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 55
        Me.Label1.Text = "From"
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(650, 284)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 23)
        Me.btnClose.TabIndex = 53
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(366, 33)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(89, 23)
        Me.btnShow.TabIndex = 4
        Me.btnShow.Text = "Show Report"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(67, 38)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.ShowUpDown = True
        Me.dtpTo.Size = New System.Drawing.Size(121, 20)
        Me.dtpTo.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 57
        Me.Label2.Text = "To"
        '
        'chkShowDisposedAsset
        '
        Me.chkShowDisposedAsset.AutoSize = True
        Me.chkShowDisposedAsset.Location = New System.Drawing.Point(194, 14)
        Me.chkShowDisposedAsset.Name = "chkShowDisposedAsset"
        Me.chkShowDisposedAsset.Size = New System.Drawing.Size(129, 17)
        Me.chkShowDisposedAsset.TabIndex = 2
        Me.chkShowDisposedAsset.Text = "Show Disposed Asset"
        Me.chkShowDisposedAsset.UseVisualStyleBackColor = True
        '
        'ChkShowNonDepreciateAsset
        '
        Me.ChkShowNonDepreciateAsset.AutoSize = True
        Me.ChkShowNonDepreciateAsset.Location = New System.Drawing.Point(194, 37)
        Me.ChkShowNonDepreciateAsset.Name = "ChkShowNonDepreciateAsset"
        Me.ChkShowNonDepreciateAsset.Size = New System.Drawing.Size(166, 17)
        Me.ChkShowNonDepreciateAsset.TabIndex = 3
        Me.ChkShowNonDepreciateAsset.Text = "Show Non-Depreciated Asset"
        Me.ChkShowNonDepreciateAsset.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(12, 64)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(860, 486)
        Me.TabControl1.TabIndex = 5
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.crv1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(852, 540)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Printable"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.DisplayStatusBar = False
        Me.crv1.DisplayToolbar = False
        Me.crv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crv1.Location = New System.Drawing.Point(3, 3)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(846, 534)
        Me.crv1.TabIndex = 6
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvItem)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(852, 460)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Results"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColPrimaryGroup, Me.ColGroup, Me.ColPrimaryAsset, Me.ColAssetCode, Me.ColAssetDesc, Me.ColQty, Me.ColUnit, Me.ColAqcValue, Me.ColAcqDate, Me.ColAccumatedDep, Me.ColDepreciation, Me.ColRevaluation, Me.ColDisposal, Me.ColRemainingValue, Me.ColShow})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(3, 3)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.Size = New System.Drawing.Size(846, 454)
        Me.dgvItem.TabIndex = 0
        '
        'ColPrimaryGroup
        '
        Me.ColPrimaryGroup.HeaderText = "PrimaryGroup"
        Me.ColPrimaryGroup.Name = "ColPrimaryGroup"
        Me.ColPrimaryGroup.ReadOnly = True
        Me.ColPrimaryGroup.Visible = False
        Me.ColPrimaryGroup.Width = 95
        '
        'ColGroup
        '
        Me.ColGroup.HeaderText = "Group"
        Me.ColGroup.Name = "ColGroup"
        Me.ColGroup.ReadOnly = True
        Me.ColGroup.Width = 61
        '
        'ColPrimaryAsset
        '
        Me.ColPrimaryAsset.HeaderText = "PrimaryAsset"
        Me.ColPrimaryAsset.Name = "ColPrimaryAsset"
        Me.ColPrimaryAsset.ReadOnly = True
        Me.ColPrimaryAsset.Visible = False
        Me.ColPrimaryAsset.Width = 92
        '
        'ColAssetCode
        '
        Me.ColAssetCode.HeaderText = "Asset Code"
        Me.ColAssetCode.Name = "ColAssetCode"
        Me.ColAssetCode.ReadOnly = True
        Me.ColAssetCode.Width = 86
        '
        'ColAssetDesc
        '
        Me.ColAssetDesc.HeaderText = "Description"
        Me.ColAssetDesc.Name = "ColAssetDesc"
        Me.ColAssetDesc.ReadOnly = True
        Me.ColAssetDesc.Width = 85
        '
        'ColQty
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.Width = 48
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.ReadOnly = True
        Me.ColUnit.Width = 51
        '
        'ColAqcValue
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAqcValue.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColAqcValue.HeaderText = "Acquisition Value"
        Me.ColAqcValue.Name = "ColAqcValue"
        Me.ColAqcValue.ReadOnly = True
        Me.ColAqcValue.Width = 113
        '
        'ColAcqDate
        '
        Me.ColAcqDate.HeaderText = "AcquisitionDate"
        Me.ColAcqDate.Name = "ColAcqDate"
        Me.ColAcqDate.ReadOnly = True
        Me.ColAcqDate.Width = 106
        '
        'ColAccumatedDep
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAccumatedDep.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColAccumatedDep.HeaderText = "Accumulated Depreciation"
        Me.ColAccumatedDep.Name = "ColAccumatedDep"
        Me.ColAccumatedDep.ReadOnly = True
        Me.ColAccumatedDep.Width = 157
        '
        'ColDepreciation
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDepreciation.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColDepreciation.HeaderText = "Depreciation"
        Me.ColDepreciation.Name = "ColDepreciation"
        Me.ColDepreciation.ReadOnly = True
        Me.ColDepreciation.Width = 92
        '
        'ColRevaluation
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRevaluation.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColRevaluation.HeaderText = "Revaluation"
        Me.ColRevaluation.Name = "ColRevaluation"
        Me.ColRevaluation.ReadOnly = True
        Me.ColRevaluation.Width = 89
        '
        'ColDisposal
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDisposal.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColDisposal.HeaderText = "Disposal"
        Me.ColDisposal.Name = "ColDisposal"
        Me.ColDisposal.ReadOnly = True
        Me.ColDisposal.Width = 72
        '
        'ColRemainingValue
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRemainingValue.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColRemainingValue.HeaderText = "Remaining Book Value"
        Me.ColRemainingValue.Name = "ColRemainingValue"
        Me.ColRemainingValue.ReadOnly = True
        Me.ColRemainingValue.Width = 140
        '
        'ColShow
        '
        Me.ColShow.HeaderText = "Show Detail"
        Me.ColShow.Name = "ColShow"
        Me.ColShow.ReadOnly = True
        Me.ColShow.Width = 70
        '
        'FrmLapAsset_DepresiasiPerPeriod
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(884, 562)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.ChkShowNonDepreciateAsset)
        Me.Controls.Add(Me.chkShowDisposedAsset)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnShow)
        Me.MinimumSize = New System.Drawing.Size(900, 600)
        Me.Name = "FrmLapAsset_DepresiasiPerPeriod"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Fixed Asset Depreciation"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkShowDisposedAsset As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowNonDepreciateAsset As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents ColPrimaryGroup As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGroup As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAsset As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAssetCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAssetDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAqcValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAcqDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccumatedDep As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDepreciation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRevaluation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDisposal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemainingValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColShow As System.Windows.Forms.DataGridViewButtonColumn
End Class
