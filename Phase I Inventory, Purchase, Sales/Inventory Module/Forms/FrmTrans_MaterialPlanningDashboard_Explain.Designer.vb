<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_Explain
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
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.cboIncludeSales = New System.Windows.Forms.CheckBox
        Me.cboIncludeBundling = New System.Windows.Forms.CheckBox
        Me.cboIncludeAdjustment = New System.Windows.Forms.CheckBox
        Me.dgvData = New System.Windows.Forms.DataGridView
        Me.ColExplain_Periode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExplain_Sales = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColExplain_Bundling = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColExplain_Adjustment = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColHistorical = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColExplain_Total = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.txtExplain = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cboIncludeSales
        '
        Me.cboIncludeSales.AutoCheck = False
        Me.cboIncludeSales.AutoSize = True
        Me.cboIncludeSales.Location = New System.Drawing.Point(12, 12)
        Me.cboIncludeSales.Name = "cboIncludeSales"
        Me.cboIncludeSales.Size = New System.Drawing.Size(90, 17)
        Me.cboIncludeSales.TabIndex = 0
        Me.cboIncludeSales.Text = "Include Sales"
        Me.cboIncludeSales.UseVisualStyleBackColor = True
        '
        'cboIncludeBundling
        '
        Me.cboIncludeBundling.AutoCheck = False
        Me.cboIncludeBundling.AutoSize = True
        Me.cboIncludeBundling.Location = New System.Drawing.Point(143, 12)
        Me.cboIncludeBundling.Name = "cboIncludeBundling"
        Me.cboIncludeBundling.Size = New System.Drawing.Size(105, 17)
        Me.cboIncludeBundling.TabIndex = 1
        Me.cboIncludeBundling.Text = "Include Bundling"
        Me.cboIncludeBundling.UseVisualStyleBackColor = True
        '
        'cboIncludeAdjustment
        '
        Me.cboIncludeAdjustment.AutoCheck = False
        Me.cboIncludeAdjustment.AutoSize = True
        Me.cboIncludeAdjustment.Location = New System.Drawing.Point(292, 12)
        Me.cboIncludeAdjustment.Name = "cboIncludeAdjustment"
        Me.cboIncludeAdjustment.Size = New System.Drawing.Size(116, 17)
        Me.cboIncludeAdjustment.TabIndex = 2
        Me.cboIncludeAdjustment.Text = "Include Adjustment"
        Me.cboIncludeAdjustment.UseVisualStyleBackColor = True
        '
        'dgvData
        '
        Me.dgvData.AllowUserToAddRows = False
        Me.dgvData.AllowUserToDeleteRows = False
        Me.dgvData.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvData.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColExplain_Periode, Me.ColExplain_Sales, Me.ColExplain_Bundling, Me.ColExplain_Adjustment, Me.ColHistorical, Me.ColExplain_Total})
        Me.dgvData.Location = New System.Drawing.Point(12, 35)
        Me.dgvData.Name = "dgvData"
        Me.dgvData.RowTemplate.Height = 24
        Me.dgvData.Size = New System.Drawing.Size(505, 164)
        Me.dgvData.TabIndex = 3
        '
        'ColExplain_Periode
        '
        Me.ColExplain_Periode.HeaderText = "Period"
        Me.ColExplain_Periode.Name = "ColExplain_Periode"
        Me.ColExplain_Periode.Width = 62
        '
        'ColExplain_Sales
        '
        Me.ColExplain_Sales.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.ColExplain_Sales.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColExplain_Sales.HeaderText = "Sales"
        Me.ColExplain_Sales.Name = "ColExplain_Sales"
        Me.ColExplain_Sales.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColExplain_Sales.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColExplain_Sales.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColExplain_Sales.Width = 58
        '
        'ColExplain_Bundling
        '
        Me.ColExplain_Bundling.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.ColExplain_Bundling.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColExplain_Bundling.HeaderText = "Bundling"
        Me.ColExplain_Bundling.Name = "ColExplain_Bundling"
        Me.ColExplain_Bundling.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColExplain_Bundling.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColExplain_Bundling.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColExplain_Bundling.Width = 73
        '
        'ColExplain_Adjustment
        '
        Me.ColExplain_Adjustment.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.ColExplain_Adjustment.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColExplain_Adjustment.HeaderText = "Adjustment"
        Me.ColExplain_Adjustment.Name = "ColExplain_Adjustment"
        Me.ColExplain_Adjustment.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColExplain_Adjustment.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColExplain_Adjustment.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColExplain_Adjustment.Width = 84
        '
        'ColHistorical
        '
        Me.ColHistorical.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.ColHistorical.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColHistorical.HeaderText = "Historical"
        Me.ColHistorical.Name = "ColHistorical"
        Me.ColHistorical.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColHistorical.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColHistorical.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColHistorical.Width = 75
        '
        'ColExplain_Total
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.ColExplain_Total.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColExplain_Total.HeaderText = "Total"
        Me.ColExplain_Total.Name = "ColExplain_Total"
        Me.ColExplain_Total.Width = 56
        '
        'txtExplain
        '
        Me.txtExplain.BackColor = System.Drawing.SystemColors.Info
        Me.txtExplain.Location = New System.Drawing.Point(12, 205)
        Me.txtExplain.Multiline = True
        Me.txtExplain.Name = "txtExplain"
        Me.txtExplain.ReadOnly = True
        Me.txtExplain.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtExplain.Size = New System.Drawing.Size(505, 195)
        Me.txtExplain.TabIndex = 4
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(12, 35)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'FrmTrans_MaterialPlanningDashboard_Explain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(529, 412)
        Me.Controls.Add(Me.txtExplain)
        Me.Controls.Add(Me.dgvData)
        Me.Controls.Add(Me.cboIncludeAdjustment)
        Me.Controls.Add(Me.cboIncludeBundling)
        Me.Controls.Add(Me.cboIncludeSales)
        Me.Controls.Add(Me.btnClose)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(545, 450)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(545, 450)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_Explain"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Explain"
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboIncludeSales As System.Windows.Forms.CheckBox
    Friend WithEvents cboIncludeBundling As System.Windows.Forms.CheckBox
    Friend WithEvents cboIncludeAdjustment As System.Windows.Forms.CheckBox
    Friend WithEvents dgvData As System.Windows.Forms.DataGridView
    Friend WithEvents txtExplain As System.Windows.Forms.TextBox
    Friend WithEvents ColExplain_Periode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExplain_Sales As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColExplain_Bundling As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColExplain_Adjustment As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColHistorical As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColExplain_Total As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class
