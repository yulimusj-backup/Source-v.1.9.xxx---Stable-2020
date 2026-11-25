<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_ExplainQtyOnHand
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTrans_MaterialPlanningDashboard_ExplainQtyOnHand))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnOK = New System.Windows.Forms.Button
        Me.llbMP_LocationSetting = New System.Windows.Forms.LinkLabel
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeColumns = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.ColumnHeadersHeight = 25
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColQty, Me.ColUnit, Me.ColBatchNo, Me.ColExpiry, Me.ColLocation, Me.ColWarehouse})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 12)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(669, 479)
        Me.DataGridView1.TabIndex = 0
        '
        'ColQty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.ReadOnly = True
        '
        'ColBatchNo
        '
        Me.ColBatchNo.HeaderText = "Batch No"
        Me.ColBatchNo.Name = "ColBatchNo"
        Me.ColBatchNo.ReadOnly = True
        '
        'ColExpiry
        '
        Me.ColExpiry.HeaderText = "Expiry"
        Me.ColExpiry.Name = "ColExpiry"
        Me.ColExpiry.ReadOnly = True
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.ReadOnly = True
        '
        'ColWarehouse
        '
        Me.ColWarehouse.HeaderText = "Warehouse"
        Me.ColWarehouse.Name = "ColWarehouse"
        Me.ColWarehouse.ReadOnly = True
        '
        'btnOK
        '
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnOK.Location = New System.Drawing.Point(606, 497)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'llbMP_LocationSetting
        '
        Me.llbMP_LocationSetting.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbMP_LocationSetting.AutoSize = True
        Me.llbMP_LocationSetting.Location = New System.Drawing.Point(13, 500)
        Me.llbMP_LocationSetting.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbMP_LocationSetting.Name = "llbMP_LocationSetting"
        Me.llbMP_LocationSetting.Size = New System.Drawing.Size(110, 17)
        Me.llbMP_LocationSetting.TabIndex = 88
        Me.llbMP_LocationSetting.TabStop = True
        Me.llbMP_LocationSetting.Text = "Location Setting"
        Me.llbMP_LocationSetting.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'FrmTrans_MaterialPlanningDashboard_ExplainQtyOnHand
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnOK
        Me.ClientSize = New System.Drawing.Size(692, 523)
        Me.Controls.Add(Me.llbMP_LocationSetting)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.DataGridView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(700, 560)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_ExplainQtyOnHand"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Qty Data Detail"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents llbMP_LocationSetting As System.Windows.Forms.LinkLabel
End Class
