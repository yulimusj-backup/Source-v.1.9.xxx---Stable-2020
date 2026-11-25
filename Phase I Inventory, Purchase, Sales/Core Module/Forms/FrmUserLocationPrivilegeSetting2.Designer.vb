<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmUserLocationPrivilegeSetting2
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
        Me.dgvLocation = New System.Windows.Forms.DataGridView
        Me.colLocation_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colLocation_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.colLocation_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colLocation_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colLocation_StockInPurchase = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colLocation_StockInSalesReturn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockInAdjust = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockOutSales = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockOutPurcahseReturn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockOutAdjust = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockTransferOut = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockTransferIn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockBundlingBefore = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_StockBundlingAfter = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_ProductionlogBefore = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_ProductionLogAfter = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colLocationPrimaryRefParty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.dgvLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvLocation
        '
        Me.dgvLocation.AllowUserToAddRows = False
        Me.dgvLocation.AllowUserToDeleteRows = False
        Me.dgvLocation.AllowUserToResizeColumns = False
        Me.dgvLocation.AllowUserToResizeRows = False
        Me.dgvLocation.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvLocation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvLocation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvLocation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colLocation_No, Me.colLocation_Level, Me.colLocation_Code, Me.colLocation_Desc, Me.colLocation_StockInPurchase, Me.colLocation_StockInSalesReturn, Me.ColLocation_StockInAdjust, Me.ColLocation_StockOutSales, Me.ColLocation_StockOutPurcahseReturn, Me.ColLocation_StockOutAdjust, Me.ColLocation_StockTransferOut, Me.ColLocation_StockTransferIn, Me.ColLocation_StockBundlingBefore, Me.ColLocation_StockBundlingAfter, Me.ColLocation_ProductionlogBefore, Me.ColLocation_ProductionLogAfter, Me.colLocationPrimaryRefParty})
        Me.dgvLocation.Location = New System.Drawing.Point(13, 30)
        Me.dgvLocation.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvLocation.MultiSelect = False
        Me.dgvLocation.Name = "dgvLocation"
        Me.dgvLocation.ReadOnly = True
        Me.dgvLocation.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvLocation.RowTemplate.Height = 24
        Me.dgvLocation.Size = New System.Drawing.Size(766, 570)
        Me.dgvLocation.TabIndex = 2
        '
        'colLocation_No
        '
        Me.colLocation_No.Frozen = True
        Me.colLocation_No.HeaderText = "No"
        Me.colLocation_No.Name = "colLocation_No"
        Me.colLocation_No.ReadOnly = True
        Me.colLocation_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.colLocation_No.Width = 32
        '
        'colLocation_Level
        '
        Me.colLocation_Level.Frozen = True
        Me.colLocation_Level.HeaderText = "Location Level"
        Me.colLocation_Level.Items.AddRange(New Object() {"Warehouse", "Location"})
        Me.colLocation_Level.Name = "colLocation_Level"
        Me.colLocation_Level.ReadOnly = True
        Me.colLocation_Level.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colLocation_Level.Width = 106
        '
        'colLocation_Code
        '
        Me.colLocation_Code.Frozen = True
        Me.colLocation_Code.HeaderText = "Code"
        Me.colLocation_Code.Name = "colLocation_Code"
        Me.colLocation_Code.ReadOnly = True
        Me.colLocation_Code.Width = 66
        '
        'colLocation_Desc
        '
        Me.colLocation_Desc.HeaderText = "Description"
        Me.colLocation_Desc.Name = "colLocation_Desc"
        Me.colLocation_Desc.ReadOnly = True
        Me.colLocation_Desc.Width = 104
        '
        'colLocation_StockInPurchase
        '
        Me.colLocation_StockInPurchase.HeaderText = "Stock in Purchase"
        Me.colLocation_StockInPurchase.Name = "colLocation_StockInPurchase"
        Me.colLocation_StockInPurchase.ReadOnly = True
        Me.colLocation_StockInPurchase.Width = 128
        '
        'colLocation_StockInSalesReturn
        '
        Me.colLocation_StockInSalesReturn.HeaderText = "Stock in Sales Return"
        Me.colLocation_StockInSalesReturn.Name = "colLocation_StockInSalesReturn"
        Me.colLocation_StockInSalesReturn.ReadOnly = True
        Me.colLocation_StockInSalesReturn.Width = 150
        '
        'ColLocation_StockInAdjust
        '
        Me.ColLocation_StockInAdjust.HeaderText = "Stock in Adjust"
        Me.ColLocation_StockInAdjust.Name = "ColLocation_StockInAdjust"
        Me.ColLocation_StockInAdjust.ReadOnly = True
        Me.ColLocation_StockInAdjust.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColLocation_StockInAdjust.Width = 107
        '
        'ColLocation_StockOutSales
        '
        Me.ColLocation_StockOutSales.HeaderText = "Stock out Sales"
        Me.ColLocation_StockOutSales.Name = "ColLocation_StockOutSales"
        Me.ColLocation_StockOutSales.ReadOnly = True
        Me.ColLocation_StockOutSales.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColLocation_StockOutSales.Width = 112
        '
        'ColLocation_StockOutPurcahseReturn
        '
        Me.ColLocation_StockOutPurcahseReturn.HeaderText = "Stock out Purchase Return"
        Me.ColLocation_StockOutPurcahseReturn.Name = "ColLocation_StockOutPurcahseReturn"
        Me.ColLocation_StockOutPurcahseReturn.ReadOnly = True
        Me.ColLocation_StockOutPurcahseReturn.Width = 184
        '
        'ColLocation_StockOutAdjust
        '
        Me.ColLocation_StockOutAdjust.HeaderText = "Stock out Adjust"
        Me.ColLocation_StockOutAdjust.Name = "ColLocation_StockOutAdjust"
        Me.ColLocation_StockOutAdjust.ReadOnly = True
        Me.ColLocation_StockOutAdjust.Width = 116
        '
        'ColLocation_StockTransferOut
        '
        Me.ColLocation_StockTransferOut.HeaderText = "Stock Transfer Out"
        Me.ColLocation_StockTransferOut.Name = "ColLocation_StockTransferOut"
        Me.ColLocation_StockTransferOut.ReadOnly = True
        Me.ColLocation_StockTransferOut.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColLocation_StockTransferOut.Width = 134
        '
        'ColLocation_StockTransferIn
        '
        Me.ColLocation_StockTransferIn.HeaderText = "Stock Transfer In"
        Me.ColLocation_StockTransferIn.Name = "ColLocation_StockTransferIn"
        Me.ColLocation_StockTransferIn.ReadOnly = True
        Me.ColLocation_StockTransferIn.Width = 122
        '
        'ColLocation_StockBundlingBefore
        '
        Me.ColLocation_StockBundlingBefore.HeaderText = "Stock Bundling Before"
        Me.ColLocation_StockBundlingBefore.Name = "ColLocation_StockBundlingBefore"
        Me.ColLocation_StockBundlingBefore.ReadOnly = True
        Me.ColLocation_StockBundlingBefore.Width = 154
        '
        'ColLocation_StockBundlingAfter
        '
        Me.ColLocation_StockBundlingAfter.HeaderText = "Stock Bundling After"
        Me.ColLocation_StockBundlingAfter.Name = "ColLocation_StockBundlingAfter"
        Me.ColLocation_StockBundlingAfter.ReadOnly = True
        Me.ColLocation_StockBundlingAfter.Width = 142
        '
        'ColLocation_ProductionlogBefore
        '
        Me.ColLocation_ProductionlogBefore.HeaderText = "Production Log RM"
        Me.ColLocation_ProductionlogBefore.Name = "ColLocation_ProductionlogBefore"
        Me.ColLocation_ProductionlogBefore.ReadOnly = True
        Me.ColLocation_ProductionlogBefore.Width = 135
        '
        'ColLocation_ProductionLogAfter
        '
        Me.ColLocation_ProductionLogAfter.HeaderText = "Production Log FG"
        Me.ColLocation_ProductionLogAfter.Name = "ColLocation_ProductionLogAfter"
        Me.ColLocation_ProductionLogAfter.ReadOnly = True
        Me.ColLocation_ProductionLogAfter.Width = 133
        '
        'colLocationPrimaryRefParty
        '
        Me.colLocationPrimaryRefParty.HeaderText = "Primary Ref Party"
        Me.colLocationPrimaryRefParty.Name = "colLocationPrimaryRefParty"
        Me.colLocationPrimaryRefParty.ReadOnly = True
        Me.colLocationPrimaryRefParty.Visible = False
        Me.colLocationPrimaryRefParty.Width = 144
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Label1"
        '
        'FrmUserLocationPrivilegeSetting2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(792, 613)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvLocation)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(800, 650)
        Me.Name = "FrmUserLocationPrivilegeSetting2"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Location Access"
        CType(Me.dgvLocation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvLocation As System.Windows.Forms.DataGridView
    Friend WithEvents colLocation_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colLocation_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colLocation_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colLocation_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colLocation_StockInPurchase As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colLocation_StockInSalesReturn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockInAdjust As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockOutSales As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockOutPurcahseReturn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockOutAdjust As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockTransferOut As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockTransferIn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockBundlingBefore As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_StockBundlingAfter As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_ProductionlogBefore As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_ProductionLogAfter As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colLocationPrimaryRefParty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
