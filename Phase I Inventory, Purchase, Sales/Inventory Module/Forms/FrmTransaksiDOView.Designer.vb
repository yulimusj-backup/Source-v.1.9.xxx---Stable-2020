<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiDOView
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
        Me.dgvItemDO2 = New System.Windows.Forms.DataGridView
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSONO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyPerMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn13 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItemDO2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItemDO2
        '
        Me.dgvItemDO2.AllowUserToAddRows = False
        Me.dgvItemDO2.AllowUserToDeleteRows = False
        Me.dgvItemDO2.AllowUserToResizeColumns = False
        Me.dgvItemDO2.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItemDO2.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItemDO2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemDO2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItemDO2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvItemDO2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn4, Me.ColCodeNo, Me.ColDescription, Me.ColCustCodeNo, Me.ColCustDesc, Me.ColSONO, Me.ColBatchNo, Me.ColExpiryDate, Me.ColQtyMaster, Me.ColQtyPerMaster, Me.ColMasterUnits, Me.ColQty, Me.ColTotalQty, Me.ColUnit, Me.ColSecQty, Me.ColSecUnit, Me.DataGridViewTextBoxColumn13, Me.ColPrimaryCodeNo, Me.ColPrimarySO, Me.ColPrimaryMasterUnits, Me.ColPrimaryUnit, Me.ColPrimarySecUnit})
        Me.dgvItemDO2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemDO2.Location = New System.Drawing.Point(0, 0)
        Me.dgvItemDO2.MultiSelect = False
        Me.dgvItemDO2.Name = "dgvItemDO2"
        Me.dgvItemDO2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItemDO2.Size = New System.Drawing.Size(609, 442)
        Me.dgvItemDO2.StandardTab = True
        Me.dgvItemDO2.TabIndex = 2
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.Frozen = True
        Me.DataGridViewTextBoxColumn4.HeaderText = "No"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.DataGridViewTextBoxColumn4.Width = 27
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.Width = 74
        '
        'ColDescription
        '
        Me.ColDescription.Frozen = True
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.Width = 85
        '
        'ColCustCodeNo
        '
        Me.ColCustCodeNo.Frozen = True
        Me.ColCustCodeNo.HeaderText = "Customer's Code No"
        Me.ColCustCodeNo.Name = "ColCustCodeNo"
        Me.ColCustCodeNo.Visible = False
        Me.ColCustCodeNo.Width = 128
        '
        'ColCustDesc
        '
        Me.ColCustDesc.Frozen = True
        Me.ColCustDesc.HeaderText = "Customer's Item Description"
        Me.ColCustDesc.Name = "ColCustDesc"
        Me.ColCustDesc.Visible = False
        Me.ColCustDesc.Width = 162
        '
        'ColSONO
        '
        Me.ColSONO.Frozen = True
        Me.ColSONO.HeaderText = "SO. No"
        Me.ColSONO.Name = "ColSONO"
        Me.ColSONO.Width = 67
        '
        'ColBatchNo
        '
        Me.ColBatchNo.Frozen = True
        Me.ColBatchNo.HeaderText = "Batch No"
        Me.ColBatchNo.Name = "ColBatchNo"
        Me.ColBatchNo.Width = 77
        '
        'ColExpiryDate
        '
        Me.ColExpiryDate.Frozen = True
        Me.ColExpiryDate.HeaderText = "Expiry Date"
        Me.ColExpiryDate.Name = "ColExpiryDate"
        Me.ColExpiryDate.Width = 86
        '
        'ColQtyMaster
        '
        Me.ColQtyMaster.HeaderText = "Qty Master"
        Me.ColQtyMaster.Name = "ColQtyMaster"
        Me.ColQtyMaster.Visible = False
        Me.ColQtyMaster.Width = 83
        '
        'ColQtyPerMaster
        '
        Me.ColQtyPerMaster.HeaderText = "Qty Per Master"
        Me.ColQtyPerMaster.Name = "ColQtyPerMaster"
        Me.ColQtyPerMaster.Visible = False
        Me.ColQtyPerMaster.Width = 102
        '
        'ColMasterUnits
        '
        Me.ColMasterUnits.HeaderText = "Master Unit"
        Me.ColMasterUnits.Name = "ColMasterUnits"
        Me.ColMasterUnits.Visible = False
        Me.ColMasterUnits.Width = 86
        '
        'ColQty
        '
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Visible = False
        Me.ColQty.Width = 48
        '
        'ColTotalQty
        '
        Me.ColTotalQty.HeaderText = "Total Qty"
        Me.ColTotalQty.Name = "ColTotalQty"
        Me.ColTotalQty.Width = 75
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 51
        '
        'ColSecQty
        '
        Me.ColSecQty.HeaderText = "@"
        Me.ColSecQty.Name = "ColSecQty"
        Me.ColSecQty.Width = 43
        '
        'ColSecUnit
        '
        Me.ColSecUnit.HeaderText = "@ Unit"
        Me.ColSecUnit.Name = "ColSecUnit"
        Me.ColSecUnit.Width = 65
        '
        'DataGridViewTextBoxColumn13
        '
        Me.DataGridViewTextBoxColumn13.HeaderText = "Remarks"
        Me.DataGridViewTextBoxColumn13.Name = "DataGridViewTextBoxColumn13"
        Me.DataGridViewTextBoxColumn13.Width = 74
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 105
        '
        'ColPrimarySO
        '
        Me.ColPrimarySO.HeaderText = "PrimarySO"
        Me.ColPrimarySO.Name = "ColPrimarySO"
        Me.ColPrimarySO.Visible = False
        Me.ColPrimarySO.Width = 81
        '
        'ColPrimaryMasterUnits
        '
        Me.ColPrimaryMasterUnits.HeaderText = "Primary Master Unit"
        Me.ColPrimaryMasterUnits.Name = "ColPrimaryMasterUnits"
        Me.ColPrimaryMasterUnits.Visible = False
        Me.ColPrimaryMasterUnits.Width = 123
        '
        'ColPrimaryUnit
        '
        Me.ColPrimaryUnit.HeaderText = "Primary Units"
        Me.ColPrimaryUnit.Name = "ColPrimaryUnit"
        Me.ColPrimaryUnit.Visible = False
        Me.ColPrimaryUnit.Width = 93
        '
        'ColPrimarySecUnit
        '
        Me.ColPrimarySecUnit.HeaderText = "Primary Sec Unit"
        Me.ColPrimarySecUnit.Name = "ColPrimarySecUnit"
        Me.ColPrimarySecUnit.Visible = False
        Me.ColPrimarySecUnit.Width = 110
        '
        'FrmTransaksiDOView
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(609, 442)
        Me.Controls.Add(Me.dgvItemDO2)
        Me.Name = "FrmTransaksiDOView"
        Me.Text = "DO Detail View"
        Me.TopMost = True
        CType(Me.dgvItemDO2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents dgvItemDO2 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSONO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyPerMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn13 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
