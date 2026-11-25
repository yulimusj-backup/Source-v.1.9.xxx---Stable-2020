<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonthlyClosing_CyclicBundlingExplain
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
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiMonthlyClosing_CyclicBundlingExplain))
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.btnOK = New System.Windows.Forms.Button
        Me.chkOutBundling = New System.Windows.Forms.CheckBox
        Me.chkInBundling = New System.Windows.Forms.CheckBox
        Me.chkInPurchase = New System.Windows.Forms.CheckBox
        Me.chkInSalesReturn = New System.Windows.Forms.CheckBox
        Me.chkInAdjust = New System.Windows.Forms.CheckBox
        Me.chkOutSales = New System.Windows.Forms.CheckBox
        Me.chkOutPurchaseReturn = New System.Windows.Forms.CheckBox
        Me.chkOutAdjust = New System.Windows.Forms.CheckBox
        Me.btnFilter = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.chkAutomatic = New System.Windows.Forms.CheckBox
        Me.ColNoAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJenisDocumentAdjustStr = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocumentNo = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColTglDocAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNoAdjust = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColDescAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDocAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeffPrec = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNoAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSetUnitPrice = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeTransaksiAsal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeItemDetail = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTanggalTransaksiAsal = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ToolTip
        '
        Me.ToolTip.AutoPopDelay = 10000
        Me.ToolTip.InitialDelay = 500
        Me.ToolTip.ReshowDelay = 100
        Me.ToolTip.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.ToolTip.ToolTipTitle = "Info"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(745, 17)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "The following item cause cyclic bundling, we can't calculate cost. Please contact" & _
            " your vendor if you see this message."
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToOrderColumns = True
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNoAdjust, Me.ColJenisDocumentAdjustStr, Me.ColDocumentNo, Me.ColTglDocAcuan, Me.ColCodeNoAdjust, Me.ColDescAdjust, Me.ColQty, Me.ColUnitAdjust, Me.ColPrimaryDocAcuan, Me.ColDeffPrec, Me.ColPrimaryCodeNoAdjust, Me.ColPrimaryUnitAdjust, Me.ColUnitPrice, Me.ColSetUnitPrice, Me.ColRemarks, Me.ColTipeTransaksiAsal, Me.ColTipeItemDetail, Me.ColTanggalTransaksiAsal})
        Me.dgvItem.Location = New System.Drawing.Point(13, 92)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersVisible = False
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(966, 528)
        Me.dgvItem.TabIndex = 9
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(626, 33)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 11
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'chkOutBundling
        '
        Me.chkOutBundling.AutoSize = True
        Me.chkOutBundling.Checked = True
        Me.chkOutBundling.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOutBundling.Location = New System.Drawing.Point(12, 35)
        Me.chkOutBundling.Name = "chkOutBundling"
        Me.chkOutBundling.Size = New System.Drawing.Size(112, 21)
        Me.chkOutBundling.TabIndex = 12
        Me.chkOutBundling.Text = "Out Bundling"
        Me.chkOutBundling.UseVisualStyleBackColor = True
        '
        'chkInBundling
        '
        Me.chkInBundling.AutoSize = True
        Me.chkInBundling.Checked = True
        Me.chkInBundling.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkInBundling.Location = New System.Drawing.Point(12, 64)
        Me.chkInBundling.Name = "chkInBundling"
        Me.chkInBundling.Size = New System.Drawing.Size(100, 21)
        Me.chkInBundling.TabIndex = 13
        Me.chkInBundling.Text = "In Bundling"
        Me.chkInBundling.UseVisualStyleBackColor = True
        '
        'chkInPurchase
        '
        Me.chkInPurchase.AutoSize = True
        Me.chkInPurchase.Location = New System.Drawing.Point(130, 35)
        Me.chkInPurchase.Name = "chkInPurchase"
        Me.chkInPurchase.Size = New System.Drawing.Size(105, 21)
        Me.chkInPurchase.TabIndex = 14
        Me.chkInPurchase.Text = "In Purchase"
        Me.chkInPurchase.UseVisualStyleBackColor = True
        '
        'chkInSalesReturn
        '
        Me.chkInSalesReturn.AutoSize = True
        Me.chkInSalesReturn.Location = New System.Drawing.Point(241, 35)
        Me.chkInSalesReturn.Name = "chkInSalesReturn"
        Me.chkInSalesReturn.Size = New System.Drawing.Size(127, 21)
        Me.chkInSalesReturn.TabIndex = 15
        Me.chkInSalesReturn.Text = "In Sales Return"
        Me.chkInSalesReturn.UseVisualStyleBackColor = True
        '
        'chkInAdjust
        '
        Me.chkInAdjust.AutoSize = True
        Me.chkInAdjust.Location = New System.Drawing.Point(411, 35)
        Me.chkInAdjust.Name = "chkInAdjust"
        Me.chkInAdjust.Size = New System.Drawing.Size(115, 21)
        Me.chkInAdjust.TabIndex = 16
        Me.chkInAdjust.Text = "In Adjustment"
        Me.chkInAdjust.UseVisualStyleBackColor = True
        '
        'chkOutSales
        '
        Me.chkOutSales.AutoSize = True
        Me.chkOutSales.Location = New System.Drawing.Point(130, 62)
        Me.chkOutSales.Name = "chkOutSales"
        Me.chkOutSales.Size = New System.Drawing.Size(92, 21)
        Me.chkOutSales.TabIndex = 17
        Me.chkOutSales.Text = "Out Sales"
        Me.chkOutSales.UseVisualStyleBackColor = True
        '
        'chkOutPurchaseReturn
        '
        Me.chkOutPurchaseReturn.AutoSize = True
        Me.chkOutPurchaseReturn.Location = New System.Drawing.Point(241, 62)
        Me.chkOutPurchaseReturn.Name = "chkOutPurchaseReturn"
        Me.chkOutPurchaseReturn.Size = New System.Drawing.Size(164, 21)
        Me.chkOutPurchaseReturn.TabIndex = 18
        Me.chkOutPurchaseReturn.Text = "Out Purchase Return"
        Me.chkOutPurchaseReturn.UseVisualStyleBackColor = True
        '
        'chkOutAdjust
        '
        Me.chkOutAdjust.AutoSize = True
        Me.chkOutAdjust.Location = New System.Drawing.Point(411, 62)
        Me.chkOutAdjust.Name = "chkOutAdjust"
        Me.chkOutAdjust.Size = New System.Drawing.Size(127, 21)
        Me.chkOutAdjust.TabIndex = 19
        Me.chkOutAdjust.Text = "Out Adjustment"
        Me.chkOutAdjust.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(544, 33)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(75, 23)
        Me.btnFilter.TabIndex = 20
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(459, 305)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 21
        Me.btnCancel.TabStop = False
        Me.btnCancel.Text = "&OK"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'chkAutomatic
        '
        Me.chkAutomatic.AutoSize = True
        Me.chkAutomatic.Location = New System.Drawing.Point(544, 62)
        Me.chkAutomatic.Name = "chkAutomatic"
        Me.chkAutomatic.Size = New System.Drawing.Size(244, 21)
        Me.chkAutomatic.TabIndex = 22
        Me.chkAutomatic.Text = "Set Reference Price Automatically"
        Me.chkAutomatic.UseVisualStyleBackColor = True
        '
        'ColNoAdjust
        '
        Me.ColNoAdjust.HeaderText = "No"
        Me.ColNoAdjust.Name = "ColNoAdjust"
        Me.ColNoAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNoAdjust.Width = 32
        '
        'ColJenisDocumentAdjustStr
        '
        Me.ColJenisDocumentAdjustStr.HeaderText = "Document"
        Me.ColJenisDocumentAdjustStr.Name = "ColJenisDocumentAdjustStr"
        Me.ColJenisDocumentAdjustStr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJenisDocumentAdjustStr.Width = 78
        '
        'ColDocumentNo
        '
        Me.ColDocumentNo.HeaderText = "Document No"
        Me.ColDocumentNo.Name = "ColDocumentNo"
        Me.ColDocumentNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'ColTglDocAcuan
        '
        Me.ColTglDocAcuan.HeaderText = "Date"
        Me.ColTglDocAcuan.Name = "ColTglDocAcuan"
        Me.ColTglDocAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTglDocAcuan.Visible = False
        Me.ColTglDocAcuan.Width = 44
        '
        'ColCodeNoAdjust
        '
        Me.ColCodeNoAdjust.HeaderText = "Code No"
        Me.ColCodeNoAdjust.Name = "ColCodeNoAdjust"
        Me.ColCodeNoAdjust.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColCodeNoAdjust.Width = 69
        '
        'ColDescAdjust
        '
        Me.ColDescAdjust.HeaderText = "Description"
        Me.ColDescAdjust.Name = "ColDescAdjust"
        Me.ColDescAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDescAdjust.Width = 85
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Width = 36
        '
        'ColUnitAdjust
        '
        Me.ColUnitAdjust.HeaderText = "Units"
        Me.ColUnitAdjust.Name = "ColUnitAdjust"
        Me.ColUnitAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitAdjust.Width = 46
        '
        'ColPrimaryDocAcuan
        '
        Me.ColPrimaryDocAcuan.HeaderText = "Primary Doc Acuan"
        Me.ColPrimaryDocAcuan.Name = "ColPrimaryDocAcuan"
        Me.ColPrimaryDocAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryDocAcuan.Visible = False
        Me.ColPrimaryDocAcuan.Width = 135
        '
        'ColDeffPrec
        '
        Me.ColDeffPrec.HeaderText = "Deff Prec"
        Me.ColDeffPrec.Name = "ColDeffPrec"
        Me.ColDeffPrec.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDeffPrec.Visible = False
        Me.ColDeffPrec.Width = 73
        '
        'ColPrimaryCodeNoAdjust
        '
        Me.ColPrimaryCodeNoAdjust.HeaderText = "Primary Code No"
        Me.ColPrimaryCodeNoAdjust.Name = "ColPrimaryCodeNoAdjust"
        Me.ColPrimaryCodeNoAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryCodeNoAdjust.Visible = False
        Me.ColPrimaryCodeNoAdjust.Width = 121
        '
        'ColPrimaryUnitAdjust
        '
        Me.ColPrimaryUnitAdjust.HeaderText = "Primary Unit"
        Me.ColPrimaryUnitAdjust.Name = "ColPrimaryUnitAdjust"
        Me.ColPrimaryUnitAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryUnitAdjust.Visible = False
        Me.ColPrimaryUnitAdjust.Width = 91
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitPrice.Width = 75
        '
        'ColSetUnitPrice
        '
        Me.ColSetUnitPrice.HeaderText = "Set Unit Price"
        Me.ColSetUnitPrice.Name = "ColSetUnitPrice"
        Me.ColSetUnitPrice.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColSetUnitPrice.Text = ""
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 70
        '
        'ColTipeTransaksiAsal
        '
        Me.ColTipeTransaksiAsal.HeaderText = "Tipe Transaki Asal"
        Me.ColTipeTransaksiAsal.Name = "ColTipeTransaksiAsal"
        Me.ColTipeTransaksiAsal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTipeTransaksiAsal.Visible = False
        Me.ColTipeTransaksiAsal.Width = 132
        '
        'ColTipeItemDetail
        '
        Me.ColTipeItemDetail.HeaderText = "Tipe Item Detail"
        Me.ColTipeItemDetail.Name = "ColTipeItemDetail"
        Me.ColTipeItemDetail.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTipeItemDetail.Visible = False
        Me.ColTipeItemDetail.Width = 112
        '
        'ColTanggalTransaksiAsal
        '
        Me.ColTanggalTransaksiAsal.HeaderText = "Tgl Transaksi Asal"
        Me.ColTanggalTransaksiAsal.Name = "ColTanggalTransaksiAsal"
        Me.ColTanggalTransaksiAsal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTanggalTransaksiAsal.Visible = False
        Me.ColTanggalTransaksiAsal.Width = 131
        '
        'FrmTransaksiMonthlyClosing_CyclicBundlingExplain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.chkAutomatic)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.chkOutAdjust)
        Me.Controls.Add(Me.chkOutPurchaseReturn)
        Me.Controls.Add(Me.chkOutSales)
        Me.Controls.Add(Me.chkInAdjust)
        Me.Controls.Add(Me.chkInSalesReturn)
        Me.Controls.Add(Me.chkInPurchase)
        Me.Controls.Add(Me.chkInBundling)
        Me.Controls.Add(Me.chkOutBundling)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOK)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmTransaksiMonthlyClosing_CyclicBundlingExplain"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Explanation"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents chkOutBundling As System.Windows.Forms.CheckBox
    Friend WithEvents chkInBundling As System.Windows.Forms.CheckBox
    Friend WithEvents chkInPurchase As System.Windows.Forms.CheckBox
    Friend WithEvents chkInSalesReturn As System.Windows.Forms.CheckBox
    Friend WithEvents chkInAdjust As System.Windows.Forms.CheckBox
    Friend WithEvents chkOutSales As System.Windows.Forms.CheckBox
    Friend WithEvents chkOutPurchaseReturn As System.Windows.Forms.CheckBox
    Friend WithEvents chkOutAdjust As System.Windows.Forms.CheckBox
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents chkAutomatic As System.Windows.Forms.CheckBox
    Friend WithEvents ColNoAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJenisDocumentAdjustStr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocumentNo As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColTglDocAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNoAdjust As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColDescAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDocAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeffPrec As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNoAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSetUnitPrice As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeTransaksiAsal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeItemDetail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTanggalTransaksiAsal As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
