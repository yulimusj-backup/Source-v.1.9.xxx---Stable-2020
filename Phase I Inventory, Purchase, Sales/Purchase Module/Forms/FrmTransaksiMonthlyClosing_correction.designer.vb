<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonthlyClosing_correction
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiMonthlyClosing_correction))
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.dgvItemAdjust = New System.Windows.Forms.DataGridView
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.LblProgress = New System.Windows.Forms.Label
        Me.prgBar = New System.Windows.Forms.ProgressBar
        Me.bgw1 = New System.ComponentModel.BackgroundWorker
        Me.bgw2 = New System.ComponentModel.BackgroundWorker
        Me.ColNoAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJenisDocumentAdjustStr = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNoAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPriceAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarksAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryStockInAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeffPrec = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks_Adjust_detail = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNoAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitAdjust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJenisDokumen = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTglDocAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeItemDetail = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAdjustmentType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryGainAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLossAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjustTypeDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeTransaksiAsal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReference = New System.Windows.Forms.DataGridViewButtonColumn
        CType(Me.dgvItemAdjust, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
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
        'dgvItemAdjust
        '
        Me.dgvItemAdjust.AllowUserToAddRows = False
        Me.dgvItemAdjust.AllowUserToDeleteRows = False
        Me.dgvItemAdjust.AllowUserToOrderColumns = True
        Me.dgvItemAdjust.AllowUserToResizeColumns = False
        Me.dgvItemAdjust.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItemAdjust.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItemAdjust.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemAdjust.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemAdjust.ColumnHeadersHeight = 25
        Me.dgvItemAdjust.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemAdjust.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNoAdjust, Me.ColJenisDocumentAdjustStr, Me.ColDocumentNo, Me.ColCodeNoAdjust, Me.ColDescAdjust, Me.ColQty, Me.ColUnitAdjust, Me.ColUnitPriceAdjust, Me.ColRemarksAdjust, Me.ColPrimaryStockInAdjust, Me.ColDeffPrec, Me.ColRemarks_Adjust_detail, Me.ColPrimaryCodeNoAdjust, Me.ColPrimaryUnitAdjust, Me.ColJenisDokumen, Me.ColTglDocAcuan, Me.ColTipeItemDetail, Me.ColPrimaryAdjustmentType, Me.ColPrimaryGainAccount, Me.ColPrimaryLossAccount, Me.ColAdjustTypeDesc, Me.ColTipeTransaksiAsal, Me.ColReference})
        Me.dgvItemAdjust.Location = New System.Drawing.Point(13, 41)
        Me.dgvItemAdjust.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItemAdjust.Name = "dgvItemAdjust"
        Me.dgvItemAdjust.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemAdjust.RowTemplate.Height = 24
        Me.dgvItemAdjust.Size = New System.Drawing.Size(966, 548)
        Me.dgvItemAdjust.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(390, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Correct your data for adjustment amount on the table below :"
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(13, 597)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 8
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(413, 12)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(77, 17)
        Me.LinkLabel1.TabIndex = 9
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "LinkLabel1"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(95, 597)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.LblProgress)
        Me.Panel1.Controls.Add(Me.prgBar)
        Me.Panel1.Location = New System.Drawing.Point(202, 284)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(588, 66)
        Me.Panel1.TabIndex = 11
        Me.Panel1.Visible = False
        '
        'LblProgress
        '
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(3, 29)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(51, 17)
        Me.LblProgress.TabIndex = 1
        Me.LblProgress.Text = "Label2"
        '
        'prgBar
        '
        Me.prgBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBar.Location = New System.Drawing.Point(3, 3)
        Me.prgBar.Name = "prgBar"
        Me.prgBar.Size = New System.Drawing.Size(582, 23)
        Me.prgBar.TabIndex = 0
        '
        'bgw1
        '
        Me.bgw1.WorkerReportsProgress = True
        '
        'bgw2
        '
        Me.bgw2.WorkerReportsProgress = True
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
        Me.ColDocumentNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColCodeNoAdjust
        '
        Me.ColCodeNoAdjust.HeaderText = "Code No"
        Me.ColCodeNoAdjust.Name = "ColCodeNoAdjust"
        Me.ColCodeNoAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
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
        Me.ColQty.HeaderText = "Qty In Adjust"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Width = 94
        '
        'ColUnitAdjust
        '
        Me.ColUnitAdjust.HeaderText = "Units"
        Me.ColUnitAdjust.Name = "ColUnitAdjust"
        Me.ColUnitAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitAdjust.Width = 46
        '
        'ColUnitPriceAdjust
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPriceAdjust.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColUnitPriceAdjust.HeaderText = "Unit Price (IDR)"
        Me.ColUnitPriceAdjust.Name = "ColUnitPriceAdjust"
        Me.ColUnitPriceAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitPriceAdjust.Width = 112
        '
        'ColRemarksAdjust
        '
        Me.ColRemarksAdjust.HeaderText = "Remarks"
        Me.ColRemarksAdjust.Name = "ColRemarksAdjust"
        Me.ColRemarksAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarksAdjust.Width = 70
        '
        'ColPrimaryStockInAdjust
        '
        Me.ColPrimaryStockInAdjust.HeaderText = "Primary Stock In"
        Me.ColPrimaryStockInAdjust.Name = "ColPrimaryStockInAdjust"
        Me.ColPrimaryStockInAdjust.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryStockInAdjust.Visible = False
        Me.ColPrimaryStockInAdjust.Width = 116
        '
        'ColDeffPrec
        '
        Me.ColDeffPrec.HeaderText = "Deff Prec"
        Me.ColDeffPrec.Name = "ColDeffPrec"
        Me.ColDeffPrec.Visible = False
        Me.ColDeffPrec.Width = 92
        '
        'ColRemarks_Adjust_detail
        '
        Me.ColRemarks_Adjust_detail.HeaderText = "Remarks Adjust"
        Me.ColRemarks_Adjust_detail.Name = "ColRemarks_Adjust_detail"
        Me.ColRemarks_Adjust_detail.Visible = False
        Me.ColRemarks_Adjust_detail.Width = 132
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
        'ColJenisDokumen
        '
        Me.ColJenisDokumen.HeaderText = "Jenis Dokumen"
        Me.ColJenisDokumen.Name = "ColJenisDokumen"
        Me.ColJenisDokumen.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJenisDokumen.Visible = False
        Me.ColJenisDokumen.Width = 111
        '
        'ColTglDocAcuan
        '
        Me.ColTglDocAcuan.HeaderText = "Tanggal Documen Acuan"
        Me.ColTglDocAcuan.Name = "ColTglDocAcuan"
        Me.ColTglDocAcuan.Visible = False
        Me.ColTglDocAcuan.Width = 193
        '
        'ColTipeItemDetail
        '
        Me.ColTipeItemDetail.HeaderText = "Tipe Item Detail"
        Me.ColTipeItemDetail.Name = "ColTipeItemDetail"
        Me.ColTipeItemDetail.Visible = False
        Me.ColTipeItemDetail.Width = 131
        '
        'ColPrimaryAdjustmentType
        '
        Me.ColPrimaryAdjustmentType.HeaderText = "Primary Adjustment Type"
        Me.ColPrimaryAdjustmentType.Name = "ColPrimaryAdjustmentType"
        Me.ColPrimaryAdjustmentType.Visible = False
        Me.ColPrimaryAdjustmentType.Width = 191
        '
        'ColPrimaryGainAccount
        '
        Me.ColPrimaryGainAccount.HeaderText = "Primary Gain Account"
        Me.ColPrimaryGainAccount.Name = "ColPrimaryGainAccount"
        Me.ColPrimaryGainAccount.Visible = False
        Me.ColPrimaryGainAccount.Width = 170
        '
        'ColPrimaryLossAccount
        '
        Me.ColPrimaryLossAccount.HeaderText = "PrimaryLossAccount"
        Me.ColPrimaryLossAccount.Name = "ColPrimaryLossAccount"
        Me.ColPrimaryLossAccount.Visible = False
        Me.ColPrimaryLossAccount.Width = 162
        '
        'ColAdjustTypeDesc
        '
        Me.ColAdjustTypeDesc.HeaderText = "Adjust Type Desc"
        Me.ColAdjustTypeDesc.Name = "ColAdjustTypeDesc"
        Me.ColAdjustTypeDesc.Visible = False
        Me.ColAdjustTypeDesc.Width = 144
        '
        'ColTipeTransaksiAsal
        '
        Me.ColTipeTransaksiAsal.HeaderText = "Tipe Transaksi Asal"
        Me.ColTipeTransaksiAsal.Name = "ColTipeTransaksiAsal"
        Me.ColTipeTransaksiAsal.Visible = False
        Me.ColTipeTransaksiAsal.Width = 158
        '
        'ColReference
        '
        Me.ColReference.HeaderText = "Unit Price Reference"
        Me.ColReference.Name = "ColReference"
        Me.ColReference.Visible = False
        Me.ColReference.Width = 145
        '
        'FrmTransaksiMonthlyClosing_correction
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvItemAdjust)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(800, 500)
        Me.Name = "FrmTransaksiMonthlyClosing_correction"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Correction for Adjustment amount"
        CType(Me.dgvItemAdjust, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents dgvItemAdjust As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents prgBar As System.Windows.Forms.ProgressBar
    Friend WithEvents bgw1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgw2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNoAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJenisDocumentAdjustStr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNoAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPriceAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarksAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryStockInAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeffPrec As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks_Adjust_detail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNoAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitAdjust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJenisDokumen As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTglDocAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeItemDetail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAdjustmentType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryGainAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLossAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjustTypeDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeTransaksiAsal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReference As System.Windows.Forms.DataGridViewButtonColumn
End Class
