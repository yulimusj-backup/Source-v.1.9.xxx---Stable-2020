<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonthlyClosing_CyclicBundlingExplain1
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvData1 = New System.Windows.Forms.DataGridView
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.btnClose = New System.Windows.Forms.Button
        Me.ColData1_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColData1_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColData1_DocNo = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColData1_QtyIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColData1_QtyOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColData1_Balance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColData1_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColData1_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColData1_PrimaryDocRef = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColData1_TipeDocRef = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColData1_ShowDocNo = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColData1_TipeItemDetail = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvData1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvData1
        '
        Me.dgvData1.AllowUserToAddRows = False
        Me.dgvData1.AllowUserToDeleteRows = False
        Me.dgvData1.AllowUserToResizeColumns = False
        Me.dgvData1.AllowUserToResizeRows = False
        Me.dgvData1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvData1.ColumnHeadersHeight = 25
        Me.dgvData1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvData1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColData1_Date, Me.ColData1_Desc, Me.ColData1_DocNo, Me.ColData1_QtyIn, Me.ColData1_QtyOut, Me.ColData1_Balance, Me.ColData1_Unit, Me.ColData1_Amount, Me.ColData1_PrimaryDocRef, Me.ColData1_TipeDocRef, Me.ColData1_ShowDocNo, Me.ColData1_TipeItemDetail})
        Me.dgvData1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvData1.Location = New System.Drawing.Point(3, 3)
        Me.dgvData1.Name = "dgvData1"
        Me.dgvData1.ReadOnly = True
        Me.dgvData1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvData1.RowTemplate.Height = 24
        Me.dgvData1.Size = New System.Drawing.Size(707, 621)
        Me.dgvData1.TabIndex = 0
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(721, 656)
        Me.TabControl1.TabIndex = 1
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvData1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(713, 627)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Detail data"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(713, 627)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Begining Balance"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'TabPage4
        '
        Me.TabPage4.Location = New System.Drawing.Point(4, 25)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(713, 627)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Inventory Pool"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(100, 207)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(138, 35)
        Me.btnClose.TabIndex = 2
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'ColData1_Date
        '
        Me.ColData1_Date.HeaderText = "Date"
        Me.ColData1_Date.Name = "ColData1_Date"
        Me.ColData1_Date.ReadOnly = True
        Me.ColData1_Date.Width = 63
        '
        'ColData1_Desc
        '
        Me.ColData1_Desc.HeaderText = "Description"
        Me.ColData1_Desc.Name = "ColData1_Desc"
        Me.ColData1_Desc.ReadOnly = True
        Me.ColData1_Desc.Width = 104
        '
        'ColData1_DocNo
        '
        Me.ColData1_DocNo.HeaderText = "Doc No"
        Me.ColData1_DocNo.Name = "ColData1_DocNo"
        Me.ColData1_DocNo.ReadOnly = True
        Me.ColData1_DocNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColData1_DocNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColData1_DocNo.Width = 80
        '
        'ColData1_QtyIn
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColData1_QtyIn.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColData1_QtyIn.HeaderText = "Qty In"
        Me.ColData1_QtyIn.Name = "ColData1_QtyIn"
        Me.ColData1_QtyIn.ReadOnly = True
        Me.ColData1_QtyIn.Width = 70
        '
        'ColData1_QtyOut
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColData1_QtyOut.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColData1_QtyOut.HeaderText = "Qty Out"
        Me.ColData1_QtyOut.Name = "ColData1_QtyOut"
        Me.ColData1_QtyOut.ReadOnly = True
        Me.ColData1_QtyOut.Width = 82
        '
        'ColData1_Balance
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColData1_Balance.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColData1_Balance.HeaderText = "Balance"
        Me.ColData1_Balance.Name = "ColData1_Balance"
        Me.ColData1_Balance.ReadOnly = True
        Me.ColData1_Balance.Width = 84
        '
        'ColData1_Unit
        '
        Me.ColData1_Unit.HeaderText = "Unit"
        Me.ColData1_Unit.Name = "ColData1_Unit"
        Me.ColData1_Unit.ReadOnly = True
        Me.ColData1_Unit.Width = 58
        '
        'ColData1_Amount
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColData1_Amount.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColData1_Amount.HeaderText = "Amount"
        Me.ColData1_Amount.Name = "ColData1_Amount"
        Me.ColData1_Amount.ReadOnly = True
        Me.ColData1_Amount.Width = 81
        '
        'ColData1_PrimaryDocRef
        '
        Me.ColData1_PrimaryDocRef.HeaderText = "Primary Doc Ref"
        Me.ColData1_PrimaryDocRef.Name = "ColData1_PrimaryDocRef"
        Me.ColData1_PrimaryDocRef.ReadOnly = True
        Me.ColData1_PrimaryDocRef.Width = 136
        '
        'ColData1_TipeDocRef
        '
        Me.ColData1_TipeDocRef.HeaderText = "Tipe Doc Ref"
        Me.ColData1_TipeDocRef.Name = "ColData1_TipeDocRef"
        Me.ColData1_TipeDocRef.ReadOnly = True
        Me.ColData1_TipeDocRef.Visible = False
        Me.ColData1_TipeDocRef.Width = 116
        '
        'ColData1_ShowDocNo
        '
        Me.ColData1_ShowDocNo.HeaderText = "Show Doc No"
        Me.ColData1_ShowDocNo.Name = "ColData1_ShowDocNo"
        Me.ColData1_ShowDocNo.ReadOnly = True
        Me.ColData1_ShowDocNo.Width = 99
        '
        'ColData1_TipeItemDetail
        '
        Me.ColData1_TipeItemDetail.HeaderText = "Tipe Item Detail"
        Me.ColData1_TipeItemDetail.Name = "ColData1_TipeItemDetail"
        Me.ColData1_TipeItemDetail.ReadOnly = True
        Me.ColData1_TipeItemDetail.Visible = False
        Me.ColData1_TipeItemDetail.Width = 131
        '
        'FrmTransaksiMonthlyClosing_CyclicBundlingExplain1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(745, 680)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnClose)
        Me.Name = "FrmTransaksiMonthlyClosing_CyclicBundlingExplain1"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Explain Code No"
        CType(Me.dgvData1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvData1 As System.Windows.Forms.DataGridView
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ColData1_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColData1_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColData1_DocNo As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColData1_QtyIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColData1_QtyOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColData1_Balance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColData1_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColData1_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColData1_PrimaryDocRef As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColData1_TipeDocRef As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColData1_ShowDocNo As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColData1_TipeItemDetail As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
