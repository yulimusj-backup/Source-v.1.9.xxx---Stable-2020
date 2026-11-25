<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanKartuAR
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.btnPrint = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker
        Me.btnFindProduct = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtCustCode = New System.Windows.Forms.TextBox
        Me.txtPrimaryCust = New System.Windows.Forms.TextBox
        Me.dgvStockCard = New System.Windows.Forms.DataGridView
        Me.gboDetail.SuspendLayout()
        CType(Me.dgvStockCard, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(740, 80)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(132, 25)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(600, 80)
        Me.btnFilter.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(132, 25)
        Me.btnFilter.TabIndex = 4
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.btnClose)
        Me.gboDetail.Controls.Add(Me.btnFilter)
        Me.gboDetail.Controls.Add(Me.btnPrint)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.dtpEndDate)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.dtpStartDate)
        Me.gboDetail.Controls.Add(Me.btnFindProduct)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.txtDescription)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.txtCustCode)
        Me.gboDetail.Controls.Add(Me.txtPrimaryCust)
        Me.gboDetail.Location = New System.Drawing.Point(16, 15)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Size = New System.Drawing.Size(881, 145)
        Me.gboDetail.TabIndex = 2
        Me.gboDetail.TabStop = False
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Location = New System.Drawing.Point(600, 112)
        Me.btnPrint.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(132, 25)
        Me.btnPrint.TabIndex = 5
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 117)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(25, 17)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "To"
        '
        'dtpEndDate
        '
        Me.dtpEndDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpEndDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEndDate.Location = New System.Drawing.Point(125, 112)
        Me.dtpEndDate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(465, 22)
        Me.dtpEndDate.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 85)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 17)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "From"
        '
        'dtpStartDate
        '
        Me.dtpStartDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpStartDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStartDate.Location = New System.Drawing.Point(125, 80)
        Me.dtpStartDate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(465, 22)
        Me.dtpStartDate.TabIndex = 2
        '
        'btnFindProduct
        '
        Me.btnFindProduct.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindProduct.Location = New System.Drawing.Point(815, 16)
        Me.btnFindProduct.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFindProduct.Name = "btnFindProduct"
        Me.btnFindProduct.Size = New System.Drawing.Size(59, 25)
        Me.btnFindProduct.TabIndex = 4
        Me.btnFindProduct.TabStop = False
        Me.btnFindProduct.Text = "...."
        Me.btnFindProduct.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 52)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(109, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Customer Name"
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(125, 48)
        Me.txtDescription.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.ReadOnly = True
        Me.txtDescription.Size = New System.Drawing.Size(745, 22)
        Me.txtDescription.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 20)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(105, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Customer Code"
        '
        'txtCustCode
        '
        Me.txtCustCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCustCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustCode.Location = New System.Drawing.Point(125, 16)
        Me.txtCustCode.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtCustCode.Name = "txtCustCode"
        Me.txtCustCode.Size = New System.Drawing.Size(679, 22)
        Me.txtCustCode.TabIndex = 0
        '
        'txtPrimaryCust
        '
        Me.txtPrimaryCust.Location = New System.Drawing.Point(415, 31)
        Me.txtPrimaryCust.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimaryCust.Name = "txtPrimaryCust"
        Me.txtPrimaryCust.Size = New System.Drawing.Size(77, 22)
        Me.txtPrimaryCust.TabIndex = 12
        Me.txtPrimaryCust.TabStop = False
        Me.txtPrimaryCust.Text = "-1"
        Me.txtPrimaryCust.Visible = False
        '
        'dgvStockCard
        '
        Me.dgvStockCard.AllowUserToAddRows = False
        Me.dgvStockCard.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvStockCard.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvStockCard.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvStockCard.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvStockCard.ColumnHeadersHeight = 25
        Me.dgvStockCard.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvStockCard.Location = New System.Drawing.Point(16, 167)
        Me.dgvStockCard.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvStockCard.MultiSelect = False
        Me.dgvStockCard.Name = "dgvStockCard"
        Me.dgvStockCard.ReadOnly = True
        Me.dgvStockCard.RowHeadersVisible = False
        Me.dgvStockCard.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvStockCard.RowTemplate.Height = 24
        Me.dgvStockCard.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvStockCard.Size = New System.Drawing.Size(881, 555)
        Me.dgvStockCard.StandardTab = True
        Me.dgvStockCard.TabIndex = 3
        '
        'FrmLaporanKartuAR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(912, 737)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.dgvStockCard)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimumSize = New System.Drawing.Size(927, 771)
        Me.Name = "FrmLaporanKartuAR"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Account Receivables Card"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        CType(Me.dgvStockCard, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnFindProduct As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCustCode As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCust As System.Windows.Forms.TextBox
    Friend WithEvents dgvStockCard As System.Windows.Forms.DataGridView
End Class
