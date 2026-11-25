<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCariTransaksiStockInOut
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
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.tslQuery = New System.Windows.Forms.ToolStripStatusLabel
    Me.txtprimarygroup = New System.Windows.Forms.TextBox
    Me.gboKriteria = New System.Windows.Forms.GroupBox
    Me.dtpSampai = New System.Windows.Forms.DateTimePicker
    Me.dtpDari = New System.Windows.Forms.DateTimePicker
    Me.lblTo = New System.Windows.Forms.Label
    Me.LblFrom = New System.Windows.Forms.Label
    Me.cboCriteria = New System.Windows.Forms.ComboBox
    Me.btnOK = New System.Windows.Forms.Button
    Me.btnClose = New System.Windows.Forms.Button
    Me.btnFind = New System.Windows.Forms.Button
    Me.txtCari = New System.Windows.Forms.TextBox
    Me.dgvHasil = New System.Windows.Forms.DataGridView
    Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.ColKet = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.Tbltransaksi_stockinout1BindingSource = New System.Windows.Forms.BindingSource(Me.components)
    Me.DseTransaksiStockInOut = New Inventory_Module.dseTransaksiStockInOut
    Me.Tbltransaksi_stockinout1TableAdapter = New Inventory_Module.dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout1TableAdapter
    Me.StatusStrip1.SuspendLayout()
    Me.gboKriteria.SuspendLayout()
    CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.Tbltransaksi_stockinout1BindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.DseTransaksiStockInOut, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslQuery})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 371)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(842, 22)
    Me.StatusStrip1.TabIndex = 19
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'tslQuery
    '
    Me.tslQuery.Name = "tslQuery"
    Me.tslQuery.Size = New System.Drawing.Size(0, 17)
    '
    'txtprimarygroup
    '
    Me.txtprimarygroup.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtprimarygroup.Location = New System.Drawing.Point(156, 38)
    Me.txtprimarygroup.Name = "txtprimarygroup"
    Me.txtprimarygroup.Size = New System.Drawing.Size(394, 20)
    Me.txtprimarygroup.TabIndex = 6
    Me.txtprimarygroup.Visible = False
    '
    'gboKriteria
    '
    Me.gboKriteria.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.gboKriteria.Controls.Add(Me.dtpSampai)
    Me.gboKriteria.Controls.Add(Me.dtpDari)
    Me.gboKriteria.Controls.Add(Me.lblTo)
    Me.gboKriteria.Controls.Add(Me.LblFrom)
    Me.gboKriteria.Controls.Add(Me.txtprimarygroup)
    Me.gboKriteria.Controls.Add(Me.cboCriteria)
    Me.gboKriteria.Controls.Add(Me.btnOK)
    Me.gboKriteria.Controls.Add(Me.btnClose)
    Me.gboKriteria.Controls.Add(Me.btnFind)
    Me.gboKriteria.Controls.Add(Me.txtCari)
    Me.gboKriteria.Location = New System.Drawing.Point(12, 1)
    Me.gboKriteria.Name = "gboKriteria"
    Me.gboKriteria.Size = New System.Drawing.Size(810, 65)
    Me.gboKriteria.TabIndex = 17
    Me.gboKriteria.TabStop = False
    '
    'dtpSampai
    '
    Me.dtpSampai.CustomFormat = "ddd, dd MMM yyyy"
    Me.dtpSampai.Format = System.Windows.Forms.DateTimePickerFormat.Custom
    Me.dtpSampai.Location = New System.Drawing.Point(169, 37)
    Me.dtpSampai.Name = "dtpSampai"
    Me.dtpSampai.Size = New System.Drawing.Size(188, 20)
    Me.dtpSampai.TabIndex = 10
    Me.dtpSampai.Visible = False
    '
    'dtpDari
    '
    Me.dtpDari.CustomFormat = "ddd, dd MMM yyyy"
    Me.dtpDari.Format = System.Windows.Forms.DateTimePickerFormat.Custom
    Me.dtpDari.Location = New System.Drawing.Point(169, 13)
    Me.dtpDari.Name = "dtpDari"
    Me.dtpDari.Size = New System.Drawing.Size(188, 20)
    Me.dtpDari.TabIndex = 9
    Me.dtpDari.Visible = False
    '
    'lblTo
    '
    Me.lblTo.AutoSize = True
    Me.lblTo.Location = New System.Drawing.Point(133, 41)
    Me.lblTo.Name = "lblTo"
    Me.lblTo.Size = New System.Drawing.Size(20, 13)
    Me.lblTo.TabIndex = 8
    Me.lblTo.Text = "To"
    Me.lblTo.Visible = False
    '
    'LblFrom
    '
    Me.LblFrom.AutoSize = True
    Me.LblFrom.Location = New System.Drawing.Point(133, 16)
    Me.LblFrom.Name = "LblFrom"
    Me.LblFrom.Size = New System.Drawing.Size(30, 13)
    Me.LblFrom.TabIndex = 7
    Me.LblFrom.Text = "From"
    Me.LblFrom.Visible = False
    '
    'cboCriteria
    '
    Me.cboCriteria.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboCriteria.FormattingEnabled = True
    Me.cboCriteria.Location = New System.Drawing.Point(6, 12)
    Me.cboCriteria.Name = "cboCriteria"
    Me.cboCriteria.Size = New System.Drawing.Size(121, 21)
    Me.cboCriteria.TabIndex = 5
    '
    'btnOK
    '
    Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.btnOK.ImeMode = System.Windows.Forms.ImeMode.NoControl
    Me.btnOK.Location = New System.Drawing.Point(678, 38)
    Me.btnOK.Name = "btnOK"
    Me.btnOK.Size = New System.Drawing.Size(60, 23)
    Me.btnOK.TabIndex = 4
    Me.btnOK.TabStop = False
    Me.btnOK.Text = "&OK"
    Me.btnOK.UseVisualStyleBackColor = True
    '
    'btnClose
    '
    Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.btnClose.ImeMode = System.Windows.Forms.ImeMode.NoControl
    Me.btnClose.Location = New System.Drawing.Point(744, 38)
    Me.btnClose.Name = "btnClose"
    Me.btnClose.Size = New System.Drawing.Size(60, 23)
    Me.btnClose.TabIndex = 2
    Me.btnClose.TabStop = False
    Me.btnClose.Text = "&Close"
    Me.btnClose.UseVisualStyleBackColor = True
    '
    'btnFind
    '
    Me.btnFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.btnFind.ImeMode = System.Windows.Forms.ImeMode.NoControl
    Me.btnFind.Location = New System.Drawing.Point(744, 10)
    Me.btnFind.Name = "btnFind"
    Me.btnFind.Size = New System.Drawing.Size(60, 23)
    Me.btnFind.TabIndex = 1
    Me.btnFind.TabStop = False
    Me.btnFind.Text = "&Find"
    Me.btnFind.UseVisualStyleBackColor = True
    '
    'txtCari
    '
    Me.txtCari.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtCari.Location = New System.Drawing.Point(133, 12)
    Me.txtCari.Name = "txtCari"
    Me.txtCari.Size = New System.Drawing.Size(605, 20)
    Me.txtCari.TabIndex = 0
    '
    'dgvHasil
    '
    Me.dgvHasil.AllowUserToAddRows = False
    Me.dgvHasil.AllowUserToDeleteRows = False
    Me.dgvHasil.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvHasil.AutoGenerateColumns = False
    Me.dgvHasil.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.ColKet, Me.DataGridViewTextBoxColumn4})
    Me.dgvHasil.DataSource = Me.Tbltransaksi_stockinout1BindingSource
    Me.dgvHasil.Location = New System.Drawing.Point(12, 72)
    Me.dgvHasil.Name = "dgvHasil"
    Me.dgvHasil.ReadOnly = True
    Me.dgvHasil.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
    Me.dgvHasil.Size = New System.Drawing.Size(810, 295)
    Me.dgvHasil.TabIndex = 20
    Me.dgvHasil.TabStop = False
    '
    'DataGridViewTextBoxColumn1
    '
    Me.DataGridViewTextBoxColumn1.DataPropertyName = "Primary"
    Me.DataGridViewTextBoxColumn1.HeaderText = "Primary"
    Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
    Me.DataGridViewTextBoxColumn1.ReadOnly = True
    Me.DataGridViewTextBoxColumn1.Visible = False
    Me.DataGridViewTextBoxColumn1.Width = 66
    '
    'DataGridViewTextBoxColumn2
    '
    Me.DataGridViewTextBoxColumn2.DataPropertyName = "No Transaksi"
    Me.DataGridViewTextBoxColumn2.HeaderText = "No Transaksi"
    Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
    Me.DataGridViewTextBoxColumn2.ReadOnly = True
    Me.DataGridViewTextBoxColumn2.Width = 150
    '
    'DataGridViewTextBoxColumn3
    '
    Me.DataGridViewTextBoxColumn3.DataPropertyName = "Tanggal Transaksi"
    Me.DataGridViewTextBoxColumn3.HeaderText = "Tanggal Transaksi"
    Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
    Me.DataGridViewTextBoxColumn3.ReadOnly = True
    Me.DataGridViewTextBoxColumn3.Width = 120
    '
    'ColKet
    '
    Me.ColKet.HeaderText = "Keterangan"
    Me.ColKet.Name = "ColKet"
    Me.ColKet.ReadOnly = True
    Me.ColKet.Visible = False
    '
    'DataGridViewTextBoxColumn4
    '
    Me.DataGridViewTextBoxColumn4.DataPropertyName = "Remarks"
    Me.DataGridViewTextBoxColumn4.HeaderText = "Remarks"
    Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
    Me.DataGridViewTextBoxColumn4.ReadOnly = True
    Me.DataGridViewTextBoxColumn4.Width = 200
    '
    'Tbltransaksi_stockinout1BindingSource
    '
    Me.Tbltransaksi_stockinout1BindingSource.DataMember = "tbltransaksi_stockinout1"
    Me.Tbltransaksi_stockinout1BindingSource.DataSource = Me.DseTransaksiStockInOut
    '
    'DseTransaksiStockInOut
    '
    Me.DseTransaksiStockInOut.DataSetName = "dseTransaksiStockInOut"
    Me.DseTransaksiStockInOut.Locale = New System.Globalization.CultureInfo("id-ID")
    Me.DseTransaksiStockInOut.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
    '
    'Tbltransaksi_stockinout1TableAdapter
    '
    Me.Tbltransaksi_stockinout1TableAdapter.ClearBeforeFill = True
    '
    'FrmCariTransaksiStockInOut
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.btnClose
    Me.ClientSize = New System.Drawing.Size(842, 393)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.gboKriteria)
    Me.Controls.Add(Me.dgvHasil)
    Me.MinimizeBox = False
    Me.MinimumSize = New System.Drawing.Size(858, 432)
    Me.Name = "FrmCariTransaksiStockInOut"
    Me.ShowIcon = False
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Find Stock In - Adjust"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.gboKriteria.ResumeLayout(False)
    Me.gboKriteria.PerformLayout()
    CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.Tbltransaksi_stockinout1BindingSource, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.DseTransaksiStockInOut, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub

    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslQuery As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents txtprimarygroup As System.Windows.Forms.TextBox
    Friend WithEvents gboKriteria As System.Windows.Forms.GroupBox
    Public WithEvents cboCriteria As System.Windows.Forms.ComboBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Public WithEvents txtCari As System.Windows.Forms.TextBox
    Friend WithEvents lblTo As System.Windows.Forms.Label
    Friend WithEvents LblFrom As System.Windows.Forms.Label
    Friend WithEvents dtpSampai As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpDari As System.Windows.Forms.DateTimePicker
    Friend WithEvents DseTransaksiStockInOut As Inventory_Module.dseTransaksiStockInOut
    Friend WithEvents Tbltransaksi_stockinout1BindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Tbltransaksi_stockinout1TableAdapter As Inventory_Module.dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout1TableAdapter
    Friend WithEvents dgvHasil As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKet As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
