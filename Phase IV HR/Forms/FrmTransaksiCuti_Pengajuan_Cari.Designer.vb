<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiCuti_Pengajuan_Cari
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.txtprimarygroup = New System.Windows.Forms.TextBox
        Me.cboCriteria = New System.Windows.Forms.ComboBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.tslQuery = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.gboKriteria = New System.Windows.Forms.GroupBox
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.txtCari = New System.Windows.Forms.TextBox
        Me.dgvHasil = New System.Windows.Forms.DataGridView
        Me.PrimaryDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NoTransaksiDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TanggalDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.EmployeeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LeaveDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.StartDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.EndDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DurationDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TbltrancutipengajuanviewBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DseTrans_Cuti_Pengajuan = New HR.DseTrans_Cuti_Pengajuan
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.TbltranscutipengajuanBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Tbltrans_cuti_pengajuanTableAdapter = New HR.DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Me.Tbltrancutipengajuan_viewTableAdapter = New HR.DseTrans_Cuti_PengajuanTableAdapters.tbltrancutipengajuan_viewTableAdapter
        Me.gboKriteria.SuspendLayout()
        CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TbltrancutipengajuanviewBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DseTrans_Cuti_Pengajuan, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.TbltranscutipengajuanBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtprimarygroup
        '
        Me.txtprimarygroup.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtprimarygroup.Location = New System.Drawing.Point(8, 47)
        Me.txtprimarygroup.Margin = New System.Windows.Forms.Padding(4)
        Me.txtprimarygroup.Name = "txtprimarygroup"
        Me.txtprimarygroup.Size = New System.Drawing.Size(208, 22)
        Me.txtprimarygroup.TabIndex = 6
        Me.txtprimarygroup.TabStop = False
        Me.txtprimarygroup.Visible = False
        '
        'cboCriteria
        '
        Me.cboCriteria.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCriteria.FormattingEnabled = True
        Me.cboCriteria.Items.AddRange(New Object() {"Document No", "Date", "Employee Code", "Leave Code"})
        Me.cboCriteria.Location = New System.Drawing.Point(8, 15)
        Me.cboCriteria.Margin = New System.Windows.Forms.Padding(4)
        Me.cboCriteria.Name = "cboCriteria"
        Me.cboCriteria.Size = New System.Drawing.Size(208, 24)
        Me.cboCriteria.TabIndex = 5
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnOK.Location = New System.Drawing.Point(696, 47)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 28)
        Me.btnOK.TabIndex = 4
        Me.btnOK.TabStop = False
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'tslQuery
        '
        Me.tslQuery.Name = "tslQuery"
        Me.tslQuery.Size = New System.Drawing.Size(0, 17)
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnClose.Location = New System.Drawing.Point(784, 47)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 28)
        Me.btnClose.TabIndex = 2
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFind.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnFind.Location = New System.Drawing.Point(784, 12)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 28)
        Me.btnFind.TabIndex = 1
        Me.btnFind.TabStop = False
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'gboKriteria
        '
        Me.gboKriteria.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboKriteria.Controls.Add(Me.DateTimePicker2)
        Me.gboKriteria.Controls.Add(Me.Label2)
        Me.gboKriteria.Controls.Add(Me.Label1)
        Me.gboKriteria.Controls.Add(Me.DateTimePicker1)
        Me.gboKriteria.Controls.Add(Me.txtprimarygroup)
        Me.gboKriteria.Controls.Add(Me.cboCriteria)
        Me.gboKriteria.Controls.Add(Me.btnOK)
        Me.gboKriteria.Controls.Add(Me.btnClose)
        Me.gboKriteria.Controls.Add(Me.btnFind)
        Me.gboKriteria.Controls.Add(Me.txtCari)
        Me.gboKriteria.Location = New System.Drawing.Point(16, 1)
        Me.gboKriteria.Margin = New System.Windows.Forms.Padding(4)
        Me.gboKriteria.Name = "gboKriteria"
        Me.gboKriteria.Padding = New System.Windows.Forms.Padding(4)
        Me.gboKriteria.Size = New System.Drawing.Size(872, 80)
        Me.gboKriteria.TabIndex = 33
        Me.gboKriteria.TabStop = False
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CustomFormat = "dd MMM yyyy"
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker2.Location = New System.Drawing.Point(544, 15)
        Me.DateTimePicker2.Margin = New System.Windows.Forms.Padding(4)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(231, 22)
        Me.DateTimePicker2.TabIndex = 10
        Me.DateTimePicker2.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(513, 18)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(25, 17)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "To"
        Me.Label2.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(225, 18)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 17)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "From"
        Me.Label1.Visible = False
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "dd MMM yyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(273, 16)
        Me.DateTimePicker1.Margin = New System.Windows.Forms.Padding(4)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(231, 22)
        Me.DateTimePicker1.TabIndex = 7
        Me.DateTimePicker1.Visible = False
        '
        'txtCari
        '
        Me.txtCari.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCari.Location = New System.Drawing.Point(225, 15)
        Me.txtCari.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCari.Name = "txtCari"
        Me.txtCari.Size = New System.Drawing.Size(549, 22)
        Me.txtCari.TabIndex = 0
        '
        'dgvHasil
        '
        Me.dgvHasil.AllowUserToAddRows = False
        Me.dgvHasil.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvHasil.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvHasil.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvHasil.AutoGenerateColumns = False
        Me.dgvHasil.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvHasil.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.PrimaryDataGridViewTextBoxColumn, Me.NoTransaksiDataGridViewTextBoxColumn, Me.TanggalDataGridViewTextBoxColumn, Me.EmployeeDataGridViewTextBoxColumn, Me.LeaveDataGridViewTextBoxColumn, Me.StartDateDataGridViewTextBoxColumn, Me.EndDateDataGridViewTextBoxColumn, Me.DurationDataGridViewTextBoxColumn})
        Me.dgvHasil.DataSource = Me.TbltrancutipengajuanviewBindingSource
        Me.dgvHasil.Location = New System.Drawing.Point(16, 89)
        Me.dgvHasil.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvHasil.MultiSelect = False
        Me.dgvHasil.Name = "dgvHasil"
        Me.dgvHasil.ReadOnly = True
        Me.dgvHasil.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvHasil.Size = New System.Drawing.Size(872, 380)
        Me.dgvHasil.TabIndex = 32
        Me.dgvHasil.TabStop = False
        '
        'PrimaryDataGridViewTextBoxColumn
        '
        Me.PrimaryDataGridViewTextBoxColumn.DataPropertyName = "primary"
        Me.PrimaryDataGridViewTextBoxColumn.HeaderText = "primary"
        Me.PrimaryDataGridViewTextBoxColumn.Name = "PrimaryDataGridViewTextBoxColumn"
        Me.PrimaryDataGridViewTextBoxColumn.ReadOnly = True
        Me.PrimaryDataGridViewTextBoxColumn.Visible = False
        Me.PrimaryDataGridViewTextBoxColumn.Width = 65
        '
        'NoTransaksiDataGridViewTextBoxColumn
        '
        Me.NoTransaksiDataGridViewTextBoxColumn.DataPropertyName = "no transaksi"
        Me.NoTransaksiDataGridViewTextBoxColumn.HeaderText = "Document No"
        Me.NoTransaksiDataGridViewTextBoxColumn.Name = "NoTransaksiDataGridViewTextBoxColumn"
        Me.NoTransaksiDataGridViewTextBoxColumn.ReadOnly = True
        Me.NoTransaksiDataGridViewTextBoxColumn.Width = 119
        '
        'TanggalDataGridViewTextBoxColumn
        '
        Me.TanggalDataGridViewTextBoxColumn.DataPropertyName = "tanggal"
        DataGridViewCellStyle2.Format = "dd MMM yyyy"
        DataGridViewCellStyle2.NullValue = Nothing
        Me.TanggalDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle2
        Me.TanggalDataGridViewTextBoxColumn.HeaderText = "Date"
        Me.TanggalDataGridViewTextBoxColumn.Name = "TanggalDataGridViewTextBoxColumn"
        Me.TanggalDataGridViewTextBoxColumn.ReadOnly = True
        Me.TanggalDataGridViewTextBoxColumn.Width = 63
        '
        'EmployeeDataGridViewTextBoxColumn
        '
        Me.EmployeeDataGridViewTextBoxColumn.DataPropertyName = "Employee"
        Me.EmployeeDataGridViewTextBoxColumn.HeaderText = "Employee"
        Me.EmployeeDataGridViewTextBoxColumn.Name = "EmployeeDataGridViewTextBoxColumn"
        Me.EmployeeDataGridViewTextBoxColumn.ReadOnly = True
        Me.EmployeeDataGridViewTextBoxColumn.Width = 95
        '
        'LeaveDataGridViewTextBoxColumn
        '
        Me.LeaveDataGridViewTextBoxColumn.DataPropertyName = "Leave"
        Me.LeaveDataGridViewTextBoxColumn.HeaderText = "Leave"
        Me.LeaveDataGridViewTextBoxColumn.Name = "LeaveDataGridViewTextBoxColumn"
        Me.LeaveDataGridViewTextBoxColumn.ReadOnly = True
        Me.LeaveDataGridViewTextBoxColumn.Width = 72
        '
        'StartDateDataGridViewTextBoxColumn
        '
        Me.StartDateDataGridViewTextBoxColumn.DataPropertyName = "start date"
        DataGridViewCellStyle3.Format = "dd MMM yyyy"
        DataGridViewCellStyle3.NullValue = Nothing
        Me.StartDateDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle3
        Me.StartDateDataGridViewTextBoxColumn.HeaderText = "Start date"
        Me.StartDateDataGridViewTextBoxColumn.Name = "StartDateDataGridViewTextBoxColumn"
        Me.StartDateDataGridViewTextBoxColumn.ReadOnly = True
        Me.StartDateDataGridViewTextBoxColumn.Width = 95
        '
        'EndDateDataGridViewTextBoxColumn
        '
        Me.EndDateDataGridViewTextBoxColumn.DataPropertyName = "end date"
        DataGridViewCellStyle4.Format = "dd MMM yyyy"
        DataGridViewCellStyle4.NullValue = Nothing
        Me.EndDateDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle4
        Me.EndDateDataGridViewTextBoxColumn.HeaderText = "End date"
        Me.EndDateDataGridViewTextBoxColumn.Name = "EndDateDataGridViewTextBoxColumn"
        Me.EndDateDataGridViewTextBoxColumn.ReadOnly = True
        Me.EndDateDataGridViewTextBoxColumn.Width = 90
        '
        'DurationDataGridViewTextBoxColumn
        '
        Me.DurationDataGridViewTextBoxColumn.DataPropertyName = "duration"
        DataGridViewCellStyle5.Format = "N1"
        DataGridViewCellStyle5.NullValue = Nothing
        Me.DurationDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle5
        Me.DurationDataGridViewTextBoxColumn.HeaderText = "Duration"
        Me.DurationDataGridViewTextBoxColumn.Name = "DurationDataGridViewTextBoxColumn"
        Me.DurationDataGridViewTextBoxColumn.ReadOnly = True
        Me.DurationDataGridViewTextBoxColumn.Width = 87
        '
        'TbltrancutipengajuanviewBindingSource
        '
        Me.TbltrancutipengajuanviewBindingSource.DataMember = "tbltrancutipengajuan_view"
        Me.TbltrancutipengajuanviewBindingSource.DataSource = Me.DseTrans_Cuti_Pengajuan
        '
        'DseTrans_Cuti_Pengajuan
        '
        Me.DseTrans_Cuti_Pengajuan.DataSetName = "DseTrans_Cuti_Pengajuan"
        Me.DseTrans_Cuti_Pengajuan.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslQuery})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 480)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(904, 22)
        Me.StatusStrip1.TabIndex = 34
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'TbltranscutipengajuanBindingSource
        '
        Me.TbltranscutipengajuanBindingSource.DataMember = "tbltrans_cuti_pengajuan"
        Me.TbltranscutipengajuanBindingSource.DataSource = Me.DseTrans_Cuti_Pengajuan
        '
        'Tbltrans_cuti_pengajuanTableAdapter
        '
        Me.Tbltrans_cuti_pengajuanTableAdapter.ClearBeforeFill = True
        '
        'Tbltrancutipengajuan_viewTableAdapter
        '
        Me.Tbltrancutipengajuan_viewTableAdapter.ClearBeforeFill = True
        '
        'FrmTransaksiCuti_Pengajuan_Cari
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(904, 502)
        Me.Controls.Add(Me.gboKriteria)
        Me.Controls.Add(Me.dgvHasil)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "FrmTransaksiCuti_Pengajuan_Cari"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Find Leave Application"
        Me.gboKriteria.ResumeLayout(False)
        Me.gboKriteria.PerformLayout()
        CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TbltrancutipengajuanviewBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DseTrans_Cuti_Pengajuan, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.TbltranscutipengajuanBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtprimarygroup As System.Windows.Forms.TextBox
    Public WithEvents cboCriteria As System.Windows.Forms.ComboBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents tslQuery As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents gboKriteria As System.Windows.Forms.GroupBox
    Public WithEvents txtCari As System.Windows.Forms.TextBox
    Friend WithEvents dgvHasil As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents DseTrans_Cuti_Pengajuan As HR.DseTrans_Cuti_Pengajuan
    Friend WithEvents TbltranscutipengajuanBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Tbltrans_cuti_pengajuanTableAdapter As HR.DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents TbltrancutipengajuanviewBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Tbltrancutipengajuan_viewTableAdapter As HR.DseTrans_Cuti_PengajuanTableAdapters.tbltrancutipengajuan_viewTableAdapter
    Friend WithEvents PrimaryDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NoTransaksiDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TanggalDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EmployeeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LeaveDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StartDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EndDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DurationDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
