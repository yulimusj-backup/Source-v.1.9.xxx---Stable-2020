<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterSalaryFormula_Cari
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.gboKriteria = New System.Windows.Forms.GroupBox
        Me.txtprimarygroup = New System.Windows.Forms.TextBox
        Me.cboCriteria = New System.Windows.Forms.ComboBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.txtCari = New System.Windows.Forms.TextBox
        Me.dgvHasil = New System.Windows.Forms.DataGridView
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslQuery = New System.Windows.Forms.ToolStripStatusLabel
        Me.DseMasterSalaryFormula = New Core_Module.DseMasterSalaryFormula
        Me.TblmastersalaryformulaBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Tblmaster_salary_formulaTableAdapter = New Core_Module.DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter
        Me.PrimaryDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DescDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CreatedByDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CreatedDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ModifiedByDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ModifiedDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TimestampDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RemarksDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.gboKriteria.SuspendLayout()
        CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.DseMasterSalaryFormula, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TblmastersalaryformulaBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gboKriteria
        '
        Me.gboKriteria.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboKriteria.Controls.Add(Me.txtprimarygroup)
        Me.gboKriteria.Controls.Add(Me.cboCriteria)
        Me.gboKriteria.Controls.Add(Me.btnOK)
        Me.gboKriteria.Controls.Add(Me.btnClose)
        Me.gboKriteria.Controls.Add(Me.btnFind)
        Me.gboKriteria.Controls.Add(Me.txtCari)
        Me.gboKriteria.Location = New System.Drawing.Point(12, 12)
        Me.gboKriteria.Name = "gboKriteria"
        Me.gboKriteria.Size = New System.Drawing.Size(569, 65)
        Me.gboKriteria.TabIndex = 29
        Me.gboKriteria.TabStop = False
        '
        'txtprimarygroup
        '
        Me.txtprimarygroup.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtprimarygroup.Location = New System.Drawing.Point(133, 38)
        Me.txtprimarygroup.Name = "txtprimarygroup"
        Me.txtprimarygroup.Size = New System.Drawing.Size(194, 20)
        Me.txtprimarygroup.TabIndex = 6
        Me.txtprimarygroup.TabStop = False
        Me.txtprimarygroup.Visible = False
        '
        'cboCriteria
        '
        Me.cboCriteria.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCriteria.FormattingEnabled = True
        Me.cboCriteria.Items.AddRange(New Object() {"Code", "Description", "Remarks"})
        Me.cboCriteria.Location = New System.Drawing.Point(6, 12)
        Me.cboCriteria.Name = "cboCriteria"
        Me.cboCriteria.Size = New System.Drawing.Size(121, 21)
        Me.cboCriteria.TabIndex = 5
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnOK.Location = New System.Drawing.Point(437, 38)
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
        Me.btnClose.Location = New System.Drawing.Point(503, 38)
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
        Me.btnFind.Location = New System.Drawing.Point(503, 10)
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
        Me.txtCari.Size = New System.Drawing.Size(364, 20)
        Me.txtCari.TabIndex = 0
        '
        'dgvHasil
        '
        Me.dgvHasil.AllowUserToAddRows = False
        Me.dgvHasil.AllowUserToDeleteRows = False
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvHasil.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvHasil.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvHasil.AutoGenerateColumns = False
        Me.dgvHasil.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvHasil.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.PrimaryDataGridViewTextBoxColumn, Me.CodeDataGridViewTextBoxColumn, Me.DescDataGridViewTextBoxColumn, Me.CreatedByDataGridViewTextBoxColumn, Me.CreatedDateDataGridViewTextBoxColumn, Me.ModifiedByDataGridViewTextBoxColumn, Me.ModifiedDateDataGridViewTextBoxColumn, Me.TimestampDataGridViewTextBoxColumn, Me.RemarksDataGridViewTextBoxColumn})
        Me.dgvHasil.DataSource = Me.TblmastersalaryformulaBindingSource
        Me.dgvHasil.Location = New System.Drawing.Point(12, 79)
        Me.dgvHasil.MultiSelect = False
        Me.dgvHasil.Name = "dgvHasil"
        Me.dgvHasil.ReadOnly = True
        Me.dgvHasil.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvHasil.Size = New System.Drawing.Size(569, 349)
        Me.dgvHasil.TabIndex = 30
        Me.dgvHasil.TabStop = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslQuery})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 435)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(593, 22)
        Me.StatusStrip1.TabIndex = 31
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslQuery
        '
        Me.tslQuery.Name = "tslQuery"
        Me.tslQuery.Size = New System.Drawing.Size(0, 17)
        '
        'DseMasterSalaryFormula
        '
        Me.DseMasterSalaryFormula.DataSetName = "DseMasterSalaryFormula"
        Me.DseMasterSalaryFormula.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'TblmastersalaryformulaBindingSource
        '
        Me.TblmastersalaryformulaBindingSource.DataMember = "tblmaster_salary_formula"
        Me.TblmastersalaryformulaBindingSource.DataSource = Me.DseMasterSalaryFormula
        '
        'Tblmaster_salary_formulaTableAdapter
        '
        Me.Tblmaster_salary_formulaTableAdapter.ClearBeforeFill = True
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
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "Code"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        Me.CodeDataGridViewTextBoxColumn.ReadOnly = True
        Me.CodeDataGridViewTextBoxColumn.Width = 57
        '
        'DescDataGridViewTextBoxColumn
        '
        Me.DescDataGridViewTextBoxColumn.DataPropertyName = "desc"
        Me.DescDataGridViewTextBoxColumn.HeaderText = "Description"
        Me.DescDataGridViewTextBoxColumn.Name = "DescDataGridViewTextBoxColumn"
        Me.DescDataGridViewTextBoxColumn.ReadOnly = True
        Me.DescDataGridViewTextBoxColumn.Width = 85
        '
        'CreatedByDataGridViewTextBoxColumn
        '
        Me.CreatedByDataGridViewTextBoxColumn.DataPropertyName = "created by"
        Me.CreatedByDataGridViewTextBoxColumn.HeaderText = "created by"
        Me.CreatedByDataGridViewTextBoxColumn.Name = "CreatedByDataGridViewTextBoxColumn"
        Me.CreatedByDataGridViewTextBoxColumn.ReadOnly = True
        Me.CreatedByDataGridViewTextBoxColumn.Visible = False
        Me.CreatedByDataGridViewTextBoxColumn.Width = 82
        '
        'CreatedDateDataGridViewTextBoxColumn
        '
        Me.CreatedDateDataGridViewTextBoxColumn.DataPropertyName = "created date"
        Me.CreatedDateDataGridViewTextBoxColumn.HeaderText = "created date"
        Me.CreatedDateDataGridViewTextBoxColumn.Name = "CreatedDateDataGridViewTextBoxColumn"
        Me.CreatedDateDataGridViewTextBoxColumn.ReadOnly = True
        Me.CreatedDateDataGridViewTextBoxColumn.Visible = False
        Me.CreatedDateDataGridViewTextBoxColumn.Width = 92
        '
        'ModifiedByDataGridViewTextBoxColumn
        '
        Me.ModifiedByDataGridViewTextBoxColumn.DataPropertyName = "modified by"
        Me.ModifiedByDataGridViewTextBoxColumn.HeaderText = "modified by"
        Me.ModifiedByDataGridViewTextBoxColumn.Name = "ModifiedByDataGridViewTextBoxColumn"
        Me.ModifiedByDataGridViewTextBoxColumn.ReadOnly = True
        Me.ModifiedByDataGridViewTextBoxColumn.Visible = False
        Me.ModifiedByDataGridViewTextBoxColumn.Width = 85
        '
        'ModifiedDateDataGridViewTextBoxColumn
        '
        Me.ModifiedDateDataGridViewTextBoxColumn.DataPropertyName = "modified date"
        Me.ModifiedDateDataGridViewTextBoxColumn.HeaderText = "modified date"
        Me.ModifiedDateDataGridViewTextBoxColumn.Name = "ModifiedDateDataGridViewTextBoxColumn"
        Me.ModifiedDateDataGridViewTextBoxColumn.ReadOnly = True
        Me.ModifiedDateDataGridViewTextBoxColumn.Visible = False
        Me.ModifiedDateDataGridViewTextBoxColumn.Width = 95
        '
        'TimestampDataGridViewTextBoxColumn
        '
        Me.TimestampDataGridViewTextBoxColumn.DataPropertyName = "timestamp"
        Me.TimestampDataGridViewTextBoxColumn.HeaderText = "timestamp"
        Me.TimestampDataGridViewTextBoxColumn.Name = "TimestampDataGridViewTextBoxColumn"
        Me.TimestampDataGridViewTextBoxColumn.ReadOnly = True
        Me.TimestampDataGridViewTextBoxColumn.Visible = False
        Me.TimestampDataGridViewTextBoxColumn.Width = 79
        '
        'RemarksDataGridViewTextBoxColumn
        '
        Me.RemarksDataGridViewTextBoxColumn.DataPropertyName = "Remarks"
        Me.RemarksDataGridViewTextBoxColumn.HeaderText = "Remarks"
        Me.RemarksDataGridViewTextBoxColumn.Name = "RemarksDataGridViewTextBoxColumn"
        Me.RemarksDataGridViewTextBoxColumn.ReadOnly = True
        Me.RemarksDataGridViewTextBoxColumn.Width = 74
        '
        'FrmMasterSalaryFormula_Cari
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(593, 457)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.dgvHasil)
        Me.Controls.Add(Me.gboKriteria)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(609, 495)
        Me.Name = "FrmMasterSalaryFormula_Cari"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Find Salary Formula"
        Me.gboKriteria.ResumeLayout(False)
        Me.gboKriteria.PerformLayout()
        CType(Me.dgvHasil, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.DseMasterSalaryFormula, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TblmastersalaryformulaBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboKriteria As System.Windows.Forms.GroupBox
    Friend WithEvents txtprimarygroup As System.Windows.Forms.TextBox
    Public WithEvents cboCriteria As System.Windows.Forms.ComboBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Public WithEvents txtCari As System.Windows.Forms.TextBox
    Friend WithEvents dgvHasil As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslQuery As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents DseMasterSalaryFormula As Core_Module.DseMasterSalaryFormula
    Friend WithEvents TblmastersalaryformulaBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Tblmaster_salary_formulaTableAdapter As Core_Module.DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formulaTableAdapter
    Friend WithEvents PrimaryDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DescDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CreatedByDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CreatedDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ModifiedByDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ModifiedDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TimestampDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RemarksDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
