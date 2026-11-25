<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterKaryawan_Kompetensi_Upd_ImportOpt
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
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.DgvImport = New System.Windows.Forms.DataGridView
        Me.ColImp_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColImp_Competence = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColImp_EmpoloyeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColImp_Value = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColImp_NextCertificateDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColImp_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColImp_Import = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.LlbChkAll = New System.Windows.Forms.LinkLabel
        Me.LlbUnchkAll = New System.Windows.Forms.LinkLabel
        CType(Me.DgvImport, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(420, 319)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(56, 20)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "&OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(481, 319)
        Me.Button2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(56, 20)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "&Close"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'DgvImport
        '
        Me.DgvImport.AllowUserToAddRows = False
        Me.DgvImport.AllowUserToDeleteRows = False
        Me.DgvImport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvImport.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvImport.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColImp_No, Me.ColImp_Competence, Me.ColImp_EmpoloyeeName, Me.ColImp_Value, Me.ColImp_NextCertificateDate, Me.ColImp_Remarks, Me.ColImp_Import})
        Me.DgvImport.Location = New System.Drawing.Point(11, 29)
        Me.DgvImport.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DgvImport.MultiSelect = False
        Me.DgvImport.Name = "DgvImport"
        Me.DgvImport.RowTemplate.Height = 24
        Me.DgvImport.Size = New System.Drawing.Size(526, 280)
        Me.DgvImport.TabIndex = 2
        '
        'ColImp_No
        '
        Me.ColImp_No.HeaderText = "No"
        Me.ColImp_No.Name = "ColImp_No"
        Me.ColImp_No.ReadOnly = True
        Me.ColImp_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColImp_No.Width = 50
        '
        'ColImp_Competence
        '
        Me.ColImp_Competence.HeaderText = "Competence"
        Me.ColImp_Competence.Name = "ColImp_Competence"
        Me.ColImp_Competence.ReadOnly = True
        '
        'ColImp_EmpoloyeeName
        '
        Me.ColImp_EmpoloyeeName.HeaderText = "Employee Name"
        Me.ColImp_EmpoloyeeName.Name = "ColImp_EmpoloyeeName"
        Me.ColImp_EmpoloyeeName.ReadOnly = True
        '
        'ColImp_Value
        '
        Me.ColImp_Value.HeaderText = "Value"
        Me.ColImp_Value.Name = "ColImp_Value"
        Me.ColImp_Value.ReadOnly = True
        '
        'ColImp_NextCertificateDate
        '
        Me.ColImp_NextCertificateDate.HeaderText = "Next Certificate Date"
        Me.ColImp_NextCertificateDate.Name = "ColImp_NextCertificateDate"
        Me.ColImp_NextCertificateDate.ReadOnly = True
        '
        'ColImp_Remarks
        '
        Me.ColImp_Remarks.HeaderText = "Remarks"
        Me.ColImp_Remarks.Name = "ColImp_Remarks"
        Me.ColImp_Remarks.ReadOnly = True
        '
        'ColImp_Import
        '
        Me.ColImp_Import.HeaderText = "Import"
        Me.ColImp_Import.Name = "ColImp_Import"
        Me.ColImp_Import.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColImp_Import.Width = 50
        '
        'LlbChkAll
        '
        Me.LlbChkAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbChkAll.AutoSize = True
        Me.LlbChkAll.Location = New System.Drawing.Point(412, 10)
        Me.LlbChkAll.Name = "LlbChkAll"
        Me.LlbChkAll.Size = New System.Drawing.Size(52, 13)
        Me.LlbChkAll.TabIndex = 3
        Me.LlbChkAll.TabStop = True
        Me.LlbChkAll.Text = "Check All"
        '
        'LlbUnchkAll
        '
        Me.LlbUnchkAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUnchkAll.AutoSize = True
        Me.LlbUnchkAll.Location = New System.Drawing.Point(468, 10)
        Me.LlbUnchkAll.Name = "LlbUnchkAll"
        Me.LlbUnchkAll.Size = New System.Drawing.Size(65, 13)
        Me.LlbUnchkAll.TabIndex = 4
        Me.LlbUnchkAll.TabStop = True
        Me.LlbUnchkAll.Text = "Uncheck All"
        '
        'FrmMasterKaryawan_Kompetensi_Upd_ImportOpt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(546, 350)
        Me.ControlBox = False
        Me.Controls.Add(Me.LlbUnchkAll)
        Me.Controls.Add(Me.LlbChkAll)
        Me.Controls.Add(Me.DgvImport)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimumSize = New System.Drawing.Size(562, 388)
        Me.Name = "FrmMasterKaryawan_Kompetensi_Upd_ImportOpt"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Import .CSV"
        CType(Me.DgvImport, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents DgvImport As System.Windows.Forms.DataGridView
    Friend WithEvents ColImp_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColImp_Competence As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColImp_EmpoloyeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColImp_Value As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColImp_NextCertificateDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColImp_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColImp_Import As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents LlbChkAll As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbUnchkAll As System.Windows.Forms.LinkLabel
End Class
