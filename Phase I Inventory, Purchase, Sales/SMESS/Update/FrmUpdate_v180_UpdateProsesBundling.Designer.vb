<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmUpdate_v180_UpdateProsesBundling
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
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.ColPrimaryWorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Button1 = New System.Windows.Forms.Button
        Me.ColPrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgv1 = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCategoryCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkUnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCategory = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LblProgress = New System.Windows.Forms.Label
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(93, 12)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(721, 23)
        Me.ProgressBar1.TabIndex = 6
        '
        'ColPrimaryWorkUnit
        '
        Me.ColPrimaryWorkUnit.HeaderText = "Primary Work Unit"
        Me.ColPrimaryWorkUnit.Name = "ColPrimaryWorkUnit"
        Me.ColPrimaryWorkUnit.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Start"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ColPrimaryProcess
        '
        Me.ColPrimaryProcess.HeaderText = "Primary Process"
        Me.ColPrimaryProcess.Name = "ColPrimaryProcess"
        Me.ColPrimaryProcess.Visible = False
        '
        'dgv1
        '
        Me.dgv1.AllowUserToAddRows = False
        Me.dgv1.AllowUserToDeleteRows = False
        Me.dgv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgv1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCategoryCode, Me.ColWorkUnitCode, Me.ColPrimaryCategory, Me.ColProcessCode, Me.ColPrimaryProcess, Me.ColPrimaryWorkUnit})
        Me.dgv1.Location = New System.Drawing.Point(12, 63)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgv1.RowTemplate.Height = 24
        Me.dgv1.Size = New System.Drawing.Size(802, 452)
        Me.dgv1.TabIndex = 4
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        '
        'ColCategoryCode
        '
        Me.ColCategoryCode.HeaderText = "Category Code"
        Me.ColCategoryCode.Name = "ColCategoryCode"
        '
        'ColWorkUnitCode
        '
        Me.ColWorkUnitCode.HeaderText = "Work Unit"
        Me.ColWorkUnitCode.Name = "ColWorkUnitCode"
        '
        'ColPrimaryCategory
        '
        Me.ColPrimaryCategory.HeaderText = "Primary Category"
        Me.ColPrimaryCategory.Name = "ColPrimaryCategory"
        Me.ColPrimaryCategory.Visible = False
        '
        'ColProcessCode
        '
        Me.ColProcessCode.HeaderText = "Process Code"
        Me.ColProcessCode.Name = "ColProcessCode"
        '
        'LblProgress
        '
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(90, 43)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(51, 17)
        Me.LblProgress.TabIndex = 7
        Me.LblProgress.Text = "Label1"
        '
        'FrmUpdate_v180_UpdateProsesBundling
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(826, 527)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.dgv1)
        Me.Controls.Add(Me.LblProgress)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(834, 564)
        Me.Name = "FrmUpdate_v180_UpdateProsesBundling"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Bundling Data Update"
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents ColPrimaryWorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ColPrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCategoryCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkUnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCategory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LblProgress As System.Windows.Forms.Label
End Class
