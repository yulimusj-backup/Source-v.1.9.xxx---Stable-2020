<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockInAdj_QuickInput
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
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtQty = New System.Windows.Forms.TextBox
        Me.TxtRef = New System.Windows.Forms.TextBox
        Me.chkImportFromDB = New System.Windows.Forms.CheckBox
        Me.ChkMirror = New System.Windows.Forms.CheckBox
        Me.chkCoverMinus = New System.Windows.Forms.CheckBox
        Me.dtpCoverMinus_RefDate = New System.Windows.Forms.DateTimePicker
        Me.TxtCoverMinus_LimitData = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeight = 25
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(13, 14)
        Me.dgvItemProductGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(659, 241)
        Me.dgvItemProductGroup.TabIndex = 3
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.Frozen = True
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_No.Width = 32
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.Frozen = True
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.Width = 48
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.Frozen = True
        Me.ColProductGroup_CodeNo.HeaderText = "Group Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_CodeNo.Width = 91
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Desc.Width = 85
        '
        'ColProductGroup_Remarks
        '
        Me.ColProductGroup_Remarks.HeaderText = "Remarks"
        Me.ColProductGroup_Remarks.Name = "ColProductGroup_Remarks"
        Me.ColProductGroup_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Remarks.Width = 70
        '
        'ColProductGroup_PrimaryAcuan
        '
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary Acuan"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 106
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(605, 261)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(524, 261)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'BackgroundWorker3
        '
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 263)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(177, 17)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Please specify default qty: "
        '
        'TxtQty
        '
        Me.TxtQty.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtQty.Location = New System.Drawing.Point(195, 261)
        Me.TxtQty.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtQty.Name = "TxtQty"
        Me.TxtQty.Size = New System.Drawing.Size(111, 22)
        Me.TxtQty.TabIndex = 7
        Me.TxtQty.Text = "0"
        Me.TxtQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtRef
        '
        Me.TxtRef.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtRef.Location = New System.Drawing.Point(312, 261)
        Me.TxtRef.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtRef.Name = "TxtRef"
        Me.TxtRef.Size = New System.Drawing.Size(111, 22)
        Me.TxtRef.TabIndex = 8
        Me.TxtRef.TabStop = False
        Me.TxtRef.Text = "0"
        Me.TxtRef.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtRef.Visible = False
        '
        'chkImportFromDB
        '
        Me.chkImportFromDB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkImportFromDB.AutoSize = True
        Me.chkImportFromDB.Location = New System.Drawing.Point(312, 261)
        Me.chkImportFromDB.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkImportFromDB.Name = "chkImportFromDB"
        Me.chkImportFromDB.Size = New System.Drawing.Size(194, 21)
        Me.chkImportFromDB.TabIndex = 9
        Me.chkImportFromDB.Text = "Take data from access db"
        Me.chkImportFromDB.UseVisualStyleBackColor = True
        '
        'ChkMirror
        '
        Me.ChkMirror.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ChkMirror.AutoSize = True
        Me.ChkMirror.Location = New System.Drawing.Point(312, 288)
        Me.ChkMirror.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ChkMirror.Name = "ChkMirror"
        Me.ChkMirror.Size = New System.Drawing.Size(101, 21)
        Me.ChkMirror.TabIndex = 10
        Me.ChkMirror.Text = "Mirror Data"
        Me.ChkMirror.UseVisualStyleBackColor = True
        '
        'chkCoverMinus
        '
        Me.chkCoverMinus.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkCoverMinus.AutoSize = True
        Me.chkCoverMinus.Location = New System.Drawing.Point(312, 313)
        Me.chkCoverMinus.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkCoverMinus.Name = "chkCoverMinus"
        Me.chkCoverMinus.Size = New System.Drawing.Size(142, 21)
        Me.chkCoverMinus.TabIndex = 11
        Me.chkCoverMinus.Text = "Cover Minus Data"
        Me.chkCoverMinus.UseVisualStyleBackColor = True
        '
        'dtpCoverMinus_RefDate
        '
        Me.dtpCoverMinus_RefDate.CustomFormat = "dd MMM yyyy"
        Me.dtpCoverMinus_RefDate.Enabled = False
        Me.dtpCoverMinus_RefDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpCoverMinus_RefDate.Location = New System.Drawing.Point(460, 312)
        Me.dtpCoverMinus_RefDate.Name = "dtpCoverMinus_RefDate"
        Me.dtpCoverMinus_RefDate.Size = New System.Drawing.Size(196, 22)
        Me.dtpCoverMinus_RefDate.TabIndex = 12
        '
        'TxtCoverMinus_LimitData
        '
        Me.TxtCoverMinus_LimitData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtCoverMinus_LimitData.Enabled = False
        Me.TxtCoverMinus_LimitData.Location = New System.Drawing.Point(195, 288)
        Me.TxtCoverMinus_LimitData.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtCoverMinus_LimitData.Name = "TxtCoverMinus_LimitData"
        Me.TxtCoverMinus_LimitData.Size = New System.Drawing.Size(111, 22)
        Me.TxtCoverMinus_LimitData.TabIndex = 14
        Me.TxtCoverMinus_LimitData.Text = "0"
        Me.TxtCoverMinus_LimitData.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 290)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(167, 17)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Please specify limit data: "
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.Enabled = False
        Me.TextBox1.Location = New System.Drawing.Point(160, 314)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(111, 22)
        Me.TextBox1.TabIndex = 15
        Me.TextBox1.TabStop = False
        Me.TextBox1.Text = "0"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox1.Visible = False
        '
        'FrmTransaksiStockInAdj_QuickInput
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(695, 345)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.TxtCoverMinus_LimitData)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpCoverMinus_RefDate)
        Me.Controls.Add(Me.chkCoverMinus)
        Me.Controls.Add(Me.ChkMirror)
        Me.Controls.Add(Me.chkImportFromDB)
        Me.Controls.Add(Me.TxtRef)
        Me.Controls.Add(Me.TxtQty)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.dgvItemProductGroup)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(703, 351)
        Me.Name = "FrmTransaksiStockInAdj_QuickInput"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Please choose scope"
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtQty As System.Windows.Forms.TextBox
    Friend WithEvents TxtRef As System.Windows.Forms.TextBox
    Friend WithEvents chkImportFromDB As System.Windows.Forms.CheckBox
    Friend WithEvents ChkMirror As System.Windows.Forms.CheckBox
    Friend WithEvents chkCoverMinus As System.Windows.Forms.CheckBox
    Friend WithEvents dtpCoverMinus_RefDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtCoverMinus_LimitData As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
End Class
