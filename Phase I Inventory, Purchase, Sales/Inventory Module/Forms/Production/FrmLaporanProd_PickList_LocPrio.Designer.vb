<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanProd_PickList_LocPrio
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmLaporanProd_PickList_LocPrio))
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.dgvLocation = New System.Windows.Forms.DataGridView
        Me.ColLocation_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation_Group = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColLocation_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation_PrimaryW = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation_Urutan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.BtnUp1 = New System.Windows.Forms.Button
        Me.dgvExcludeLocation = New System.Windows.Forms.DataGridView
        Me.ColE_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColE_Group = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColE_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColE_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColE_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColE_PrimaryW = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColE_Urutan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BtnDown1 = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.dgvLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvExcludeLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(562, 346)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(94, 27)
        Me.btnClose.TabIndex = 0
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(462, 346)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(94, 27)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'dgvLocation
        '
        Me.dgvLocation.AllowUserToResizeColumns = False
        Me.dgvLocation.AllowUserToResizeRows = False
        Me.dgvLocation.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvLocation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvLocation.ColumnHeadersHeight = 25
        Me.dgvLocation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvLocation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColLocation_No, Me.ColLocation_Group, Me.ColLocation_Code, Me.ColLocation_Desc, Me.ColLocation_Primary, Me.ColLocation_PrimaryW, Me.ColLocation_Urutan})
        Me.dgvLocation.Location = New System.Drawing.Point(7, 7)
        Me.dgvLocation.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvLocation.Name = "dgvLocation"
        Me.dgvLocation.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvLocation.RowTemplate.Height = 24
        Me.dgvLocation.Size = New System.Drawing.Size(626, 211)
        Me.dgvLocation.TabIndex = 4
        '
        'ColLocation_No
        '
        Me.ColLocation_No.HeaderText = "No"
        Me.ColLocation_No.Name = "ColLocation_No"
        Me.ColLocation_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLocation_No.Width = 32
        '
        'ColLocation_Group
        '
        Me.ColLocation_Group.HeaderText = "Location Level"
        Me.ColLocation_Group.Items.AddRange(New Object() {"Location", "Warehouse"})
        Me.ColLocation_Group.Name = "ColLocation_Group"
        Me.ColLocation_Group.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColLocation_Group.Width = 106
        '
        'ColLocation_Code
        '
        Me.ColLocation_Code.HeaderText = "Code"
        Me.ColLocation_Code.Name = "ColLocation_Code"
        Me.ColLocation_Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLocation_Code.Width = 47
        '
        'ColLocation_Desc
        '
        Me.ColLocation_Desc.HeaderText = "Desc"
        Me.ColLocation_Desc.Name = "ColLocation_Desc"
        Me.ColLocation_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLocation_Desc.Width = 46
        '
        'ColLocation_Primary
        '
        Me.ColLocation_Primary.HeaderText = "Primary"
        Me.ColLocation_Primary.Name = "ColLocation_Primary"
        Me.ColLocation_Primary.Visible = False
        Me.ColLocation_Primary.Width = 81
        '
        'ColLocation_PrimaryW
        '
        Me.ColLocation_PrimaryW.HeaderText = "Primary W"
        Me.ColLocation_PrimaryW.Name = "ColLocation_PrimaryW"
        Me.ColLocation_PrimaryW.ReadOnly = True
        Me.ColLocation_PrimaryW.Visible = False
        Me.ColLocation_PrimaryW.Width = 98
        '
        'ColLocation_Urutan
        '
        Me.ColLocation_Urutan.HeaderText = "Urutan"
        Me.ColLocation_Urutan.Name = "ColLocation_Urutan"
        Me.ColLocation_Urutan.ReadOnly = True
        Me.ColLocation_Urutan.Visible = False
        Me.ColLocation_Urutan.Width = 76
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 50)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(648, 294)
        Me.TabControl1.TabIndex = 5
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btnUP)
        Me.TabPage1.Controls.Add(Me.dgvLocation)
        Me.TabPage1.Controls.Add(Me.btnDown)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(640, 265)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Location"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = CType(resources.GetObject("btnUP.BackgroundImage"), System.Drawing.Image)
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(7, 226)
        Me.btnUP.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(37, 32)
        Me.btnUP.TabIndex = 13
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = CType(resources.GetObject("btnDown.BackgroundImage"), System.Drawing.Image)
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(52, 226)
        Me.btnDown.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(37, 32)
        Me.btnDown.TabIndex = 14
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.BtnUp1)
        Me.TabPage2.Controls.Add(Me.dgvExcludeLocation)
        Me.TabPage2.Controls.Add(Me.BtnDown1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(640, 265)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Exclude Location"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'BtnUp1
        '
        Me.BtnUp1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnUp1.BackColor = System.Drawing.Color.Transparent
        Me.BtnUp1.BackgroundImage = CType(resources.GetObject("BtnUp1.BackgroundImage"), System.Drawing.Image)
        Me.BtnUp1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.BtnUp1.ForeColor = System.Drawing.SystemColors.Control
        Me.BtnUp1.Location = New System.Drawing.Point(7, 226)
        Me.BtnUp1.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnUp1.Name = "BtnUp1"
        Me.BtnUp1.Size = New System.Drawing.Size(37, 32)
        Me.BtnUp1.TabIndex = 15
        Me.BtnUp1.TabStop = False
        Me.BtnUp1.UseVisualStyleBackColor = True
        '
        'dgvExcludeLocation
        '
        Me.dgvExcludeLocation.AllowUserToResizeColumns = False
        Me.dgvExcludeLocation.AllowUserToResizeRows = False
        Me.dgvExcludeLocation.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvExcludeLocation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvExcludeLocation.ColumnHeadersHeight = 25
        Me.dgvExcludeLocation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvExcludeLocation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColE_No, Me.ColE_Group, Me.ColE_Code, Me.ColE_Desc, Me.ColE_Primary, Me.ColE_PrimaryW, Me.ColE_Urutan})
        Me.dgvExcludeLocation.Location = New System.Drawing.Point(7, 7)
        Me.dgvExcludeLocation.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvExcludeLocation.Name = "dgvExcludeLocation"
        Me.dgvExcludeLocation.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvExcludeLocation.RowTemplate.Height = 24
        Me.dgvExcludeLocation.Size = New System.Drawing.Size(626, 211)
        Me.dgvExcludeLocation.TabIndex = 5
        '
        'ColE_No
        '
        Me.ColE_No.HeaderText = "No"
        Me.ColE_No.Name = "ColE_No"
        Me.ColE_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColE_No.Width = 32
        '
        'ColE_Group
        '
        Me.ColE_Group.HeaderText = "Location Level"
        Me.ColE_Group.Items.AddRange(New Object() {"Location", "Warehouse"})
        Me.ColE_Group.Name = "ColE_Group"
        Me.ColE_Group.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColE_Group.Width = 106
        '
        'ColE_Code
        '
        Me.ColE_Code.HeaderText = "Code"
        Me.ColE_Code.Name = "ColE_Code"
        Me.ColE_Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColE_Code.Width = 47
        '
        'ColE_Desc
        '
        Me.ColE_Desc.HeaderText = "Desc"
        Me.ColE_Desc.Name = "ColE_Desc"
        Me.ColE_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColE_Desc.Width = 46
        '
        'ColE_Primary
        '
        Me.ColE_Primary.HeaderText = "Primary"
        Me.ColE_Primary.Name = "ColE_Primary"
        Me.ColE_Primary.Visible = False
        Me.ColE_Primary.Width = 81
        '
        'ColE_PrimaryW
        '
        Me.ColE_PrimaryW.HeaderText = "Primary W"
        Me.ColE_PrimaryW.Name = "ColE_PrimaryW"
        Me.ColE_PrimaryW.ReadOnly = True
        Me.ColE_PrimaryW.Visible = False
        Me.ColE_PrimaryW.Width = 98
        '
        'ColE_Urutan
        '
        Me.ColE_Urutan.HeaderText = "Urutan"
        Me.ColE_Urutan.Name = "ColE_Urutan"
        Me.ColE_Urutan.ReadOnly = True
        Me.ColE_Urutan.Visible = False
        Me.ColE_Urutan.Width = 76
        '
        'BtnDown1
        '
        Me.BtnDown1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnDown1.BackColor = System.Drawing.Color.Transparent
        Me.BtnDown1.BackgroundImage = CType(resources.GetObject("BtnDown1.BackgroundImage"), System.Drawing.Image)
        Me.BtnDown1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.BtnDown1.ForeColor = System.Drawing.SystemColors.Control
        Me.BtnDown1.Location = New System.Drawing.Point(52, 226)
        Me.BtnDown1.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnDown1.Name = "BtnDown1"
        Me.BtnDown1.Size = New System.Drawing.Size(37, 32)
        Me.BtnDown1.TabIndex = 16
        Me.BtnDown1.TabStop = False
        Me.BtnDown1.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(199, 17)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Please specify location priority"
        '
        'FrmLaporanProd_PickList_LocPrio
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(672, 385)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnClose)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(662, 300)
        Me.Name = "FrmLaporanProd_PickList_LocPrio"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Location Priority"
        CType(Me.dgvLocation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvExcludeLocation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents dgvLocation As System.Windows.Forms.DataGridView
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvExcludeLocation As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColLocation_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Group As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColLocation_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_PrimaryW As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Urutan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColE_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColE_Group As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColE_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColE_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColE_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColE_PrimaryW As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColE_Urutan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents BtnUp1 As System.Windows.Forms.Button
    Friend WithEvents BtnDown1 As System.Windows.Forms.Button
End Class
