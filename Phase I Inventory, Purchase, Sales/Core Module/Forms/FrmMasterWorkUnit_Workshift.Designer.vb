<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterWorkUnit_Workshift
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
        Me.DataGridView2 = New System.Windows.Forms.DataGridView
        Me.Label7 = New System.Windows.Forms.Label
        Me.dtpToday = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.DataGridView3 = New System.Windows.Forms.DataGridView
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ColWorkUnit_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMachine_Allocate = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColMachine_PrimaryWorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkshift_WorkshiftCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkshift_Allocate = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColWorkshift_PrimaryWorkShift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllocate_WorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllocate_WorkShift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllocate_Allocate = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColAllocate_PrimaryWorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllocate_PrimaryWorkshift = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView2
        '
        Me.DataGridView2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.DataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWorkUnit_Code, Me.ColMachine_Allocate, Me.ColMachine_PrimaryWorkUnit})
        Me.DataGridView2.Location = New System.Drawing.Point(7, 80)
        Me.DataGridView2.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView2.MultiSelect = False
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowHeadersVisible = False
        Me.DataGridView2.RowTemplate.Height = 24
        Me.DataGridView2.Size = New System.Drawing.Size(297, 357)
        Me.DataGridView2.TabIndex = 3
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(4, 54)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(33, 17)
        Me.Label7.TabIndex = 90
        Me.Label7.Text = "End"
        '
        'dtpToday
        '
        Me.dtpToday.CustomFormat = "dd MMM yyyy"
        Me.dtpToday.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToday.Location = New System.Drawing.Point(120, 50)
        Me.dtpToday.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpToday.Name = "dtpToday"
        Me.dtpToday.Size = New System.Drawing.Size(184, 22)
        Me.dtpToday.TabIndex = 89
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(4, 26)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(38, 17)
        Me.Label5.TabIndex = 88
        Me.Label5.Text = "Start"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(120, 22)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(184, 22)
        Me.dtpFrom.TabIndex = 87
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColWorkshift_WorkshiftCode, Me.ColWorkshift_Allocate, Me.ColWorkshift_PrimaryWorkShift})
        Me.DataGridView1.Location = New System.Drawing.Point(7, 445)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(297, 160)
        Me.DataGridView1.TabIndex = 91
        '
        'DataGridView3
        '
        Me.DataGridView3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView3.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView3.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataGridView3.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAllocate_WorkUnit, Me.ColAllocate_WorkShift, Me.ColAllocate_Allocate, Me.ColAllocate_PrimaryWorkUnit, Me.ColAllocate_PrimaryWorkshift})
        Me.DataGridView3.Location = New System.Drawing.Point(339, 30)
        Me.DataGridView3.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView3.MultiSelect = False
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.RowHeadersVisible = False
        Me.DataGridView3.RowTemplate.Height = 24
        Me.DataGridView3.Size = New System.Drawing.Size(552, 561)
        Me.DataGridView3.TabIndex = 92
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(815, 9)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(82, 17)
        Me.LinkLabel1.TabIndex = 93
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Uncheck All"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(743, 9)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(66, 17)
        Me.LinkLabel2.TabIndex = 94
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Check All"
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(816, 598)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 95
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DataGridView2)
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.dtpFrom)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.dtpToday)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 9)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(312, 612)
        Me.GroupBox1.TabIndex = 96
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter"
        '
        'ColWorkUnit_Code
        '
        Me.ColWorkUnit_Code.HeaderText = "Work Unit"
        Me.ColWorkUnit_Code.Name = "ColWorkUnit_Code"
        Me.ColWorkUnit_Code.Width = 95
        '
        'ColMachine_Allocate
        '
        Me.ColMachine_Allocate.HeaderText = "Allocate"
        Me.ColMachine_Allocate.Name = "ColMachine_Allocate"
        Me.ColMachine_Allocate.Width = 64
        '
        'ColMachine_PrimaryWorkUnit
        '
        Me.ColMachine_PrimaryWorkUnit.HeaderText = "Primary WorkUnit"
        Me.ColMachine_PrimaryWorkUnit.Name = "ColMachine_PrimaryWorkUnit"
        Me.ColMachine_PrimaryWorkUnit.Visible = False
        Me.ColMachine_PrimaryWorkUnit.Width = 143
        '
        'ColWorkshift_WorkshiftCode
        '
        Me.ColWorkshift_WorkshiftCode.HeaderText = "Workshift"
        Me.ColWorkshift_WorkshiftCode.Name = "ColWorkshift_WorkshiftCode"
        Me.ColWorkshift_WorkshiftCode.Width = 92
        '
        'ColWorkshift_Allocate
        '
        Me.ColWorkshift_Allocate.HeaderText = "Allocate"
        Me.ColWorkshift_Allocate.Name = "ColWorkshift_Allocate"
        Me.ColWorkshift_Allocate.Width = 64
        '
        'ColWorkshift_PrimaryWorkShift
        '
        Me.ColWorkshift_PrimaryWorkShift.HeaderText = "Primary Workshift"
        Me.ColWorkshift_PrimaryWorkShift.Name = "ColWorkshift_PrimaryWorkShift"
        Me.ColWorkshift_PrimaryWorkShift.Visible = False
        Me.ColWorkshift_PrimaryWorkShift.Width = 144
        '
        'ColAllocate_WorkUnit
        '
        Me.ColAllocate_WorkUnit.HeaderText = "Work Unit"
        Me.ColAllocate_WorkUnit.Name = "ColAllocate_WorkUnit"
        Me.ColAllocate_WorkUnit.Width = 95
        '
        'ColAllocate_WorkShift
        '
        Me.ColAllocate_WorkShift.HeaderText = "Workshift"
        Me.ColAllocate_WorkShift.Name = "ColAllocate_WorkShift"
        Me.ColAllocate_WorkShift.Width = 92
        '
        'ColAllocate_Allocate
        '
        Me.ColAllocate_Allocate.HeaderText = "Allocate"
        Me.ColAllocate_Allocate.Name = "ColAllocate_Allocate"
        Me.ColAllocate_Allocate.Width = 64
        '
        'ColAllocate_PrimaryWorkUnit
        '
        Me.ColAllocate_PrimaryWorkUnit.HeaderText = "Primary WorkUnit"
        Me.ColAllocate_PrimaryWorkUnit.Name = "ColAllocate_PrimaryWorkUnit"
        Me.ColAllocate_PrimaryWorkUnit.Visible = False
        Me.ColAllocate_PrimaryWorkUnit.Width = 143
        '
        'ColAllocate_PrimaryWorkshift
        '
        Me.ColAllocate_PrimaryWorkshift.HeaderText = "Primary Workshift"
        Me.ColAllocate_PrimaryWorkshift.Name = "ColAllocate_PrimaryWorkshift"
        Me.ColAllocate_PrimaryWorkshift.Visible = False
        Me.ColAllocate_PrimaryWorkshift.Width = 144
        '
        'FrmMasterWorkUnit_Workshift
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(904, 633)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.LinkLabel2)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.DataGridView3)
        Me.Name = "FrmMasterWorkUnit_Workshift"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Workshift"
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents dtpToday As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView3 As System.Windows.Forms.DataGridView
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ColWorkUnit_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMachine_Allocate As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColMachine_PrimaryWorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkshift_WorkshiftCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkshift_Allocate As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColWorkshift_PrimaryWorkShift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllocate_WorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllocate_WorkShift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllocate_Allocate As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColAllocate_PrimaryWorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllocate_PrimaryWorkshift As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
