<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ManufacturingOrder_Exp_QtyRMReplace
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTrans_ManufacturingOrder_Exp_QtyRMReplace))
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.TxtWONumber = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtProcessCode = New System.Windows.Forms.TextBox
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.LblUnit3 = New System.Windows.Forms.Label
        Me.LblUnit2 = New System.Windows.Forms.Label
        Me.LblUnit1 = New System.Windows.Forms.Label
        Me.TxtQtyReplaced = New System.Windows.Forms.TextBox
        Me.TxtQtyAdditional = New System.Windows.Forms.TextBox
        Me.TxtQtyRequired = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.DgvItemRMReplace = New System.Windows.Forms.DataGridView
        Me.bgw1 = New System.ComponentModel.BackgroundWorker
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage_Data = New System.Windows.Forms.TabPage
        Me.TabPage_History = New System.Windows.Forms.TabPage
        Me.TabPage_DefaultData = New System.Windows.Forms.TabPage
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.ColDefault_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefault_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefault_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefault_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefault_RatioDesc = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColDefault_RatioReplacement = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefault_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefault_RatioReplaced = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefault_UnitReplaced = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefault_Add = New System.Windows.Forms.DataGridViewButtonColumn
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn9 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn10 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn11 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtPriority = New System.Windows.Forms.TextBox
        Me.btnUp = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDesc = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColQtyReplaced = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitReplaced = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReplace = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReplacement = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitReplaced = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DgvItemRMReplace, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage_Data.SuspendLayout()
        Me.TabPage_DefaultData.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TxtWONumber
        '
        Me.TxtWONumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtWONumber.Location = New System.Drawing.Point(98, 61)
        Me.TxtWONumber.Name = "TxtWONumber"
        Me.TxtWONumber.ReadOnly = True
        Me.TxtWONumber.Size = New System.Drawing.Size(279, 22)
        Me.TxtWONumber.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 17)
        Me.Label4.TabIndex = 70
        Me.Label4.Text = "Code No"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 8)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtWONumber)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtProcessCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCodeNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDesc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPriority)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblUnit3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblUnit2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblUnit1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtQtyReplaced)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtQtyAdditional)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtQtyRequired)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Size = New System.Drawing.Size(761, 119)
        Me.SplitContainer1.SplitterDistance = 380
        Me.SplitContainer1.TabIndex = 4
        '
        'TxtProcessCode
        '
        Me.TxtProcessCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtProcessCode.Location = New System.Drawing.Point(98, 89)
        Me.TxtProcessCode.Name = "TxtProcessCode"
        Me.TxtProcessCode.ReadOnly = True
        Me.TxtProcessCode.Size = New System.Drawing.Size(279, 22)
        Me.TxtProcessCode.TabIndex = 1
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeNo.Location = New System.Drawing.Point(98, 5)
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.ReadOnly = True
        Me.TxtCodeNo.Size = New System.Drawing.Size(279, 22)
        Me.TxtCodeNo.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 17)
        Me.Label1.TabIndex = 66
        Me.Label1.Text = "WO Number"
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(98, 33)
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.ReadOnly = True
        Me.TxtDesc.Size = New System.Drawing.Size(279, 22)
        Me.TxtDesc.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 92)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 17)
        Me.Label2.TabIndex = 68
        Me.Label2.Text = "Process"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 36)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(79, 17)
        Me.Label3.TabIndex = 72
        Me.Label3.Text = "Description"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(1, 36)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 17)
        Me.Label5.TabIndex = 89
        Me.Label5.Text = "Qty Additional"
        '
        'LblUnit3
        '
        Me.LblUnit3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnit3.AutoSize = True
        Me.LblUnit3.Location = New System.Drawing.Point(314, 64)
        Me.LblUnit3.Name = "LblUnit3"
        Me.LblUnit3.Size = New System.Drawing.Size(33, 17)
        Me.LblUnit3.TabIndex = 87
        Me.LblUnit3.Text = "Unit"
        '
        'LblUnit2
        '
        Me.LblUnit2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnit2.AutoSize = True
        Me.LblUnit2.Location = New System.Drawing.Point(314, 36)
        Me.LblUnit2.Name = "LblUnit2"
        Me.LblUnit2.Size = New System.Drawing.Size(33, 17)
        Me.LblUnit2.TabIndex = 86
        Me.LblUnit2.Text = "Unit"
        '
        'LblUnit1
        '
        Me.LblUnit1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnit1.AutoSize = True
        Me.LblUnit1.Location = New System.Drawing.Point(314, 8)
        Me.LblUnit1.Name = "LblUnit1"
        Me.LblUnit1.Size = New System.Drawing.Size(33, 17)
        Me.LblUnit1.TabIndex = 85
        Me.LblUnit1.Text = "Unit"
        '
        'TxtQtyReplaced
        '
        Me.TxtQtyReplaced.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyReplaced.Location = New System.Drawing.Point(101, 61)
        Me.TxtQtyReplaced.Name = "TxtQtyReplaced"
        Me.TxtQtyReplaced.ReadOnly = True
        Me.TxtQtyReplaced.Size = New System.Drawing.Size(207, 22)
        Me.TxtQtyReplaced.TabIndex = 81
        Me.TxtQtyReplaced.Text = "0"
        Me.TxtQtyReplaced.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtQtyAdditional
        '
        Me.TxtQtyAdditional.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyAdditional.Location = New System.Drawing.Point(101, 33)
        Me.TxtQtyAdditional.Name = "TxtQtyAdditional"
        Me.TxtQtyAdditional.ReadOnly = True
        Me.TxtQtyAdditional.Size = New System.Drawing.Size(207, 22)
        Me.TxtQtyAdditional.TabIndex = 76
        Me.TxtQtyAdditional.Text = "0"
        Me.TxtQtyAdditional.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtQtyRequired
        '
        Me.TxtQtyRequired.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyRequired.Location = New System.Drawing.Point(101, 5)
        Me.TxtQtyRequired.Name = "TxtQtyRequired"
        Me.TxtQtyRequired.ReadOnly = True
        Me.TxtQtyRequired.Size = New System.Drawing.Size(207, 22)
        Me.TxtQtyRequired.TabIndex = 73
        Me.TxtQtyRequired.Text = "0"
        Me.TxtQtyRequired.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(1, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(94, 17)
        Me.Label6.TabIndex = 78
        Me.Label6.Text = "Qty Replaced"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(1, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(92, 17)
        Me.Label7.TabIndex = 79
        Me.Label7.Text = "Qty Required"
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(617, 386)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 6
        Me.btnOK.TabStop = False
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(698, 386)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 7
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Cancel"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'DgvItemRMReplace
        '
        Me.DgvItemRMReplace.AllowUserToResizeColumns = False
        Me.DgvItemRMReplace.AllowUserToResizeRows = False
        Me.DgvItemRMReplace.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DgvItemRMReplace.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DgvItemRMReplace.ColumnHeadersHeight = 25
        Me.DgvItemRMReplace.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItemRMReplace.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColPrimaryCodeNo, Me.ColCodeNo, Me.ColDesc, Me.ColQty, Me.ColUnit, Me.ColRatioDesc, Me.ColQtyReplaced, Me.ColUnitReplaced, Me.ColRatioReplace, Me.ColRatioReplacement, Me.ColPrimaryUnitReplaced})
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DgvItemRMReplace.DefaultCellStyle = DataGridViewCellStyle4
        Me.DgvItemRMReplace.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvItemRMReplace.Location = New System.Drawing.Point(3, 3)
        Me.DgvItemRMReplace.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DgvItemRMReplace.Name = "DgvItemRMReplace"
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DgvItemRMReplace.RowHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.DgvItemRMReplace.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItemRMReplace.RowTemplate.Height = 24
        Me.DgvItemRMReplace.Size = New System.Drawing.Size(747, 212)
        Me.DgvItemRMReplace.TabIndex = 5
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage_Data)
        Me.TabControl1.Controls.Add(Me.TabPage_History)
        Me.TabControl1.Controls.Add(Me.TabPage_DefaultData)
        Me.TabControl1.Location = New System.Drawing.Point(12, 133)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(761, 247)
        Me.TabControl1.TabIndex = 9
        '
        'TabPage_Data
        '
        Me.TabPage_Data.Controls.Add(Me.DgvItemRMReplace)
        Me.TabPage_Data.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_Data.Name = "TabPage_Data"
        Me.TabPage_Data.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_Data.Size = New System.Drawing.Size(753, 218)
        Me.TabPage_Data.TabIndex = 0
        Me.TabPage_Data.Text = "Detail Replacement"
        Me.TabPage_Data.UseVisualStyleBackColor = True
        '
        'TabPage_History
        '
        Me.TabPage_History.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_History.Name = "TabPage_History"
        Me.TabPage_History.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_History.Size = New System.Drawing.Size(753, 218)
        Me.TabPage_History.TabIndex = 1
        Me.TabPage_History.Text = "Replacement History"
        Me.TabPage_History.UseVisualStyleBackColor = True
        '
        'TabPage_DefaultData
        '
        Me.TabPage_DefaultData.Controls.Add(Me.DataGridView1)
        Me.TabPage_DefaultData.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_DefaultData.Name = "TabPage_DefaultData"
        Me.TabPage_DefaultData.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_DefaultData.Size = New System.Drawing.Size(753, 218)
        Me.TabPage_DefaultData.TabIndex = 2
        Me.TabPage_DefaultData.Text = "Default Replacement Data"
        Me.TabPage_DefaultData.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToResizeColumns = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.DataGridView1.ColumnHeadersHeight = 25
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDefault_No, Me.ColDefault_PrimaryCodeNo, Me.ColDefault_CodeNo, Me.ColDefault_Description, Me.ColDefault_RatioDesc, Me.ColDefault_RatioReplacement, Me.ColDefault_Unit, Me.ColDefault_RatioReplaced, Me.ColDefault_UnitReplaced, Me.ColDefault_Add})
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle7
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DataGridView1.Name = "DataGridView1"
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.RowHeadersDefaultCellStyle = DataGridViewCellStyle8
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(747, 212)
        Me.DataGridView1.TabIndex = 6
        '
        'ColDefault_No
        '
        Me.ColDefault_No.HeaderText = "No"
        Me.ColDefault_No.Name = "ColDefault_No"
        Me.ColDefault_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDefault_No.Width = 32
        '
        'ColDefault_PrimaryCodeNo
        '
        Me.ColDefault_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColDefault_PrimaryCodeNo.Name = "ColDefault_PrimaryCodeNo"
        Me.ColDefault_PrimaryCodeNo.Visible = False
        Me.ColDefault_PrimaryCodeNo.Width = 140
        '
        'ColDefault_CodeNo
        '
        Me.ColDefault_CodeNo.HeaderText = "Code No"
        Me.ColDefault_CodeNo.Name = "ColDefault_CodeNo"
        Me.ColDefault_CodeNo.Width = 88
        '
        'ColDefault_Description
        '
        Me.ColDefault_Description.HeaderText = "Description"
        Me.ColDefault_Description.Name = "ColDefault_Description"
        Me.ColDefault_Description.Width = 104
        '
        'ColDefault_RatioDesc
        '
        Me.ColDefault_RatioDesc.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColDefault_RatioDesc.HeaderText = "Ratio"
        Me.ColDefault_RatioDesc.Name = "ColDefault_RatioDesc"
        Me.ColDefault_RatioDesc.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColDefault_RatioDesc.Width = 47
        '
        'ColDefault_RatioReplacement
        '
        Me.ColDefault_RatioReplacement.HeaderText = "Ratio Replacement"
        Me.ColDefault_RatioReplacement.Name = "ColDefault_RatioReplacement"
        Me.ColDefault_RatioReplacement.Width = 153
        '
        'ColDefault_Unit
        '
        Me.ColDefault_Unit.HeaderText = "Unit"
        Me.ColDefault_Unit.Name = "ColDefault_Unit"
        Me.ColDefault_Unit.Width = 58
        '
        'ColDefault_RatioReplaced
        '
        Me.ColDefault_RatioReplaced.HeaderText = "Ratio Replaced"
        Me.ColDefault_RatioReplaced.Name = "ColDefault_RatioReplaced"
        Me.ColDefault_RatioReplaced.Width = 130
        '
        'ColDefault_UnitReplaced
        '
        Me.ColDefault_UnitReplaced.HeaderText = "Unit"
        Me.ColDefault_UnitReplaced.Name = "ColDefault_UnitReplaced"
        Me.ColDefault_UnitReplaced.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDefault_UnitReplaced.Width = 39
        '
        'ColDefault_Add
        '
        Me.ColDefault_Add.HeaderText = "Use"
        Me.ColDefault_Add.Name = "ColDefault_Add"
        Me.ColDefault_Add.Width = 39
        '
        'BackgroundWorker1
        '
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "No"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.DataGridViewTextBoxColumn1.Width = 32
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Primary Code No"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.Visible = False
        Me.DataGridViewTextBoxColumn2.Width = 140
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "Code No"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Width = 88
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.HeaderText = "Description"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.Width = 104
        '
        'DataGridViewTextBoxColumn5
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn5.DefaultCellStyle = DataGridViewCellStyle9
        Me.DataGridViewTextBoxColumn5.HeaderText = "Qty"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.DataGridViewTextBoxColumn5.Width = 36
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.HeaderText = "Unit"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.Width = 58
        '
        'DataGridViewTextBoxColumn7
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn7.DefaultCellStyle = DataGridViewCellStyle10
        Me.DataGridViewTextBoxColumn7.HeaderText = "Qty Replaced"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.HeaderText = "Unit"
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        Me.DataGridViewTextBoxColumn8.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.DataGridViewTextBoxColumn8.Width = 39
        '
        'DataGridViewTextBoxColumn9
        '
        Me.DataGridViewTextBoxColumn9.HeaderText = "Ratio Replace"
        Me.DataGridViewTextBoxColumn9.Name = "DataGridViewTextBoxColumn9"
        Me.DataGridViewTextBoxColumn9.Visible = False
        Me.DataGridViewTextBoxColumn9.Width = 122
        '
        'DataGridViewTextBoxColumn10
        '
        Me.DataGridViewTextBoxColumn10.HeaderText = "Ratio Replacement"
        Me.DataGridViewTextBoxColumn10.Name = "DataGridViewTextBoxColumn10"
        Me.DataGridViewTextBoxColumn10.Visible = False
        Me.DataGridViewTextBoxColumn10.Width = 153
        '
        'DataGridViewTextBoxColumn11
        '
        Me.DataGridViewTextBoxColumn11.HeaderText = "Primary Unit Replaced"
        Me.DataGridViewTextBoxColumn11.Name = "DataGridViewTextBoxColumn11"
        Me.DataGridViewTextBoxColumn11.Visible = False
        Me.DataGridViewTextBoxColumn11.Width = 174
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(1, 92)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(52, 17)
        Me.Label8.TabIndex = 90
        Me.Label8.Text = "Priority"
        '
        'TxtPriority
        '
        Me.TxtPriority.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPriority.Location = New System.Drawing.Point(101, 89)
        Me.TxtPriority.Name = "TxtPriority"
        Me.TxtPriority.Size = New System.Drawing.Size(207, 22)
        Me.TxtPriority.TabIndex = 91
        Me.TxtPriority.Text = "0"
        Me.TxtPriority.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnUp
        '
        Me.btnUp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUp.BackColor = System.Drawing.Color.Transparent
        Me.btnUp.BackgroundImage = CType(resources.GetObject("btnUp.BackgroundImage"), System.Drawing.Image)
        Me.btnUp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUp.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUp.Location = New System.Drawing.Point(12, 383)
        Me.btnUp.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUp.Name = "btnUp"
        Me.btnUp.Size = New System.Drawing.Size(37, 32)
        Me.btnUp.TabIndex = 10
        Me.btnUp.TabStop = False
        Me.btnUp.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = CType(resources.GetObject("btnDown.BackgroundImage"), System.Drawing.Image)
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(54, 383)
        Me.btnDown.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(37, 32)
        Me.btnDown.TabIndex = 11
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "Priority"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 58
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 121
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCodeNo.Width = 69
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Width = 85
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty Replacement"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Width = 123
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 39
        '
        'ColRatioDesc
        '
        Me.ColRatioDesc.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColRatioDesc.HeaderText = "Ratio"
        Me.ColRatioDesc.Name = "ColRatioDesc"
        Me.ColRatioDesc.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRatioDesc.Width = 47
        '
        'ColQtyReplaced
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyReplaced.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColQtyReplaced.HeaderText = "Qty Replaced"
        Me.ColQtyReplaced.Name = "ColQtyReplaced"
        Me.ColQtyReplaced.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColUnitReplaced
        '
        Me.ColUnitReplaced.HeaderText = "Unit"
        Me.ColUnitReplaced.Name = "ColUnitReplaced"
        Me.ColUnitReplaced.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitReplaced.Width = 39
        '
        'ColRatioReplace
        '
        Me.ColRatioReplace.HeaderText = "Ratio Replace"
        Me.ColRatioReplace.Name = "ColRatioReplace"
        Me.ColRatioReplace.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRatioReplace.Visible = False
        Me.ColRatioReplace.Width = 103
        '
        'ColRatioReplacement
        '
        Me.ColRatioReplacement.HeaderText = "Ratio Replacement"
        Me.ColRatioReplacement.Name = "ColRatioReplacement"
        Me.ColRatioReplacement.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRatioReplacement.Visible = False
        Me.ColRatioReplacement.Width = 134
        '
        'ColPrimaryUnitReplaced
        '
        Me.ColPrimaryUnitReplaced.HeaderText = "Primary Unit Replaced"
        Me.ColPrimaryUnitReplaced.Name = "ColPrimaryUnitReplaced"
        Me.ColPrimaryUnitReplaced.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryUnitReplaced.Visible = False
        Me.ColPrimaryUnitReplaced.Width = 155
        '
        'FrmTrans_ManufacturingOrder_Exp_QtyRMReplace
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(785, 421)
        Me.Controls.Add(Me.btnUp)
        Me.Controls.Add(Me.btnDown)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnClose)
        Me.MinimumSize = New System.Drawing.Size(793, 334)
        Me.Name = "FrmTrans_ManufacturingOrder_Exp_QtyRMReplace"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Qty Replaced"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DgvItemRMReplace, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage_Data.ResumeLayout(False)
        Me.TabPage_DefaultData.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TxtWONumber As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtProcessCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents DgvItemRMReplace As System.Windows.Forms.DataGridView
    Friend WithEvents bgw1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage_Data As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_History As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_DefaultData As System.Windows.Forms.TabPage
    Friend WithEvents TxtQtyAdditional As System.Windows.Forms.TextBox
    Friend WithEvents TxtQtyRequired As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtQtyReplaced As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents LblUnit3 As System.Windows.Forms.Label
    Friend WithEvents LblUnit2 As System.Windows.Forms.Label
    Friend WithEvents LblUnit1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents ColDefault_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefault_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefault_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefault_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefault_RatioDesc As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColDefault_RatioReplacement As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefault_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefault_RatioReplaced As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefault_UnitReplaced As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefault_Add As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents TxtPriority As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnUp As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDesc As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColQtyReplaced As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitReplaced As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReplace As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReplacement As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitReplaced As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
