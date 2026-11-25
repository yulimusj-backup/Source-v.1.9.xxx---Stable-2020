<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SPL_Authorize
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.TboAuthorize = New System.Windows.Forms.TabControl
        Me.TpAuthorize = New System.Windows.Forms.TabPage
        Me.DgvOtorisasi = New System.Windows.Forms.DataGridView
        Me.ColDgvAuth_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDgvAuth_Authorize = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDgvAuth_SPLNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDgvAuth_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LlbUncheckAll_Athorize = New System.Windows.Forms.LinkLabel
        Me.LlbCheckAll_Authorize = New System.Windows.Forms.LinkLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.TpUnauthorize = New System.Windows.Forms.TabPage
        Me.DgvCancelOtorisasi = New System.Windows.Forms.DataGridView
        Me.ColDgvUnAuth_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDgvUnAuth_Cancel = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDgvUnAuth_SPLNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDgvUnAuth_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LlbUncheckAll_Cancel = New System.Windows.Forms.LinkLabel
        Me.LlbCheckAll_Cancel = New System.Windows.Forms.LinkLabel
        Me.TpAllauthorize = New System.Windows.Forms.TabPage
        Me.LlbUncheckAll_All = New System.Windows.Forms.LinkLabel
        Me.LlbCheckAll_ALL = New System.Windows.Forms.LinkLabel
        Me.DgvAllautorisasi = New System.Windows.Forms.DataGridView
        Me.col3_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col3_DocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col3_Print = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.col3_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col3_Status = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GboFilter = New System.Windows.Forms.GroupBox
        Me.CboStatus = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtNoSPL = New System.Windows.Forms.TextBox
        Me.BtnFilter = New System.Windows.Forms.Button
        Me.Label10 = New System.Windows.Forms.Label
        Me.BtnFindSPL = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.ChkSPLDate = New System.Windows.Forms.CheckBox
        Me.DgvUser = New System.Windows.Forms.DataGridView
        Me.ColCreatedBy_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TxtPrimarySPL = New System.Windows.Forms.TextBox
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.DgvWorkshift = New System.Windows.Forms.DataGridView
        Me.colNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colShiftCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEarlyOT = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colEarlyOTDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEarlyDelayDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEalryBreak = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDelayBeforeEarlyBreak = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colOT = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colOTDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDelayDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colBreak = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDelayBeforeBreak = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPrimaryWorkshift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DtpFromDetail = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.DtpToDetail = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.GboItem2 = New System.Windows.Forms.GroupBox
        Me.DgvWorkunit = New System.Windows.Forms.DataGridView
        Me.colWU_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWU_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.txtJobDesc = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.txtDocumentNo = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.tsbExit = New System.Windows.Forms.ToolStripButton
        Me.TsbAuthorize = New System.Windows.Forms.ToolStripButton
        Me.tsbPrintAuthorize = New System.Windows.Forms.ToolStripButton
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.TboAuthorize.SuspendLayout()
        Me.TpAuthorize.SuspendLayout()
        CType(Me.DgvOtorisasi, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TpUnauthorize.SuspendLayout()
        CType(Me.DgvCancelOtorisasi, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TpAllauthorize.SuspendLayout()
        CType(Me.DgvAllautorisasi, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GboFilter.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboDetail.SuspendLayout()
        Me.gboItem.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.DgvWorkshift, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GboItem2.SuspendLayout()
        CType(Me.DgvWorkunit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer3)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.gboDetail)
        Me.SplitContainer1.Size = New System.Drawing.Size(904, 419)
        Me.SplitContainer1.SplitterDistance = 257
        Me.SplitContainer1.TabIndex = 0
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.TboAuthorize)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.GboFilter)
        Me.SplitContainer3.Size = New System.Drawing.Size(257, 419)
        Me.SplitContainer3.SplitterDistance = 184
        Me.SplitContainer3.SplitterWidth = 3
        Me.SplitContainer3.TabIndex = 1
        '
        'TboAuthorize
        '
        Me.TboAuthorize.Controls.Add(Me.TpAuthorize)
        Me.TboAuthorize.Controls.Add(Me.TpUnauthorize)
        Me.TboAuthorize.Controls.Add(Me.TpAllauthorize)
        Me.TboAuthorize.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TboAuthorize.Location = New System.Drawing.Point(0, 0)
        Me.TboAuthorize.Margin = New System.Windows.Forms.Padding(2)
        Me.TboAuthorize.Name = "TboAuthorize"
        Me.TboAuthorize.SelectedIndex = 0
        Me.TboAuthorize.Size = New System.Drawing.Size(257, 184)
        Me.TboAuthorize.TabIndex = 0
        '
        'TpAuthorize
        '
        Me.TpAuthorize.Controls.Add(Me.DgvOtorisasi)
        Me.TpAuthorize.Controls.Add(Me.LlbUncheckAll_Athorize)
        Me.TpAuthorize.Controls.Add(Me.LlbCheckAll_Authorize)
        Me.TpAuthorize.Controls.Add(Me.btnExit)
        Me.TpAuthorize.Location = New System.Drawing.Point(4, 22)
        Me.TpAuthorize.Margin = New System.Windows.Forms.Padding(2)
        Me.TpAuthorize.Name = "TpAuthorize"
        Me.TpAuthorize.Padding = New System.Windows.Forms.Padding(2)
        Me.TpAuthorize.Size = New System.Drawing.Size(249, 158)
        Me.TpAuthorize.TabIndex = 0
        Me.TpAuthorize.Text = "Authorize"
        Me.TpAuthorize.UseVisualStyleBackColor = True
        '
        'DgvOtorisasi
        '
        Me.DgvOtorisasi.AllowUserToAddRows = False
        Me.DgvOtorisasi.AllowUserToDeleteRows = False
        Me.DgvOtorisasi.AllowUserToResizeColumns = False
        Me.DgvOtorisasi.AllowUserToResizeRows = False
        Me.DgvOtorisasi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvOtorisasi.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvOtorisasi.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvOtorisasi.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDgvAuth_Primary, Me.ColDgvAuth_Authorize, Me.ColDgvAuth_SPLNo, Me.ColDgvAuth_Date})
        Me.DgvOtorisasi.Location = New System.Drawing.Point(2, 2)
        Me.DgvOtorisasi.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvOtorisasi.MultiSelect = False
        Me.DgvOtorisasi.Name = "DgvOtorisasi"
        Me.DgvOtorisasi.RowTemplate.Height = 24
        Me.DgvOtorisasi.Size = New System.Drawing.Size(247, 140)
        Me.DgvOtorisasi.TabIndex = 0
        '
        'ColDgvAuth_Primary
        '
        Me.ColDgvAuth_Primary.HeaderText = "Primary"
        Me.ColDgvAuth_Primary.Name = "ColDgvAuth_Primary"
        Me.ColDgvAuth_Primary.ReadOnly = True
        Me.ColDgvAuth_Primary.Visible = False
        Me.ColDgvAuth_Primary.Width = 66
        '
        'ColDgvAuth_Authorize
        '
        Me.ColDgvAuth_Authorize.HeaderText = "Authorize"
        Me.ColDgvAuth_Authorize.Name = "ColDgvAuth_Authorize"
        Me.ColDgvAuth_Authorize.Width = 57
        '
        'ColDgvAuth_SPLNo
        '
        Me.ColDgvAuth_SPLNo.HeaderText = "Doc No"
        Me.ColDgvAuth_SPLNo.MinimumWidth = 64
        Me.ColDgvAuth_SPLNo.Name = "ColDgvAuth_SPLNo"
        Me.ColDgvAuth_SPLNo.ReadOnly = True
        Me.ColDgvAuth_SPLNo.Width = 69
        '
        'ColDgvAuth_Date
        '
        Me.ColDgvAuth_Date.HeaderText = "Date"
        Me.ColDgvAuth_Date.Name = "ColDgvAuth_Date"
        Me.ColDgvAuth_Date.ReadOnly = True
        Me.ColDgvAuth_Date.Width = 55
        '
        'LlbUncheckAll_Athorize
        '
        Me.LlbUncheckAll_Athorize.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUncheckAll_Athorize.AutoSize = True
        Me.LlbUncheckAll_Athorize.Location = New System.Drawing.Point(184, 145)
        Me.LlbUncheckAll_Athorize.Name = "LlbUncheckAll_Athorize"
        Me.LlbUncheckAll_Athorize.Size = New System.Drawing.Size(65, 13)
        Me.LlbUncheckAll_Athorize.TabIndex = 1
        Me.LlbUncheckAll_Athorize.TabStop = True
        Me.LlbUncheckAll_Athorize.Text = "Uncheck All"
        '
        'LlbCheckAll_Authorize
        '
        Me.LlbCheckAll_Authorize.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbCheckAll_Authorize.AutoSize = True
        Me.LlbCheckAll_Authorize.Location = New System.Drawing.Point(129, 145)
        Me.LlbCheckAll_Authorize.Name = "LlbCheckAll_Authorize"
        Me.LlbCheckAll_Authorize.Size = New System.Drawing.Size(52, 13)
        Me.LlbCheckAll_Authorize.TabIndex = 0
        Me.LlbCheckAll_Authorize.TabStop = True
        Me.LlbCheckAll_Authorize.Text = "Check All"
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(39, 101)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "Button1"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'TpUnauthorize
        '
        Me.TpUnauthorize.Controls.Add(Me.DgvCancelOtorisasi)
        Me.TpUnauthorize.Controls.Add(Me.LlbUncheckAll_Cancel)
        Me.TpUnauthorize.Controls.Add(Me.LlbCheckAll_Cancel)
        Me.TpUnauthorize.Location = New System.Drawing.Point(4, 22)
        Me.TpUnauthorize.Margin = New System.Windows.Forms.Padding(2)
        Me.TpUnauthorize.Name = "TpUnauthorize"
        Me.TpUnauthorize.Padding = New System.Windows.Forms.Padding(2)
        Me.TpUnauthorize.Size = New System.Drawing.Size(249, 158)
        Me.TpUnauthorize.TabIndex = 1
        Me.TpUnauthorize.Text = "Cancel Authorize"
        Me.TpUnauthorize.UseVisualStyleBackColor = True
        '
        'DgvCancelOtorisasi
        '
        Me.DgvCancelOtorisasi.AllowUserToAddRows = False
        Me.DgvCancelOtorisasi.AllowUserToDeleteRows = False
        Me.DgvCancelOtorisasi.AllowUserToResizeColumns = False
        Me.DgvCancelOtorisasi.AllowUserToResizeRows = False
        Me.DgvCancelOtorisasi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvCancelOtorisasi.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCancelOtorisasi.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCancelOtorisasi.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDgvUnAuth_Primary, Me.ColDgvUnAuth_Cancel, Me.ColDgvUnAuth_SPLNo, Me.ColDgvUnAuth_Date})
        Me.DgvCancelOtorisasi.Location = New System.Drawing.Point(2, 2)
        Me.DgvCancelOtorisasi.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvCancelOtorisasi.MultiSelect = False
        Me.DgvCancelOtorisasi.Name = "DgvCancelOtorisasi"
        Me.DgvCancelOtorisasi.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvCancelOtorisasi.RowTemplate.Height = 24
        Me.DgvCancelOtorisasi.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvCancelOtorisasi.Size = New System.Drawing.Size(247, 140)
        Me.DgvCancelOtorisasi.TabIndex = 173
        '
        'ColDgvUnAuth_Primary
        '
        Me.ColDgvUnAuth_Primary.HeaderText = "Primary"
        Me.ColDgvUnAuth_Primary.Name = "ColDgvUnAuth_Primary"
        Me.ColDgvUnAuth_Primary.ReadOnly = True
        Me.ColDgvUnAuth_Primary.Visible = False
        Me.ColDgvUnAuth_Primary.Width = 66
        '
        'ColDgvUnAuth_Cancel
        '
        Me.ColDgvUnAuth_Cancel.HeaderText = "Cancel Authorize"
        Me.ColDgvUnAuth_Cancel.Name = "ColDgvUnAuth_Cancel"
        Me.ColDgvUnAuth_Cancel.Width = 93
        '
        'ColDgvUnAuth_SPLNo
        '
        Me.ColDgvUnAuth_SPLNo.HeaderText = "Doc No"
        Me.ColDgvUnAuth_SPLNo.Name = "ColDgvUnAuth_SPLNo"
        Me.ColDgvUnAuth_SPLNo.ReadOnly = True
        Me.ColDgvUnAuth_SPLNo.Width = 69
        '
        'ColDgvUnAuth_Date
        '
        Me.ColDgvUnAuth_Date.HeaderText = "Date"
        Me.ColDgvUnAuth_Date.Name = "ColDgvUnAuth_Date"
        Me.ColDgvUnAuth_Date.Width = 55
        '
        'LlbUncheckAll_Cancel
        '
        Me.LlbUncheckAll_Cancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUncheckAll_Cancel.AutoSize = True
        Me.LlbUncheckAll_Cancel.Location = New System.Drawing.Point(184, 145)
        Me.LlbUncheckAll_Cancel.Name = "LlbUncheckAll_Cancel"
        Me.LlbUncheckAll_Cancel.Size = New System.Drawing.Size(65, 13)
        Me.LlbUncheckAll_Cancel.TabIndex = 172
        Me.LlbUncheckAll_Cancel.TabStop = True
        Me.LlbUncheckAll_Cancel.Text = "Uncheck All"
        '
        'LlbCheckAll_Cancel
        '
        Me.LlbCheckAll_Cancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbCheckAll_Cancel.AutoSize = True
        Me.LlbCheckAll_Cancel.Location = New System.Drawing.Point(129, 145)
        Me.LlbCheckAll_Cancel.Name = "LlbCheckAll_Cancel"
        Me.LlbCheckAll_Cancel.Size = New System.Drawing.Size(52, 13)
        Me.LlbCheckAll_Cancel.TabIndex = 171
        Me.LlbCheckAll_Cancel.TabStop = True
        Me.LlbCheckAll_Cancel.Text = "Check All"
        '
        'TpAllauthorize
        '
        Me.TpAllauthorize.Controls.Add(Me.LlbUncheckAll_All)
        Me.TpAllauthorize.Controls.Add(Me.LlbCheckAll_ALL)
        Me.TpAllauthorize.Controls.Add(Me.DgvAllautorisasi)
        Me.TpAllauthorize.Location = New System.Drawing.Point(4, 22)
        Me.TpAllauthorize.Name = "TpAllauthorize"
        Me.TpAllauthorize.Padding = New System.Windows.Forms.Padding(3)
        Me.TpAllauthorize.Size = New System.Drawing.Size(249, 158)
        Me.TpAllauthorize.TabIndex = 2
        Me.TpAllauthorize.Text = "All Overtime Order"
        Me.TpAllauthorize.UseVisualStyleBackColor = True
        '
        'LlbUncheckAll_All
        '
        Me.LlbUncheckAll_All.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUncheckAll_All.AutoSize = True
        Me.LlbUncheckAll_All.Location = New System.Drawing.Point(184, 145)
        Me.LlbUncheckAll_All.Name = "LlbUncheckAll_All"
        Me.LlbUncheckAll_All.Size = New System.Drawing.Size(65, 13)
        Me.LlbUncheckAll_All.TabIndex = 176
        Me.LlbUncheckAll_All.TabStop = True
        Me.LlbUncheckAll_All.Text = "Uncheck All"
        '
        'LlbCheckAll_ALL
        '
        Me.LlbCheckAll_ALL.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbCheckAll_ALL.AutoSize = True
        Me.LlbCheckAll_ALL.Location = New System.Drawing.Point(129, 145)
        Me.LlbCheckAll_ALL.Name = "LlbCheckAll_ALL"
        Me.LlbCheckAll_ALL.Size = New System.Drawing.Size(52, 13)
        Me.LlbCheckAll_ALL.TabIndex = 175
        Me.LlbCheckAll_ALL.TabStop = True
        Me.LlbCheckAll_ALL.Text = "Check All"
        '
        'DgvAllautorisasi
        '
        Me.DgvAllautorisasi.AllowUserToAddRows = False
        Me.DgvAllautorisasi.AllowUserToDeleteRows = False
        Me.DgvAllautorisasi.AllowUserToResizeColumns = False
        Me.DgvAllautorisasi.AllowUserToResizeRows = False
        Me.DgvAllautorisasi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvAllautorisasi.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvAllautorisasi.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvAllautorisasi.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.col3_Primary, Me.col3_DocNo, Me.col3_Print, Me.col3_Date, Me.col3_Status})
        Me.DgvAllautorisasi.Location = New System.Drawing.Point(3, 3)
        Me.DgvAllautorisasi.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvAllautorisasi.MultiSelect = False
        Me.DgvAllautorisasi.Name = "DgvAllautorisasi"
        Me.DgvAllautorisasi.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvAllautorisasi.RowTemplate.Height = 24
        Me.DgvAllautorisasi.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvAllautorisasi.Size = New System.Drawing.Size(246, 140)
        Me.DgvAllautorisasi.TabIndex = 174
        '
        'col3_Primary
        '
        Me.col3_Primary.HeaderText = "Primary"
        Me.col3_Primary.Name = "col3_Primary"
        Me.col3_Primary.ReadOnly = True
        Me.col3_Primary.Visible = False
        Me.col3_Primary.Width = 66
        '
        'col3_DocNo
        '
        Me.col3_DocNo.HeaderText = "Doc No"
        Me.col3_DocNo.Name = "col3_DocNo"
        Me.col3_DocNo.ReadOnly = True
        Me.col3_DocNo.Width = 69
        '
        'col3_Print
        '
        Me.col3_Print.HeaderText = "Print"
        Me.col3_Print.Name = "col3_Print"
        Me.col3_Print.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.col3_Print.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.col3_Print.Width = 53
        '
        'col3_Date
        '
        Me.col3_Date.HeaderText = "Date"
        Me.col3_Date.Name = "col3_Date"
        Me.col3_Date.Width = 55
        '
        'col3_Status
        '
        Me.col3_Status.HeaderText = "Status"
        Me.col3_Status.Name = "col3_Status"
        Me.col3_Status.Width = 62
        '
        'GboFilter
        '
        Me.GboFilter.Controls.Add(Me.CboStatus)
        Me.GboFilter.Controls.Add(Me.Label8)
        Me.GboFilter.Controls.Add(Me.TxtNoSPL)
        Me.GboFilter.Controls.Add(Me.BtnFilter)
        Me.GboFilter.Controls.Add(Me.Label10)
        Me.GboFilter.Controls.Add(Me.BtnFindSPL)
        Me.GboFilter.Controls.Add(Me.GroupBox2)
        Me.GboFilter.Controls.Add(Me.DgvUser)
        Me.GboFilter.Controls.Add(Me.TxtPrimarySPL)
        Me.GboFilter.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GboFilter.Location = New System.Drawing.Point(0, 0)
        Me.GboFilter.Margin = New System.Windows.Forms.Padding(2)
        Me.GboFilter.Name = "GboFilter"
        Me.GboFilter.Padding = New System.Windows.Forms.Padding(2)
        Me.GboFilter.Size = New System.Drawing.Size(257, 232)
        Me.GboFilter.TabIndex = 0
        Me.GboFilter.TabStop = False
        Me.GboFilter.Text = "Filter"
        '
        'CboStatus
        '
        Me.CboStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboStatus.FormattingEnabled = True
        Me.CboStatus.Items.AddRange(New Object() {"All Overtime Order", "Already Authorized only", "Unauthorized only"})
        Me.CboStatus.Location = New System.Drawing.Point(60, 13)
        Me.CboStatus.Name = "CboStatus"
        Me.CboStatus.Size = New System.Drawing.Size(144, 21)
        Me.CboStatus.TabIndex = 0
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 17)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(37, 13)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Status"
        '
        'TxtNoSPL
        '
        Me.TxtNoSPL.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNoSPL.Location = New System.Drawing.Point(60, 40)
        Me.TxtNoSPL.Name = "TxtNoSPL"
        Me.TxtNoSPL.Size = New System.Drawing.Size(144, 20)
        Me.TxtNoSPL.TabIndex = 1
        '
        'BtnFilter
        '
        Me.BtnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFilter.Location = New System.Drawing.Point(210, 13)
        Me.BtnFilter.Name = "BtnFilter"
        Me.BtnFilter.Size = New System.Drawing.Size(43, 21)
        Me.BtnFilter.TabIndex = 4
        Me.BtnFilter.Text = "Filter"
        Me.BtnFilter.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(8, 44)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(44, 13)
        Me.Label10.TabIndex = 1
        Me.Label10.Text = "Doc No"
        '
        'BtnFindSPL
        '
        Me.BtnFindSPL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindSPL.Location = New System.Drawing.Point(210, 40)
        Me.BtnFindSPL.Name = "BtnFindSPL"
        Me.BtnFindSPL.Size = New System.Drawing.Size(43, 21)
        Me.BtnFindSPL.TabIndex = 2
        Me.BtnFindSPL.Text = "...."
        Me.BtnFindSPL.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.DtpTo)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.DtpFrom)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.ChkSPLDate)
        Me.GroupBox2.Location = New System.Drawing.Point(4, 66)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(249, 64)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        '
        'DtpTo
        '
        Me.DtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(56, 40)
        Me.DtpTo.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(188, 20)
        Me.DtpTo.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 43)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "To"
        '
        'DtpFrom
        '
        Me.DtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(56, 17)
        Me.DtpFrom.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(188, 20)
        Me.DtpFrom.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 20)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "From"
        '
        'ChkSPLDate
        '
        Me.ChkSPLDate.AutoSize = True
        Me.ChkSPLDate.Location = New System.Drawing.Point(5, 0)
        Me.ChkSPLDate.Name = "ChkSPLDate"
        Me.ChkSPLDate.Size = New System.Drawing.Size(49, 17)
        Me.ChkSPLDate.TabIndex = 0
        Me.ChkSPLDate.Text = "Date"
        Me.ChkSPLDate.UseVisualStyleBackColor = True
        '
        'DgvUser
        '
        Me.DgvUser.AllowUserToResizeColumns = False
        Me.DgvUser.AllowUserToResizeRows = False
        Me.DgvUser.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvUser.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCreatedBy_Primary, Me.ColCreatedBy, Me.ColCreatedBy_Name})
        Me.DgvUser.Location = New System.Drawing.Point(4, 137)
        Me.DgvUser.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvUser.Name = "DgvUser"
        Me.DgvUser.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvUser.RowTemplate.Height = 24
        Me.DgvUser.Size = New System.Drawing.Size(249, 89)
        Me.DgvUser.TabIndex = 3
        '
        'ColCreatedBy_Primary
        '
        Me.ColCreatedBy_Primary.HeaderText = "Primary"
        Me.ColCreatedBy_Primary.Name = "ColCreatedBy_Primary"
        Me.ColCreatedBy_Primary.ReadOnly = True
        Me.ColCreatedBy_Primary.Visible = False
        Me.ColCreatedBy_Primary.Width = 66
        '
        'ColCreatedBy
        '
        Me.ColCreatedBy.HeaderText = "User Code"
        Me.ColCreatedBy.Name = "ColCreatedBy"
        Me.ColCreatedBy.Width = 82
        '
        'ColCreatedBy_Name
        '
        Me.ColCreatedBy_Name.HeaderText = "Name"
        Me.ColCreatedBy_Name.Name = "ColCreatedBy_Name"
        Me.ColCreatedBy_Name.ReadOnly = True
        Me.ColCreatedBy_Name.Width = 60
        '
        'TxtPrimarySPL
        '
        Me.TxtPrimarySPL.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimarySPL.Location = New System.Drawing.Point(9, 26)
        Me.TxtPrimarySPL.MaxLength = 50
        Me.TxtPrimarySPL.Name = "TxtPrimarySPL"
        Me.TxtPrimarySPL.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimarySPL.TabIndex = 4
        Me.TxtPrimarySPL.TabStop = False
        Me.TxtPrimarySPL.Text = "-1"
        Me.TxtPrimarySPL.Visible = False
        '
        'gboDetail
        '
        Me.gboDetail.Controls.Add(Me.gboItem)
        Me.gboDetail.Controls.Add(Me.SplitContainer4)
        Me.gboDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboDetail.Location = New System.Drawing.Point(0, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(643, 419)
        Me.gboDetail.TabIndex = 4
        Me.gboDetail.TabStop = False
        '
        'gboItem
        '
        Me.gboItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItem.Controls.Add(Me.SplitContainer2)
        Me.gboItem.Location = New System.Drawing.Point(5, 81)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Size = New System.Drawing.Size(631, 332)
        Me.gboItem.TabIndex = 1
        Me.gboItem.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.DgvWorkshift)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DtpFromDetail)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DtpToDetail)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label5)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GboItem2)
        Me.SplitContainer2.Size = New System.Drawing.Size(625, 313)
        Me.SplitContainer2.SplitterDistance = 215
        Me.SplitContainer2.TabIndex = 19
        '
        'DgvWorkshift
        '
        Me.DgvWorkshift.AllowUserToAddRows = False
        Me.DgvWorkshift.AllowUserToDeleteRows = False
        Me.DgvWorkshift.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvWorkshift.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvWorkshift.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvWorkshift.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvWorkshift.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colNo, Me.colShiftCode, Me.colDesc, Me.colEarlyOT, Me.colEarlyOTDuration, Me.colEarlyDelayDuration, Me.colEalryBreak, Me.colDelayBeforeEarlyBreak, Me.colOT, Me.colOTDuration, Me.colDelayDuration, Me.colBreak, Me.colDelayBeforeBreak, Me.colRemarks, Me.colPrimaryWorkshift})
        Me.DgvWorkshift.Location = New System.Drawing.Point(4, 42)
        Me.DgvWorkshift.Name = "DgvWorkshift"
        Me.DgvWorkshift.ReadOnly = True
        Me.DgvWorkshift.Size = New System.Drawing.Size(617, 170)
        Me.DgvWorkshift.TabIndex = 2
        '
        'colNo
        '
        Me.colNo.HeaderText = "No"
        Me.colNo.Name = "colNo"
        Me.colNo.ReadOnly = True
        Me.colNo.Width = 46
        '
        'colShiftCode
        '
        Me.colShiftCode.HeaderText = "Shift"
        Me.colShiftCode.Name = "colShiftCode"
        Me.colShiftCode.ReadOnly = True
        Me.colShiftCode.Width = 53
        '
        'colDesc
        '
        Me.colDesc.HeaderText = "Desc"
        Me.colDesc.Name = "colDesc"
        Me.colDesc.ReadOnly = True
        Me.colDesc.Width = 57
        '
        'colEarlyOT
        '
        Me.colEarlyOT.HeaderText = "Early OT"
        Me.colEarlyOT.Name = "colEarlyOT"
        Me.colEarlyOT.ReadOnly = True
        Me.colEarlyOT.Width = 49
        '
        'colEarlyOTDuration
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colEarlyOTDuration.DefaultCellStyle = DataGridViewCellStyle1
        Me.colEarlyOTDuration.HeaderText = "Duration"
        Me.colEarlyOTDuration.Name = "colEarlyOTDuration"
        Me.colEarlyOTDuration.ReadOnly = True
        Me.colEarlyOTDuration.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colEarlyOTDuration.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.colEarlyOTDuration.Width = 53
        '
        'colEarlyDelayDuration
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colEarlyDelayDuration.DefaultCellStyle = DataGridViewCellStyle2
        Me.colEarlyDelayDuration.HeaderText = "Delay"
        Me.colEarlyDelayDuration.Name = "colEarlyDelayDuration"
        Me.colEarlyDelayDuration.ReadOnly = True
        Me.colEarlyDelayDuration.Width = 59
        '
        'colEalryBreak
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colEalryBreak.DefaultCellStyle = DataGridViewCellStyle3
        Me.colEalryBreak.HeaderText = "Break"
        Me.colEalryBreak.Name = "colEalryBreak"
        Me.colEalryBreak.ReadOnly = True
        Me.colEalryBreak.Width = 60
        '
        'colDelayBeforeEarlyBreak
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colDelayBeforeEarlyBreak.DefaultCellStyle = DataGridViewCellStyle4
        Me.colDelayBeforeEarlyBreak.HeaderText = "Delay Before Break"
        Me.colDelayBeforeEarlyBreak.Name = "colDelayBeforeEarlyBreak"
        Me.colDelayBeforeEarlyBreak.ReadOnly = True
        Me.colDelayBeforeEarlyBreak.Width = 114
        '
        'colOT
        '
        Me.colOT.HeaderText = "OT"
        Me.colOT.Name = "colOT"
        Me.colOT.ReadOnly = True
        Me.colOT.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colOT.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.colOT.Width = 47
        '
        'colOTDuration
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colOTDuration.DefaultCellStyle = DataGridViewCellStyle5
        Me.colOTDuration.HeaderText = "OT Duration"
        Me.colOTDuration.Name = "colOTDuration"
        Me.colOTDuration.ReadOnly = True
        Me.colOTDuration.Width = 83
        '
        'colDelayDuration
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colDelayDuration.DefaultCellStyle = DataGridViewCellStyle6
        Me.colDelayDuration.HeaderText = "Delay"
        Me.colDelayDuration.Name = "colDelayDuration"
        Me.colDelayDuration.ReadOnly = True
        Me.colDelayDuration.Width = 59
        '
        'colBreak
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colBreak.DefaultCellStyle = DataGridViewCellStyle7
        Me.colBreak.HeaderText = "Break"
        Me.colBreak.Name = "colBreak"
        Me.colBreak.ReadOnly = True
        Me.colBreak.Width = 60
        '
        'colDelayBeforeBreak
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colDelayBeforeBreak.DefaultCellStyle = DataGridViewCellStyle8
        Me.colDelayBeforeBreak.HeaderText = "Delay Before Break"
        Me.colDelayBeforeBreak.Name = "colDelayBeforeBreak"
        Me.colDelayBeforeBreak.ReadOnly = True
        Me.colDelayBeforeBreak.Width = 114
        '
        'colRemarks
        '
        Me.colRemarks.HeaderText = "Remarks"
        Me.colRemarks.Name = "colRemarks"
        Me.colRemarks.ReadOnly = True
        Me.colRemarks.Width = 74
        '
        'colPrimaryWorkshift
        '
        Me.colPrimaryWorkshift.HeaderText = "Primary Workshift"
        Me.colPrimaryWorkshift.Name = "colPrimaryWorkshift"
        Me.colPrimaryWorkshift.ReadOnly = True
        Me.colPrimaryWorkshift.Visible = False
        Me.colPrimaryWorkshift.Width = 105
        '
        'DtpFromDetail
        '
        Me.DtpFromDetail.CustomFormat = "dd MMM yyyy"
        Me.DtpFromDetail.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFromDetail.Location = New System.Drawing.Point(62, 3)
        Me.DtpFromDetail.Name = "DtpFromDetail"
        Me.DtpFromDetail.Size = New System.Drawing.Size(149, 20)
        Me.DtpFromDetail.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 7)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "From"
        '
        'DtpToDetail
        '
        Me.DtpToDetail.CustomFormat = "dd MMM yyyy"
        Me.DtpToDetail.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpToDetail.Location = New System.Drawing.Point(287, 3)
        Me.DtpToDetail.Name = "DtpToDetail"
        Me.DtpToDetail.Size = New System.Drawing.Size(149, 20)
        Me.DtpToDetail.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(238, 7)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(20, 13)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "To"
        '
        'GboItem2
        '
        Me.GboItem2.Controls.Add(Me.DgvWorkunit)
        Me.GboItem2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GboItem2.Location = New System.Drawing.Point(0, 0)
        Me.GboItem2.Name = "GboItem2"
        Me.GboItem2.Size = New System.Drawing.Size(625, 94)
        Me.GboItem2.TabIndex = 0
        Me.GboItem2.TabStop = False
        Me.GboItem2.Text = "Apply To"
        '
        'DgvWorkunit
        '
        Me.DgvWorkunit.AllowUserToAddRows = False
        Me.DgvWorkunit.AllowUserToDeleteRows = False
        Me.DgvWorkunit.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvWorkunit.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvWorkunit.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvWorkunit.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colWU_No, Me.colWU_Code, Me.colWU_Desc, Me.colWU_Primary})
        Me.DgvWorkunit.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvWorkunit.Location = New System.Drawing.Point(3, 16)
        Me.DgvWorkunit.Name = "DgvWorkunit"
        Me.DgvWorkunit.ReadOnly = True
        Me.DgvWorkunit.Size = New System.Drawing.Size(619, 75)
        Me.DgvWorkunit.TabIndex = 0
        '
        'colWU_No
        '
        Me.colWU_No.Frozen = True
        Me.colWU_No.HeaderText = "No"
        Me.colWU_No.Name = "colWU_No"
        Me.colWU_No.ReadOnly = True
        Me.colWU_No.Width = 46
        '
        'colWU_Code
        '
        Me.colWU_Code.Frozen = True
        Me.colWU_Code.HeaderText = "Work Unit"
        Me.colWU_Code.MaxInputLength = 250
        Me.colWU_Code.Name = "colWU_Code"
        Me.colWU_Code.ReadOnly = True
        Me.colWU_Code.Width = 80
        '
        'colWU_Desc
        '
        Me.colWU_Desc.HeaderText = "Desc"
        Me.colWU_Desc.Name = "colWU_Desc"
        Me.colWU_Desc.ReadOnly = True
        Me.colWU_Desc.Width = 57
        '
        'colWU_Primary
        '
        Me.colWU_Primary.HeaderText = "Primary"
        Me.colWU_Primary.Name = "colWU_Primary"
        Me.colWU_Primary.ReadOnly = True
        Me.colWU_Primary.Visible = False
        Me.colWU_Primary.Width = 66
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer4.Location = New System.Drawing.Point(6, 10)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.txtJobDesc)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer4.Panel1.Controls.Add(Me.txtPrimary)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer4.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer4.Panel1.Controls.Add(Me.txtDocumentNo)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label6)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.txtRemarks)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer4.Size = New System.Drawing.Size(631, 74)
        Me.SplitContainer4.SplitterDistance = 313
        Me.SplitContainer4.TabIndex = 0
        '
        'txtJobDesc
        '
        Me.txtJobDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtJobDesc.BackColor = System.Drawing.SystemColors.Control
        Me.txtJobDesc.Location = New System.Drawing.Point(99, 51)
        Me.txtJobDesc.MaxLength = 250
        Me.txtJobDesc.Name = "txtJobDesc"
        Me.txtJobDesc.ReadOnly = True
        Me.txtJobDesc.Size = New System.Drawing.Size(211, 20)
        Me.txtJobDesc.TabIndex = 22
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 55)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(44, 13)
        Me.Label9.TabIndex = 23
        Me.Label9.Text = "Reason"
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(64, 0)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(29, 20)
        Me.txtPrimary.TabIndex = 19
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Date"
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(99, 26)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(211, 20)
        Me.dtpDate.TabIndex = 1
        '
        'txtDocumentNo
        '
        Me.txtDocumentNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDocumentNo.Location = New System.Drawing.Point(99, 0)
        Me.txtDocumentNo.MaxLength = 250
        Me.txtDocumentNo.Name = "txtDocumentNo"
        Me.txtDocumentNo.ReadOnly = True
        Me.txtDocumentNo.Size = New System.Drawing.Size(211, 20)
        Me.txtDocumentNo.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 4)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(73, 13)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Document No"
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.Location = New System.Drawing.Point(58, 0)
        Me.txtRemarks.MaxLength = 250
        Me.txtRemarks.Multiline = True
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.ReadOnly = True
        Me.txtRemarks.Size = New System.Drawing.Size(249, 71)
        Me.txtRemarks.TabIndex = 0
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 4)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(49, 13)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "Remarks"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslModifiedBy, Me.tslCreatedDate, Me.tslModifiedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 456)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(904, 22)
        Me.StatusStrip2.TabIndex = 2
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(222, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(222, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(222, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(222, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'ToolStrip1
        '
        Me.ToolStrip1.AutoSize = False
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(30, 30)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbExit, Me.TsbAuthorize, Me.tsbPrintAuthorize})
        Me.ToolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 423)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip1.Size = New System.Drawing.Size(904, 33)
        Me.ToolStrip1.TabIndex = 3
        '
        'tsbExit
        '
        Me.tsbExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbExit.Name = "tsbExit"
        Me.tsbExit.Size = New System.Drawing.Size(29, 30)
        Me.tsbExit.Text = "E&xit"
        Me.tsbExit.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'TsbAuthorize
        '
        Me.TsbAuthorize.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TsbAuthorize.Name = "TsbAuthorize"
        Me.TsbAuthorize.Size = New System.Drawing.Size(62, 30)
        Me.TsbAuthorize.Text = "Authorize"
        Me.TsbAuthorize.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'tsbPrintAuthorize
        '
        Me.tsbPrintAuthorize.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbPrintAuthorize.Name = "tsbPrintAuthorize"
        Me.tsbPrintAuthorize.Size = New System.Drawing.Size(36, 30)
        Me.tsbPrintAuthorize.Text = "&Print"
        Me.tsbPrintAuthorize.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.tsbPrintAuthorize.Visible = False
        '
        'BackgroundWorker1
        '
        '
        'FrmTrans_SPL_Authorize
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(904, 478)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimumSize = New System.Drawing.Size(920, 516)
        Me.Name = "FrmTrans_SPL_Authorize"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Overtime Order Authorize"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.TboAuthorize.ResumeLayout(False)
        Me.TpAuthorize.ResumeLayout(False)
        Me.TpAuthorize.PerformLayout()
        CType(Me.DgvOtorisasi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TpUnauthorize.ResumeLayout(False)
        Me.TpUnauthorize.PerformLayout()
        CType(Me.DgvCancelOtorisasi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TpAllauthorize.ResumeLayout(False)
        Me.TpAllauthorize.PerformLayout()
        CType(Me.DgvAllautorisasi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GboFilter.ResumeLayout(False)
        Me.GboFilter.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboDetail.ResumeLayout(False)
        Me.gboItem.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.DgvWorkshift, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GboItem2.ResumeLayout(False)
        CType(Me.DgvWorkunit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.Panel2.PerformLayout()
        Me.SplitContainer4.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbExit As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbPrintAuthorize As System.Windows.Forms.ToolStripButton
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents TboAuthorize As System.Windows.Forms.TabControl
    Friend WithEvents TpAuthorize As System.Windows.Forms.TabPage
    Friend WithEvents DgvOtorisasi As System.Windows.Forms.DataGridView
    Friend WithEvents LlbUncheckAll_Athorize As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbCheckAll_Authorize As System.Windows.Forms.LinkLabel
    Friend WithEvents TpUnauthorize As System.Windows.Forms.TabPage
    Friend WithEvents DgvCancelOtorisasi As System.Windows.Forms.DataGridView
    Friend WithEvents ColDgvUnAuth_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvUnAuth_Cancel As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDgvUnAuth_SPLNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvUnAuth_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LlbUncheckAll_Cancel As System.Windows.Forms.LinkLabel
    Friend WithEvents GboFilter As System.Windows.Forms.GroupBox
    Friend WithEvents TxtNoSPL As System.Windows.Forms.TextBox
    Friend WithEvents BtnFilter As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents BtnFindSPL As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ChkSPLDate As System.Windows.Forms.CheckBox
    Friend WithEvents DgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents ColCreatedBy_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtPrimarySPL As System.Windows.Forms.TextBox
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents DtpFromDetail As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DtpToDetail As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GboItem2 As System.Windows.Forms.GroupBox
    Friend WithEvents DgvWorkunit As System.Windows.Forms.DataGridView
    Friend WithEvents colWU_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWU_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtDocumentNo As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents DgvWorkshift As System.Windows.Forms.DataGridView
    Friend WithEvents TpAllauthorize As System.Windows.Forms.TabPage
    Friend WithEvents DgvAllautorisasi As System.Windows.Forms.DataGridView
    Friend WithEvents CboStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TsbAuthorize As System.Windows.Forms.ToolStripButton
    Friend WithEvents col3_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col3_DocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col3_Print As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents col3_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col3_Status As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txtJobDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents colNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colShiftCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEarlyOT As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colEarlyOTDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEarlyDelayDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEalryBreak As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDelayBeforeEarlyBreak As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colOT As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colOTDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDelayDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colBreak As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDelayBeforeBreak As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPrimaryWorkshift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LlbCheckAll_Cancel As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbUncheckAll_All As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbCheckAll_ALL As System.Windows.Forms.LinkLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents ColDgvAuth_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvAuth_Authorize As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDgvAuth_SPLNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvAuth_Date As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
