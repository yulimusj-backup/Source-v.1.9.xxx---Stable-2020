<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanLedger
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
        Me.btnFindAcc = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtCodeAcc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAcc = New System.Windows.Forms.TextBox
        Me.btnShow = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.chkShowReverseAccount = New System.Windows.Forms.CheckBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.LblJob = New System.Windows.Forms.Label
        Me.LblDept = New System.Windows.Forms.Label
        Me.TxtPrimaryJob = New System.Windows.Forms.TextBox
        Me.btnFindJob = New System.Windows.Forms.Button
        Me.TxtJobCode = New System.Windows.Forms.TextBox
        Me.TxtPrimaryDept = New System.Windows.Forms.TextBox
        Me.btnFindDept = New System.Windows.Forms.Button
        Me.TxtDeptCode = New System.Windows.Forms.TextBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.clbAccount = New System.Windows.Forms.CheckedListBox
        Me.chkAccountFilter = New System.Windows.Forms.CheckBox
        Me.chkOriginal = New System.Windows.Forms.CheckBox
        Me.ChkShowHeaderDesc = New System.Windows.Forms.CheckBox
        Me.chkShowItemRemarks = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.chkPageAfter = New System.Windows.Forms.CheckBox
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvAccount = New System.Windows.Forms.DataGridView
        Me.ColPrimaryAcc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNoAcc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNameAcc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.DgvUser = New System.Windows.Forms.DataGridView
        Me.ColCreatedBy_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabJurnalType = New System.Windows.Forms.TabPage
        Me.llblunchekall = New System.Windows.Forms.LinkLabel
        Me.llbchekalltipe = New System.Windows.Forms.LinkLabel
        Me.clbTipeJournal = New System.Windows.Forms.CheckedListBox
        Me.BgwUser = New System.ComponentModel.BackgroundWorker
        Me.BgwAccount = New System.ComponentModel.BackgroundWorker
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvAccount, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabJurnalType.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnFindAcc
        '
        Me.btnFindAcc.Location = New System.Drawing.Point(217, 277)
        Me.btnFindAcc.Name = "btnFindAcc"
        Me.btnFindAcc.Size = New System.Drawing.Size(38, 23)
        Me.btnFindAcc.TabIndex = 1
        Me.btnFindAcc.TabStop = False
        Me.btnFindAcc.Text = "Find"
        Me.btnFindAcc.UseVisualStyleBackColor = True
        Me.btnFindAcc.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 279)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 13)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Account Number"
        Me.Label1.Visible = False
        '
        'TxtDesc
        '
        Me.TxtDesc.Location = New System.Drawing.Point(97, 302)
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.ReadOnly = True
        Me.TxtDesc.Size = New System.Drawing.Size(158, 20)
        Me.TxtDesc.TabIndex = 2
        Me.TxtDesc.TabStop = False
        Me.TxtDesc.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 305)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Description"
        Me.Label2.Visible = False
        '
        'TxtCodeAcc
        '
        Me.TxtCodeAcc.Location = New System.Drawing.Point(97, 276)
        Me.TxtCodeAcc.Name = "TxtCodeAcc"
        Me.TxtCodeAcc.Size = New System.Drawing.Size(114, 20)
        Me.TxtCodeAcc.TabIndex = 0
        Me.TxtCodeAcc.Visible = False
        '
        'TxtPrimaryAcc
        '
        Me.TxtPrimaryAcc.Location = New System.Drawing.Point(48, 276)
        Me.TxtPrimaryAcc.Name = "TxtPrimaryAcc"
        Me.TxtPrimaryAcc.Size = New System.Drawing.Size(43, 20)
        Me.TxtPrimaryAcc.TabIndex = 12
        Me.TxtPrimaryAcc.TabStop = False
        Me.TxtPrimaryAcc.Text = "-1"
        Me.TxtPrimaryAcc.Visible = False
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(195, 16)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(66, 23)
        Me.btnShow.TabIndex = 7
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(195, 42)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(66, 23)
        Me.btnClose.TabIndex = 8
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.EnableDrillDown = False
        Me.crv1.Location = New System.Drawing.Point(285, 12)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(592, 542)
        Me.crv1.TabIndex = 0
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'chkShowReverseAccount
        '
        Me.chkShowReverseAccount.AutoSize = True
        Me.chkShowReverseAccount.Location = New System.Drawing.Point(197, 103)
        Me.chkShowReverseAccount.Name = "chkShowReverseAccount"
        Me.chkShowReverseAccount.Size = New System.Drawing.Size(109, 17)
        Me.chkShowReverseAccount.TabIndex = 8
        Me.chkShowReverseAccount.TabStop = False
        Me.chkShowReverseAccount.Text = "Reverse Account"
        Me.chkShowReverseAccount.UseVisualStyleBackColor = True
        Me.chkShowReverseAccount.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 47)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "To"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 22)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "From"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(72, 19)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(117, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(72, 45)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(117, 20)
        Me.dtpTo.TabIndex = 1
        '
        'LblJob
        '
        Me.LblJob.AutoSize = True
        Me.LblJob.Location = New System.Drawing.Point(7, 250)
        Me.LblJob.Name = "LblJob"
        Me.LblJob.Size = New System.Drawing.Size(24, 13)
        Me.LblJob.TabIndex = 66
        Me.LblJob.Text = "Job"
        '
        'LblDept
        '
        Me.LblDept.AutoSize = True
        Me.LblDept.Location = New System.Drawing.Point(7, 224)
        Me.LblDept.Name = "LblDept"
        Me.LblDept.Size = New System.Drawing.Size(68, 13)
        Me.LblDept.TabIndex = 64
        Me.LblDept.Text = "Departement"
        '
        'TxtPrimaryJob
        '
        Me.TxtPrimaryJob.Location = New System.Drawing.Point(34, 235)
        Me.TxtPrimaryJob.Name = "TxtPrimaryJob"
        Me.TxtPrimaryJob.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryJob.TabIndex = 67
        Me.TxtPrimaryJob.TabStop = False
        Me.TxtPrimaryJob.Text = "-1"
        Me.TxtPrimaryJob.Visible = False
        '
        'btnFindJob
        '
        Me.btnFindJob.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindJob.Location = New System.Drawing.Point(192, 250)
        Me.btnFindJob.Name = "btnFindJob"
        Me.btnFindJob.Size = New System.Drawing.Size(50, 21)
        Me.btnFindJob.TabIndex = 12
        Me.btnFindJob.TabStop = False
        Me.btnFindJob.Text = "Find"
        Me.btnFindJob.UseVisualStyleBackColor = True
        '
        'TxtJobCode
        '
        Me.TxtJobCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtJobCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtJobCode.Location = New System.Drawing.Point(81, 243)
        Me.TxtJobCode.Name = "TxtJobCode"
        Me.TxtJobCode.Size = New System.Drawing.Size(105, 20)
        Me.TxtJobCode.TabIndex = 11
        Me.TxtJobCode.TabStop = False
        '
        'TxtPrimaryDept
        '
        Me.TxtPrimaryDept.Location = New System.Drawing.Point(34, 212)
        Me.TxtPrimaryDept.Name = "TxtPrimaryDept"
        Me.TxtPrimaryDept.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryDept.TabIndex = 65
        Me.TxtPrimaryDept.TabStop = False
        Me.TxtPrimaryDept.Text = "-1"
        Me.TxtPrimaryDept.Visible = False
        '
        'btnFindDept
        '
        Me.btnFindDept.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindDept.Location = New System.Drawing.Point(192, 215)
        Me.btnFindDept.Name = "btnFindDept"
        Me.btnFindDept.Size = New System.Drawing.Size(50, 21)
        Me.btnFindDept.TabIndex = 10
        Me.btnFindDept.TabStop = False
        Me.btnFindDept.Text = "Find"
        Me.btnFindDept.UseVisualStyleBackColor = True
        '
        'TxtDeptCode
        '
        Me.TxtDeptCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDeptCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDeptCode.Location = New System.Drawing.Point(81, 217)
        Me.TxtDeptCode.Name = "TxtDeptCode"
        Me.TxtDeptCode.Size = New System.Drawing.Size(105, 20)
        Me.TxtDeptCode.TabIndex = 9
        Me.TxtDeptCode.TabStop = False
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"Potrait", "Landscape"})
        Me.ComboBox1.Location = New System.Drawing.Point(72, 76)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(117, 21)
        Me.ComboBox1.TabIndex = 2
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.llbUncheckAll)
        Me.GroupBox1.Controls.Add(Me.llbCheckAll)
        Me.GroupBox1.Controls.Add(Me.clbAccount)
        Me.GroupBox1.Controls.Add(Me.chkAccountFilter)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 345)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(254, 184)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.Location = New System.Drawing.Point(183, 0)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAll.TabIndex = 2
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        '
        'llbCheckAll
        '
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(135, 1)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAll.TabIndex = 1
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        '
        'clbAccount
        '
        Me.clbAccount.CheckOnClick = True
        Me.clbAccount.ColumnWidth = 250
        Me.clbAccount.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbAccount.FormattingEnabled = True
        Me.clbAccount.Location = New System.Drawing.Point(3, 16)
        Me.clbAccount.MultiColumn = True
        Me.clbAccount.Name = "clbAccount"
        Me.clbAccount.Size = New System.Drawing.Size(248, 154)
        Me.clbAccount.TabIndex = 3
        '
        'chkAccountFilter
        '
        Me.chkAccountFilter.AutoSize = True
        Me.chkAccountFilter.Location = New System.Drawing.Point(8, 0)
        Me.chkAccountFilter.Name = "chkAccountFilter"
        Me.chkAccountFilter.Size = New System.Drawing.Size(91, 17)
        Me.chkAccountFilter.TabIndex = 0
        Me.chkAccountFilter.Text = "Account Filter"
        Me.chkAccountFilter.UseVisualStyleBackColor = True
        '
        'chkOriginal
        '
        Me.chkOriginal.AutoSize = True
        Me.chkOriginal.Location = New System.Drawing.Point(72, 103)
        Me.chkOriginal.Name = "chkOriginal"
        Me.chkOriginal.Size = New System.Drawing.Size(112, 17)
        Me.chkOriginal.TabIndex = 3
        Me.chkOriginal.Text = "In Original Amount"
        Me.chkOriginal.UseVisualStyleBackColor = True
        '
        'ChkShowHeaderDesc
        '
        Me.ChkShowHeaderDesc.AutoSize = True
        Me.ChkShowHeaderDesc.Checked = True
        Me.ChkShowHeaderDesc.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkShowHeaderDesc.Location = New System.Drawing.Point(72, 123)
        Me.ChkShowHeaderDesc.Name = "ChkShowHeaderDesc"
        Me.ChkShowHeaderDesc.Size = New System.Drawing.Size(147, 17)
        Me.ChkShowHeaderDesc.TabIndex = 4
        Me.ChkShowHeaderDesc.Text = "Show Header Description"
        Me.ChkShowHeaderDesc.UseVisualStyleBackColor = True
        '
        'chkShowItemRemarks
        '
        Me.chkShowItemRemarks.AutoSize = True
        Me.chkShowItemRemarks.Location = New System.Drawing.Point(72, 170)
        Me.chkShowItemRemarks.Name = "chkShowItemRemarks"
        Me.chkShowItemRemarks.Size = New System.Drawing.Size(121, 17)
        Me.chkShowItemRemarks.TabIndex = 6
        Me.chkShowItemRemarks.TabStop = False
        Me.chkShowItemRemarks.Text = "Show Item Remarks"
        Me.chkShowItemRemarks.UseVisualStyleBackColor = True
        Me.chkShowItemRemarks.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.chkPageAfter)
        Me.GroupBox2.Controls.Add(Me.TabControl2)
        Me.GroupBox2.Controls.Add(Me.btnClose)
        Me.GroupBox2.Controls.Add(Me.LblJob)
        Me.GroupBox2.Controls.Add(Me.LblDept)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.TxtPrimaryJob)
        Me.GroupBox2.Controls.Add(Me.btnFindJob)
        Me.GroupBox2.Controls.Add(Me.TxtJobCode)
        Me.GroupBox2.Controls.Add(Me.ChkShowHeaderDesc)
        Me.GroupBox2.Controls.Add(Me.TxtPrimaryDept)
        Me.GroupBox2.Controls.Add(Me.btnFindDept)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.TxtDeptCode)
        Me.GroupBox2.Controls.Add(Me.dtpFrom)
        Me.GroupBox2.Controls.Add(Me.TxtPrimaryAcc)
        Me.GroupBox2.Controls.Add(Me.chkOriginal)
        Me.GroupBox2.Controls.Add(Me.TxtDesc)
        Me.GroupBox2.Controls.Add(Me.dtpTo)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.btnShow)
        Me.GroupBox2.Controls.Add(Me.TxtCodeAcc)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.ComboBox1)
        Me.GroupBox2.Controls.Add(Me.btnFindAcc)
        Me.GroupBox2.Controls.Add(Me.chkShowItemRemarks)
        Me.GroupBox2.Controls.Add(Me.chkShowReverseAccount)
        Me.GroupBox2.Controls.Add(Me.GroupBox1)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(267, 542)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Filter By"
        '
        'chkPageAfter
        '
        Me.chkPageAfter.AutoSize = True
        Me.chkPageAfter.Checked = True
        Me.chkPageAfter.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPageAfter.Location = New System.Drawing.Point(72, 147)
        Me.chkPageAfter.Name = "chkPageAfter"
        Me.chkPageAfter.Size = New System.Drawing.Size(162, 17)
        Me.chkPageAfter.TabIndex = 5
        Me.chkPageAfter.Text = "New Page for Each Account"
        Me.chkPageAfter.UseVisualStyleBackColor = True
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage2)
        Me.TabControl2.Controls.Add(Me.TabPage3)
        Me.TabControl2.Controls.Add(Me.TabJurnalType)
        Me.TabControl2.Location = New System.Drawing.Point(7, 192)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(259, 340)
        Me.TabControl2.TabIndex = 7
        Me.TabControl2.TabStop = False
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvAccount)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Size = New System.Drawing.Size(251, 314)
        Me.TabPage2.TabIndex = 0
        Me.TabPage2.Text = "Account"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvAccount
        '
        Me.dgvAccount.AllowUserToResizeColumns = False
        Me.dgvAccount.AllowUserToResizeRows = False
        Me.dgvAccount.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAccount.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAccount.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvAccount.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColPrimaryAcc, Me.ColNoAcc, Me.ColNameAcc})
        Me.dgvAccount.Location = New System.Drawing.Point(3, 4)
        Me.dgvAccount.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvAccount.Name = "dgvAccount"
        Me.dgvAccount.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvAccount.RowTemplate.Height = 24
        Me.dgvAccount.Size = New System.Drawing.Size(245, 308)
        Me.dgvAccount.TabIndex = 0
        Me.dgvAccount.TabStop = False
        '
        'ColPrimaryAcc
        '
        Me.ColPrimaryAcc.HeaderText = "Primary Account"
        Me.ColPrimaryAcc.Name = "ColPrimaryAcc"
        Me.ColPrimaryAcc.ReadOnly = True
        Me.ColPrimaryAcc.Visible = False
        Me.ColPrimaryAcc.Width = 109
        '
        'ColNoAcc
        '
        Me.ColNoAcc.HeaderText = "Account Code"
        Me.ColNoAcc.Name = "ColNoAcc"
        '
        'ColNameAcc
        '
        Me.ColNameAcc.HeaderText = "Account Name"
        Me.ColNameAcc.Name = "ColNameAcc"
        Me.ColNameAcc.ReadOnly = True
        Me.ColNameAcc.Width = 103
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.DgvUser)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage3.Size = New System.Drawing.Size(251, 315)
        Me.TabPage3.TabIndex = 1
        Me.TabPage3.Text = "Created By"
        Me.TabPage3.UseVisualStyleBackColor = True
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
        Me.DgvUser.Location = New System.Drawing.Point(2, 4)
        Me.DgvUser.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DgvUser.Name = "DgvUser"
        Me.DgvUser.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvUser.RowTemplate.Height = 24
        Me.DgvUser.Size = New System.Drawing.Size(248, 334)
        Me.DgvUser.TabIndex = 1
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
        'TabJurnalType
        '
        Me.TabJurnalType.Controls.Add(Me.llblunchekall)
        Me.TabJurnalType.Controls.Add(Me.llbchekalltipe)
        Me.TabJurnalType.Controls.Add(Me.clbTipeJournal)
        Me.TabJurnalType.Location = New System.Drawing.Point(4, 22)
        Me.TabJurnalType.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabJurnalType.Name = "TabJurnalType"
        Me.TabJurnalType.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabJurnalType.Size = New System.Drawing.Size(251, 315)
        Me.TabJurnalType.TabIndex = 2
        Me.TabJurnalType.Text = "Journal Type"
        Me.TabJurnalType.UseVisualStyleBackColor = True
        '
        'llblunchekall
        '
        Me.llblunchekall.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llblunchekall.AutoSize = True
        Me.llblunchekall.BackColor = System.Drawing.SystemColors.Window
        Me.llblunchekall.Location = New System.Drawing.Point(188, 300)
        Me.llblunchekall.Name = "llblunchekall"
        Me.llblunchekall.Size = New System.Drawing.Size(65, 13)
        Me.llblunchekall.TabIndex = 1
        Me.llblunchekall.TabStop = True
        Me.llblunchekall.Text = "Uncheck All"
        '
        'llbchekalltipe
        '
        Me.llbchekalltipe.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbchekalltipe.AutoSize = True
        Me.llbchekalltipe.BackColor = System.Drawing.SystemColors.Window
        Me.llbchekalltipe.Location = New System.Drawing.Point(133, 300)
        Me.llbchekalltipe.Name = "llbchekalltipe"
        Me.llbchekalltipe.Size = New System.Drawing.Size(52, 13)
        Me.llbchekalltipe.TabIndex = 0
        Me.llbchekalltipe.TabStop = True
        Me.llbchekalltipe.Text = "Check All"
        '
        'clbTipeJournal
        '
        Me.clbTipeJournal.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.clbTipeJournal.CheckOnClick = True
        Me.clbTipeJournal.ColumnWidth = 250
        Me.clbTipeJournal.FormattingEnabled = True
        Me.clbTipeJournal.Location = New System.Drawing.Point(2, 2)
        Me.clbTipeJournal.MultiColumn = True
        Me.clbTipeJournal.Name = "clbTipeJournal"
        Me.clbTipeJournal.Size = New System.Drawing.Size(249, 304)
        Me.clbTipeJournal.TabIndex = 15
        '
        'BgwUser
        '
        '
        'BgwAccount
        '
        '
        'FrmLaporanLedger
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(890, 559)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.crv1)
        Me.MinimumSize = New System.Drawing.Size(906, 597)
        Me.Name = "FrmLaporanLedger"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ledger"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvAccount, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabJurnalType.ResumeLayout(False)
        Me.TabJurnalType.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnFindAcc As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtCodeAcc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAcc As System.Windows.Forms.TextBox
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents chkShowReverseAccount As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblJob As System.Windows.Forms.Label
    Friend WithEvents LblDept As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryJob As System.Windows.Forms.TextBox
    Friend WithEvents btnFindJob As System.Windows.Forms.Button
    Friend WithEvents TxtJobCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryDept As System.Windows.Forms.TextBox
    Friend WithEvents btnFindDept As System.Windows.Forms.Button
    Friend WithEvents TxtDeptCode As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents clbAccount As System.Windows.Forms.CheckedListBox
    Friend WithEvents chkAccountFilter As System.Windows.Forms.CheckBox
    Friend WithEvents chkOriginal As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowHeaderDesc As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowItemRemarks As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents DgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents ColCreatedBy_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BgwUser As System.ComponentModel.BackgroundWorker
    Friend WithEvents dgvAccount As System.Windows.Forms.DataGridView
    Friend WithEvents BgwAccount As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColPrimaryAcc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNoAcc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNameAcc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkPageAfter As System.Windows.Forms.CheckBox
    Friend WithEvents TabJurnalType As System.Windows.Forms.TabPage
    Friend WithEvents clbTipeJournal As System.Windows.Forms.CheckedListBox
    Friend WithEvents llblunchekall As System.Windows.Forms.LinkLabel
    Friend WithEvents llbchekalltipe As System.Windows.Forms.LinkLabel
End Class
