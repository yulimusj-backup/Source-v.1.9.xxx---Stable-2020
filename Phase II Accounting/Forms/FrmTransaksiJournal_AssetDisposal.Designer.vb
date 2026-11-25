<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiJournal_AssetDisposal
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
        Me.btnViewJurnal = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TxtDisposed_Accum_Dep_Amount = New System.Windows.Forms.TextBox
        Me.TxtDisposed_Asset_Amount = New System.Windows.Forms.TextBox
        Me.TxtUnbalance = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtAssetDesc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAsset = New System.Windows.Forms.TextBox
        Me.BtnFindAsset = New System.Windows.Forms.Button
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtAssetCode = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtPrimaryJournal = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.chkSell = New System.Windows.Forms.CheckBox
        Me.TxtRefNo = New System.Windows.Forms.TextBox
        Me.cboDisposalType = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtDisposeAccountDesc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryDisposeAccount = New System.Windows.Forms.TextBox
        Me.btnFindDisposeAccount = New System.Windows.Forms.Button
        Me.TxtDisposeAccountCode = New System.Windows.Forms.TextBox
        Me.TxtIncomeAmount = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtIncomeAccountDesc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryIncomeAccount = New System.Windows.Forms.TextBox
        Me.BtnFindIncomeAccount = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtIncomeAccount_Code = New System.Windows.Forms.TextBox
        Me.LblUnit = New System.Windows.Forms.Label
        Me.TxtDisposeQty = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDebet = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKredit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExchangeRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOriginalCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDebet_OriginalCurr = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCredit_OriginalCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJob = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryJob = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnViewJurnal
        '
        Me.btnViewJurnal.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnViewJurnal.Location = New System.Drawing.Point(12, 346)
        Me.btnViewJurnal.Name = "btnViewJurnal"
        Me.btnViewJurnal.Size = New System.Drawing.Size(80, 26)
        Me.btnViewJurnal.TabIndex = 7
        Me.btnViewJurnal.TabStop = False
        Me.btnViewJurnal.Text = "View Journal"
        Me.btnViewJurnal.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(622, 346)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(446, 346)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(358, 346)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(182, 346)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(270, 346)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(534, 346)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 385)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(884, 22)
        Me.StatusStrip1.TabIndex = 57
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TxtDisposed_Accum_Dep_Amount)
        Me.gboDetail.Controls.Add(Me.TxtDisposed_Asset_Amount)
        Me.gboDetail.Controls.Add(Me.TxtUnbalance)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(12, 2)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(860, 338)
        Me.gboDetail.TabIndex = 6
        Me.gboDetail.TabStop = False
        '
        'TxtDisposed_Accum_Dep_Amount
        '
        Me.TxtDisposed_Accum_Dep_Amount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDisposed_Accum_Dep_Amount.Location = New System.Drawing.Point(336, 202)
        Me.TxtDisposed_Accum_Dep_Amount.MaxLength = 150
        Me.TxtDisposed_Accum_Dep_Amount.Name = "TxtDisposed_Accum_Dep_Amount"
        Me.TxtDisposed_Accum_Dep_Amount.Size = New System.Drawing.Size(131, 20)
        Me.TxtDisposed_Accum_Dep_Amount.TabIndex = 186
        Me.TxtDisposed_Accum_Dep_Amount.TabStop = False
        Me.TxtDisposed_Accum_Dep_Amount.Text = "0"
        Me.TxtDisposed_Accum_Dep_Amount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtDisposed_Asset_Amount
        '
        Me.TxtDisposed_Asset_Amount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDisposed_Asset_Amount.Location = New System.Drawing.Point(336, 176)
        Me.TxtDisposed_Asset_Amount.MaxLength = 150
        Me.TxtDisposed_Asset_Amount.Name = "TxtDisposed_Asset_Amount"
        Me.TxtDisposed_Asset_Amount.Size = New System.Drawing.Size(131, 20)
        Me.TxtDisposed_Asset_Amount.TabIndex = 185
        Me.TxtDisposed_Asset_Amount.TabStop = False
        Me.TxtDisposed_Asset_Amount.Text = "0"
        Me.TxtDisposed_Asset_Amount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtUnbalance
        '
        Me.TxtUnbalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUnbalance.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtUnbalance.Location = New System.Drawing.Point(202, 267)
        Me.TxtUnbalance.MaxLength = 50
        Me.TxtUnbalance.Name = "TxtUnbalance"
        Me.TxtUnbalance.ReadOnly = True
        Me.TxtUnbalance.Size = New System.Drawing.Size(312, 20)
        Me.TxtUnbalance.TabIndex = 1
        Me.TxtUnbalance.TabStop = False
        Me.TxtUnbalance.Text = "0"
        Me.TxtUnbalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtUnbalance.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(101, 272)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 164
        Me.Label2.Text = "Unbalance"
        Me.Label2.Visible = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 19)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtAssetDesc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryAsset)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindAsset)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtReadOnly)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtAssetCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryJournal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkSell)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRefNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboDisposalType)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDisposeAccountDesc)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryDisposeAccount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindDisposeAccount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDisposeAccountCode)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtIncomeAmount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtIncomeAccountDesc)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryIncomeAccount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnFindIncomeAccount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtIncomeAccount_Code)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblUnit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDisposeQty)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Size = New System.Drawing.Size(848, 105)
        Me.SplitContainer1.SplitterDistance = 423
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'TxtAssetDesc
        '
        Me.TxtAssetDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAssetDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAssetDesc.Location = New System.Drawing.Point(218, 55)
        Me.TxtAssetDesc.MaxLength = 50
        Me.TxtAssetDesc.Name = "TxtAssetDesc"
        Me.TxtAssetDesc.ReadOnly = True
        Me.TxtAssetDesc.Size = New System.Drawing.Size(202, 20)
        Me.TxtAssetDesc.TabIndex = 5
        Me.TxtAssetDesc.TabStop = False
        '
        'TxtPrimaryAsset
        '
        Me.TxtPrimaryAsset.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryAsset.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryAsset.Location = New System.Drawing.Point(24, 56)
        Me.TxtPrimaryAsset.MaxLength = 50
        Me.TxtPrimaryAsset.Name = "TxtPrimaryAsset"
        Me.TxtPrimaryAsset.Size = New System.Drawing.Size(39, 20)
        Me.TxtPrimaryAsset.TabIndex = 172
        Me.TxtPrimaryAsset.TabStop = False
        Me.TxtPrimaryAsset.Text = "-1"
        Me.TxtPrimaryAsset.Visible = False
        '
        'BtnFindAsset
        '
        Me.BtnFindAsset.Location = New System.Drawing.Point(168, 55)
        Me.BtnFindAsset.Name = "BtnFindAsset"
        Me.BtnFindAsset.Size = New System.Drawing.Size(44, 20)
        Me.BtnFindAsset.TabIndex = 4
        Me.BtnFindAsset.TabStop = False
        Me.BtnFindAsset.Text = "Find"
        Me.BtnFindAsset.UseVisualStyleBackColor = True
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(82, 29)
        Me.TxtRemarks.MaxLength = 150
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(338, 20)
        Me.TxtRemarks.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 61)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(61, 13)
        Me.Label5.TabIndex = 171
        Me.Label5.Text = "Asset Code"
        '
        'txtReadOnly
        '
        Me.txtReadOnly.Location = New System.Drawing.Point(82, 3)
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(23, 20)
        Me.txtReadOnly.TabIndex = 211
        Me.txtReadOnly.TabStop = False
        Me.txtReadOnly.Visible = False
        '
        'TxtAssetCode
        '
        Me.TxtAssetCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAssetCode.Location = New System.Drawing.Point(82, 55)
        Me.TxtAssetCode.MaxLength = 50
        Me.TxtAssetCode.Name = "TxtAssetCode"
        Me.TxtAssetCode.Size = New System.Drawing.Size(80, 20)
        Me.TxtAssetCode.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 13)
        Me.Label6.TabIndex = 143
        Me.Label6.Text = "Description"
        '
        'TxtPrimaryJournal
        '
        Me.TxtPrimaryJournal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryJournal.Location = New System.Drawing.Point(29, 3)
        Me.TxtPrimaryJournal.Name = "TxtPrimaryJournal"
        Me.TxtPrimaryJournal.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimaryJournal.TabIndex = 163
        Me.TxtPrimaryJournal.TabStop = False
        Me.TxtPrimaryJournal.Text = "-1"
        Me.TxtPrimaryJournal.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 7)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Document No"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(258, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Date"
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(294, 3)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(126, 20)
        Me.dtpDate.TabIndex = 1
        '
        'chkSell
        '
        Me.chkSell.AutoSize = True
        Me.chkSell.Location = New System.Drawing.Point(218, 83)
        Me.chkSell.Name = "chkSell"
        Me.chkSell.Size = New System.Drawing.Size(95, 17)
        Me.chkSell.TabIndex = 7
        Me.chkSell.Text = "Sell This Asset"
        Me.chkSell.UseVisualStyleBackColor = True
        '
        'TxtRefNo
        '
        Me.TxtRefNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRefNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRefNo.Location = New System.Drawing.Point(82, 3)
        Me.TxtRefNo.MaxLength = 50
        Me.TxtRefNo.Name = "TxtRefNo"
        Me.TxtRefNo.ReadOnly = True
        Me.TxtRefNo.Size = New System.Drawing.Size(170, 20)
        Me.TxtRefNo.TabIndex = 0
        Me.TxtRefNo.TabStop = False
        '
        'cboDisposalType
        '
        Me.cboDisposalType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDisposalType.FormattingEnabled = True
        Me.cboDisposalType.Items.AddRange(New Object() {"Full Disposal", "Partial Disposal"})
        Me.cboDisposalType.Location = New System.Drawing.Point(82, 81)
        Me.cboDisposalType.Name = "cboDisposalType"
        Me.cboDisposalType.Size = New System.Drawing.Size(130, 21)
        Me.cboDisposalType.TabIndex = 6
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(1, 10)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(90, 13)
        Me.Label10.TabIndex = 183
        Me.Label10.Text = "Disposal Account"
        '
        'TxtDisposeAccountDesc
        '
        Me.TxtDisposeAccountDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDisposeAccountDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDisposeAccountDesc.Location = New System.Drawing.Point(226, 4)
        Me.TxtDisposeAccountDesc.MaxLength = 50
        Me.TxtDisposeAccountDesc.Name = "TxtDisposeAccountDesc"
        Me.TxtDisposeAccountDesc.ReadOnly = True
        Me.TxtDisposeAccountDesc.Size = New System.Drawing.Size(192, 20)
        Me.TxtDisposeAccountDesc.TabIndex = 1
        Me.TxtDisposeAccountDesc.TabStop = False
        '
        'TxtPrimaryDisposeAccount
        '
        Me.TxtPrimaryDisposeAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryDisposeAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryDisposeAccount.Location = New System.Drawing.Point(19, 7)
        Me.TxtPrimaryDisposeAccount.MaxLength = 50
        Me.TxtPrimaryDisposeAccount.Name = "TxtPrimaryDisposeAccount"
        Me.TxtPrimaryDisposeAccount.Size = New System.Drawing.Size(39, 20)
        Me.TxtPrimaryDisposeAccount.TabIndex = 184
        Me.TxtPrimaryDisposeAccount.TabStop = False
        Me.TxtPrimaryDisposeAccount.Text = "-1"
        Me.TxtPrimaryDisposeAccount.Visible = False
        '
        'btnFindDisposeAccount
        '
        Me.btnFindDisposeAccount.Location = New System.Drawing.Point(176, 4)
        Me.btnFindDisposeAccount.Name = "btnFindDisposeAccount"
        Me.btnFindDisposeAccount.Size = New System.Drawing.Size(44, 20)
        Me.btnFindDisposeAccount.TabIndex = 182
        Me.btnFindDisposeAccount.TabStop = False
        Me.btnFindDisposeAccount.Text = "Find"
        Me.btnFindDisposeAccount.UseVisualStyleBackColor = True
        '
        'TxtDisposeAccountCode
        '
        Me.TxtDisposeAccountCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDisposeAccountCode.Location = New System.Drawing.Point(89, 4)
        Me.TxtDisposeAccountCode.MaxLength = 50
        Me.TxtDisposeAccountCode.Name = "TxtDisposeAccountCode"
        Me.TxtDisposeAccountCode.Size = New System.Drawing.Size(81, 20)
        Me.TxtDisposeAccountCode.TabIndex = 0
        '
        'TxtIncomeAmount
        '
        Me.TxtIncomeAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtIncomeAmount.Location = New System.Drawing.Point(89, 80)
        Me.TxtIncomeAmount.MaxLength = 150
        Me.TxtIncomeAmount.Name = "TxtIncomeAmount"
        Me.TxtIncomeAmount.Size = New System.Drawing.Size(131, 20)
        Me.TxtIncomeAmount.TabIndex = 5
        Me.TxtIncomeAmount.Text = "0"
        Me.TxtIncomeAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(1, 83)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(81, 13)
        Me.Label9.TabIndex = 178
        Me.Label9.Text = "Income Amount"
        '
        'TxtIncomeAccountDesc
        '
        Me.TxtIncomeAccountDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtIncomeAccountDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtIncomeAccountDesc.Location = New System.Drawing.Point(226, 54)
        Me.TxtIncomeAccountDesc.MaxLength = 50
        Me.TxtIncomeAccountDesc.Name = "TxtIncomeAccountDesc"
        Me.TxtIncomeAccountDesc.ReadOnly = True
        Me.TxtIncomeAccountDesc.Size = New System.Drawing.Size(192, 20)
        Me.TxtIncomeAccountDesc.TabIndex = 4
        Me.TxtIncomeAccountDesc.TabStop = False
        '
        'TxtPrimaryIncomeAccount
        '
        Me.TxtPrimaryIncomeAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryIncomeAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryIncomeAccount.Location = New System.Drawing.Point(19, 57)
        Me.TxtPrimaryIncomeAccount.MaxLength = 50
        Me.TxtPrimaryIncomeAccount.Name = "TxtPrimaryIncomeAccount"
        Me.TxtPrimaryIncomeAccount.Size = New System.Drawing.Size(39, 20)
        Me.TxtPrimaryIncomeAccount.TabIndex = 177
        Me.TxtPrimaryIncomeAccount.TabStop = False
        Me.TxtPrimaryIncomeAccount.Text = "-1"
        Me.TxtPrimaryIncomeAccount.Visible = False
        '
        'BtnFindIncomeAccount
        '
        Me.BtnFindIncomeAccount.Location = New System.Drawing.Point(176, 54)
        Me.BtnFindIncomeAccount.Name = "BtnFindIncomeAccount"
        Me.BtnFindIncomeAccount.Size = New System.Drawing.Size(44, 20)
        Me.BtnFindIncomeAccount.TabIndex = 175
        Me.BtnFindIncomeAccount.TabStop = False
        Me.BtnFindIncomeAccount.Text = "Find"
        Me.BtnFindIncomeAccount.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(1, 60)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(85, 13)
        Me.Label8.TabIndex = 176
        Me.Label8.Text = "Income Account"
        '
        'TxtIncomeAccount_Code
        '
        Me.TxtIncomeAccount_Code.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtIncomeAccount_Code.Location = New System.Drawing.Point(89, 54)
        Me.TxtIncomeAccount_Code.MaxLength = 50
        Me.TxtIncomeAccount_Code.Name = "TxtIncomeAccount_Code"
        Me.TxtIncomeAccount_Code.Size = New System.Drawing.Size(81, 20)
        Me.TxtIncomeAccount_Code.TabIndex = 3
        '
        'LblUnit
        '
        Me.LblUnit.AutoSize = True
        Me.LblUnit.Location = New System.Drawing.Point(176, 31)
        Me.LblUnit.Name = "LblUnit"
        Me.LblUnit.Size = New System.Drawing.Size(26, 13)
        Me.LblUnit.TabIndex = 146
        Me.LblUnit.Text = "Unit"
        '
        'TxtDisposeQty
        '
        Me.TxtDisposeQty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDisposeQty.Location = New System.Drawing.Point(89, 28)
        Me.TxtDisposeQty.MaxLength = 150
        Me.TxtDisposeQty.Name = "TxtDisposeQty"
        Me.TxtDisposeQty.Size = New System.Drawing.Size(81, 20)
        Me.TxtDisposeQty.TabIndex = 2
        Me.TxtDisposeQty.Text = "0"
        Me.TxtDisposeQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(1, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 13)
        Me.Label1.TabIndex = 145
        Me.Label1.Text = "Disposed Qty"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount, Me.ColDescAccount, Me.ColDebet, Me.ColKredit, Me.ColExchangeRate, Me.ColOriginalCurrency, Me.ColDebet_OriginalCurr, Me.ColCredit_OriginalCurrency, Me.ColRemarks, Me.ColDept, Me.ColJob, Me.ColPrimaryDept, Me.ColPrimaryJob, Me.ColPrimaryAkun, Me.ColPrimaryCurrency})
        Me.dgvItem.Location = New System.Drawing.Point(6, 130)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.Size = New System.Drawing.Size(848, 202)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 1
        '
        'ColAccount
        '
        Me.ColAccount.HeaderText = "Account"
        Me.ColAccount.MaxInputLength = 50
        Me.ColAccount.Name = "ColAccount"
        Me.ColAccount.ReadOnly = True
        Me.ColAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccount.Width = 53
        '
        'ColDescAccount
        '
        Me.ColDescAccount.HeaderText = "Description"
        Me.ColDescAccount.MaxInputLength = 150
        Me.ColDescAccount.Name = "ColDescAccount"
        Me.ColDescAccount.ReadOnly = True
        Me.ColDescAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDescAccount.Width = 66
        '
        'ColDebet
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle2.NullValue = "0"
        Me.ColDebet.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColDebet.HeaderText = "Debet"
        Me.ColDebet.MaxInputLength = 25
        Me.ColDebet.Name = "ColDebet"
        Me.ColDebet.ReadOnly = True
        Me.ColDebet.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDebet.Width = 42
        '
        'ColKredit
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.NullValue = "0"
        Me.ColKredit.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColKredit.HeaderText = "Credit"
        Me.ColKredit.MaxInputLength = 25
        Me.ColKredit.Name = "ColKredit"
        Me.ColKredit.ReadOnly = True
        Me.ColKredit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColKredit.Width = 40
        '
        'ColExchangeRate
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.NullValue = "1"
        Me.ColExchangeRate.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColExchangeRate.HeaderText = "Exchange Rate"
        Me.ColExchangeRate.Name = "ColExchangeRate"
        Me.ColExchangeRate.ReadOnly = True
        Me.ColExchangeRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColExchangeRate.Visible = False
        Me.ColExchangeRate.Width = 87
        '
        'ColOriginalCurrency
        '
        Me.ColOriginalCurrency.HeaderText = "Currency"
        Me.ColOriginalCurrency.Name = "ColOriginalCurrency"
        Me.ColOriginalCurrency.ReadOnly = True
        Me.ColOriginalCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColOriginalCurrency.Visible = False
        Me.ColOriginalCurrency.Width = 55
        '
        'ColDebet_OriginalCurr
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.NullValue = "0"
        Me.ColDebet_OriginalCurr.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColDebet_OriginalCurr.HeaderText = "Debet"
        Me.ColDebet_OriginalCurr.Name = "ColDebet_OriginalCurr"
        Me.ColDebet_OriginalCurr.ReadOnly = True
        Me.ColDebet_OriginalCurr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDebet_OriginalCurr.Visible = False
        Me.ColDebet_OriginalCurr.Width = 42
        '
        'ColCredit_OriginalCurrency
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle6.NullValue = "0"
        Me.ColCredit_OriginalCurrency.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColCredit_OriginalCurrency.HeaderText = "Credit"
        Me.ColCredit_OriginalCurrency.Name = "ColCredit_OriginalCurrency"
        Me.ColCredit_OriginalCurrency.ReadOnly = True
        Me.ColCredit_OriginalCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCredit_OriginalCurrency.Visible = False
        Me.ColCredit_OriginalCurrency.Width = 40
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.MaxInputLength = 150
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.ReadOnly = True
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 55
        '
        'ColDept
        '
        Me.ColDept.HeaderText = "Departement"
        Me.ColDept.Name = "ColDept"
        Me.ColDept.ReadOnly = True
        Me.ColDept.Visible = False
        Me.ColDept.Width = 93
        '
        'ColJob
        '
        Me.ColJob.HeaderText = "Job"
        Me.ColJob.Name = "ColJob"
        Me.ColJob.ReadOnly = True
        Me.ColJob.Visible = False
        Me.ColJob.Width = 49
        '
        'ColPrimaryDept
        '
        Me.ColPrimaryDept.HeaderText = "Primary Dept"
        Me.ColPrimaryDept.Name = "ColPrimaryDept"
        Me.ColPrimaryDept.ReadOnly = True
        Me.ColPrimaryDept.Visible = False
        Me.ColPrimaryDept.Width = 92
        '
        'ColPrimaryJob
        '
        Me.ColPrimaryJob.HeaderText = "Primary Job"
        Me.ColPrimaryJob.Name = "ColPrimaryJob"
        Me.ColPrimaryJob.ReadOnly = True
        Me.ColPrimaryJob.Visible = False
        Me.ColPrimaryJob.Width = 86
        '
        'ColPrimaryAkun
        '
        Me.ColPrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColPrimaryAkun.Name = "ColPrimaryAkun"
        Me.ColPrimaryAkun.ReadOnly = True
        Me.ColPrimaryAkun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryAkun.Visible = False
        Me.ColPrimaryAkun.Width = 72
        '
        'ColPrimaryCurrency
        '
        Me.ColPrimaryCurrency.HeaderText = "Primary Currency"
        Me.ColPrimaryCurrency.Name = "ColPrimaryCurrency"
        Me.ColPrimaryCurrency.ReadOnly = True
        Me.ColPrimaryCurrency.Visible = False
        Me.ColPrimaryCurrency.Width = 111
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(436, 306)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(320, 306)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmTransaksiJournal_AssetDisposal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(884, 407)
        Me.Controls.Add(Me.btnViewJurnal)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.gboDetail)
        Me.Name = "FrmTransaksiJournal_AssetDisposal"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Asset Disposal"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnViewJurnal As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TxtUnbalance As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryJournal As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtRefNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtAssetDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAsset As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindAsset As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtAssetCode As System.Windows.Forms.TextBox
    Friend WithEvents LblUnit As System.Windows.Forms.Label
    Friend WithEvents TxtDisposeQty As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboDisposalType As System.Windows.Forms.ComboBox
    Friend WithEvents TxtIncomeAccountDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryIncomeAccount As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindIncomeAccount As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtIncomeAccount_Code As System.Windows.Forms.TextBox
    Friend WithEvents chkSell As System.Windows.Forms.CheckBox
    Friend WithEvents TxtIncomeAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ColAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDebet As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKredit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExchangeRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOriginalCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDebet_OriginalCurr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCredit_OriginalCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJob As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryJob As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtDisposeAccountDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryDisposeAccount As System.Windows.Forms.TextBox
    Friend WithEvents btnFindDisposeAccount As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtDisposeAccountCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtDisposed_Accum_Dep_Amount As System.Windows.Forms.TextBox
    Friend WithEvents TxtDisposed_Asset_Amount As System.Windows.Forms.TextBox
End Class
