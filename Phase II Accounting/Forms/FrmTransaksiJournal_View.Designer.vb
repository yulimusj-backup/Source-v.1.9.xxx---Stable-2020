<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiJournal_View
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TxtRemarksItem = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtJournalDescription = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.chkJournalType = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.clbTipeJournal = New System.Windows.Forms.CheckedListBox
        Me.BtnFindTo = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnFindFrom = New System.Windows.Forms.Button
        Me.TxtTo = New System.Windows.Forms.TextBox
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryFrom = New System.Windows.Forms.TextBox
        Me.TxtPrimaryTo = New System.Windows.Forms.TextBox
        Me.TxtFrom = New System.Windows.Forms.TextBox
        Me.btnShow = New System.Windows.Forms.Button
        Me.cboFilterBy = New System.Windows.Forms.ComboBox
        Me.crvJournal = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboFormat = New System.Windows.Forms.ComboBox
        Me.chkShowRemarks = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(94, 503)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dgvItem)
        Me.GroupBox1.Controls.Add(Me.TxtRemarksItem)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.TxtJournalDescription)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.chkJournalType)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.BtnFindTo)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.dtpTo)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.btnFindFrom)
        Me.GroupBox1.Controls.Add(Me.TxtTo)
        Me.GroupBox1.Controls.Add(Me.dtpFrom)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryFrom)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryTo)
        Me.GroupBox1.Controls.Add(Me.TxtFrom)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(314, 458)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter By"
        '
        'dgvItem
        '
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount, Me.ColDescAccount, Me.ColPrimaryAkun})
        Me.dgvItem.Location = New System.Drawing.Point(5, 349)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(303, 104)
        Me.dgvItem.TabIndex = 8
        '
        'ColAccount
        '
        Me.ColAccount.HeaderText = "Account Code"
        Me.ColAccount.Name = "ColAccount"
        '
        'ColDescAccount
        '
        Me.ColDescAccount.HeaderText = "Description"
        Me.ColDescAccount.Name = "ColDescAccount"
        Me.ColDescAccount.Width = 85
        '
        'ColPrimaryAkun
        '
        Me.ColPrimaryAkun.HeaderText = "Primary Account"
        Me.ColPrimaryAkun.Name = "ColPrimaryAkun"
        Me.ColPrimaryAkun.Visible = False
        Me.ColPrimaryAkun.Width = 109
        '
        'TxtRemarksItem
        '
        Me.TxtRemarksItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarksItem.Location = New System.Drawing.Point(81, 321)
        Me.TxtRemarksItem.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtRemarksItem.Name = "TxtRemarksItem"
        Me.TxtRemarksItem.Size = New System.Drawing.Size(229, 20)
        Me.TxtRemarksItem.TabIndex = 14
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 324)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 13)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Item Remarks"
        '
        'TxtJournalDescription
        '
        Me.TxtJournalDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtJournalDescription.Location = New System.Drawing.Point(81, 299)
        Me.TxtJournalDescription.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtJournalDescription.Name = "TxtJournalDescription"
        Me.TxtJournalDescription.Size = New System.Drawing.Size(229, 20)
        Me.TxtJournalDescription.TabIndex = 8
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 301)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Description"
        '
        'chkJournalType
        '
        Me.chkJournalType.AutoSize = True
        Me.chkJournalType.Location = New System.Drawing.Point(16, 67)
        Me.chkJournalType.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkJournalType.Name = "chkJournalType"
        Me.chkJournalType.Size = New System.Drawing.Size(87, 17)
        Me.chkJournalType.TabIndex = 8
        Me.chkJournalType.Text = "Journal Type"
        Me.chkJournalType.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.llbUncheckAll)
        Me.GroupBox2.Controls.Add(Me.llbCheckAll)
        Me.GroupBox2.Controls.Add(Me.clbTipeJournal)
        Me.GroupBox2.Location = New System.Drawing.Point(5, 71)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox2.Size = New System.Drawing.Size(303, 228)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.BackColor = System.Drawing.SystemColors.Window
        Me.llbUncheckAll.Location = New System.Drawing.Point(233, 214)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAll.TabIndex = 4
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        '
        'llbCheckAll
        '
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.BackColor = System.Drawing.SystemColors.Window
        Me.llbCheckAll.Location = New System.Drawing.Point(174, 214)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAll.TabIndex = 13
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        '
        'clbTipeJournal
        '
        Me.clbTipeJournal.CheckOnClick = True
        Me.clbTipeJournal.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbTipeJournal.FormattingEnabled = True
        Me.clbTipeJournal.Items.AddRange(New Object() {"General Journal", "Cash And Bank", "End of Month Journal", "Sales Journal", "Sales Return Journal", "Sales Down Payment", "Purchase Journal", "Purchase DownPayment Journal", "Purchase Return Journal", "Sales Receivement Journal", "Purchase Payment Journal", "Credit Note", "Debit Note"})
        Me.clbTipeJournal.Location = New System.Drawing.Point(2, 15)
        Me.clbTipeJournal.Name = "clbTipeJournal"
        Me.clbTipeJournal.Size = New System.Drawing.Size(299, 199)
        Me.clbTipeJournal.TabIndex = 2
        '
        'BtnFindTo
        '
        Me.BtnFindTo.Location = New System.Drawing.Point(197, 45)
        Me.BtnFindTo.Name = "BtnFindTo"
        Me.BtnFindTo.Size = New System.Drawing.Size(32, 20)
        Me.BtnFindTo.TabIndex = 10
        Me.BtnFindTo.Text = "..."
        Me.BtnFindTo.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "To"
        '
        'dtpTo
        '
        Me.dtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(42, 46)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(263, 20)
        Me.dtpTo.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "From"
        '
        'btnFindFrom
        '
        Me.btnFindFrom.Location = New System.Drawing.Point(197, 19)
        Me.btnFindFrom.Name = "btnFindFrom"
        Me.btnFindFrom.Size = New System.Drawing.Size(32, 20)
        Me.btnFindFrom.TabIndex = 8
        Me.btnFindFrom.Text = "..."
        Me.btnFindFrom.UseVisualStyleBackColor = True
        '
        'TxtTo
        '
        Me.TxtTo.Location = New System.Drawing.Point(42, 45)
        Me.TxtTo.Name = "TxtTo"
        Me.TxtTo.Size = New System.Drawing.Size(149, 20)
        Me.TxtTo.TabIndex = 9
        Me.TxtTo.TabStop = False
        '
        'dtpFrom
        '
        Me.dtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(42, 19)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(263, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'TxtPrimaryFrom
        '
        Me.TxtPrimaryFrom.Location = New System.Drawing.Point(42, 39)
        Me.TxtPrimaryFrom.Name = "TxtPrimaryFrom"
        Me.TxtPrimaryFrom.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimaryFrom.TabIndex = 11
        Me.TxtPrimaryFrom.TabStop = False
        Me.TxtPrimaryFrom.Text = "-1"
        Me.TxtPrimaryFrom.Visible = False
        '
        'TxtPrimaryTo
        '
        Me.TxtPrimaryTo.Location = New System.Drawing.Point(42, 65)
        Me.TxtPrimaryTo.Name = "TxtPrimaryTo"
        Me.TxtPrimaryTo.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimaryTo.TabIndex = 12
        Me.TxtPrimaryTo.TabStop = False
        Me.TxtPrimaryTo.Text = "-1"
        Me.TxtPrimaryTo.Visible = False
        '
        'TxtFrom
        '
        Me.TxtFrom.Location = New System.Drawing.Point(42, 19)
        Me.TxtFrom.Name = "TxtFrom"
        Me.TxtFrom.Size = New System.Drawing.Size(149, 20)
        Me.TxtFrom.TabIndex = 0
        Me.TxtFrom.TabStop = False
        '
        'btnShow
        '
        Me.btnShow.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnShow.Location = New System.Drawing.Point(13, 503)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(75, 23)
        Me.btnShow.TabIndex = 3
        Me.btnShow.Text = "&Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'cboFilterBy
        '
        Me.cboFilterBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFilterBy.FormattingEnabled = True
        Me.cboFilterBy.Items.AddRange(New Object() {"Period", "Journal Number"})
        Me.cboFilterBy.Location = New System.Drawing.Point(399, 101)
        Me.cboFilterBy.Name = "cboFilterBy"
        Me.cboFilterBy.Size = New System.Drawing.Size(223, 21)
        Me.cboFilterBy.TabIndex = 0
        Me.cboFilterBy.TabStop = False
        '
        'crvJournal
        '
        Me.crvJournal.ActiveViewIndex = -1
        Me.crvJournal.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvJournal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvJournal.DisplayGroupTree = False
        Me.crvJournal.Location = New System.Drawing.Point(332, 12)
        Me.crvJournal.Name = "crvJournal"
        Me.crvJournal.SelectionFormula = ""
        Me.crvJournal.ShowCloseButton = False
        Me.crvJournal.ShowGroupTreeButton = False
        Me.crvJournal.ShowRefreshButton = False
        Me.crvJournal.Size = New System.Drawing.Size(540, 514)
        Me.crvJournal.TabIndex = 0
        Me.crvJournal.TabStop = False
        Me.crvJournal.ViewTimeSelectionFormula = ""
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 478)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Format"
        '
        'cboFormat
        '
        Me.cboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat.FormattingEnabled = True
        Me.cboFormat.Items.AddRange(New Object() {"Complete", "Original Currency", "Base Currency (IDR)"})
        Me.cboFormat.Location = New System.Drawing.Point(52, 475)
        Me.cboFormat.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboFormat.Name = "cboFormat"
        Me.cboFormat.Size = New System.Drawing.Size(147, 21)
        Me.cboFormat.TabIndex = 6
        '
        'chkShowRemarks
        '
        Me.chkShowRemarks.AutoSize = True
        Me.chkShowRemarks.Location = New System.Drawing.Point(206, 477)
        Me.chkShowRemarks.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkShowRemarks.Name = "chkShowRemarks"
        Me.chkShowRemarks.Size = New System.Drawing.Size(121, 17)
        Me.chkShowRemarks.TabIndex = 7
        Me.chkShowRemarks.Text = "Show Item Remarks"
        Me.chkShowRemarks.UseVisualStyleBackColor = True
        '
        'FrmTransaksiJournal_View
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(884, 539)
        Me.Controls.Add(Me.chkShowRemarks)
        Me.Controls.Add(Me.cboFormat)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.crvJournal)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cboFilterBy)
        Me.Controls.Add(Me.btnShow)
        Me.Controls.Add(Me.btnClose)
        Me.MinimumSize = New System.Drawing.Size(856, 576)
        Me.Name = "FrmTransaksiJournal_View"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Journal View"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnFindTo As System.Windows.Forms.Button
    Friend WithEvents TxtTo As System.Windows.Forms.TextBox
    Friend WithEvents btnFindFrom As System.Windows.Forms.Button
    Friend WithEvents TxtFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboFilterBy As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPrimaryTo As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryFrom As System.Windows.Forms.TextBox
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents crvJournal As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents clbTipeJournal As System.Windows.Forms.CheckedListBox
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents chkShowRemarks As System.Windows.Forms.CheckBox
    Friend WithEvents chkJournalType As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtRemarksItem As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtJournalDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
