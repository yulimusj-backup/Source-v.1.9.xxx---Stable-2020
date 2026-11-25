<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmEmailSettings_LogReport
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnPrint = New System.Windows.Forms.Button
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.dgvSortBy = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColField = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOrder = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDirection = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.cboDeliveryStatus = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.dgvCustomer = New System.Windows.Forms.DataGridView
        Me.ColNo2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustCode2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustName2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCust2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.chkShowAllCustomer = New System.Windows.Forms.CheckBox
        Me.btnShow = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.chkListEmailAccounts = New System.Windows.Forms.CheckedListBox
        Me.dgvData = New System.Windows.Forms.DataGridView
        Me.ColNo3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccount3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustCode3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustName3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTransaction3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocNo3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus3 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColAction3 = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColPrimaryAccount3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCust3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryTransAcuan3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccountCode3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccountName3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmailFrom3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmailTo3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCc3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSubject3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMessage3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAttachmentFilename3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.dgvSortBy, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.dgvCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.btnPrint)
        Me.GroupBox1.Controls.Add(Me.GroupBox5)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.btnShow)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(928, 279)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(736, 209)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(59, 21)
        Me.btnPrint.TabIndex = 2
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.dgvSortBy)
        Me.GroupBox5.Location = New System.Drawing.Point(303, 117)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(362, 151)
        Me.GroupBox5.TabIndex = 8
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Sort by"
        '
        'dgvSortBy
        '
        Me.dgvSortBy.AllowUserToAddRows = False
        Me.dgvSortBy.AllowUserToDeleteRows = False
        Me.dgvSortBy.AllowUserToOrderColumns = True
        Me.dgvSortBy.AllowUserToResizeColumns = False
        Me.dgvSortBy.AllowUserToResizeRows = False
        Me.dgvSortBy.ColumnHeadersHeight = 25
        Me.dgvSortBy.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColField, Me.ColOrder, Me.ColDirection})
        Me.dgvSortBy.Location = New System.Drawing.Point(6, 14)
        Me.dgvSortBy.Name = "dgvSortBy"
        Me.dgvSortBy.Size = New System.Drawing.Size(350, 131)
        Me.dgvSortBy.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 30
        '
        'ColField
        '
        Me.ColField.HeaderText = "Data"
        Me.ColField.Name = "ColField"
        Me.ColField.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColOrder
        '
        Me.ColOrder.HeaderText = "Order"
        Me.ColOrder.Name = "ColOrder"
        Me.ColOrder.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColOrder.Width = 50
        '
        'ColDirection
        '
        Me.ColDirection.HeaderText = "Type"
        Me.ColDirection.Items.AddRange(New Object() {"Ascending", "Descending"})
        Me.ColDirection.Name = "ColDirection"
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(801, 209)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(59, 21)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.cboDeliveryStatus)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.dtpTo)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Controls.Add(Me.dtpFrom)
        Me.GroupBox4.Controls.Add(Me.CheckBox1)
        Me.GroupBox4.Controls.Add(Me.CheckBox2)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Location = New System.Drawing.Point(671, 117)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(250, 86)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        '
        'cboDeliveryStatus
        '
        Me.cboDeliveryStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDeliveryStatus.FormattingEnabled = True
        Me.cboDeliveryStatus.Items.AddRange(New Object() {"All", "Success", "Failed"})
        Me.cboDeliveryStatus.Location = New System.Drawing.Point(93, 56)
        Me.cboDeliveryStatus.Name = "cboDeliveryStatus"
        Me.cboDeliveryStatus.Size = New System.Drawing.Size(147, 21)
        Me.cboDeliveryStatus.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 60)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 13)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Delivery Status"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 36)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "To"
        '
        'dtpTo
        '
        Me.dtpTo.Location = New System.Drawing.Point(93, 32)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(147, 20)
        Me.dtpTo.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 14)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(33, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "From "
        '
        'dtpFrom
        '
        Me.dtpFrom.Location = New System.Drawing.Point(93, 10)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(147, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(102, 13)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(102, 17)
        Me.CheckBox1.TabIndex = 2
        Me.CheckBox1.Text = "Collection Letter"
        Me.CheckBox1.UseVisualStyleBackColor = True
        Me.CheckBox1.Visible = False
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(102, 36)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(61, 17)
        Me.CheckBox2.TabIndex = 4
        Me.CheckBox2.Text = "Invoice"
        Me.CheckBox2.UseVisualStyleBackColor = True
        Me.CheckBox2.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Transaction Type"
        Me.Label1.Visible = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.dgvCustomer)
        Me.GroupBox3.Controls.Add(Me.chkShowAllCustomer)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 117)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(291, 151)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        '
        'dgvCustomer
        '
        Me.dgvCustomer.ColumnHeadersHeight = 25
        Me.dgvCustomer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo2, Me.ColCustCode2, Me.ColCustName2, Me.ColPrimaryCust2})
        Me.dgvCustomer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCustomer.Location = New System.Drawing.Point(3, 16)
        Me.dgvCustomer.Name = "dgvCustomer"
        Me.dgvCustomer.Size = New System.Drawing.Size(285, 132)
        Me.dgvCustomer.TabIndex = 1
        '
        'ColNo2
        '
        Me.ColNo2.HeaderText = "No"
        Me.ColNo2.Name = "ColNo2"
        Me.ColNo2.Width = 30
        '
        'ColCustCode2
        '
        Me.ColCustCode2.HeaderText = "Cust Code"
        Me.ColCustCode2.Name = "ColCustCode2"
        '
        'ColCustName2
        '
        Me.ColCustName2.HeaderText = "Cust Name"
        Me.ColCustName2.Name = "ColCustName2"
        '
        'ColPrimaryCust2
        '
        Me.ColPrimaryCust2.HeaderText = "Primary Cust"
        Me.ColPrimaryCust2.Name = "ColPrimaryCust2"
        Me.ColPrimaryCust2.Visible = False
        '
        'chkShowAllCustomer
        '
        Me.chkShowAllCustomer.AutoSize = True
        Me.chkShowAllCustomer.Location = New System.Drawing.Point(10, -1)
        Me.chkShowAllCustomer.Name = "chkShowAllCustomer"
        Me.chkShowAllCustomer.Size = New System.Drawing.Size(114, 17)
        Me.chkShowAllCustomer.TabIndex = 0
        Me.chkShowAllCustomer.Text = "Show All Customer"
        Me.chkShowAllCustomer.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(671, 209)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(59, 21)
        Me.btnShow.TabIndex = 1
        Me.btnShow.Text = "&Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.LinkLabel2)
        Me.GroupBox2.Controls.Add(Me.LinkLabel1)
        Me.GroupBox2.Controls.Add(Me.chkListEmailAccounts)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 10)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(916, 101)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Email Accounts"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(792, 1)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel2.TabIndex = 0
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Check All"
        Me.LinkLabel2.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(845, 1)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel1.TabIndex = 1
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Uncheck All"
        Me.LinkLabel1.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'chkListEmailAccounts
        '
        Me.chkListEmailAccounts.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkListEmailAccounts.ColumnWidth = 250
        Me.chkListEmailAccounts.FormattingEnabled = True
        Me.chkListEmailAccounts.Location = New System.Drawing.Point(6, 16)
        Me.chkListEmailAccounts.MultiColumn = True
        Me.chkListEmailAccounts.Name = "chkListEmailAccounts"
        Me.chkListEmailAccounts.Size = New System.Drawing.Size(904, 79)
        Me.chkListEmailAccounts.TabIndex = 2
        '
        'dgvData
        '
        Me.dgvData.AllowUserToAddRows = False
        Me.dgvData.AllowUserToDeleteRows = False
        Me.dgvData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvData.ColumnHeadersHeight = 25
        Me.dgvData.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo3, Me.ColAccount3, Me.ColCustCode3, Me.ColCustName3, Me.ColTransaction3, Me.ColDate3, Me.ColDocNo3, Me.ColStatus3, Me.ColAction3, Me.ColPrimaryAccount3, Me.ColPrimaryCust3, Me.ColPrimaryTransAcuan3, Me.ColAccountCode3, Me.ColAccountName3, Me.ColEmailFrom3, Me.ColEmailTo3, Me.ColCc3, Me.ColSubject3, Me.ColMessage3, Me.ColAttachmentFilename3})
        Me.dgvData.Location = New System.Drawing.Point(8, 288)
        Me.dgvData.Name = "dgvData"
        Me.dgvData.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvData.Size = New System.Drawing.Size(928, 246)
        Me.dgvData.TabIndex = 0
        '
        'ColNo3
        '
        Me.ColNo3.HeaderText = "No"
        Me.ColNo3.Name = "ColNo3"
        Me.ColNo3.Width = 30
        '
        'ColAccount3
        '
        Me.ColAccount3.HeaderText = "Account"
        Me.ColAccount3.Name = "ColAccount3"
        '
        'ColCustCode3
        '
        Me.ColCustCode3.HeaderText = "Cust Code"
        Me.ColCustCode3.Name = "ColCustCode3"
        '
        'ColCustName3
        '
        Me.ColCustName3.HeaderText = "Cust Name"
        Me.ColCustName3.Name = "ColCustName3"
        '
        'ColTransaction3
        '
        Me.ColTransaction3.HeaderText = "Transaction"
        Me.ColTransaction3.Name = "ColTransaction3"
        '
        'ColDate3
        '
        Me.ColDate3.HeaderText = "Date"
        Me.ColDate3.Name = "ColDate3"
        '
        'ColDocNo3
        '
        Me.ColDocNo3.HeaderText = "Document No"
        Me.ColDocNo3.Name = "ColDocNo3"
        '
        'ColStatus3
        '
        Me.ColStatus3.HeaderText = "Delivery"
        Me.ColStatus3.Name = "ColStatus3"
        Me.ColStatus3.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColStatus3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'ColAction3
        '
        Me.ColAction3.HeaderText = "View Email"
        Me.ColAction3.Name = "ColAction3"
        Me.ColAction3.Text = "View Email"
        Me.ColAction3.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'ColPrimaryAccount3
        '
        Me.ColPrimaryAccount3.HeaderText = "PrimaryAccount"
        Me.ColPrimaryAccount3.Name = "ColPrimaryAccount3"
        Me.ColPrimaryAccount3.Visible = False
        '
        'ColPrimaryCust3
        '
        Me.ColPrimaryCust3.HeaderText = "PrimaryCustomer"
        Me.ColPrimaryCust3.Name = "ColPrimaryCust3"
        Me.ColPrimaryCust3.Visible = False
        '
        'ColPrimaryTransAcuan3
        '
        Me.ColPrimaryTransAcuan3.HeaderText = "PrimaryTransAcuan"
        Me.ColPrimaryTransAcuan3.Name = "ColPrimaryTransAcuan3"
        Me.ColPrimaryTransAcuan3.Visible = False
        '
        'ColAccountCode3
        '
        Me.ColAccountCode3.HeaderText = "Account Code"
        Me.ColAccountCode3.Name = "ColAccountCode3"
        Me.ColAccountCode3.Visible = False
        '
        'ColAccountName3
        '
        Me.ColAccountName3.HeaderText = "Account Name"
        Me.ColAccountName3.Name = "ColAccountName3"
        Me.ColAccountName3.Visible = False
        '
        'ColEmailFrom3
        '
        Me.ColEmailFrom3.HeaderText = "Email From"
        Me.ColEmailFrom3.Name = "ColEmailFrom3"
        Me.ColEmailFrom3.Visible = False
        '
        'ColEmailTo3
        '
        Me.ColEmailTo3.HeaderText = "Email To"
        Me.ColEmailTo3.Name = "ColEmailTo3"
        Me.ColEmailTo3.Visible = False
        '
        'ColCc3
        '
        Me.ColCc3.HeaderText = "CC"
        Me.ColCc3.Name = "ColCc3"
        Me.ColCc3.Visible = False
        '
        'ColSubject3
        '
        Me.ColSubject3.HeaderText = "Subject"
        Me.ColSubject3.Name = "ColSubject3"
        Me.ColSubject3.Visible = False
        '
        'ColMessage3
        '
        Me.ColMessage3.HeaderText = "Message"
        Me.ColMessage3.Name = "ColMessage3"
        Me.ColMessage3.Visible = False
        '
        'ColAttachmentFilename3
        '
        Me.ColAttachmentFilename3.HeaderText = "Attachment Filename"
        Me.ColAttachmentFilename3.Name = "ColAttachmentFilename3"
        Me.ColAttachmentFilename3.Visible = False
        '
        'FrmEmailSettings_LogReport
        '
        Me.AcceptButton = Me.btnShow
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(948, 546)
        Me.Controls.Add(Me.dgvData)
        Me.Controls.Add(Me.GroupBox1)
        Me.MinimumSize = New System.Drawing.Size(895, 584)
        Me.Name = "FrmEmailSettings_LogReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Email Log Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        CType(Me.dgvSortBy, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.dgvCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkListEmailAccounts As System.Windows.Forms.CheckedListBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboDeliveryStatus As System.Windows.Forms.ComboBox
    Friend WithEvents dgvSortBy As System.Windows.Forms.DataGridView
    Friend WithEvents dgvCustomer As System.Windows.Forms.DataGridView
    Friend WithEvents chkShowAllCustomer As System.Windows.Forms.CheckBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents dgvData As System.Windows.Forms.DataGridView
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents ColNo2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustCode2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustName2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCust2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColField As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOrder As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDirection As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColNo3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccount3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustCode3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustName3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTransaction3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocNo3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus3 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColAction3 As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColPrimaryAccount3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCust3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryTransAcuan3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccountCode3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccountName3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmailFrom3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmailTo3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCc3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSubject3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMessage3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAttachmentFilename3 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
