<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanChequePerPeriod
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.dgvCheque = New System.Windows.Forms.DataGridView
        Me.ColChequeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBank = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDueDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurr = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSPaid = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurr = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColViewJurnal = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColPaid = New System.Windows.Forms.DataGridViewButtonColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkPaid = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnFilter = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.chkDate = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker
        Me.cboType = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        CType(Me.dgvCheque, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(801, 427)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(71, 23)
        Me.btnClose.TabIndex = 12
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'dgvCheque
        '
        Me.dgvCheque.AllowUserToAddRows = False
        Me.dgvCheque.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvCheque.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvCheque.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCheque.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColChequeNo, Me.ColBank, Me.ColDate, Me.ColDueDate, Me.ColCurr, Me.ColAmount, Me.ColRemarks, Me.ColSPaid, Me.ColIdentifier, Me.ColPrimaryCurr, Me.ColViewJurnal, Me.ColPaid})
        Me.dgvCheque.Location = New System.Drawing.Point(12, 58)
        Me.dgvCheque.MultiSelect = False
        Me.dgvCheque.Name = "dgvCheque"
        Me.dgvCheque.Size = New System.Drawing.Size(860, 363)
        Me.dgvCheque.TabIndex = 11
        '
        'ColChequeNo
        '
        Me.ColChequeNo.HeaderText = "Cheque No."
        Me.ColChequeNo.Name = "ColChequeNo"
        '
        'ColBank
        '
        Me.ColBank.HeaderText = "Bank"
        Me.ColBank.Name = "ColBank"
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Cheque Date"
        Me.ColDate.Name = "ColDate"
        '
        'ColDueDate
        '
        Me.ColDueDate.HeaderText = "Due Date"
        Me.ColDueDate.Name = "ColDueDate"
        '
        'ColCurr
        '
        Me.ColCurr.HeaderText = "Currency"
        Me.ColCurr.Name = "ColCurr"
        '
        'ColAmount
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        '
        'ColSPaid
        '
        Me.ColSPaid.HeaderText = "Paid"
        Me.ColSPaid.Name = "ColSPaid"
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.Visible = False
        '
        'ColPrimaryCurr
        '
        Me.ColPrimaryCurr.HeaderText = "Primary Currency"
        Me.ColPrimaryCurr.Name = "ColPrimaryCurr"
        Me.ColPrimaryCurr.Visible = False
        '
        'ColViewJurnal
        '
        Me.ColViewJurnal.HeaderText = "View Jurnal"
        Me.ColViewJurnal.Name = "ColViewJurnal"
        Me.ColViewJurnal.Text = "View"
        Me.ColViewJurnal.UseColumnTextForButtonValue = True
        '
        'ColPaid
        '
        Me.ColPaid.HeaderText = "Modify Payment"
        Me.ColPaid.Name = "ColPaid"
        Me.ColPaid.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPaid.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColPaid.Text = "Modify"
        Me.ColPaid.UseColumnTextForButtonValue = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(187, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Status"
        '
        'chkPaid
        '
        Me.chkPaid.AutoSize = True
        Me.chkPaid.Location = New System.Drawing.Point(230, 12)
        Me.chkPaid.Name = "chkPaid"
        Me.chkPaid.Size = New System.Drawing.Size(47, 17)
        Me.chkPaid.TabIndex = 1
        Me.chkPaid.Text = "Paid"
        Me.chkPaid.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(230, 35)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(60, 17)
        Me.CheckBox1.TabIndex = 2
        Me.CheckBox1.Text = "Unpaid"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Enabled = False
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(407, 9)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(125, 20)
        Me.dtpFrom.TabIndex = 4
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Enabled = False
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(560, 9)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(125, 20)
        Me.dtpTo.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(538, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(16, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "to"
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(691, 9)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(75, 23)
        Me.btnFilter.TabIndex = 9
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(691, 32)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 10
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'chkDate
        '
        Me.chkDate.AutoSize = True
        Me.chkDate.Location = New System.Drawing.Point(304, 11)
        Me.chkDate.Name = "chkDate"
        Me.chkDate.Size = New System.Drawing.Size(74, 17)
        Me.chkDate.TabIndex = 3
        Me.chkDate.Text = "Date Filter"
        Me.chkDate.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(304, 37)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(97, 17)
        Me.CheckBox2.TabIndex = 6
        Me.CheckBox2.Text = "Due Date Filter"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "dd MMM yyyy"
        Me.DateTimePicker1.Enabled = False
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(560, 35)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(125, 20)
        Me.DateTimePicker1.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(538, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(16, 13)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "to"
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CustomFormat = "dd MMM yyyy"
        Me.DateTimePicker2.Enabled = False
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker2.Location = New System.Drawing.Point(407, 35)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(125, 20)
        Me.DateTimePicker2.TabIndex = 7
        '
        'cboType
        '
        Me.cboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboType.FormattingEnabled = True
        Me.cboType.Items.AddRange(New Object() {"In Coming", "Out Going"})
        Me.cboType.Location = New System.Drawing.Point(53, 9)
        Me.cboType.Name = "cboType"
        Me.cboType.Size = New System.Drawing.Size(121, 21)
        Me.cboType.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 13)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(31, 13)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "Type"
        '
        'FrmLaporanChequePerPeriod
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(884, 462)
        Me.Controls.Add(Me.cboType)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DateTimePicker2)
        Me.Controls.Add(Me.chkDate)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.chkPaid)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvCheque)
        Me.Controls.Add(Me.btnClose)
        Me.MinimumSize = New System.Drawing.Size(900, 500)
        Me.Name = "FrmLaporanChequePerPeriod"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cheque Report Per Period"
        CType(Me.dgvCheque, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents dgvCheque As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkPaid As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents chkDate As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboType As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ColChequeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBank As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDueDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSPaid As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColViewJurnal As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColPaid As System.Windows.Forms.DataGridViewButtonColumn
End Class
