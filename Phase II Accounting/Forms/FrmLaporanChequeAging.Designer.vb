<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanChequeAging
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
        Me.cboType = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.dtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.dgvCheque = New System.Windows.Forms.DataGridView
        Me.ColChequeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBank = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDueDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPaymentTerms = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOverDue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPeriod1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPeriod2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPeriod3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPeriod4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColViewJurnal = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColPay = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColPrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        CType(Me.dgvCheque, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cboType
        '
        Me.cboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboType.FormattingEnabled = True
        Me.cboType.Items.AddRange(New Object() {"In Coming", "Out Going"})
        Me.cboType.Location = New System.Drawing.Point(53, 11)
        Me.cboType.Name = "cboType"
        Me.cboType.Size = New System.Drawing.Size(121, 21)
        Me.cboType.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 15)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(31, 13)
        Me.Label4.TabIndex = 37
        Me.Label4.Text = "Type"
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(354, 34)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 3
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(354, 11)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(75, 23)
        Me.btnFilter.TabIndex = 2
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'dtpPeriod
        '
        Me.dtpPeriod.CustomFormat = "MMM yyyy"
        Me.dtpPeriod.Enabled = False
        Me.dtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPeriod.Location = New System.Drawing.Point(223, 11)
        Me.dtpPeriod.Name = "dtpPeriod"
        Me.dtpPeriod.Size = New System.Drawing.Size(125, 20)
        Me.dtpPeriod.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(180, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(37, 13)
        Me.Label3.TabIndex = 31
        Me.Label3.Text = "Period"
        '
        'dgvCheque
        '
        Me.dgvCheque.AllowUserToAddRows = False
        Me.dgvCheque.AllowUserToDeleteRows = False
        Me.dgvCheque.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCheque.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColChequeNo, Me.ColBank, Me.ColDate, Me.ColDueDate, Me.ColPaymentTerms, Me.ColCurrency, Me.ColOverDue, Me.ColPeriod1, Me.ColPeriod2, Me.ColPeriod3, Me.ColPeriod4, Me.ColIdentifier, Me.ColViewJurnal, Me.ColPay, Me.ColPrimaryCurrency})
        Me.dgvCheque.Location = New System.Drawing.Point(12, 60)
        Me.dgvCheque.Name = "dgvCheque"
        Me.dgvCheque.Size = New System.Drawing.Size(860, 363)
        Me.dgvCheque.TabIndex = 4
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
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        '
        'ColDueDate
        '
        Me.ColDueDate.HeaderText = "Due Date"
        Me.ColDueDate.Name = "ColDueDate"
        '
        'ColPaymentTerms
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPaymentTerms.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColPaymentTerms.HeaderText = "Terms (Days)"
        Me.ColPaymentTerms.Name = "ColPaymentTerms"
        '
        'ColCurrency
        '
        Me.ColCurrency.HeaderText = "Currency"
        Me.ColCurrency.Name = "ColCurrency"
        '
        'ColOverDue
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColOverDue.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColOverDue.HeaderText = "Over Due"
        Me.ColOverDue.Name = "ColOverDue"
        '
        'ColPeriod1
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPeriod1.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColPeriod1.HeaderText = "Period 1"
        Me.ColPeriod1.Name = "ColPeriod1"
        '
        'ColPeriod2
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPeriod2.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColPeriod2.HeaderText = "Period 2"
        Me.ColPeriod2.Name = "ColPeriod2"
        '
        'ColPeriod3
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPeriod3.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColPeriod3.HeaderText = "Period 3"
        Me.ColPeriod3.Name = "ColPeriod3"
        '
        'ColPeriod4
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPeriod4.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColPeriod4.HeaderText = ">3 months"
        Me.ColPeriod4.Name = "ColPeriod4"
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.Visible = False
        '
        'ColViewJurnal
        '
        Me.ColViewJurnal.HeaderText = "View Jurnal"
        Me.ColViewJurnal.Name = "ColViewJurnal"
        Me.ColViewJurnal.Text = "View"
        Me.ColViewJurnal.UseColumnTextForButtonValue = True
        '
        'ColPay
        '
        Me.ColPay.HeaderText = "Pay"
        Me.ColPay.Name = "ColPay"
        Me.ColPay.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPay.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColPay.Text = "Pay"
        Me.ColPay.UseColumnTextForButtonValue = True
        '
        'ColPrimaryCurrency
        '
        Me.ColPrimaryCurrency.HeaderText = "Primary Currency"
        Me.ColPrimaryCurrency.Name = "ColPrimaryCurrency"
        Me.ColPrimaryCurrency.Visible = False
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(797, 429)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'FrmLaporanChequeAging
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(884, 462)
        Me.Controls.Add(Me.cboType)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.dtpPeriod)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dgvCheque)
        Me.Controls.Add(Me.btnClose)
        Me.MinimumSize = New System.Drawing.Size(900, 500)
        Me.Name = "FrmLaporanChequeAging"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cheque Aging Report"
        CType(Me.dgvCheque, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboType As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents dtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dgvCheque As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ColChequeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBank As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDueDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPaymentTerms As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOverDue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPeriod1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPeriod2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPeriod3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPeriod4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColViewJurnal As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColPay As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColPrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
