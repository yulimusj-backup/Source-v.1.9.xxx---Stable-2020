<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiJournal_EndOfMonth
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
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.dgvItemAP = New System.Windows.Forms.DataGridView
        Me.ColItemAP_PrimaryTransaksiAsal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAP_TipeTransaksiAsal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAP_PrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAP_Saldo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAP_SaldoOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAP_ExchangeRateLama = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAP_PrimaryAkunAP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAP_Selisih = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAP_Tanggal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAP_PrimaryPihakAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAP_PaymentTermAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvItemAR = New System.Windows.Forms.DataGridView
        Me.ColItemAR_PrimaryTransaksiAsal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAR_TipeTransaksiAsal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAR_PrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAR_Saldo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAR_SaldoOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAR_ExchangeRateLama = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAR_PrimaryAkunAR = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAR_Selisih = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAR_Tanggal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAR_PrimaryPihakAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAR_PaymentTermAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label6 = New System.Windows.Forms.Label
        Me.LblQuickJurnal = New System.Windows.Forms.Label
        Me.TxtUnbalance = New System.Windows.Forms.TextBox
        Me.llbEditQuickJournal = New System.Windows.Forms.LinkLabel
        Me.cboQuickJournal = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtPrimaryJournal = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtRefNo = New System.Windows.Forms.TextBox
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.dgvExchangeRate = New System.Windows.Forms.DataGridView
        Me.ColItemExc_Currency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemExc_ExchangeRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemExc_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemExc_PrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.ColPrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColShowDetail = New System.Windows.Forms.DataGridViewButtonColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnViewJurnal = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.ReccuringTab = New System.Windows.Forms.TabPage
        Me.AssetDepreciationTab = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail.SuspendLayout()
        CType(Me.dgvItemAP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItemAR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvExchangeRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.dgvItemAP)
        Me.gboDetail.Controls.Add(Me.dgvItemAR)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(8, 7)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.gboDetail.Size = New System.Drawing.Size(1157, 677)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'dgvItemAP
        '
        Me.dgvItemAP.AllowUserToAddRows = False
        Me.dgvItemAP.AllowUserToDeleteRows = False
        Me.dgvItemAP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemAP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItemAP_PrimaryTransaksiAsal, Me.ColItemAP_TipeTransaksiAsal, Me.ColItemAP_PrimaryCurrency, Me.ColItemAP_Saldo, Me.ColItemAP_SaldoOriginal, Me.ColItemAP_ExchangeRateLama, Me.ColItemAP_PrimaryAkunAP, Me.ColItemAP_Selisih, Me.ColItemAP_Tanggal, Me.ColItemAP_PrimaryPihakAcuan, Me.ColItemAP_PaymentTermAcuan})
        Me.dgvItemAP.Location = New System.Drawing.Point(56, 188)
        Me.dgvItemAP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvItemAP.Name = "dgvItemAP"
        Me.dgvItemAP.ReadOnly = True
        Me.dgvItemAP.Size = New System.Drawing.Size(1035, 206)
        Me.dgvItemAP.TabIndex = 2
        Me.dgvItemAP.TabStop = False
        Me.dgvItemAP.Visible = False
        '
        'ColItemAP_PrimaryTransaksiAsal
        '
        Me.ColItemAP_PrimaryTransaksiAsal.HeaderText = "Primary Transaksi Asal"
        Me.ColItemAP_PrimaryTransaksiAsal.Name = "ColItemAP_PrimaryTransaksiAsal"
        Me.ColItemAP_PrimaryTransaksiAsal.ReadOnly = True
        '
        'ColItemAP_TipeTransaksiAsal
        '
        Me.ColItemAP_TipeTransaksiAsal.HeaderText = "TipeTransaksiAsal"
        Me.ColItemAP_TipeTransaksiAsal.Name = "ColItemAP_TipeTransaksiAsal"
        Me.ColItemAP_TipeTransaksiAsal.ReadOnly = True
        '
        'ColItemAP_PrimaryCurrency
        '
        Me.ColItemAP_PrimaryCurrency.HeaderText = "Primary Currency"
        Me.ColItemAP_PrimaryCurrency.Name = "ColItemAP_PrimaryCurrency"
        Me.ColItemAP_PrimaryCurrency.ReadOnly = True
        '
        'ColItemAP_Saldo
        '
        Me.ColItemAP_Saldo.HeaderText = "Saldo"
        Me.ColItemAP_Saldo.Name = "ColItemAP_Saldo"
        Me.ColItemAP_Saldo.ReadOnly = True
        '
        'ColItemAP_SaldoOriginal
        '
        Me.ColItemAP_SaldoOriginal.HeaderText = "Saldo Original"
        Me.ColItemAP_SaldoOriginal.Name = "ColItemAP_SaldoOriginal"
        Me.ColItemAP_SaldoOriginal.ReadOnly = True
        '
        'ColItemAP_ExchangeRateLama
        '
        Me.ColItemAP_ExchangeRateLama.HeaderText = "ExchangeRateLama"
        Me.ColItemAP_ExchangeRateLama.Name = "ColItemAP_ExchangeRateLama"
        Me.ColItemAP_ExchangeRateLama.ReadOnly = True
        '
        'ColItemAP_PrimaryAkunAP
        '
        Me.ColItemAP_PrimaryAkunAP.HeaderText = "Primary Akun AP"
        Me.ColItemAP_PrimaryAkunAP.Name = "ColItemAP_PrimaryAkunAP"
        Me.ColItemAP_PrimaryAkunAP.ReadOnly = True
        '
        'ColItemAP_Selisih
        '
        Me.ColItemAP_Selisih.HeaderText = "Selisih"
        Me.ColItemAP_Selisih.Name = "ColItemAP_Selisih"
        Me.ColItemAP_Selisih.ReadOnly = True
        '
        'ColItemAP_Tanggal
        '
        Me.ColItemAP_Tanggal.HeaderText = "Tanggal"
        Me.ColItemAP_Tanggal.Name = "ColItemAP_Tanggal"
        Me.ColItemAP_Tanggal.ReadOnly = True
        '
        'ColItemAP_PrimaryPihakAcuan
        '
        Me.ColItemAP_PrimaryPihakAcuan.HeaderText = "Primary Pihak Acuan"
        Me.ColItemAP_PrimaryPihakAcuan.Name = "ColItemAP_PrimaryPihakAcuan"
        Me.ColItemAP_PrimaryPihakAcuan.ReadOnly = True
        '
        'ColItemAP_PaymentTermAcuan
        '
        Me.ColItemAP_PaymentTermAcuan.HeaderText = "Payment Term Acuan"
        Me.ColItemAP_PaymentTermAcuan.Name = "ColItemAP_PaymentTermAcuan"
        Me.ColItemAP_PaymentTermAcuan.ReadOnly = True
        '
        'dgvItemAR
        '
        Me.dgvItemAR.AllowUserToAddRows = False
        Me.dgvItemAR.AllowUserToDeleteRows = False
        Me.dgvItemAR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemAR.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItemAR_PrimaryTransaksiAsal, Me.ColItemAR_TipeTransaksiAsal, Me.ColItemAR_PrimaryCurrency, Me.ColItemAR_Saldo, Me.ColItemAR_SaldoOriginal, Me.ColItemAR_ExchangeRateLama, Me.ColItemAR_PrimaryAkunAR, Me.ColItemAR_Selisih, Me.ColItemAR_Tanggal, Me.ColItemAR_PrimaryPihakAcuan, Me.ColItemAR_PaymentTermAcuan})
        Me.dgvItemAR.Location = New System.Drawing.Point(56, 401)
        Me.dgvItemAR.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvItemAR.Name = "dgvItemAR"
        Me.dgvItemAR.ReadOnly = True
        Me.dgvItemAR.Size = New System.Drawing.Size(1035, 206)
        Me.dgvItemAR.TabIndex = 1
        Me.dgvItemAR.TabStop = False
        Me.dgvItemAR.Visible = False
        '
        'ColItemAR_PrimaryTransaksiAsal
        '
        Me.ColItemAR_PrimaryTransaksiAsal.HeaderText = "Primary Transaksi Asal"
        Me.ColItemAR_PrimaryTransaksiAsal.Name = "ColItemAR_PrimaryTransaksiAsal"
        Me.ColItemAR_PrimaryTransaksiAsal.ReadOnly = True
        '
        'ColItemAR_TipeTransaksiAsal
        '
        Me.ColItemAR_TipeTransaksiAsal.HeaderText = "TipeTransaksiAsal"
        Me.ColItemAR_TipeTransaksiAsal.Name = "ColItemAR_TipeTransaksiAsal"
        Me.ColItemAR_TipeTransaksiAsal.ReadOnly = True
        '
        'ColItemAR_PrimaryCurrency
        '
        Me.ColItemAR_PrimaryCurrency.HeaderText = "Primary Currency"
        Me.ColItemAR_PrimaryCurrency.Name = "ColItemAR_PrimaryCurrency"
        Me.ColItemAR_PrimaryCurrency.ReadOnly = True
        '
        'ColItemAR_Saldo
        '
        Me.ColItemAR_Saldo.HeaderText = "Saldo"
        Me.ColItemAR_Saldo.Name = "ColItemAR_Saldo"
        Me.ColItemAR_Saldo.ReadOnly = True
        '
        'ColItemAR_SaldoOriginal
        '
        Me.ColItemAR_SaldoOriginal.HeaderText = "Saldo Original"
        Me.ColItemAR_SaldoOriginal.Name = "ColItemAR_SaldoOriginal"
        Me.ColItemAR_SaldoOriginal.ReadOnly = True
        '
        'ColItemAR_ExchangeRateLama
        '
        Me.ColItemAR_ExchangeRateLama.HeaderText = "ExchangeRateLama"
        Me.ColItemAR_ExchangeRateLama.Name = "ColItemAR_ExchangeRateLama"
        Me.ColItemAR_ExchangeRateLama.ReadOnly = True
        '
        'ColItemAR_PrimaryAkunAR
        '
        Me.ColItemAR_PrimaryAkunAR.HeaderText = "Primary Akun AR"
        Me.ColItemAR_PrimaryAkunAR.Name = "ColItemAR_PrimaryAkunAR"
        Me.ColItemAR_PrimaryAkunAR.ReadOnly = True
        '
        'ColItemAR_Selisih
        '
        Me.ColItemAR_Selisih.HeaderText = "Selisih"
        Me.ColItemAR_Selisih.Name = "ColItemAR_Selisih"
        Me.ColItemAR_Selisih.ReadOnly = True
        '
        'ColItemAR_Tanggal
        '
        Me.ColItemAR_Tanggal.HeaderText = "Tanggal"
        Me.ColItemAR_Tanggal.Name = "ColItemAR_Tanggal"
        Me.ColItemAR_Tanggal.ReadOnly = True
        '
        'ColItemAR_PrimaryPihakAcuan
        '
        Me.ColItemAR_PrimaryPihakAcuan.HeaderText = "Primary Pihak Acuan"
        Me.ColItemAR_PrimaryPihakAcuan.Name = "ColItemAR_PrimaryPihakAcuan"
        Me.ColItemAR_PrimaryPihakAcuan.ReadOnly = True
        '
        'ColItemAR_PaymentTermAcuan
        '
        Me.ColItemAR_PaymentTermAcuan.HeaderText = "Payment Term Acuan"
        Me.ColItemAR_PaymentTermAcuan.Name = "ColItemAR_PaymentTermAcuan"
        Me.ColItemAR_PaymentTermAcuan.ReadOnly = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(8, 10)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblQuickJurnal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtUnbalance)
        Me.SplitContainer1.Panel1.Controls.Add(Me.llbEditQuickJournal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboQuickJournal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtReadOnly)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryJournal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRefNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRemarks)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvExchangeRate)
        Me.SplitContainer1.Size = New System.Drawing.Size(1134, 138)
        Me.SplitContainer1.SplitterDistance = 566
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(4, 39)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(79, 17)
        Me.Label6.TabIndex = 143
        Me.Label6.Text = "Description"
        '
        'LblQuickJurnal
        '
        Me.LblQuickJurnal.AutoSize = True
        Me.LblQuickJurnal.Location = New System.Drawing.Point(4, 38)
        Me.LblQuickJurnal.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblQuickJurnal.Name = "LblQuickJurnal"
        Me.LblQuickJurnal.Size = New System.Drawing.Size(95, 17)
        Me.LblQuickJurnal.TabIndex = 214
        Me.LblQuickJurnal.Text = "Quick Journal"
        Me.LblQuickJurnal.Visible = False
        '
        'TxtUnbalance
        '
        Me.TxtUnbalance.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUnbalance.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtUnbalance.Location = New System.Drawing.Point(424, 36)
        Me.TxtUnbalance.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtUnbalance.MaxLength = 50
        Me.TxtUnbalance.Name = "TxtUnbalance"
        Me.TxtUnbalance.ReadOnly = True
        Me.TxtUnbalance.Size = New System.Drawing.Size(137, 22)
        Me.TxtUnbalance.TabIndex = 1
        Me.TxtUnbalance.TabStop = False
        Me.TxtUnbalance.Text = "0"
        Me.TxtUnbalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtUnbalance.Visible = False
        '
        'llbEditQuickJournal
        '
        Me.llbEditQuickJournal.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbEditQuickJournal.AutoSize = True
        Me.llbEditQuickJournal.Location = New System.Drawing.Point(505, 66)
        Me.llbEditQuickJournal.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbEditQuickJournal.Name = "llbEditQuickJournal"
        Me.llbEditQuickJournal.Size = New System.Drawing.Size(32, 17)
        Me.llbEditQuickJournal.TabIndex = 3
        Me.llbEditQuickJournal.TabStop = True
        Me.llbEditQuickJournal.Text = "Edit"
        Me.llbEditQuickJournal.Visible = False
        '
        'cboQuickJournal
        '
        Me.cboQuickJournal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboQuickJournal.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboQuickJournal.FormattingEnabled = True
        Me.cboQuickJournal.Location = New System.Drawing.Point(139, 36)
        Me.cboQuickJournal.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboQuickJournal.Name = "cboQuickJournal"
        Me.cboQuickJournal.Size = New System.Drawing.Size(354, 24)
        Me.cboQuickJournal.TabIndex = 2
        Me.cboQuickJournal.TabStop = False
        Me.cboQuickJournal.Visible = False
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(337, 36)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 17)
        Me.Label2.TabIndex = 164
        Me.Label2.Text = "Unbalance"
        Me.Label2.Visible = False
        '
        'txtReadOnly
        '
        Me.txtReadOnly.Location = New System.Drawing.Point(100, 4)
        Me.txtReadOnly.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(29, 22)
        Me.txtReadOnly.TabIndex = 211
        Me.txtReadOnly.TabStop = False
        Me.txtReadOnly.Visible = False
        '
        'TxtPrimaryJournal
        '
        Me.TxtPrimaryJournal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryJournal.Location = New System.Drawing.Point(92, 32)
        Me.TxtPrimaryJournal.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryJournal.Name = "TxtPrimaryJournal"
        Me.TxtPrimaryJournal.Size = New System.Drawing.Size(45, 22)
        Me.TxtPrimaryJournal.TabIndex = 163
        Me.TxtPrimaryJournal.TabStop = False
        Me.TxtPrimaryJournal.Text = "-1"
        Me.TxtPrimaryJournal.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 10)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 17)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "No Journal"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(337, 7)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 17)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Date"
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(424, 2)
        Me.dtpDate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(137, 22)
        Me.dtpDate.TabIndex = 1
        '
        'TxtRefNo
        '
        Me.TxtRefNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRefNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRefNo.Location = New System.Drawing.Point(139, 4)
        Me.TxtRefNo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtRefNo.MaxLength = 50
        Me.TxtRefNo.Name = "TxtRefNo"
        Me.TxtRefNo.Size = New System.Drawing.Size(190, 22)
        Me.TxtRefNo.TabIndex = 0
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(139, 36)
        Me.TxtRemarks.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtRemarks.MaxLength = 150
        Me.TxtRemarks.Multiline = True
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(422, 98)
        Me.TxtRemarks.TabIndex = 2
        '
        'dgvExchangeRate
        '
        Me.dgvExchangeRate.AllowUserToAddRows = False
        Me.dgvExchangeRate.AllowUserToDeleteRows = False
        Me.dgvExchangeRate.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvExchangeRate.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItemExc_Currency, Me.ColItemExc_ExchangeRate, Me.ColItemExc_Remarks, Me.ColItemExc_PrimaryCurrency})
        Me.dgvExchangeRate.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvExchangeRate.Location = New System.Drawing.Point(0, 0)
        Me.dgvExchangeRate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvExchangeRate.Name = "dgvExchangeRate"
        Me.dgvExchangeRate.Size = New System.Drawing.Size(563, 138)
        Me.dgvExchangeRate.TabIndex = 0
        '
        'ColItemExc_Currency
        '
        Me.ColItemExc_Currency.HeaderText = "Currency"
        Me.ColItemExc_Currency.Name = "ColItemExc_Currency"
        Me.ColItemExc_Currency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItemExc_Currency.Width = 71
        '
        'ColItemExc_ExchangeRate
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItemExc_ExchangeRate.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColItemExc_ExchangeRate.HeaderText = "Exchange Rate"
        Me.ColItemExc_ExchangeRate.Name = "ColItemExc_ExchangeRate"
        Me.ColItemExc_ExchangeRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItemExc_ExchangeRate.Width = 110
        '
        'ColItemExc_Remarks
        '
        Me.ColItemExc_Remarks.HeaderText = "Remarks"
        Me.ColItemExc_Remarks.Name = "ColItemExc_Remarks"
        Me.ColItemExc_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItemExc_Remarks.Width = 70
        '
        'ColItemExc_PrimaryCurrency
        '
        Me.ColItemExc_PrimaryCurrency.HeaderText = "Primary Currency"
        Me.ColItemExc_PrimaryCurrency.Name = "ColItemExc_PrimaryCurrency"
        Me.ColItemExc_PrimaryCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItemExc_PrimaryCurrency.Visible = False
        Me.ColItemExc_PrimaryCurrency.Width = 92
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount, Me.ColDescAccount, Me.ColDebet, Me.ColKredit, Me.ColExchangeRate, Me.ColOriginalCurrency, Me.ColDebet_OriginalCurr, Me.ColCredit_OriginalCurrency, Me.ColRemarks, Me.ColPrimaryAkun, Me.ColPrimaryCurrency, Me.ColShowDetail})
        Me.dgvItem.Location = New System.Drawing.Point(8, 155)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(1134, 475)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 0
        '
        'ColAccount
        '
        Me.ColAccount.HeaderText = "Account"
        Me.ColAccount.MaxInputLength = 50
        Me.ColAccount.Name = "ColAccount"
        Me.ColAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccount.Width = 65
        '
        'ColDescAccount
        '
        Me.ColDescAccount.HeaderText = "Description"
        Me.ColDescAccount.MaxInputLength = 150
        Me.ColDescAccount.Name = "ColDescAccount"
        Me.ColDescAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDescAccount.Width = 85
        '
        'ColDebet
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.NullValue = "0"
        Me.ColDebet.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColDebet.HeaderText = "Debet"
        Me.ColDebet.MaxInputLength = 25
        Me.ColDebet.Name = "ColDebet"
        Me.ColDebet.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDebet.Width = 52
        '
        'ColKredit
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.NullValue = "0"
        Me.ColKredit.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColKredit.HeaderText = "Credit"
        Me.ColKredit.MaxInputLength = 25
        Me.ColKredit.Name = "ColKredit"
        Me.ColKredit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColKredit.Width = 51
        '
        'ColExchangeRate
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.NullValue = "1"
        Me.ColExchangeRate.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColExchangeRate.HeaderText = "Exchange Rate"
        Me.ColExchangeRate.Name = "ColExchangeRate"
        Me.ColExchangeRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColExchangeRate.Visible = False
        Me.ColExchangeRate.Width = 87
        '
        'ColOriginalCurrency
        '
        Me.ColOriginalCurrency.HeaderText = "Currency"
        Me.ColOriginalCurrency.Name = "ColOriginalCurrency"
        Me.ColOriginalCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColOriginalCurrency.Visible = False
        Me.ColOriginalCurrency.Width = 55
        '
        'ColDebet_OriginalCurr
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle6.NullValue = "0"
        Me.ColDebet_OriginalCurr.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColDebet_OriginalCurr.HeaderText = "Debet"
        Me.ColDebet_OriginalCurr.Name = "ColDebet_OriginalCurr"
        Me.ColDebet_OriginalCurr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDebet_OriginalCurr.Visible = False
        Me.ColDebet_OriginalCurr.Width = 42
        '
        'ColCredit_OriginalCurrency
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle7.NullValue = "0"
        Me.ColCredit_OriginalCurrency.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColCredit_OriginalCurrency.HeaderText = "Credit"
        Me.ColCredit_OriginalCurrency.Name = "ColCredit_OriginalCurrency"
        Me.ColCredit_OriginalCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCredit_OriginalCurrency.Visible = False
        Me.ColCredit_OriginalCurrency.Width = 40
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.MaxInputLength = 150
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 70
        '
        'ColPrimaryAkun
        '
        Me.ColPrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColPrimaryAkun.Name = "ColPrimaryAkun"
        Me.ColPrimaryAkun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryAkun.Visible = False
        Me.ColPrimaryAkun.Width = 72
        '
        'ColPrimaryCurrency
        '
        Me.ColPrimaryCurrency.HeaderText = "Primary Currency"
        Me.ColPrimaryCurrency.Name = "ColPrimaryCurrency"
        Me.ColPrimaryCurrency.Visible = False
        Me.ColPrimaryCurrency.Width = 111
        '
        'ColShowDetail
        '
        Me.ColShowDetail.HeaderText = "Detail"
        Me.ColShowDetail.Name = "ColShowDetail"
        Me.ColShowDetail.Width = 50
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(586, 638)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(139, 32)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(432, 638)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(139, 32)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(592, 692)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(107, 32)
        Me.btnDelete.TabIndex = 4
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(826, 692)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(107, 32)
        Me.btnExit.TabIndex = 6
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(474, 692)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(107, 32)
        Me.btnEdit.TabIndex = 3
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(240, 692)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(107, 32)
        Me.btnFind.TabIndex = 1
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(357, 692)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(107, 32)
        Me.btnNew.TabIndex = 2
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnViewJurnal
        '
        Me.btnViewJurnal.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnViewJurnal.Location = New System.Drawing.Point(8, 692)
        Me.btnViewJurnal.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnViewJurnal.Name = "btnViewJurnal"
        Me.btnViewJurnal.Size = New System.Drawing.Size(107, 32)
        Me.btnViewJurnal.TabIndex = 7
        Me.btnViewJurnal.TabStop = False
        Me.btnViewJurnal.Text = "View Journal"
        Me.btnViewJurnal.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(709, 692)
        Me.btnPrint.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(107, 32)
        Me.btnPrint.TabIndex = 5
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.ReccuringTab)
        Me.TabControl2.Controls.Add(Me.AssetDepreciationTab)
        Me.TabControl2.Controls.Add(Me.TabPage2)
        Me.TabControl2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl2.Location = New System.Drawing.Point(0, 0)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(1189, 790)
        Me.TabControl2.TabIndex = 40
        '
        'ReccuringTab
        '
        Me.ReccuringTab.Location = New System.Drawing.Point(4, 25)
        Me.ReccuringTab.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ReccuringTab.Name = "ReccuringTab"
        Me.ReccuringTab.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ReccuringTab.Size = New System.Drawing.Size(1181, 761)
        Me.ReccuringTab.TabIndex = 0
        Me.ReccuringTab.Text = "Recurring"
        Me.ReccuringTab.UseVisualStyleBackColor = True
        '
        'AssetDepreciationTab
        '
        Me.AssetDepreciationTab.Location = New System.Drawing.Point(4, 25)
        Me.AssetDepreciationTab.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.AssetDepreciationTab.Name = "AssetDepreciationTab"
        Me.AssetDepreciationTab.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.AssetDepreciationTab.Size = New System.Drawing.Size(1171, 761)
        Me.AssetDepreciationTab.TabIndex = 2
        Me.AssetDepreciationTab.Text = "Asset Depreciation"
        Me.AssetDepreciationTab.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.StatusStrip1)
        Me.TabPage2.Controls.Add(Me.btnPrint)
        Me.TabPage2.Controls.Add(Me.gboDetail)
        Me.TabPage2.Controls.Add(Me.btnViewJurnal)
        Me.TabPage2.Controls.Add(Me.btnNew)
        Me.TabPage2.Controls.Add(Me.btnFind)
        Me.TabPage2.Controls.Add(Me.btnEdit)
        Me.TabPage2.Controls.Add(Me.btnDelete)
        Me.TabPage2.Controls.Add(Me.btnExit)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TabPage2.Size = New System.Drawing.Size(1181, 761)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Revaluation"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(4, 732)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(1173, 25)
        Me.StatusStrip1.TabIndex = 40
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(288, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(288, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(288, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(288, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'FrmTransaksiJournal_EndOfMonth
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(1189, 790)
        Me.Controls.Add(Me.TabControl2)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1197, 826)
        Me.Name = "FrmTransaksiJournal_EndOfMonth"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "End Of Month Journal"
        Me.gboDetail.ResumeLayout(False)
        CType(Me.dgvItemAP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItemAR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvExchangeRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtRefNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents TxtUnbalance As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnViewJurnal As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryJournal As System.Windows.Forms.TextBox
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents LblQuickJurnal As System.Windows.Forms.Label
    Friend WithEvents llbEditQuickJournal As System.Windows.Forms.LinkLabel
    Friend WithEvents cboQuickJournal As System.Windows.Forms.ComboBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents ColAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDebet As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKredit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExchangeRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOriginalCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDebet_OriginalCurr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCredit_OriginalCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColShowDetail As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents dgvExchangeRate As System.Windows.Forms.DataGridView
    Friend WithEvents dgvItemAR As System.Windows.Forms.DataGridView
    Friend WithEvents ColItemAR_PrimaryTransaksiAsal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAR_TipeTransaksiAsal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAR_PrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAR_Saldo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAR_SaldoOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAR_ExchangeRateLama As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAR_PrimaryAkunAR As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAR_Selisih As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAR_Tanggal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAR_PrimaryPihakAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAR_PaymentTermAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemExc_Currency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemExc_ExchangeRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemExc_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemExc_PrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvItemAP As System.Windows.Forms.DataGridView
    Friend WithEvents ColItemAP_PrimaryTransaksiAsal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAP_TipeTransaksiAsal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAP_PrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAP_Saldo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAP_SaldoOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAP_ExchangeRateLama As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAP_PrimaryAkunAP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAP_Selisih As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAP_Tanggal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAP_PrimaryPihakAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAP_PaymentTermAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents ReccuringTab As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents AssetDepreciationTab As System.Windows.Forms.TabPage
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
End Class
