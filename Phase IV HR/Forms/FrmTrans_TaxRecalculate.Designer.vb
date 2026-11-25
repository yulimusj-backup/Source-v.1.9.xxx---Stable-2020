<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_TaxRecalculate
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
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.LblNo = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipePajak = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRecalculatedAmount = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColDifference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col21Pesangon = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col21TunjPPhPesangon = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col21PPhPesangon = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col21TarifPesagon = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col21NewTunjPesangon = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col21NewPPhPesangon = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col21NewTarifPesangon = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySalary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryEmployee = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col01Gaji = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col01GajiRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col02TunjPPh = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col02TunjPPHNonRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col02TunjPPhRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col02TunjPPhDisetahunkan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col02TunjPPhDisetahunkanRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col02NewTunjPPh = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col02NewTunjPPhNonRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col02NewTunjPPhRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col02NewTunjPPhDisetahunkan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col02NewTunjPPhDisetahunkanRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col03TunjLainnya = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col03TunjLainnyaRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col04Honorarium = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col04HonorariumRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col05PremiAsuransi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col05PremiAsuransiRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col06Natura = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col06NaturaRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col07Tantiem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col07TantiemRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col08PenghasilanBruto = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col08PenghasilanBrutoRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col08NewPenghasilanBruto = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col08NewPenghasilanBrutoRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col09BiayaJabatan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col09BiayaJabatanRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col09NewBiayaJabatan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col09NewBiayaJabatanRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col10IuranPensiun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col10IuranPensiunRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col11Jumlahengurang = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col11JumlahPengurangRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col11NewJumlahPengurang = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col11NewJumlahPengurangRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col14PenghasilanNettoUntPPH = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col14PenghasilanNettoPPhRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col14NewPenghasilanNettoUntPPh = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col14NewPenghasilanNettoUntPPhRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col15PTKP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col15NewPTKP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col16PKPSetahun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col16PKPSetahunRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col16NewPKPSetahun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col16NewPKPSetahunRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col17PPH21 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col17PPH21Rutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col17NewPPh21 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col17NewPPh21Rutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col18PPh21YgTelahDipotongMasaPjkSblmnya = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col18NewPPh21YgTelahDipotongMasaPjkSblmnya = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col19PPHTerutang = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col19PPHTerutangDisetahunakan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col19PPhTerutangDisetahunkanRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col19PPhTerutangSebulanRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col19PPHTerutangSebulanNonRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col19NewPPhTerutang = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col19NewPPhTerutangDisetahunkan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col19NewPPHTerutangDisetahunkanRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col19NewPPhTerutangSebulanRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col19NewPPhTerutangSebulanNonRutin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPenghasilBruto = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNewPenghasilanBruto = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNPWP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNewNPWP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlamat1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlamat2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNewAlamat1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNewAlamat2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatusNPWP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNewStatusNPWP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIncomeTax = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNewIncomeTax = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKurangBayar = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNewKurangBayar = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.btnDelete = New System.Windows.Forms.Button
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnNew = New System.Windows.Forms.Button
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.TxtDocumentNo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtDifference = New System.Windows.Forms.TextBox
        Me.btnReload = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtRemark = New System.Windows.Forms.TextBox
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SuspendLayout()
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(160, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(11, 3)
        Me.TxtReadOnly.MaxLength = 30
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(29, 20)
        Me.TxtReadOnly.TabIndex = 4
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'LblNo
        '
        Me.LblNo.AutoSize = True
        Me.LblNo.Location = New System.Drawing.Point(9, 3)
        Me.LblNo.Name = "LblNo"
        Me.LblNo.Size = New System.Drawing.Size(73, 13)
        Me.LblNo.TabIndex = 45
        Me.LblNo.Text = "Document No"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColDocNo, Me.ColDate, Me.ColEmployee, Me.ColTipePajak, Me.ColAmount, Me.ColRecalculatedAmount, Me.ColDifference, Me.Col21Pesangon, Me.Col21TunjPPhPesangon, Me.Col21PPhPesangon, Me.Col21TarifPesagon, Me.Col21NewTunjPesangon, Me.Col21NewPPhPesangon, Me.Col21NewTarifPesangon, Me.ColPrimarySalary, Me.ColPrimaryEmployee, Me.Col01Gaji, Me.Col01GajiRutin, Me.Col02TunjPPh, Me.Col02TunjPPHNonRutin, Me.Col02TunjPPhRutin, Me.Col02TunjPPhDisetahunkan, Me.Col02TunjPPhDisetahunkanRutin, Me.Col02NewTunjPPh, Me.Col02NewTunjPPhNonRutin, Me.Col02NewTunjPPhRutin, Me.Col02NewTunjPPhDisetahunkan, Me.Col02NewTunjPPhDisetahunkanRutin, Me.Col03TunjLainnya, Me.Col03TunjLainnyaRutin, Me.Col04Honorarium, Me.Col04HonorariumRutin, Me.Col05PremiAsuransi, Me.Col05PremiAsuransiRutin, Me.Col06Natura, Me.Col06NaturaRutin, Me.Col07Tantiem, Me.Col07TantiemRutin, Me.Col08PenghasilanBruto, Me.Col08PenghasilanBrutoRutin, Me.Col08NewPenghasilanBruto, Me.Col08NewPenghasilanBrutoRutin, Me.Col09BiayaJabatan, Me.Col09BiayaJabatanRutin, Me.Col09NewBiayaJabatan, Me.Col09NewBiayaJabatanRutin, Me.Col10IuranPensiun, Me.Col10IuranPensiunRutin, Me.Col11Jumlahengurang, Me.Col11JumlahPengurangRutin, Me.Col11NewJumlahPengurang, Me.Col11NewJumlahPengurangRutin, Me.Col14PenghasilanNettoUntPPH, Me.Col14PenghasilanNettoPPhRutin, Me.Col14NewPenghasilanNettoUntPPh, Me.Col14NewPenghasilanNettoUntPPhRutin, Me.Col15PTKP, Me.Col15NewPTKP, Me.Col16PKPSetahun, Me.Col16PKPSetahunRutin, Me.Col16NewPKPSetahun, Me.Col16NewPKPSetahunRutin, Me.Col17PPH21, Me.Col17PPH21Rutin, Me.Col17NewPPh21, Me.Col17NewPPh21Rutin, Me.Col18PPh21YgTelahDipotongMasaPjkSblmnya, Me.Col18NewPPh21YgTelahDipotongMasaPjkSblmnya, Me.Col19PPHTerutang, Me.Col19PPHTerutangDisetahunakan, Me.Col19PPhTerutangDisetahunkanRutin, Me.Col19PPhTerutangSebulanRutin, Me.Col19PPHTerutangSebulanNonRutin, Me.Col19NewPPhTerutang, Me.Col19NewPPhTerutangDisetahunkan, Me.Col19NewPPHTerutangDisetahunkanRutin, Me.Col19NewPPhTerutangSebulanRutin, Me.Col19NewPPhTerutangSebulanNonRutin, Me.ColPenghasilBruto, Me.ColNewPenghasilanBruto, Me.ColNPWP, Me.ColNewNPWP, Me.ColAlamat1, Me.ColAlamat2, Me.ColNewAlamat1, Me.ColNewAlamat2, Me.ColStatusNPWP, Me.ColNewStatusNPWP, Me.ColIncomeTax, Me.ColNewIncomeTax, Me.ColKurangBayar, Me.ColNewKurangBayar})
        Me.dgvItem.Location = New System.Drawing.Point(12, 64)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(623, 252)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 1
        Me.dgvItem.TabStop = False
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 46
        '
        'ColDocNo
        '
        Me.ColDocNo.HeaderText = "Document No"
        Me.ColDocNo.Name = "ColDocNo"
        Me.ColDocNo.ReadOnly = True
        Me.ColDocNo.Width = 98
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.Width = 55
        '
        'ColEmployee
        '
        Me.ColEmployee.HeaderText = "Employee"
        Me.ColEmployee.Name = "ColEmployee"
        Me.ColEmployee.ReadOnly = True
        Me.ColEmployee.Width = 78
        '
        'ColTipePajak
        '
        Me.ColTipePajak.HeaderText = "Tax Calculation"
        Me.ColTipePajak.Name = "ColTipePajak"
        Me.ColTipePajak.ReadOnly = True
        Me.ColTipePajak.Visible = False
        Me.ColTipePajak.Width = 141
        '
        'ColAmount
        '
        Me.ColAmount.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColAmount.HeaderText = "Tax Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.ReadOnly = True
        Me.ColAmount.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColAmount.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColAmount.Width = 89
        '
        'ColRecalculatedAmount
        '
        Me.ColRecalculatedAmount.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRecalculatedAmount.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColRecalculatedAmount.HeaderText = "Recalculated Tax Amount"
        Me.ColRecalculatedAmount.Name = "ColRecalculatedAmount"
        Me.ColRecalculatedAmount.ReadOnly = True
        Me.ColRecalculatedAmount.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRecalculatedAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRecalculatedAmount.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRecalculatedAmount.Width = 155
        '
        'ColDifference
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDifference.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColDifference.HeaderText = "Difference"
        Me.ColDifference.Name = "ColDifference"
        Me.ColDifference.ReadOnly = True
        Me.ColDifference.Width = 81
        '
        'Col21Pesangon
        '
        Me.Col21Pesangon.HeaderText = "21 Pesangon"
        Me.Col21Pesangon.Name = "Col21Pesangon"
        Me.Col21Pesangon.ReadOnly = True
        Me.Col21Pesangon.Visible = False
        Me.Col21Pesangon.Width = 128
        '
        'Col21TunjPPhPesangon
        '
        Me.Col21TunjPPhPesangon.HeaderText = "21 Tunj PPh Pesangon"
        Me.Col21TunjPPhPesangon.Name = "Col21TunjPPhPesangon"
        Me.Col21TunjPPhPesangon.ReadOnly = True
        Me.Col21TunjPPhPesangon.Visible = False
        Me.Col21TunjPPhPesangon.Width = 195
        '
        'Col21PPhPesangon
        '
        Me.Col21PPhPesangon.HeaderText = "21 PPh Pesangon"
        Me.Col21PPhPesangon.Name = "Col21PPhPesangon"
        Me.Col21PPhPesangon.ReadOnly = True
        Me.Col21PPhPesangon.Visible = False
        Me.Col21PPhPesangon.Width = 161
        '
        'Col21TarifPesagon
        '
        Me.Col21TarifPesagon.HeaderText = "21 Tarif Pesangon"
        Me.Col21TarifPesagon.Name = "Col21TarifPesagon"
        Me.Col21TarifPesagon.ReadOnly = True
        Me.Col21TarifPesagon.Visible = False
        Me.Col21TarifPesagon.Width = 163
        '
        'Col21NewTunjPesangon
        '
        Me.Col21NewTunjPesangon.HeaderText = "21 New Tunj Pesangon"
        Me.Col21NewTunjPesangon.Name = "Col21NewTunjPesangon"
        Me.Col21NewTunjPesangon.ReadOnly = True
        Me.Col21NewTunjPesangon.Visible = False
        Me.Col21NewTunjPesangon.Width = 197
        '
        'Col21NewPPhPesangon
        '
        Me.Col21NewPPhPesangon.HeaderText = "21 New PPh Pesangon"
        Me.Col21NewPPhPesangon.Name = "Col21NewPPhPesangon"
        Me.Col21NewPPhPesangon.ReadOnly = True
        Me.Col21NewPPhPesangon.Visible = False
        Me.Col21NewPPhPesangon.Width = 196
        '
        'Col21NewTarifPesangon
        '
        Me.Col21NewTarifPesangon.HeaderText = "21 New Tarif Pesangon"
        Me.Col21NewTarifPesangon.Name = "Col21NewTarifPesangon"
        Me.Col21NewTarifPesangon.ReadOnly = True
        Me.Col21NewTarifPesangon.Visible = False
        Me.Col21NewTarifPesangon.Width = 198
        '
        'ColPrimarySalary
        '
        Me.ColPrimarySalary.HeaderText = "Primary Salary"
        Me.ColPrimarySalary.Name = "ColPrimarySalary"
        Me.ColPrimarySalary.ReadOnly = True
        Me.ColPrimarySalary.Visible = False
        Me.ColPrimarySalary.Width = 134
        '
        'ColPrimaryEmployee
        '
        Me.ColPrimaryEmployee.HeaderText = "Primary Employee"
        Me.ColPrimaryEmployee.Name = "ColPrimaryEmployee"
        Me.ColPrimaryEmployee.ReadOnly = True
        Me.ColPrimaryEmployee.Visible = False
        Me.ColPrimaryEmployee.Width = 160
        '
        'Col01Gaji
        '
        Me.Col01Gaji.HeaderText = "01 Gaji"
        Me.Col01Gaji.Name = "Col01Gaji"
        Me.Col01Gaji.ReadOnly = True
        Me.Col01Gaji.Visible = False
        Me.Col01Gaji.Width = 84
        '
        'Col01GajiRutin
        '
        Me.Col01GajiRutin.HeaderText = "01 Gaji Rutin"
        Me.Col01GajiRutin.Name = "Col01GajiRutin"
        Me.Col01GajiRutin.ReadOnly = True
        Me.Col01GajiRutin.Visible = False
        Me.Col01GajiRutin.Width = 126
        '
        'Col02TunjPPh
        '
        Me.Col02TunjPPh.HeaderText = "02 Tunj PPH"
        Me.Col02TunjPPh.Name = "Col02TunjPPh"
        Me.Col02TunjPPh.ReadOnly = True
        Me.Col02TunjPPh.Visible = False
        Me.Col02TunjPPh.Width = 122
        '
        'Col02TunjPPHNonRutin
        '
        Me.Col02TunjPPHNonRutin.HeaderText = "02 Tunj PPh Non Rutin"
        Me.Col02TunjPPHNonRutin.Name = "Col02TunjPPHNonRutin"
        Me.Col02TunjPPHNonRutin.ReadOnly = True
        Me.Col02TunjPPHNonRutin.Visible = False
        Me.Col02TunjPPHNonRutin.Width = 194
        '
        'Col02TunjPPhRutin
        '
        Me.Col02TunjPPhRutin.HeaderText = "02 Tunj PPh Rutin"
        Me.Col02TunjPPhRutin.Name = "Col02TunjPPhRutin"
        Me.Col02TunjPPhRutin.ReadOnly = True
        Me.Col02TunjPPhRutin.Visible = False
        Me.Col02TunjPPhRutin.Width = 161
        '
        'Col02TunjPPhDisetahunkan
        '
        Me.Col02TunjPPhDisetahunkan.HeaderText = "02 Tunj PPh Disetahunkan"
        Me.Col02TunjPPhDisetahunkan.Name = "Col02TunjPPhDisetahunkan"
        Me.Col02TunjPPhDisetahunkan.ReadOnly = True
        Me.Col02TunjPPhDisetahunkan.Visible = False
        Me.Col02TunjPPhDisetahunkan.Width = 222
        '
        'Col02TunjPPhDisetahunkanRutin
        '
        Me.Col02TunjPPhDisetahunkanRutin.HeaderText = "02 Tunj PPh Disetahunkan Rutin"
        Me.Col02TunjPPhDisetahunkanRutin.Name = "Col02TunjPPhDisetahunkanRutin"
        Me.Col02TunjPPhDisetahunkanRutin.ReadOnly = True
        Me.Col02TunjPPhDisetahunkanRutin.Visible = False
        Me.Col02TunjPPhDisetahunkanRutin.Width = 264
        '
        'Col02NewTunjPPh
        '
        Me.Col02NewTunjPPh.HeaderText = "02 New Tunj PPh"
        Me.Col02NewTunjPPh.Name = "Col02NewTunjPPh"
        Me.Col02NewTunjPPh.ReadOnly = True
        Me.Col02NewTunjPPh.Visible = False
        Me.Col02NewTunjPPh.Width = 154
        '
        'Col02NewTunjPPhNonRutin
        '
        Me.Col02NewTunjPPhNonRutin.HeaderText = "02 New Tunj PPh Non Rutin"
        Me.Col02NewTunjPPhNonRutin.Name = "Col02NewTunjPPhNonRutin"
        Me.Col02NewTunjPPhNonRutin.ReadOnly = True
        Me.Col02NewTunjPPhNonRutin.Visible = False
        Me.Col02NewTunjPPhNonRutin.Width = 229
        '
        'Col02NewTunjPPhRutin
        '
        Me.Col02NewTunjPPhRutin.HeaderText = "02 New Tunj PPh Rutin"
        Me.Col02NewTunjPPhRutin.Name = "Col02NewTunjPPhRutin"
        Me.Col02NewTunjPPhRutin.ReadOnly = True
        Me.Col02NewTunjPPhRutin.Visible = False
        Me.Col02NewTunjPPhRutin.Width = 196
        '
        'Col02NewTunjPPhDisetahunkan
        '
        Me.Col02NewTunjPPhDisetahunkan.HeaderText = "02 New Tunj PPh Disetahunkan"
        Me.Col02NewTunjPPhDisetahunkan.Name = "Col02NewTunjPPhDisetahunkan"
        Me.Col02NewTunjPPhDisetahunkan.ReadOnly = True
        Me.Col02NewTunjPPhDisetahunkan.Visible = False
        Me.Col02NewTunjPPhDisetahunkan.Width = 257
        '
        'Col02NewTunjPPhDisetahunkanRutin
        '
        Me.Col02NewTunjPPhDisetahunkanRutin.HeaderText = "02 New Tunj PPh Disetahunkan Rutin"
        Me.Col02NewTunjPPhDisetahunkanRutin.Name = "Col02NewTunjPPhDisetahunkanRutin"
        Me.Col02NewTunjPPhDisetahunkanRutin.ReadOnly = True
        Me.Col02NewTunjPPhDisetahunkanRutin.Visible = False
        Me.Col02NewTunjPPhDisetahunkanRutin.Width = 299
        '
        'Col03TunjLainnya
        '
        Me.Col03TunjLainnya.HeaderText = "03 Tunj Lainny"
        Me.Col03TunjLainnya.Name = "Col03TunjLainnya"
        Me.Col03TunjLainnya.ReadOnly = True
        Me.Col03TunjLainnya.Visible = False
        Me.Col03TunjLainnya.Width = 136
        '
        'Col03TunjLainnyaRutin
        '
        Me.Col03TunjLainnyaRutin.HeaderText = "03 Tunj Lainnya Rutin"
        Me.Col03TunjLainnyaRutin.Name = "Col03TunjLainnyaRutin"
        Me.Col03TunjLainnyaRutin.ReadOnly = True
        Me.Col03TunjLainnyaRutin.Visible = False
        Me.Col03TunjLainnyaRutin.Width = 187
        '
        'Col04Honorarium
        '
        Me.Col04Honorarium.HeaderText = "04 Honorarium"
        Me.Col04Honorarium.Name = "Col04Honorarium"
        Me.Col04Honorarium.ReadOnly = True
        Me.Col04Honorarium.Visible = False
        Me.Col04Honorarium.Width = 139
        '
        'Col04HonorariumRutin
        '
        Me.Col04HonorariumRutin.HeaderText = "04 Honorarium Rutin"
        Me.Col04HonorariumRutin.Name = "Col04HonorariumRutin"
        Me.Col04HonorariumRutin.ReadOnly = True
        Me.Col04HonorariumRutin.Visible = False
        Me.Col04HonorariumRutin.Width = 181
        '
        'Col05PremiAsuransi
        '
        Me.Col05PremiAsuransi.HeaderText = "05 Premi Asuransi"
        Me.Col05PremiAsuransi.Name = "Col05PremiAsuransi"
        Me.Col05PremiAsuransi.ReadOnly = True
        Me.Col05PremiAsuransi.Visible = False
        Me.Col05PremiAsuransi.Width = 162
        '
        'Col05PremiAsuransiRutin
        '
        Me.Col05PremiAsuransiRutin.HeaderText = "05 Premi Asuransi Rutin"
        Me.Col05PremiAsuransiRutin.Name = "Col05PremiAsuransiRutin"
        Me.Col05PremiAsuransiRutin.ReadOnly = True
        Me.Col05PremiAsuransiRutin.Visible = False
        Me.Col05PremiAsuransiRutin.Width = 204
        '
        'Col06Natura
        '
        Me.Col06Natura.HeaderText = "06 Natura"
        Me.Col06Natura.Name = "Col06Natura"
        Me.Col06Natura.ReadOnly = True
        Me.Col06Natura.Visible = False
        Me.Col06Natura.Width = 104
        '
        'Col06NaturaRutin
        '
        Me.Col06NaturaRutin.HeaderText = "06 Natura Rutin"
        Me.Col06NaturaRutin.Name = "Col06NaturaRutin"
        Me.Col06NaturaRutin.ReadOnly = True
        Me.Col06NaturaRutin.Visible = False
        Me.Col06NaturaRutin.Width = 146
        '
        'Col07Tantiem
        '
        Me.Col07Tantiem.HeaderText = "07 Tantiem"
        Me.Col07Tantiem.Name = "Col07Tantiem"
        Me.Col07Tantiem.ReadOnly = True
        Me.Col07Tantiem.Visible = False
        Me.Col07Tantiem.Width = 113
        '
        'Col07TantiemRutin
        '
        Me.Col07TantiemRutin.HeaderText = "07 Tantiem Rutin"
        Me.Col07TantiemRutin.Name = "Col07TantiemRutin"
        Me.Col07TantiemRutin.ReadOnly = True
        Me.Col07TantiemRutin.Visible = False
        Me.Col07TantiemRutin.Width = 155
        '
        'Col08PenghasilanBruto
        '
        Me.Col08PenghasilanBruto.HeaderText = "08 Penghasilan Bruto"
        Me.Col08PenghasilanBruto.Name = "Col08PenghasilanBruto"
        Me.Col08PenghasilanBruto.ReadOnly = True
        Me.Col08PenghasilanBruto.Visible = False
        Me.Col08PenghasilanBruto.Width = 186
        '
        'Col08PenghasilanBrutoRutin
        '
        Me.Col08PenghasilanBrutoRutin.HeaderText = "08 Penghasilan Bruto Rutin"
        Me.Col08PenghasilanBrutoRutin.Name = "Col08PenghasilanBrutoRutin"
        Me.Col08PenghasilanBrutoRutin.ReadOnly = True
        Me.Col08PenghasilanBrutoRutin.Visible = False
        Me.Col08PenghasilanBrutoRutin.Width = 228
        '
        'Col08NewPenghasilanBruto
        '
        Me.Col08NewPenghasilanBruto.HeaderText = "08 New Penghasilan Bruto"
        Me.Col08NewPenghasilanBruto.Name = "Col08NewPenghasilanBruto"
        Me.Col08NewPenghasilanBruto.ReadOnly = True
        Me.Col08NewPenghasilanBruto.Visible = False
        Me.Col08NewPenghasilanBruto.Width = 221
        '
        'Col08NewPenghasilanBrutoRutin
        '
        Me.Col08NewPenghasilanBrutoRutin.HeaderText = "08 New Penghasilan Bruto Rutin"
        Me.Col08NewPenghasilanBrutoRutin.Name = "Col08NewPenghasilanBrutoRutin"
        Me.Col08NewPenghasilanBrutoRutin.ReadOnly = True
        Me.Col08NewPenghasilanBrutoRutin.Visible = False
        Me.Col08NewPenghasilanBrutoRutin.Width = 263
        '
        'Col09BiayaJabatan
        '
        Me.Col09BiayaJabatan.HeaderText = "09 Biaya Jabatan"
        Me.Col09BiayaJabatan.Name = "Col09BiayaJabatan"
        Me.Col09BiayaJabatan.ReadOnly = True
        Me.Col09BiayaJabatan.Visible = False
        Me.Col09BiayaJabatan.Width = 157
        '
        'Col09BiayaJabatanRutin
        '
        Me.Col09BiayaJabatanRutin.HeaderText = "09 Biaya Jabatan Rutin"
        Me.Col09BiayaJabatanRutin.Name = "Col09BiayaJabatanRutin"
        Me.Col09BiayaJabatanRutin.ReadOnly = True
        Me.Col09BiayaJabatanRutin.Visible = False
        Me.Col09BiayaJabatanRutin.Width = 199
        '
        'Col09NewBiayaJabatan
        '
        Me.Col09NewBiayaJabatan.HeaderText = "09 New Biaya Jabatan"
        Me.Col09NewBiayaJabatan.Name = "Col09NewBiayaJabatan"
        Me.Col09NewBiayaJabatan.ReadOnly = True
        Me.Col09NewBiayaJabatan.Visible = False
        Me.Col09NewBiayaJabatan.Width = 192
        '
        'Col09NewBiayaJabatanRutin
        '
        Me.Col09NewBiayaJabatanRutin.HeaderText = "09 New Biaya Jabatan Rutin"
        Me.Col09NewBiayaJabatanRutin.Name = "Col09NewBiayaJabatanRutin"
        Me.Col09NewBiayaJabatanRutin.ReadOnly = True
        Me.Col09NewBiayaJabatanRutin.Visible = False
        Me.Col09NewBiayaJabatanRutin.Width = 234
        '
        'Col10IuranPensiun
        '
        Me.Col10IuranPensiun.HeaderText = "10 Iuran Pensiun"
        Me.Col10IuranPensiun.Name = "Col10IuranPensiun"
        Me.Col10IuranPensiun.ReadOnly = True
        Me.Col10IuranPensiun.Visible = False
        Me.Col10IuranPensiun.Width = 154
        '
        'Col10IuranPensiunRutin
        '
        Me.Col10IuranPensiunRutin.HeaderText = "10 Iuran Pensiun Rutin"
        Me.Col10IuranPensiunRutin.Name = "Col10IuranPensiunRutin"
        Me.Col10IuranPensiunRutin.ReadOnly = True
        Me.Col10IuranPensiunRutin.Visible = False
        Me.Col10IuranPensiunRutin.Width = 196
        '
        'Col11Jumlahengurang
        '
        Me.Col11Jumlahengurang.HeaderText = "11 Jumlah Pengurang"
        Me.Col11Jumlahengurang.Name = "Col11Jumlahengurang"
        Me.Col11Jumlahengurang.ReadOnly = True
        Me.Col11Jumlahengurang.Visible = False
        Me.Col11Jumlahengurang.Width = 189
        '
        'Col11JumlahPengurangRutin
        '
        Me.Col11JumlahPengurangRutin.HeaderText = "11 Jumlah Pengurang Rutin"
        Me.Col11JumlahPengurangRutin.Name = "Col11JumlahPengurangRutin"
        Me.Col11JumlahPengurangRutin.ReadOnly = True
        Me.Col11JumlahPengurangRutin.Visible = False
        Me.Col11JumlahPengurangRutin.Width = 231
        '
        'Col11NewJumlahPengurang
        '
        Me.Col11NewJumlahPengurang.HeaderText = "11 New Jumlah Pengurang"
        Me.Col11NewJumlahPengurang.Name = "Col11NewJumlahPengurang"
        Me.Col11NewJumlahPengurang.ReadOnly = True
        Me.Col11NewJumlahPengurang.Visible = False
        Me.Col11NewJumlahPengurang.Width = 224
        '
        'Col11NewJumlahPengurangRutin
        '
        Me.Col11NewJumlahPengurangRutin.HeaderText = "11 New Jumlah Pengurang Rutin"
        Me.Col11NewJumlahPengurangRutin.Name = "Col11NewJumlahPengurangRutin"
        Me.Col11NewJumlahPengurangRutin.ReadOnly = True
        Me.Col11NewJumlahPengurangRutin.Visible = False
        Me.Col11NewJumlahPengurangRutin.Width = 266
        '
        'Col14PenghasilanNettoUntPPH
        '
        Me.Col14PenghasilanNettoUntPPH.HeaderText = "14 Penghasilan Netto Unt PPh"
        Me.Col14PenghasilanNettoUntPPH.Name = "Col14PenghasilanNettoUntPPH"
        Me.Col14PenghasilanNettoUntPPH.ReadOnly = True
        Me.Col14PenghasilanNettoUntPPH.Visible = False
        Me.Col14PenghasilanNettoUntPPH.Width = 249
        '
        'Col14PenghasilanNettoPPhRutin
        '
        Me.Col14PenghasilanNettoPPhRutin.HeaderText = "14 Penghasilan Nettot Untuk PPH Rutin"
        Me.Col14PenghasilanNettoPPhRutin.Name = "Col14PenghasilanNettoPPhRutin"
        Me.Col14PenghasilanNettoPPhRutin.ReadOnly = True
        Me.Col14PenghasilanNettoPPhRutin.Visible = False
        Me.Col14PenghasilanNettoPPhRutin.Width = 316
        '
        'Col14NewPenghasilanNettoUntPPh
        '
        Me.Col14NewPenghasilanNettoUntPPh.HeaderText = "14 New Penghasilan Netto Unt PPh"
        Me.Col14NewPenghasilanNettoUntPPh.Name = "Col14NewPenghasilanNettoUntPPh"
        Me.Col14NewPenghasilanNettoUntPPh.ReadOnly = True
        Me.Col14NewPenghasilanNettoUntPPh.Visible = False
        Me.Col14NewPenghasilanNettoUntPPh.Width = 284
        '
        'Col14NewPenghasilanNettoUntPPhRutin
        '
        Me.Col14NewPenghasilanNettoUntPPhRutin.HeaderText = "14 New Penghasilan Netto Unt PPh Rutin"
        Me.Col14NewPenghasilanNettoUntPPhRutin.Name = "Col14NewPenghasilanNettoUntPPhRutin"
        Me.Col14NewPenghasilanNettoUntPPhRutin.ReadOnly = True
        Me.Col14NewPenghasilanNettoUntPPhRutin.Visible = False
        Me.Col14NewPenghasilanNettoUntPPhRutin.Width = 326
        '
        'Col15PTKP
        '
        Me.Col15PTKP.HeaderText = "15 PTKP"
        Me.Col15PTKP.Name = "Col15PTKP"
        Me.Col15PTKP.ReadOnly = True
        Me.Col15PTKP.Visible = False
        Me.Col15PTKP.Width = 95
        '
        'Col15NewPTKP
        '
        Me.Col15NewPTKP.HeaderText = "15 New PTKP"
        Me.Col15NewPTKP.Name = "Col15NewPTKP"
        Me.Col15NewPTKP.ReadOnly = True
        Me.Col15NewPTKP.Visible = False
        Me.Col15NewPTKP.Width = 130
        '
        'Col16PKPSetahun
        '
        Me.Col16PKPSetahun.HeaderText = "16 PKP Setahun"
        Me.Col16PKPSetahun.Name = "Col16PKPSetahun"
        Me.Col16PKPSetahun.ReadOnly = True
        Me.Col16PKPSetahun.Visible = False
        Me.Col16PKPSetahun.Width = 151
        '
        'Col16PKPSetahunRutin
        '
        Me.Col16PKPSetahunRutin.HeaderText = "16 PKP Setahun Rutin"
        Me.Col16PKPSetahunRutin.Name = "Col16PKPSetahunRutin"
        Me.Col16PKPSetahunRutin.ReadOnly = True
        Me.Col16PKPSetahunRutin.Visible = False
        Me.Col16PKPSetahunRutin.Width = 193
        '
        'Col16NewPKPSetahun
        '
        Me.Col16NewPKPSetahun.HeaderText = "16 New PKP Setahun"
        Me.Col16NewPKPSetahun.Name = "Col16NewPKPSetahun"
        Me.Col16NewPKPSetahun.ReadOnly = True
        Me.Col16NewPKPSetahun.Visible = False
        Me.Col16NewPKPSetahun.Width = 186
        '
        'Col16NewPKPSetahunRutin
        '
        Me.Col16NewPKPSetahunRutin.HeaderText = "16 New PKP Setahun Rutin"
        Me.Col16NewPKPSetahunRutin.Name = "Col16NewPKPSetahunRutin"
        Me.Col16NewPKPSetahunRutin.ReadOnly = True
        Me.Col16NewPKPSetahunRutin.Visible = False
        Me.Col16NewPKPSetahunRutin.Width = 228
        '
        'Col17PPH21
        '
        Me.Col17PPH21.HeaderText = "17 PPh 21"
        Me.Col17PPH21.Name = "Col17PPH21"
        Me.Col17PPH21.ReadOnly = True
        Me.Col17PPH21.Visible = False
        Me.Col17PPH21.Width = 107
        '
        'Col17PPH21Rutin
        '
        Me.Col17PPH21Rutin.HeaderText = "17 PPh 21 Rutin"
        Me.Col17PPH21Rutin.Name = "Col17PPH21Rutin"
        Me.Col17PPH21Rutin.ReadOnly = True
        Me.Col17PPH21Rutin.Visible = False
        Me.Col17PPH21Rutin.Width = 149
        '
        'Col17NewPPh21
        '
        Me.Col17NewPPh21.HeaderText = "17 New PPh 21"
        Me.Col17NewPPh21.Name = "Col17NewPPh21"
        Me.Col17NewPPh21.ReadOnly = True
        Me.Col17NewPPh21.Visible = False
        Me.Col17NewPPh21.Width = 142
        '
        'Col17NewPPh21Rutin
        '
        Me.Col17NewPPh21Rutin.HeaderText = "17 New PPh 21 Rutin"
        Me.Col17NewPPh21Rutin.Name = "Col17NewPPh21Rutin"
        Me.Col17NewPPh21Rutin.ReadOnly = True
        Me.Col17NewPPh21Rutin.Visible = False
        Me.Col17NewPPh21Rutin.Width = 184
        '
        'Col18PPh21YgTelahDipotongMasaPjkSblmnya
        '
        Me.Col18PPh21YgTelahDipotongMasaPjkSblmnya.HeaderText = "18 PPh21YgTelahDipotongMasaPjkSblmnya"
        Me.Col18PPh21YgTelahDipotongMasaPjkSblmnya.Name = "Col18PPh21YgTelahDipotongMasaPjkSblmnya"
        Me.Col18PPh21YgTelahDipotongMasaPjkSblmnya.ReadOnly = True
        Me.Col18PPh21YgTelahDipotongMasaPjkSblmnya.Visible = False
        Me.Col18PPh21YgTelahDipotongMasaPjkSblmnya.Width = 348
        '
        'Col18NewPPh21YgTelahDipotongMasaPjkSblmnya
        '
        Me.Col18NewPPh21YgTelahDipotongMasaPjkSblmnya.HeaderText = "18 New PPh21YgTelahDipotongMasaPjkSblmnya"
        Me.Col18NewPPh21YgTelahDipotongMasaPjkSblmnya.Name = "Col18NewPPh21YgTelahDipotongMasaPjkSblmnya"
        Me.Col18NewPPh21YgTelahDipotongMasaPjkSblmnya.ReadOnly = True
        Me.Col18NewPPh21YgTelahDipotongMasaPjkSblmnya.Visible = False
        Me.Col18NewPPh21YgTelahDipotongMasaPjkSblmnya.Width = 383
        '
        'Col19PPHTerutang
        '
        Me.Col19PPHTerutang.HeaderText = "19 PPh Terutang"
        Me.Col19PPHTerutang.Name = "Col19PPHTerutang"
        Me.Col19PPHTerutang.ReadOnly = True
        Me.Col19PPHTerutang.Visible = False
        Me.Col19PPHTerutang.Width = 153
        '
        'Col19PPHTerutangDisetahunakan
        '
        Me.Col19PPHTerutangDisetahunakan.HeaderText = "19 PPh Terutang Disetahunkan"
        Me.Col19PPHTerutangDisetahunakan.Name = "Col19PPHTerutangDisetahunakan"
        Me.Col19PPHTerutangDisetahunakan.ReadOnly = True
        Me.Col19PPHTerutangDisetahunakan.Visible = False
        Me.Col19PPHTerutangDisetahunakan.Width = 256
        '
        'Col19PPhTerutangDisetahunkanRutin
        '
        Me.Col19PPhTerutangDisetahunkanRutin.HeaderText = "19 PPh Terutang Disetahunkan Rutin"
        Me.Col19PPhTerutangDisetahunkanRutin.Name = "Col19PPhTerutangDisetahunkanRutin"
        Me.Col19PPhTerutangDisetahunkanRutin.ReadOnly = True
        Me.Col19PPhTerutangDisetahunkanRutin.Visible = False
        Me.Col19PPhTerutangDisetahunkanRutin.Width = 298
        '
        'Col19PPhTerutangSebulanRutin
        '
        Me.Col19PPhTerutangSebulanRutin.HeaderText = "19 PPh Terutang Sebulan Rutin"
        Me.Col19PPhTerutangSebulanRutin.Name = "Col19PPhTerutangSebulanRutin"
        Me.Col19PPhTerutangSebulanRutin.ReadOnly = True
        Me.Col19PPhTerutangSebulanRutin.Visible = False
        Me.Col19PPhTerutangSebulanRutin.Width = 258
        '
        'Col19PPHTerutangSebulanNonRutin
        '
        Me.Col19PPHTerutangSebulanNonRutin.HeaderText = "19 PPh Terutang Sebulan Non Rutin"
        Me.Col19PPHTerutangSebulanNonRutin.Name = "Col19PPHTerutangSebulanNonRutin"
        Me.Col19PPHTerutangSebulanNonRutin.ReadOnly = True
        Me.Col19PPHTerutangSebulanNonRutin.Visible = False
        Me.Col19PPHTerutangSebulanNonRutin.Width = 291
        '
        'Col19NewPPhTerutang
        '
        Me.Col19NewPPhTerutang.HeaderText = "19 New PPh Terutang"
        Me.Col19NewPPhTerutang.Name = "Col19NewPPhTerutang"
        Me.Col19NewPPhTerutang.ReadOnly = True
        Me.Col19NewPPhTerutang.Visible = False
        Me.Col19NewPPhTerutang.Width = 188
        '
        'Col19NewPPhTerutangDisetahunkan
        '
        Me.Col19NewPPhTerutangDisetahunkan.HeaderText = "19 New PPh Terutang Disetahunkan"
        Me.Col19NewPPhTerutangDisetahunkan.Name = "Col19NewPPhTerutangDisetahunkan"
        Me.Col19NewPPhTerutangDisetahunkan.ReadOnly = True
        Me.Col19NewPPhTerutangDisetahunkan.Visible = False
        Me.Col19NewPPhTerutangDisetahunkan.Width = 291
        '
        'Col19NewPPHTerutangDisetahunkanRutin
        '
        Me.Col19NewPPHTerutangDisetahunkanRutin.HeaderText = "19 New PPh Terutang Disetahunkan Rutin"
        Me.Col19NewPPHTerutangDisetahunkanRutin.Name = "Col19NewPPHTerutangDisetahunkanRutin"
        Me.Col19NewPPHTerutangDisetahunkanRutin.ReadOnly = True
        Me.Col19NewPPHTerutangDisetahunkanRutin.Visible = False
        Me.Col19NewPPHTerutangDisetahunkanRutin.Width = 333
        '
        'Col19NewPPhTerutangSebulanRutin
        '
        Me.Col19NewPPhTerutangSebulanRutin.HeaderText = "19 New PPh Sebulan Rutin"
        Me.Col19NewPPhTerutangSebulanRutin.Name = "Col19NewPPhTerutangSebulanRutin"
        Me.Col19NewPPhTerutangSebulanRutin.ReadOnly = True
        Me.Col19NewPPhTerutangSebulanRutin.Visible = False
        Me.Col19NewPPhTerutangSebulanRutin.Width = 225
        '
        'Col19NewPPhTerutangSebulanNonRutin
        '
        Me.Col19NewPPhTerutangSebulanNonRutin.HeaderText = "19 New PPh Terutang Sebulan Non Rutin"
        Me.Col19NewPPhTerutangSebulanNonRutin.Name = "Col19NewPPhTerutangSebulanNonRutin"
        Me.Col19NewPPhTerutangSebulanNonRutin.ReadOnly = True
        Me.Col19NewPPhTerutangSebulanNonRutin.Visible = False
        Me.Col19NewPPhTerutangSebulanNonRutin.Width = 326
        '
        'ColPenghasilBruto
        '
        Me.ColPenghasilBruto.HeaderText = "Penghasilan Bruto"
        Me.ColPenghasilBruto.Name = "ColPenghasilBruto"
        Me.ColPenghasilBruto.ReadOnly = True
        Me.ColPenghasilBruto.Visible = False
        Me.ColPenghasilBruto.Width = 164
        '
        'ColNewPenghasilanBruto
        '
        Me.ColNewPenghasilanBruto.HeaderText = "New Penghasilan Bruto"
        Me.ColNewPenghasilanBruto.Name = "ColNewPenghasilanBruto"
        Me.ColNewPenghasilanBruto.ReadOnly = True
        Me.ColNewPenghasilanBruto.Visible = False
        Me.ColNewPenghasilanBruto.Width = 199
        '
        'ColNPWP
        '
        Me.ColNPWP.HeaderText = "NPWP"
        Me.ColNPWP.Name = "ColNPWP"
        Me.ColNPWP.ReadOnly = True
        Me.ColNPWP.Visible = False
        Me.ColNPWP.Width = 80
        '
        'ColNewNPWP
        '
        Me.ColNewNPWP.HeaderText = "New NPWP"
        Me.ColNewNPWP.Name = "ColNewNPWP"
        Me.ColNewNPWP.ReadOnly = True
        Me.ColNewNPWP.Visible = False
        Me.ColNewNPWP.Width = 115
        '
        'ColAlamat1
        '
        Me.ColAlamat1.HeaderText = "Alamat 1"
        Me.ColAlamat1.Name = "ColAlamat1"
        Me.ColAlamat1.ReadOnly = True
        Me.ColAlamat1.Visible = False
        Me.ColAlamat1.Width = 97
        '
        'ColAlamat2
        '
        Me.ColAlamat2.HeaderText = "Alamat 2"
        Me.ColAlamat2.Name = "ColAlamat2"
        Me.ColAlamat2.ReadOnly = True
        Me.ColAlamat2.Visible = False
        Me.ColAlamat2.Width = 97
        '
        'ColNewAlamat1
        '
        Me.ColNewAlamat1.HeaderText = "New Alamat 1"
        Me.ColNewAlamat1.Name = "ColNewAlamat1"
        Me.ColNewAlamat1.ReadOnly = True
        Me.ColNewAlamat1.Visible = False
        Me.ColNewAlamat1.Width = 132
        '
        'ColNewAlamat2
        '
        Me.ColNewAlamat2.HeaderText = "New Alamat 2"
        Me.ColNewAlamat2.Name = "ColNewAlamat2"
        Me.ColNewAlamat2.ReadOnly = True
        Me.ColNewAlamat2.Visible = False
        Me.ColNewAlamat2.Width = 132
        '
        'ColStatusNPWP
        '
        Me.ColStatusNPWP.HeaderText = "Status NPWP"
        Me.ColStatusNPWP.Name = "ColStatusNPWP"
        Me.ColStatusNPWP.ReadOnly = True
        Me.ColStatusNPWP.Visible = False
        Me.ColStatusNPWP.Width = 131
        '
        'ColNewStatusNPWP
        '
        Me.ColNewStatusNPWP.HeaderText = "New Status NPWP"
        Me.ColNewStatusNPWP.Name = "ColNewStatusNPWP"
        Me.ColNewStatusNPWP.ReadOnly = True
        Me.ColNewStatusNPWP.Visible = False
        Me.ColNewStatusNPWP.Width = 166
        '
        'ColIncomeTax
        '
        Me.ColIncomeTax.HeaderText = "Income Tax"
        Me.ColIncomeTax.Name = "ColIncomeTax"
        Me.ColIncomeTax.ReadOnly = True
        Me.ColIncomeTax.Visible = False
        Me.ColIncomeTax.Width = 116
        '
        'ColNewIncomeTax
        '
        Me.ColNewIncomeTax.HeaderText = "New Income Tax"
        Me.ColNewIncomeTax.Name = "ColNewIncomeTax"
        Me.ColNewIncomeTax.ReadOnly = True
        Me.ColNewIncomeTax.Visible = False
        Me.ColNewIncomeTax.Width = 151
        '
        'ColKurangBayar
        '
        Me.ColKurangBayar.HeaderText = "Kurang Bayar"
        Me.ColKurangBayar.Name = "ColKurangBayar"
        Me.ColKurangBayar.ReadOnly = True
        Me.ColKurangBayar.Visible = False
        Me.ColKurangBayar.Width = 130
        '
        'ColNewKurangBayar
        '
        Me.ColNewKurangBayar.HeaderText = "New Kurang Bayar"
        Me.ColNewKurangBayar.Name = "ColNewKurangBayar"
        Me.ColNewKurangBayar.ReadOnly = True
        Me.ColNewKurangBayar.Visible = False
        Me.ColNewKurangBayar.Width = 165
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(92, 3)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(40, 20)
        Me.TxtPrimary.TabIndex = 47
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(198, 3)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(111, 20)
        Me.dtpDate.TabIndex = 1
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(373, 359)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(160, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(210, 320)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(288, 359)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(327, 320)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(117, 359)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(160, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(203, 359)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(160, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(459, 359)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 389)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(656, 22)
        Me.StatusStrip1.TabIndex = 22
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.SplitContainer2)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(7, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(641, 353)
        Me.gboDetail.TabIndex = 6
        Me.gboDetail.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(12, 12)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.dtpPeriod)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtReadOnly)
        Me.SplitContainer2.Panel1.Controls.Add(Me.LblNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtPrimary)
        Me.SplitContainer2.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtDocumentNo)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtDifference)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnReload)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRemark)
        Me.SplitContainer2.Size = New System.Drawing.Size(623, 49)
        Me.SplitContainer2.SplitterDistance = 312
        Me.SplitContainer2.TabIndex = 0
        Me.SplitContainer2.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 49
        Me.Label1.Text = "Period"
        '
        'dtpPeriod
        '
        Me.dtpPeriod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpPeriod.CustomFormat = "MMM yyyy"
        Me.dtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPeriod.Location = New System.Drawing.Point(92, 27)
        Me.dtpPeriod.Name = "dtpPeriod"
        Me.dtpPeriod.ShowUpDown = True
        Me.dtpPeriod.Size = New System.Drawing.Size(217, 20)
        Me.dtpPeriod.TabIndex = 2
        '
        'TxtDocumentNo
        '
        Me.TxtDocumentNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDocumentNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDocumentNo.Location = New System.Drawing.Point(92, 3)
        Me.TxtDocumentNo.Name = "TxtDocumentNo"
        Me.TxtDocumentNo.ReadOnly = True
        Me.TxtDocumentNo.Size = New System.Drawing.Size(101, 20)
        Me.TxtDocumentNo.TabIndex = 0
        Me.TxtDocumentNo.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 30)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 68
        Me.Label3.Text = "Remarks"
        '
        'TxtDifference
        '
        Me.TxtDifference.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDifference.Location = New System.Drawing.Point(77, 5)
        Me.TxtDifference.MaxLength = 150
        Me.TxtDifference.Name = "TxtDifference"
        Me.TxtDifference.ReadOnly = True
        Me.TxtDifference.Size = New System.Drawing.Size(147, 20)
        Me.TxtDifference.TabIndex = 0
        Me.TxtDifference.TabStop = False
        Me.TxtDifference.Text = "0"
        Me.TxtDifference.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnReload
        '
        Me.btnReload.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnReload.Location = New System.Drawing.Point(228, 3)
        Me.btnReload.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnReload.Name = "btnReload"
        Me.btnReload.Size = New System.Drawing.Size(74, 19)
        Me.btnReload.TabIndex = 2
        Me.btnReload.Text = "Recalculate"
        Me.btnReload.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 6)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 13)
        Me.Label6.TabIndex = 65
        Me.Label6.Text = "Difference"
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(77, 27)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(227, 20)
        Me.TxtRemark.TabIndex = 1
        '
        'FrmTrans_TaxRecalculate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(656, 411)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.gboDetail)
        Me.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.MinimumSize = New System.Drawing.Size(672, 449)
        Me.Name = "FrmTrans_TaxRecalculate"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Recalculate Tax"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents LblNo As System.Windows.Forms.Label
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Public WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents dtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtDocumentNo As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents btnReload As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents TxtDifference As System.Windows.Forms.TextBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipePajak As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRecalculatedAmount As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColDifference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col21Pesangon As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col21TunjPPhPesangon As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col21PPhPesangon As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col21TarifPesagon As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col21NewTunjPesangon As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col21NewPPhPesangon As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col21NewTarifPesangon As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySalary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryEmployee As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col01Gaji As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col01GajiRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col02TunjPPh As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col02TunjPPHNonRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col02TunjPPhRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col02TunjPPhDisetahunkan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col02TunjPPhDisetahunkanRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col02NewTunjPPh As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col02NewTunjPPhNonRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col02NewTunjPPhRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col02NewTunjPPhDisetahunkan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col02NewTunjPPhDisetahunkanRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col03TunjLainnya As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col03TunjLainnyaRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col04Honorarium As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col04HonorariumRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col05PremiAsuransi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col05PremiAsuransiRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col06Natura As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col06NaturaRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col07Tantiem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col07TantiemRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col08PenghasilanBruto As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col08PenghasilanBrutoRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col08NewPenghasilanBruto As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col08NewPenghasilanBrutoRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col09BiayaJabatan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col09BiayaJabatanRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col09NewBiayaJabatan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col09NewBiayaJabatanRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col10IuranPensiun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col10IuranPensiunRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col11Jumlahengurang As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col11JumlahPengurangRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col11NewJumlahPengurang As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col11NewJumlahPengurangRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col14PenghasilanNettoUntPPH As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col14PenghasilanNettoPPhRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col14NewPenghasilanNettoUntPPh As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col14NewPenghasilanNettoUntPPhRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col15PTKP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col15NewPTKP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col16PKPSetahun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col16PKPSetahunRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col16NewPKPSetahun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col16NewPKPSetahunRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col17PPH21 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col17PPH21Rutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col17NewPPh21 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col17NewPPh21Rutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col18PPh21YgTelahDipotongMasaPjkSblmnya As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col18NewPPh21YgTelahDipotongMasaPjkSblmnya As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col19PPHTerutang As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col19PPHTerutangDisetahunakan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col19PPhTerutangDisetahunkanRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col19PPhTerutangSebulanRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col19PPHTerutangSebulanNonRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col19NewPPhTerutang As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col19NewPPhTerutangDisetahunkan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col19NewPPHTerutangDisetahunkanRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col19NewPPhTerutangSebulanRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col19NewPPhTerutangSebulanNonRutin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPenghasilBruto As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNewPenghasilanBruto As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNPWP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNewNPWP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlamat1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlamat2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNewAlamat1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNewAlamat2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatusNPWP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNewStatusNPWP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIncomeTax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNewIncomeTax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKurangBayar As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNewKurangBayar As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
