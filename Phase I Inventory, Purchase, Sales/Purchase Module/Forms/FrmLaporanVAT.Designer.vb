<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanVAT
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
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.btnCreate = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColTipeTransaksiString = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppCust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNPWPNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvoiceNumber = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColTaxInvoiceNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurchaseSales = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColVAT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDocAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEnumType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEnumInOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.btnCreateDataImpor = New System.Windows.Forms.Button
        Me.btnImportData = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(216, 38)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(144, 20)
        Me.dtpTo.TabIndex = 2
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(216, 12)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(144, 20)
        Me.dtpFrom.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(180, 41)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "To"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(180, 15)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "From"
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.DisplayStatusBar = False
        Me.crv1.DisplayToolbar = False
        Me.crv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crv1.Location = New System.Drawing.Point(3, 3)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(746, 354)
        Me.crv1.TabIndex = 6
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(366, 36)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 23)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(366, 12)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(76, 23)
        Me.btnShow.TabIndex = 3
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"VAT In", "VAT OUT", "Both"})
        Me.ComboBox1.Location = New System.Drawing.Point(12, 14)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(162, 21)
        Me.ComboBox1.TabIndex = 0
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(448, 12)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(109, 23)
        Me.btnCreate.TabIndex = 5
        Me.btnCreate.Text = "Create csv data"
        Me.btnCreate.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 64)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(760, 386)
        Me.TabControl1.TabIndex = 25
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvItem)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(752, 360)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Data"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColTipeTransaksiString, Me.ColSuppCust, Me.ColNPWPNumber, Me.ColInvoiceNumber, Me.ColTaxInvoiceNumber, Me.ColDate, Me.ColPurchaseSales, Me.ColVAT, Me.ColPrimaryDocAcuan, Me.ColEnumType, Me.ColEnumInOut})
        Me.dgvItem.Location = New System.Drawing.Point(3, 3)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.Size = New System.Drawing.Size(746, 354)
        Me.dgvItem.TabIndex = 0
        '
        'ColTipeTransaksiString
        '
        Me.ColTipeTransaksiString.HeaderText = "Type"
        Me.ColTipeTransaksiString.Name = "ColTipeTransaksiString"
        Me.ColTipeTransaksiString.ReadOnly = True
        Me.ColTipeTransaksiString.Width = 56
        '
        'ColSuppCust
        '
        Me.ColSuppCust.HeaderText = "Supplier/Customer"
        Me.ColSuppCust.Name = "ColSuppCust"
        Me.ColSuppCust.ReadOnly = True
        Me.ColSuppCust.Width = 119
        '
        'ColNPWPNumber
        '
        Me.ColNPWPNumber.HeaderText = "NPWP"
        Me.ColNPWPNumber.Name = "ColNPWPNumber"
        Me.ColNPWPNumber.ReadOnly = True
        Me.ColNPWPNumber.Width = 65
        '
        'ColInvoiceNumber
        '
        Me.ColInvoiceNumber.HeaderText = "Invoice Number"
        Me.ColInvoiceNumber.Name = "ColInvoiceNumber"
        Me.ColInvoiceNumber.ReadOnly = True
        Me.ColInvoiceNumber.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColInvoiceNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColInvoiceNumber.Width = 107
        '
        'ColTaxInvoiceNumber
        '
        Me.ColTaxInvoiceNumber.HeaderText = "Tax Invoice Number"
        Me.ColTaxInvoiceNumber.Name = "ColTaxInvoiceNumber"
        Me.ColTaxInvoiceNumber.ReadOnly = True
        Me.ColTaxInvoiceNumber.Width = 128
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.Width = 55
        '
        'ColPurchaseSales
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPurchaseSales.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColPurchaseSales.HeaderText = "Purchase/Sales"
        Me.ColPurchaseSales.Name = "ColPurchaseSales"
        Me.ColPurchaseSales.ReadOnly = True
        Me.ColPurchaseSales.Width = 108
        '
        'ColVAT
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColVAT.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColVAT.HeaderText = "VAT"
        Me.ColVAT.Name = "ColVAT"
        Me.ColVAT.ReadOnly = True
        Me.ColVAT.Width = 53
        '
        'ColPrimaryDocAcuan
        '
        Me.ColPrimaryDocAcuan.HeaderText = "Primary Doc Acuan"
        Me.ColPrimaryDocAcuan.Name = "ColPrimaryDocAcuan"
        Me.ColPrimaryDocAcuan.ReadOnly = True
        Me.ColPrimaryDocAcuan.Visible = False
        Me.ColPrimaryDocAcuan.Width = 123
        '
        'ColEnumType
        '
        Me.ColEnumType.HeaderText = "Enum Type"
        Me.ColEnumType.Name = "ColEnumType"
        Me.ColEnumType.ReadOnly = True
        Me.ColEnumType.Visible = False
        Me.ColEnumType.Width = 86
        '
        'ColEnumInOut
        '
        Me.ColEnumInOut.HeaderText = "Enum In Out"
        Me.ColEnumInOut.Name = "ColEnumInOut"
        Me.ColEnumInOut.ReadOnly = True
        Me.ColEnumInOut.Visible = False
        Me.ColEnumInOut.Width = 91
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.crv1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(752, 360)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Printable Version"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "csv"
        Me.SaveFileDialog1.Filter = "csv files|*.csv"
        Me.SaveFileDialog1.Title = "Save e-SPT data"
        '
        'btnCreateDataImpor
        '
        Me.btnCreateDataImpor.Location = New System.Drawing.Point(448, 36)
        Me.btnCreateDataImpor.Name = "btnCreateDataImpor"
        Me.btnCreateDataImpor.Size = New System.Drawing.Size(109, 23)
        Me.btnCreateDataImpor.TabIndex = 26
        Me.btnCreateDataImpor.TabStop = False
        Me.btnCreateDataImpor.Text = "Create Data Import"
        Me.btnCreateDataImpor.UseVisualStyleBackColor = True
        Me.btnCreateDataImpor.Visible = False
        '
        'btnImportData
        '
        Me.btnImportData.Location = New System.Drawing.Point(563, 36)
        Me.btnImportData.Name = "btnImportData"
        Me.btnImportData.Size = New System.Drawing.Size(109, 23)
        Me.btnImportData.TabIndex = 27
        Me.btnImportData.TabStop = False
        Me.btnImportData.Text = "Import Data"
        Me.btnImportData.UseVisualStyleBackColor = True
        Me.btnImportData.Visible = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "csv files|*.csv"
        Me.OpenFileDialog1.Title = "Import Tax Invoice Number"
        '
        'FrmLaporanVAT
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(784, 462)
        Me.Controls.Add(Me.btnImportData)
        Me.Controls.Add(Me.btnCreateDataImpor)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnCreate)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnShow)
        Me.MinimumSize = New System.Drawing.Size(800, 500)
        Me.Name = "FrmLaporanVAT"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VAT Report"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents btnCreate As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ColTipeTransaksiString As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppCust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNPWPNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvoiceNumber As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColTaxInvoiceNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurchaseSales As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVAT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDocAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEnumType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEnumInOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnCreateDataImpor As System.Windows.Forms.Button
    Friend WithEvents btnImportData As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
End Class
