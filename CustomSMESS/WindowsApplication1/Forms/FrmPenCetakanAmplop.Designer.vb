<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmPenCetakanAmplop
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
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboFormatLaporan = New System.Windows.Forms.ComboBox
        Me.rboSupplier = New System.Windows.Forms.RadioButton
        Me.rboCustomer = New System.Windows.Forms.RadioButton
        Me.TxtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.btnFindSupplier = New System.Windows.Forms.Button
        Me.TxtSupplierCode = New System.Windows.Forms.TextBox
        Me.LblSupp = New System.Windows.Forms.Label
        Me.BtnFindCustomer = New System.Windows.Forms.Button
        Me.TxtCustomerCode = New System.Windows.Forms.TextBox
        Me.LblCust = New System.Windows.Forms.Label
        Me.TxtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.CBOOffice = New System.Windows.Forms.ComboBox
        Me.CBOcp = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtIdentifierOffice = New System.Windows.Forms.TextBox
        Me.TxtIdentifierCP = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.BtnPrint = New System.Windows.Forms.Button
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.CRV_Biasa_Potrait = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.CRV_Biasa_Landscape = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.CboPrint = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.CRV_Matrix_Potrait = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.CRV_Matrix_Landscape = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.Location = New System.Drawing.Point(12, 119)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowPrintButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(800, 391)
        Me.crv1.TabIndex = 10
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(407, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 62
        Me.Label3.Text = "Format"
        '
        'cboFormatLaporan
        '
        Me.cboFormatLaporan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormatLaporan.FormattingEnabled = True
        Me.cboFormatLaporan.Items.AddRange(New Object() {"Potrait", "Landscape"})
        Me.cboFormatLaporan.Location = New System.Drawing.Point(452, 32)
        Me.cboFormatLaporan.Name = "cboFormatLaporan"
        Me.cboFormatLaporan.Size = New System.Drawing.Size(276, 21)
        Me.cboFormatLaporan.TabIndex = 7
        '
        'rboSupplier
        '
        Me.rboSupplier.AutoSize = True
        Me.rboSupplier.Location = New System.Drawing.Point(112, 12)
        Me.rboSupplier.Name = "rboSupplier"
        Me.rboSupplier.Size = New System.Drawing.Size(63, 17)
        Me.rboSupplier.TabIndex = 1
        Me.rboSupplier.Text = "Supplier"
        Me.rboSupplier.UseVisualStyleBackColor = True
        '
        'rboCustomer
        '
        Me.rboCustomer.AutoSize = True
        Me.rboCustomer.Checked = True
        Me.rboCustomer.Location = New System.Drawing.Point(12, 12)
        Me.rboCustomer.Name = "rboCustomer"
        Me.rboCustomer.Size = New System.Drawing.Size(69, 17)
        Me.rboCustomer.TabIndex = 0
        Me.rboCustomer.TabStop = True
        Me.rboCustomer.Text = "Customer"
        Me.rboCustomer.UseVisualStyleBackColor = True
        '
        'TxtPrimarySupplier
        '
        Me.TxtPrimarySupplier.Location = New System.Drawing.Point(114, 35)
        Me.TxtPrimarySupplier.Name = "TxtPrimarySupplier"
        Me.TxtPrimarySupplier.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimarySupplier.TabIndex = 68
        Me.TxtPrimarySupplier.TabStop = False
        Me.TxtPrimarySupplier.Text = "-1"
        Me.TxtPrimarySupplier.Visible = False
        '
        'btnFindSupplier
        '
        Me.btnFindSupplier.Location = New System.Drawing.Point(352, 32)
        Me.btnFindSupplier.Name = "btnFindSupplier"
        Me.btnFindSupplier.Size = New System.Drawing.Size(38, 23)
        Me.btnFindSupplier.TabIndex = 66
        Me.btnFindSupplier.TabStop = False
        Me.btnFindSupplier.Text = "Find"
        Me.btnFindSupplier.UseVisualStyleBackColor = True
        '
        'TxtSupplierCode
        '
        Me.TxtSupplierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSupplierCode.Location = New System.Drawing.Point(114, 35)
        Me.TxtSupplierCode.Name = "TxtSupplierCode"
        Me.TxtSupplierCode.Size = New System.Drawing.Size(232, 20)
        Me.TxtSupplierCode.TabIndex = 2
        '
        'LblSupp
        '
        Me.LblSupp.AutoSize = True
        Me.LblSupp.Location = New System.Drawing.Point(12, 38)
        Me.LblSupp.Name = "LblSupp"
        Me.LblSupp.Size = New System.Drawing.Size(73, 13)
        Me.LblSupp.TabIndex = 67
        Me.LblSupp.Text = "Supplier Code"
        '
        'BtnFindCustomer
        '
        Me.BtnFindCustomer.Location = New System.Drawing.Point(352, 32)
        Me.BtnFindCustomer.Name = "BtnFindCustomer"
        Me.BtnFindCustomer.Size = New System.Drawing.Size(38, 23)
        Me.BtnFindCustomer.TabIndex = 3
        Me.BtnFindCustomer.TabStop = False
        Me.BtnFindCustomer.Text = "Find"
        Me.BtnFindCustomer.UseVisualStyleBackColor = True
        '
        'TxtCustomerCode
        '
        Me.TxtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerCode.Location = New System.Drawing.Point(114, 34)
        Me.TxtCustomerCode.Name = "TxtCustomerCode"
        Me.TxtCustomerCode.Size = New System.Drawing.Size(232, 20)
        Me.TxtCustomerCode.TabIndex = 2
        '
        'LblCust
        '
        Me.LblCust.AutoSize = True
        Me.LblCust.Location = New System.Drawing.Point(12, 40)
        Me.LblCust.Name = "LblCust"
        Me.LblCust.Size = New System.Drawing.Size(79, 13)
        Me.LblCust.TabIndex = 71
        Me.LblCust.Text = "Customer Code"
        '
        'TxtPrimaryCustomer
        '
        Me.TxtPrimaryCustomer.Location = New System.Drawing.Point(114, 33)
        Me.TxtPrimaryCustomer.Name = "TxtPrimaryCustomer"
        Me.TxtPrimaryCustomer.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimaryCustomer.TabIndex = 72
        Me.TxtPrimaryCustomer.TabStop = False
        Me.TxtPrimaryCustomer.Text = "-1"
        Me.TxtPrimaryCustomer.Visible = False
        '
        'CBOOffice
        '
        Me.CBOOffice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBOOffice.FormattingEnabled = True
        Me.CBOOffice.Location = New System.Drawing.Point(114, 61)
        Me.CBOOffice.Name = "CBOOffice"
        Me.CBOOffice.Size = New System.Drawing.Size(276, 21)
        Me.CBOOffice.TabIndex = 4
        '
        'CBOcp
        '
        Me.CBOcp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBOcp.FormattingEnabled = True
        Me.CBOcp.Location = New System.Drawing.Point(114, 88)
        Me.CBOcp.Name = "CBOcp"
        Me.CBOcp.Size = New System.Drawing.Size(276, 21)
        Me.CBOcp.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 13)
        Me.Label4.TabIndex = 75
        Me.Label4.Text = "Address"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(13, 91)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 13)
        Me.Label5.TabIndex = 76
        Me.Label5.Text = "Contact Person"
        '
        'TxtIdentifierOffice
        '
        Me.TxtIdentifierOffice.Location = New System.Drawing.Point(82, 61)
        Me.TxtIdentifierOffice.Name = "TxtIdentifierOffice"
        Me.TxtIdentifierOffice.Size = New System.Drawing.Size(26, 20)
        Me.TxtIdentifierOffice.TabIndex = 77
        Me.TxtIdentifierOffice.TabStop = False
        Me.TxtIdentifierOffice.Text = "-1"
        Me.TxtIdentifierOffice.Visible = False
        '
        'TxtIdentifierCP
        '
        Me.TxtIdentifierCP.Location = New System.Drawing.Point(82, 87)
        Me.TxtIdentifierCP.Name = "TxtIdentifierCP"
        Me.TxtIdentifierCP.Size = New System.Drawing.Size(26, 20)
        Me.TxtIdentifierCP.TabIndex = 78
        Me.TxtIdentifierCP.TabStop = False
        Me.TxtIdentifierCP.Text = "-1"
        Me.TxtIdentifierCP.Visible = False
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(734, 90)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(734, 32)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(75, 23)
        Me.btnFilter.TabIndex = 8
        Me.btnFilter.Text = "Preview"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'BtnPrint
        '
        Me.BtnPrint.Location = New System.Drawing.Point(734, 61)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(75, 23)
        Me.BtnPrint.TabIndex = 9
        Me.BtnPrint.Text = "Print"
        Me.BtnPrint.UseVisualStyleBackColor = True
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(181, 9)
        Me.TxtReadOnly.MaxLength = 150
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(50, 20)
        Me.TxtReadOnly.TabIndex = 79
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'CRV_Biasa_Potrait
        '
        Me.CRV_Biasa_Potrait.ActiveViewIndex = -1
        Me.CRV_Biasa_Potrait.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CRV_Biasa_Potrait.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CRV_Biasa_Potrait.DisplayGroupTree = False
        Me.CRV_Biasa_Potrait.Location = New System.Drawing.Point(15, 169)
        Me.CRV_Biasa_Potrait.Name = "CRV_Biasa_Potrait"
        Me.CRV_Biasa_Potrait.SelectionFormula = ""
        Me.CRV_Biasa_Potrait.ShowCloseButton = False
        Me.CRV_Biasa_Potrait.ShowGroupTreeButton = False
        Me.CRV_Biasa_Potrait.ShowPrintButton = False
        Me.CRV_Biasa_Potrait.ShowRefreshButton = False
        Me.CRV_Biasa_Potrait.Size = New System.Drawing.Size(794, 341)
        Me.CRV_Biasa_Potrait.TabIndex = 80
        Me.CRV_Biasa_Potrait.TabStop = False
        Me.CRV_Biasa_Potrait.ViewTimeSelectionFormula = ""
        '
        'CRV_Biasa_Landscape
        '
        Me.CRV_Biasa_Landscape.ActiveViewIndex = -1
        Me.CRV_Biasa_Landscape.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CRV_Biasa_Landscape.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CRV_Biasa_Landscape.DisplayGroupTree = False
        Me.CRV_Biasa_Landscape.Location = New System.Drawing.Point(12, 141)
        Me.CRV_Biasa_Landscape.Name = "CRV_Biasa_Landscape"
        Me.CRV_Biasa_Landscape.SelectionFormula = ""
        Me.CRV_Biasa_Landscape.ShowCloseButton = False
        Me.CRV_Biasa_Landscape.ShowGroupTreeButton = False
        Me.CRV_Biasa_Landscape.ShowPrintButton = False
        Me.CRV_Biasa_Landscape.ShowRefreshButton = False
        Me.CRV_Biasa_Landscape.Size = New System.Drawing.Size(797, 369)
        Me.CRV_Biasa_Landscape.TabIndex = 81
        Me.CRV_Biasa_Landscape.TabStop = False
        Me.CRV_Biasa_Landscape.ViewTimeSelectionFormula = ""
        '
        'CboPrint
        '
        Me.CboPrint.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPrint.FormattingEnabled = True
        Me.CboPrint.Items.AddRange(New Object() {"Non Dot Matrix", "Dot Matrix"})
        Me.CboPrint.Location = New System.Drawing.Point(452, 32)
        Me.CboPrint.Name = "CboPrint"
        Me.CboPrint.Size = New System.Drawing.Size(276, 21)
        Me.CboPrint.TabIndex = 6
        Me.CboPrint.TabStop = False
        Me.CboPrint.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(407, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 83
        Me.Label1.Text = "Printer"
        Me.Label1.Visible = False
        '
        'CRV_Matrix_Potrait
        '
        Me.CRV_Matrix_Potrait.ActiveViewIndex = -1
        Me.CRV_Matrix_Potrait.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CRV_Matrix_Potrait.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CRV_Matrix_Potrait.DisplayGroupTree = False
        Me.CRV_Matrix_Potrait.Location = New System.Drawing.Point(14, 119)
        Me.CRV_Matrix_Potrait.Name = "CRV_Matrix_Potrait"
        Me.CRV_Matrix_Potrait.SelectionFormula = ""
        Me.CRV_Matrix_Potrait.ShowCloseButton = False
        Me.CRV_Matrix_Potrait.ShowGroupTreeButton = False
        Me.CRV_Matrix_Potrait.ShowPrintButton = False
        Me.CRV_Matrix_Potrait.ShowRefreshButton = False
        Me.CRV_Matrix_Potrait.Size = New System.Drawing.Size(797, 295)
        Me.CRV_Matrix_Potrait.TabIndex = 84
        Me.CRV_Matrix_Potrait.TabStop = False
        Me.CRV_Matrix_Potrait.ViewTimeSelectionFormula = ""
        '
        'CRV_Matrix_Landscape
        '
        Me.CRV_Matrix_Landscape.ActiveViewIndex = -1
        Me.CRV_Matrix_Landscape.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CRV_Matrix_Landscape.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CRV_Matrix_Landscape.DisplayGroupTree = False
        Me.CRV_Matrix_Landscape.Location = New System.Drawing.Point(12, 141)
        Me.CRV_Matrix_Landscape.Name = "CRV_Matrix_Landscape"
        Me.CRV_Matrix_Landscape.SelectionFormula = ""
        Me.CRV_Matrix_Landscape.ShowCloseButton = False
        Me.CRV_Matrix_Landscape.ShowGroupTreeButton = False
        Me.CRV_Matrix_Landscape.ShowPrintButton = False
        Me.CRV_Matrix_Landscape.ShowRefreshButton = False
        Me.CRV_Matrix_Landscape.Size = New System.Drawing.Size(797, 306)
        Me.CRV_Matrix_Landscape.TabIndex = 85
        Me.CRV_Matrix_Landscape.TabStop = False
        Me.CRV_Matrix_Landscape.ViewTimeSelectionFormula = ""
        '
        'FrmPenCetakanAmplop
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(824, 522)
        Me.Controls.Add(Me.crv1)
        Me.Controls.Add(Me.CRV_Matrix_Landscape)
        Me.Controls.Add(Me.CRV_Matrix_Potrait)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CboPrint)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.BtnPrint)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.TxtIdentifierCP)
        Me.Controls.Add(Me.TxtIdentifierOffice)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.CBOcp)
        Me.Controls.Add(Me.CBOOffice)
        Me.Controls.Add(Me.TxtPrimaryCustomer)
        Me.Controls.Add(Me.BtnFindCustomer)
        Me.Controls.Add(Me.TxtCustomerCode)
        Me.Controls.Add(Me.LblCust)
        Me.Controls.Add(Me.TxtPrimarySupplier)
        Me.Controls.Add(Me.btnFindSupplier)
        Me.Controls.Add(Me.TxtSupplierCode)
        Me.Controls.Add(Me.LblSupp)
        Me.Controls.Add(Me.rboSupplier)
        Me.Controls.Add(Me.rboCustomer)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboFormatLaporan)
        Me.Controls.Add(Me.CRV_Biasa_Potrait)
        Me.Controls.Add(Me.CRV_Biasa_Landscape)
        Me.MinimumSize = New System.Drawing.Size(840, 560)
        Me.Name = "FrmPenCetakanAmplop"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Envelope Printing"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboFormatLaporan As System.Windows.Forms.ComboBox
    Friend WithEvents rboSupplier As System.Windows.Forms.RadioButton
    Friend WithEvents rboCustomer As System.Windows.Forms.RadioButton
    Friend WithEvents TxtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents btnFindSupplier As System.Windows.Forms.Button
    Friend WithEvents TxtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents LblSupp As System.Windows.Forms.Label
    Friend WithEvents BtnFindCustomer As System.Windows.Forms.Button
    Friend WithEvents TxtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents LblCust As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents CBOOffice As System.Windows.Forms.ComboBox
    Friend WithEvents CBOcp As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtIdentifierOffice As System.Windows.Forms.TextBox
    Friend WithEvents TxtIdentifierCP As System.Windows.Forms.TextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents CRV_Biasa_Potrait As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents CRV_Biasa_Landscape As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents CboPrint As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CRV_Matrix_Potrait As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents CRV_Matrix_Landscape As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
