<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPurchaseDP
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.RboAOP = New System.Windows.Forms.RadioButton
        Me.Label7 = New System.Windows.Forms.Label
        Me.DtpAOPDateTo = New System.Windows.Forms.DateTimePicker
        Me.DtpAOPDateFrom = New System.Windows.Forms.DateTimePicker
        Me.rboAOPDP = New System.Windows.Forms.RadioButton
        Me.rboPODate = New System.Windows.Forms.RadioButton
        Me.TxtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtSupplierCode = New System.Windows.Forms.TextBox
        Me.TxtSupplierName = New System.Windows.Forms.TextBox
        Me.btnFindSupplier = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpAOPDPTo = New System.Windows.Forms.DateTimePicker
        Me.dtpAOPDPFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpPOTo = New System.Windows.Forms.DateTimePicker
        Me.dtpPOFrom = New System.Windows.Forms.DateTimePicker
        Me.ChkUse = New System.Windows.Forms.CheckBox
        Me.BtnFilter = New System.Windows.Forms.Button
        Me.cboRealizationStatus = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cboPaymentStatus = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboDPStatus = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(10, 11)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.RboAOP)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpAOPDateTo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpAOPDateFrom)
        Me.SplitContainer1.Panel1.Controls.Add(Me.rboAOPDP)
        Me.SplitContainer1.Panel1.Controls.Add(Me.rboPODate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimarySupplier)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtSupplierCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtSupplierName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindSupplier)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpAOPDPTo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpAOPDPFrom)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpPOTo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpPOFrom)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.ChkUse)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnFilter)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboRealizationStatus)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboPaymentStatus)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboDPStatus)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Size = New System.Drawing.Size(720, 100)
        Me.SplitContainer1.SplitterDistance = 378
        Me.SplitContainer1.TabIndex = 28
        Me.SplitContainer1.TabStop = False
        '
        'RboAOP
        '
        Me.RboAOP.AutoSize = True
        Me.RboAOP.Location = New System.Drawing.Point(2, 49)
        Me.RboAOP.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.RboAOP.Name = "RboAOP"
        Me.RboAOP.Size = New System.Drawing.Size(73, 17)
        Me.RboAOP.TabIndex = 34
        Me.RboAOP.Text = "AOP Date"
        Me.RboAOP.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(212, 50)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(20, 13)
        Me.Label7.TabIndex = 33
        Me.Label7.Text = "To"
        '
        'DtpAOPDateTo
        '
        Me.DtpAOPDateTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpAOPDateTo.CustomFormat = "dd MMM yyyy"
        Me.DtpAOPDateTo.Enabled = False
        Me.DtpAOPDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpAOPDateTo.Location = New System.Drawing.Point(235, 48)
        Me.DtpAOPDateTo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DtpAOPDateTo.Name = "DtpAOPDateTo"
        Me.DtpAOPDateTo.Size = New System.Drawing.Size(142, 20)
        Me.DtpAOPDateTo.TabIndex = 32
        '
        'DtpAOPDateFrom
        '
        Me.DtpAOPDateFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpAOPDateFrom.Enabled = False
        Me.DtpAOPDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpAOPDateFrom.Location = New System.Drawing.Point(94, 48)
        Me.DtpAOPDateFrom.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DtpAOPDateFrom.Name = "DtpAOPDateFrom"
        Me.DtpAOPDateFrom.Size = New System.Drawing.Size(114, 20)
        Me.DtpAOPDateFrom.TabIndex = 31
        '
        'rboAOPDP
        '
        Me.rboAOPDP.AutoSize = True
        Me.rboAOPDP.Location = New System.Drawing.Point(2, 26)
        Me.rboAOPDP.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.rboAOPDP.Name = "rboAOPDP"
        Me.rboAOPDP.Size = New System.Drawing.Size(91, 17)
        Me.rboAOPDP.TabIndex = 30
        Me.rboAOPDP.Text = "AOP DP Date"
        Me.rboAOPDP.UseVisualStyleBackColor = True
        '
        'rboPODate
        '
        Me.rboPODate.AutoSize = True
        Me.rboPODate.Checked = True
        Me.rboPODate.Location = New System.Drawing.Point(2, 3)
        Me.rboPODate.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.rboPODate.Name = "rboPODate"
        Me.rboPODate.Size = New System.Drawing.Size(66, 17)
        Me.rboPODate.TabIndex = 29
        Me.rboPODate.TabStop = True
        Me.rboPODate.Text = "PO Date"
        Me.rboPODate.UseVisualStyleBackColor = True
        '
        'TxtPrimarySupplier
        '
        Me.TxtPrimarySupplier.Location = New System.Drawing.Point(59, 71)
        Me.TxtPrimarySupplier.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimarySupplier.Name = "TxtPrimarySupplier"
        Me.TxtPrimarySupplier.Size = New System.Drawing.Size(31, 20)
        Me.TxtPrimarySupplier.TabIndex = 12
        Me.TxtPrimarySupplier.TabStop = False
        Me.TxtPrimarySupplier.Text = "-1"
        Me.TxtPrimarySupplier.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(2, 73)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 13)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Supplier"
        '
        'TxtSupplierCode
        '
        Me.TxtSupplierCode.Location = New System.Drawing.Point(94, 71)
        Me.TxtSupplierCode.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtSupplierCode.Name = "TxtSupplierCode"
        Me.TxtSupplierCode.Size = New System.Drawing.Size(67, 20)
        Me.TxtSupplierCode.TabIndex = 10
        '
        'TxtSupplierName
        '
        Me.TxtSupplierName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSupplierName.Location = New System.Drawing.Point(212, 71)
        Me.TxtSupplierName.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtSupplierName.Name = "TxtSupplierName"
        Me.TxtSupplierName.ReadOnly = True
        Me.TxtSupplierName.Size = New System.Drawing.Size(165, 20)
        Me.TxtSupplierName.TabIndex = 9
        Me.TxtSupplierName.TabStop = False
        '
        'btnFindSupplier
        '
        Me.btnFindSupplier.Location = New System.Drawing.Point(164, 70)
        Me.btnFindSupplier.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFindSupplier.Name = "btnFindSupplier"
        Me.btnFindSupplier.Size = New System.Drawing.Size(43, 20)
        Me.btnFindSupplier.TabIndex = 8
        Me.btnFindSupplier.Text = "Find"
        Me.btnFindSupplier.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(212, 28)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "To"
        '
        'dtpAOPDPTo
        '
        Me.dtpAOPDPTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpAOPDPTo.CustomFormat = "dd MMM yyyy"
        Me.dtpAOPDPTo.Enabled = False
        Me.dtpAOPDPTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpAOPDPTo.Location = New System.Drawing.Point(235, 25)
        Me.dtpAOPDPTo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpAOPDPTo.Name = "dtpAOPDPTo"
        Me.dtpAOPDPTo.Size = New System.Drawing.Size(142, 20)
        Me.dtpAOPDPTo.TabIndex = 6
        '
        'dtpAOPDPFrom
        '
        Me.dtpAOPDPFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpAOPDPFrom.Enabled = False
        Me.dtpAOPDPFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpAOPDPFrom.Location = New System.Drawing.Point(94, 25)
        Me.dtpAOPDPFrom.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpAOPDPFrom.Name = "dtpAOPDPFrom"
        Me.dtpAOPDPFrom.Size = New System.Drawing.Size(114, 20)
        Me.dtpAOPDPFrom.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(212, 5)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(20, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "To"
        '
        'dtpPOTo
        '
        Me.dtpPOTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpPOTo.CustomFormat = "dd MMM yyyy"
        Me.dtpPOTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPOTo.Location = New System.Drawing.Point(235, 2)
        Me.dtpPOTo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpPOTo.Name = "dtpPOTo"
        Me.dtpPOTo.Size = New System.Drawing.Size(142, 20)
        Me.dtpPOTo.TabIndex = 2
        '
        'dtpPOFrom
        '
        Me.dtpPOFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpPOFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPOFrom.Location = New System.Drawing.Point(94, 2)
        Me.dtpPOFrom.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpPOFrom.Name = "dtpPOFrom"
        Me.dtpPOFrom.Size = New System.Drawing.Size(114, 20)
        Me.dtpPOFrom.TabIndex = 1
        '
        'ChkUse
        '
        Me.ChkUse.AutoSize = True
        Me.ChkUse.Location = New System.Drawing.Point(4, 77)
        Me.ChkUse.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.ChkUse.Name = "ChkUse"
        Me.ChkUse.Size = New System.Drawing.Size(159, 17)
        Me.ChkUse.TabIndex = 30
        Me.ChkUse.Text = "Use supplier invoice number"
        Me.ChkUse.UseVisualStyleBackColor = True
        '
        'BtnFilter
        '
        Me.BtnFilter.Location = New System.Drawing.Point(233, 77)
        Me.BtnFilter.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.BtnFilter.Name = "BtnFilter"
        Me.BtnFilter.Size = New System.Drawing.Size(104, 20)
        Me.BtnFilter.TabIndex = 29
        Me.BtnFilter.Text = "Filter"
        Me.BtnFilter.UseVisualStyleBackColor = True
        '
        'cboRealizationStatus
        '
        Me.cboRealizationStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRealizationStatus.FormattingEnabled = True
        Me.cboRealizationStatus.Items.AddRange(New Object() {"All", "Realized", "Partially Realized/Not Realized"})
        Me.cboRealizationStatus.Location = New System.Drawing.Point(129, 49)
        Me.cboRealizationStatus.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboRealizationStatus.Name = "cboRealizationStatus"
        Me.cboRealizationStatus.Size = New System.Drawing.Size(209, 21)
        Me.cboRealizationStatus.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(2, 51)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(92, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Realization Status"
        '
        'cboPaymentStatus
        '
        Me.cboPaymentStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPaymentStatus.FormattingEnabled = True
        Me.cboPaymentStatus.Items.AddRange(New Object() {"All", "Paid", "Partially Paid/No Payment"})
        Me.cboPaymentStatus.Location = New System.Drawing.Point(129, 25)
        Me.cboPaymentStatus.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboPaymentStatus.Name = "cboPaymentStatus"
        Me.cboPaymentStatus.Size = New System.Drawing.Size(209, 21)
        Me.cboPaymentStatus.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(2, 28)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(81, 13)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "Payment Status"
        '
        'cboDPStatus
        '
        Me.cboDPStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDPStatus.FormattingEnabled = True
        Me.cboDPStatus.Items.AddRange(New Object() {"All", "Downpayment Created", "Downpayment Not Created"})
        Me.cboDPStatus.Location = New System.Drawing.Point(129, 1)
        Me.cboDPStatus.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboDPStatus.Name = "cboDPStatus"
        Me.cboDPStatus.Size = New System.Drawing.Size(209, 21)
        Me.cboDPStatus.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(2, 3)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Down Payment Status"
        '
        'crv
        '
        Me.crv.ActiveViewIndex = -1
        Me.crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv.DisplayGroupTree = False
        Me.crv.EnableDrillDown = False
        Me.crv.Location = New System.Drawing.Point(12, 117)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(719, 394)
        Me.crv.TabIndex = 27
        Me.crv.TabStop = False
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'FrmLaporanPurchaseDP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(740, 514)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.crv)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(756, 552)
        Me.Name = "FrmLaporanPurchaseDP"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Downpayment Report"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpAOPDPTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpAOPDPFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpPOTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpPOFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents btnFindSupplier As System.Windows.Forms.Button
    Friend WithEvents cboRealizationStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboPaymentStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboDPStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents rboAOPDP As System.Windows.Forms.RadioButton
    Friend WithEvents rboPODate As System.Windows.Forms.RadioButton
    Friend WithEvents RboAOP As System.Windows.Forms.RadioButton
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents DtpAOPDateTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpAOPDateFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents BtnFilter As System.Windows.Forms.Button
    Friend WithEvents ChkUse As System.Windows.Forms.CheckBox
End Class
