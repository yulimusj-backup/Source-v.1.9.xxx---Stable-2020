<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanSalesListByDO
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmLaporanSalesListByDO))
        Me.rboAll = New System.Windows.Forms.RadioButton
        Me.rboNotInvoiceDO = New System.Windows.Forms.RadioButton
        Me.rboInvoiceDO = New System.Windows.Forms.RadioButton
        Me.btnExit = New System.Windows.Forms.Button
        Me.PrintForm1 = New Microsoft.VisualBasic.PowerPacks.Printing.PrintForm(Me.components)
        Me.btnShowReport = New System.Windows.Forms.Button
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.dtpDari = New System.Windows.Forms.DateTimePicker
        Me.dtpSampai = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.chkNonPendingDO = New System.Windows.Forms.CheckBox
        Me.cboFilterBy = New System.Windows.Forms.ComboBox
        Me.rboDONotStockedOut = New System.Windows.Forms.RadioButton
        Me.SuspendLayout()
        '
        'rboAll
        '
        Me.rboAll.AutoSize = True
        Me.rboAll.Location = New System.Drawing.Point(399, 37)
        Me.rboAll.Name = "rboAll"
        Me.rboAll.Size = New System.Drawing.Size(36, 17)
        Me.rboAll.TabIndex = 4
        Me.rboAll.Text = "All"
        Me.rboAll.UseVisualStyleBackColor = True
        '
        'rboNotInvoiceDO
        '
        Me.rboNotInvoiceDO.AutoSize = True
        Me.rboNotInvoiceDO.Location = New System.Drawing.Point(265, 36)
        Me.rboNotInvoiceDO.Name = "rboNotInvoiceDO"
        Me.rboNotInvoiceDO.Size = New System.Drawing.Size(123, 17)
        Me.rboNotInvoiceDO.TabIndex = 3
        Me.rboNotInvoiceDO.Text = "Not Invoice DO Only"
        Me.rboNotInvoiceDO.UseVisualStyleBackColor = True
        '
        'rboInvoiceDO
        '
        Me.rboInvoiceDO.AutoSize = True
        Me.rboInvoiceDO.Checked = True
        Me.rboInvoiceDO.Location = New System.Drawing.Point(265, 13)
        Me.rboInvoiceDO.Name = "rboInvoiceDO"
        Me.rboInvoiceDO.Size = New System.Drawing.Size(109, 17)
        Me.rboInvoiceDO.TabIndex = 2
        Me.rboInvoiceDO.TabStop = True
        Me.rboInvoiceDO.Text = "Invoiced DO Only"
        Me.rboInvoiceDO.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(528, 39)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(77, 24)
        Me.btnExit.TabIndex = 6
        Me.btnExit.Text = "&Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'PrintForm1
        '
        Me.PrintForm1.DocumentName = "document"
        Me.PrintForm1.Form = Me
        Me.PrintForm1.PrintAction = System.Drawing.Printing.PrintAction.PrintToPrinter
        Me.PrintForm1.PrinterSettings = CType(resources.GetObject("PrintForm1.PrinterSettings"), System.Drawing.Printing.PrinterSettings)
        Me.PrintForm1.PrintFileName = Nothing
        '
        'btnShowReport
        '
        Me.btnShowReport.Location = New System.Drawing.Point(528, 12)
        Me.btnShowReport.Name = "btnShowReport"
        Me.btnShowReport.Size = New System.Drawing.Size(77, 24)
        Me.btnShowReport.TabIndex = 5
        Me.btnShowReport.Text = "&Filter"
        Me.btnShowReport.UseVisualStyleBackColor = True
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.Location = New System.Drawing.Point(12, 89)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(878, 423)
        Me.crv1.TabIndex = 6
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 535)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(910, 22)
        Me.StatusStrip1.TabIndex = 21
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'dtpDari
        '
        Me.dtpDari.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDari.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDari.Location = New System.Drawing.Point(59, 9)
        Me.dtpDari.Name = "dtpDari"
        Me.dtpDari.Size = New System.Drawing.Size(200, 20)
        Me.dtpDari.TabIndex = 0
        '
        'dtpSampai
        '
        Me.dtpSampai.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpSampai.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSampai.Location = New System.Drawing.Point(59, 35)
        Me.dtpSampai.Name = "dtpSampai"
        Me.dtpSampai.Size = New System.Drawing.Size(200, 20)
        Me.dtpSampai.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 14)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "From"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 41)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 31
        Me.Label3.Text = "To"
        '
        'chkNonPendingDO
        '
        Me.chkNonPendingDO.AutoSize = True
        Me.chkNonPendingDO.Location = New System.Drawing.Point(265, 64)
        Me.chkNonPendingDO.Name = "chkNonPendingDO"
        Me.chkNonPendingDO.Size = New System.Drawing.Size(125, 17)
        Me.chkNonPendingDO.TabIndex = 32
        Me.chkNonPendingDO.Text = "Exclude Pending DO"
        Me.chkNonPendingDO.UseVisualStyleBackColor = True
        '
        'cboFilterBy
        '
        Me.cboFilterBy.FormattingEnabled = True
        Me.cboFilterBy.Items.AddRange(New Object() {"Stock Out Date", "DO Date"})
        Me.cboFilterBy.Location = New System.Drawing.Point(19, 61)
        Me.cboFilterBy.Name = "cboFilterBy"
        Me.cboFilterBy.Size = New System.Drawing.Size(240, 21)
        Me.cboFilterBy.TabIndex = 33
        Me.cboFilterBy.TabStop = False
        Me.cboFilterBy.Visible = False
        '
        'rboDONotStockedOut
        '
        Me.rboDONotStockedOut.AutoSize = True
        Me.rboDONotStockedOut.Location = New System.Drawing.Point(399, 13)
        Me.rboDONotStockedOut.Name = "rboDONotStockedOut"
        Me.rboDONotStockedOut.Size = New System.Drawing.Size(124, 17)
        Me.rboDONotStockedOut.TabIndex = 34
        Me.rboDONotStockedOut.Text = "DO Not Stocked Out"
        Me.rboDONotStockedOut.UseVisualStyleBackColor = True
        '
        'FrmLaporanSalesListByDO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(910, 557)
        Me.Controls.Add(Me.rboDONotStockedOut)
        Me.Controls.Add(Me.cboFilterBy)
        Me.Controls.Add(Me.chkNonPendingDO)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpSampai)
        Me.Controls.Add(Me.dtpDari)
        Me.Controls.Add(Me.rboAll)
        Me.Controls.Add(Me.rboNotInvoiceDO)
        Me.Controls.Add(Me.rboInvoiceDO)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnShowReport)
        Me.Controls.Add(Me.crv1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.MinimumSize = New System.Drawing.Size(918, 588)
        Me.Name = "FrmLaporanSalesListByDO"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales List By Delivery Order"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents rboAll As System.Windows.Forms.RadioButton
    Friend WithEvents rboNotInvoiceDO As System.Windows.Forms.RadioButton
    Friend WithEvents rboInvoiceDO As System.Windows.Forms.RadioButton
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents PrintForm1 As Microsoft.VisualBasic.PowerPacks.Printing.PrintForm
    Friend WithEvents btnShowReport As System.Windows.Forms.Button
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpSampai As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpDari As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkNonPendingDO As System.Windows.Forms.CheckBox
    Friend WithEvents cboFilterBy As System.Windows.Forms.ComboBox
    Friend WithEvents rboDONotStockedOut As System.Windows.Forms.RadioButton
End Class
