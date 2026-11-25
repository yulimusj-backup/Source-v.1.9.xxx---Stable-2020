<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanSalesOrderBalance
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmLaporanSalesOrderBalance))
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtCusomerCode = New System.Windows.Forms.TextBox
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtCustomerName = New System.Windows.Forms.TextBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.Label3 = New System.Windows.Forms.Label
        Me.PrintForm1 = New Microsoft.VisualBasic.PowerPacks.Printing.PrintForm(Me.components)
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkFilter = New System.Windows.Forms.CheckBox
        Me.rboGroupByProductCode = New System.Windows.Forms.RadioButton
        Me.rboGroupBySalesOrder = New System.Windows.Forms.RadioButton
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.txtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnFindCustomer = New System.Windows.Forms.Button
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(176, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "To"
        '
        'txtCusomerCode
        '
        Me.txtCusomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCusomerCode.Location = New System.Drawing.Point(579, 19)
        Me.txtCusomerCode.Name = "txtCusomerCode"
        Me.txtCusomerCode.Size = New System.Drawing.Size(172, 20)
        Me.txtCusomerCode.TabIndex = 5
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpFrom.Enabled = False
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(243, 19)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(210, 20)
        Me.dtpFrom.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(494, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Customer Name"
        '
        'txtCustomerName
        '
        Me.txtCustomerName.Location = New System.Drawing.Point(579, 45)
        Me.txtCustomerName.Name = "txtCustomerName"
        Me.txtCustomerName.ReadOnly = True
        Me.txtCustomerName.Size = New System.Drawing.Size(210, 20)
        Me.txtCustomerName.TabIndex = 6
        Me.txtCustomerName.TabStop = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 411)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(902, 22)
        Me.StatusStrip1.TabIndex = 5
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(494, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(79, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Customer Code"
        '
        'PrintForm1
        '
        Me.PrintForm1.DocumentName = "document"
        Me.PrintForm1.Form = Me
        Me.PrintForm1.PrintAction = System.Drawing.Printing.PrintAction.PrintToPrinter
        Me.PrintForm1.PrinterSettings = CType(resources.GetObject("PrintForm1.PrinterSettings"), System.Drawing.Printing.PrinterSettings)
        Me.PrintForm1.PrintFileName = Nothing
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpTo.Enabled = False
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(243, 45)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(210, 20)
        Me.dtpTo.TabIndex = 3
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.chkFilter)
        Me.GroupBox1.Controls.Add(Me.rboGroupByProductCode)
        Me.GroupBox1.Controls.Add(Me.rboGroupBySalesOrder)
        Me.GroupBox1.Controls.Add(Me.btnExit)
        Me.GroupBox1.Controls.Add(Me.btnFilter)
        Me.GroupBox1.Controls.Add(Me.txtPrimaryCustomer)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtCustomerName)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.dtpTo)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtCusomerCode)
        Me.GroupBox1.Controls.Add(Me.dtpFrom)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.btnFindCustomer)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(878, 83)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter"
        '
        'chkFilter
        '
        Me.chkFilter.AutoSize = True
        Me.chkFilter.Location = New System.Drawing.Point(6, 64)
        Me.chkFilter.Name = "chkFilter"
        Me.chkFilter.Size = New System.Drawing.Size(107, 17)
        Me.chkFilter.TabIndex = 12
        Me.chkFilter.Text = "Filter By SO Date"
        Me.chkFilter.UseVisualStyleBackColor = True
        '
        'rboGroupByProductCode
        '
        Me.rboGroupByProductCode.AutoSize = True
        Me.rboGroupByProductCode.Location = New System.Drawing.Point(6, 41)
        Me.rboGroupByProductCode.Name = "rboGroupByProductCode"
        Me.rboGroupByProductCode.Size = New System.Drawing.Size(137, 17)
        Me.rboGroupByProductCode.TabIndex = 1
        Me.rboGroupByProductCode.Text = "Group By Product Code"
        Me.rboGroupByProductCode.UseVisualStyleBackColor = True
        '
        'rboGroupBySalesOrder
        '
        Me.rboGroupBySalesOrder.AutoSize = True
        Me.rboGroupBySalesOrder.Checked = True
        Me.rboGroupBySalesOrder.Location = New System.Drawing.Point(6, 17)
        Me.rboGroupBySalesOrder.Name = "rboGroupBySalesOrder"
        Me.rboGroupBySalesOrder.Size = New System.Drawing.Size(127, 17)
        Me.rboGroupBySalesOrder.TabIndex = 0
        Me.rboGroupBySalesOrder.TabStop = True
        Me.rboGroupBySalesOrder.Text = "Group By Sales Order"
        Me.rboGroupBySalesOrder.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(795, 46)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(77, 24)
        Me.btnExit.TabIndex = 9
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(795, 19)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(77, 24)
        Me.btnFilter.TabIndex = 8
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'txtPrimaryCustomer
        '
        Me.txtPrimaryCustomer.Location = New System.Drawing.Point(542, 19)
        Me.txtPrimaryCustomer.Name = "txtPrimaryCustomer"
        Me.txtPrimaryCustomer.Size = New System.Drawing.Size(30, 20)
        Me.txtPrimaryCustomer.TabIndex = 4
        Me.txtPrimaryCustomer.Text = "-1"
        Me.txtPrimaryCustomer.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(176, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "From"
        '
        'btnFindCustomer
        '
        Me.btnFindCustomer.Location = New System.Drawing.Point(757, 19)
        Me.btnFindCustomer.Name = "btnFindCustomer"
        Me.btnFindCustomer.Size = New System.Drawing.Size(32, 20)
        Me.btnFindCustomer.TabIndex = 7
        Me.btnFindCustomer.TabStop = False
        Me.btnFindCustomer.Text = "...."
        Me.btnFindCustomer.UseVisualStyleBackColor = True
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.Location = New System.Drawing.Point(12, 93)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(878, 309)
        Me.crv1.TabIndex = 2
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'FrmLaporanSalesOrderBalance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(902, 433)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.crv1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(918, 400)
        Me.Name = "FrmLaporanSalesOrderBalance"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales Order Balance Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCusomerCode As System.Windows.Forms.TextBox
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents PrintForm1 As Microsoft.VisualBasic.PowerPacks.Printing.PrintForm
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnFindCustomer As System.Windows.Forms.Button
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents txtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents rboGroupByProductCode As System.Windows.Forms.RadioButton
    Friend WithEvents rboGroupBySalesOrder As System.Windows.Forms.RadioButton
    Friend WithEvents chkFilter As System.Windows.Forms.CheckBox
End Class
