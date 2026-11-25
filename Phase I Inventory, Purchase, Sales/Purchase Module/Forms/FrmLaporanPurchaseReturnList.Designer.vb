<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPurchaseReturnList
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
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.Label3 = New System.Windows.Forms.Label
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.clbBrand = New System.Windows.Forms.CheckedListBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnShowReport = New System.Windows.Forms.Button
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.gboBrand = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboGrouping = New System.Windows.Forms.ComboBox
        Me.gboSupplier = New System.Windows.Forms.GroupBox
        Me.llbUncekAll_Supplier = New System.Windows.Forms.LinkLabel
        Me.llbAll_Supplier = New System.Windows.Forms.LinkLabel
        Me.chkSupplier = New System.Windows.Forms.CheckedListBox
        Me.chkService = New System.Windows.Forms.CheckBox
        Me.chkGoods = New System.Windows.Forms.CheckBox
        Me.gboBrand.SuspendLayout()
        Me.gboSupplier.SuspendLayout()
        Me.SuspendLayout()
        '
        'dtpFrom
        '
        Me.dtpFrom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom.CustomFormat = "MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(670, 22)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.ShowUpDown = True
        Me.dtpFrom.Size = New System.Drawing.Size(141, 20)
        Me.dtpFrom.TabIndex = 1
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(451, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel1.TabIndex = 0
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Check All"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(614, 25)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(37, 13)
        Me.Label3.TabIndex = 312
        Me.Label3.Text = "Period"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(507, 0)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel2.TabIndex = 2
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Uncheck All"
        '
        'clbBrand
        '
        Me.clbBrand.CheckOnClick = True
        Me.clbBrand.ColumnWidth = 300
        Me.clbBrand.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbBrand.FormattingEnabled = True
        Me.clbBrand.Location = New System.Drawing.Point(3, 16)
        Me.clbBrand.MultiColumn = True
        Me.clbBrand.Name = "clbBrand"
        Me.clbBrand.Size = New System.Drawing.Size(590, 79)
        Me.clbBrand.TabIndex = 0
        Me.clbBrand.ThreeDCheckBoxes = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(817, 46)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(77, 24)
        Me.btnExit.TabIndex = 6
        Me.btnExit.Text = "&Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnShowReport
        '
        Me.btnShowReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnShowReport.Location = New System.Drawing.Point(817, 19)
        Me.btnShowReport.Name = "btnShowReport"
        Me.btnShowReport.Size = New System.Drawing.Size(77, 24)
        Me.btnShowReport.TabIndex = 5
        Me.btnShowReport.Text = "&Filter"
        Me.btnShowReport.UseVisualStyleBackColor = True
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.DisplayGroupTree = False
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(12, 111)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.SelectionFormula = ""
        Me.CrystalReportViewer1.ShowCloseButton = False
        Me.CrystalReportViewer1.ShowGroupTreeButton = False
        Me.CrystalReportViewer1.ShowRefreshButton = False
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(878, 387)
        Me.CrystalReportViewer1.TabIndex = 306
        Me.CrystalReportViewer1.ViewTimeSelectionFormula = ""
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 526)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(902, 22)
        Me.StatusStrip1.TabIndex = 307
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'gboBrand
        '
        Me.gboBrand.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboBrand.Controls.Add(Me.LinkLabel2)
        Me.gboBrand.Controls.Add(Me.LinkLabel1)
        Me.gboBrand.Controls.Add(Me.clbBrand)
        Me.gboBrand.Location = New System.Drawing.Point(12, 6)
        Me.gboBrand.Name = "gboBrand"
        Me.gboBrand.Size = New System.Drawing.Size(596, 99)
        Me.gboBrand.TabIndex = 310
        Me.gboBrand.TabStop = False
        Me.gboBrand.Text = "Brand Filter"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(614, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(50, 13)
        Me.Label2.TabIndex = 316
        Me.Label2.Text = "Grouping"
        Me.Label2.Visible = False
        '
        'cboGrouping
        '
        Me.cboGrouping.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboGrouping.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGrouping.FormattingEnabled = True
        Me.cboGrouping.Items.AddRange(New Object() {"By Brand", "By Supplier"})
        Me.cboGrouping.Location = New System.Drawing.Point(670, 48)
        Me.cboGrouping.Name = "cboGrouping"
        Me.cboGrouping.Size = New System.Drawing.Size(141, 21)
        Me.cboGrouping.TabIndex = 4
        Me.cboGrouping.TabStop = False
        Me.cboGrouping.Visible = False
        '
        'gboSupplier
        '
        Me.gboSupplier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboSupplier.Controls.Add(Me.llbUncekAll_Supplier)
        Me.gboSupplier.Controls.Add(Me.llbAll_Supplier)
        Me.gboSupplier.Controls.Add(Me.chkSupplier)
        Me.gboSupplier.Location = New System.Drawing.Point(12, 6)
        Me.gboSupplier.Name = "gboSupplier"
        Me.gboSupplier.Size = New System.Drawing.Size(596, 99)
        Me.gboSupplier.TabIndex = 0
        Me.gboSupplier.TabStop = False
        Me.gboSupplier.Text = "Supplier Filter"
        '
        'llbUncekAll_Supplier
        '
        Me.llbUncekAll_Supplier.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncekAll_Supplier.AutoSize = True
        Me.llbUncekAll_Supplier.Location = New System.Drawing.Point(507, 0)
        Me.llbUncekAll_Supplier.Name = "llbUncekAll_Supplier"
        Me.llbUncekAll_Supplier.Size = New System.Drawing.Size(65, 13)
        Me.llbUncekAll_Supplier.TabIndex = 0
        Me.llbUncekAll_Supplier.TabStop = True
        Me.llbUncekAll_Supplier.Text = "Uncheck All"
        '
        'llbAll_Supplier
        '
        Me.llbAll_Supplier.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAll_Supplier.AutoSize = True
        Me.llbAll_Supplier.Location = New System.Drawing.Point(451, 0)
        Me.llbAll_Supplier.Name = "llbAll_Supplier"
        Me.llbAll_Supplier.Size = New System.Drawing.Size(52, 13)
        Me.llbAll_Supplier.TabIndex = 1
        Me.llbAll_Supplier.TabStop = True
        Me.llbAll_Supplier.Text = "Check All"
        '
        'chkSupplier
        '
        Me.chkSupplier.CheckOnClick = True
        Me.chkSupplier.ColumnWidth = 300
        Me.chkSupplier.Dock = System.Windows.Forms.DockStyle.Fill
        Me.chkSupplier.FormattingEnabled = True
        Me.chkSupplier.Location = New System.Drawing.Point(3, 16)
        Me.chkSupplier.MultiColumn = True
        Me.chkSupplier.Name = "chkSupplier"
        Me.chkSupplier.Size = New System.Drawing.Size(590, 79)
        Me.chkSupplier.TabIndex = 0
        Me.chkSupplier.ThreeDCheckBoxes = True
        '
        'chkService
        '
        Me.chkService.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkService.AutoSize = True
        Me.chkService.Checked = True
        Me.chkService.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkService.Location = New System.Drawing.Point(745, 50)
        Me.chkService.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkService.Name = "chkService"
        Me.chkService.Size = New System.Drawing.Size(62, 17)
        Me.chkService.TabIndex = 3
        Me.chkService.Text = "Service"
        Me.chkService.UseVisualStyleBackColor = True
        '
        'chkGoods
        '
        Me.chkGoods.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkGoods.AutoSize = True
        Me.chkGoods.Checked = True
        Me.chkGoods.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkGoods.Location = New System.Drawing.Point(667, 50)
        Me.chkGoods.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkGoods.Name = "chkGoods"
        Me.chkGoods.Size = New System.Drawing.Size(57, 17)
        Me.chkGoods.TabIndex = 2
        Me.chkGoods.Text = "Goods"
        Me.chkGoods.UseVisualStyleBackColor = True
        '
        'FrmLaporanPurchaseReturnList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(902, 548)
        Me.Controls.Add(Me.chkService)
        Me.Controls.Add(Me.chkGoods)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboGrouping)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnShowReport)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.gboSupplier)
        Me.Controls.Add(Me.gboBrand)
        Me.MinimumSize = New System.Drawing.Size(918, 586)
        Me.Name = "FrmLaporanPurchaseReturnList"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Return List"
        Me.gboBrand.ResumeLayout(False)
        Me.gboBrand.PerformLayout()
        Me.gboSupplier.ResumeLayout(False)
        Me.gboSupplier.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents clbBrand As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnShowReport As System.Windows.Forms.Button
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents gboBrand As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboGrouping As System.Windows.Forms.ComboBox
    Friend WithEvents gboSupplier As System.Windows.Forms.GroupBox
    Friend WithEvents llbUncekAll_Supplier As System.Windows.Forms.LinkLabel
    Friend WithEvents llbAll_Supplier As System.Windows.Forms.LinkLabel
    Friend WithEvents chkSupplier As System.Windows.Forms.CheckedListBox
    Friend WithEvents chkService As System.Windows.Forms.CheckBox
    Friend WithEvents chkGoods As System.Windows.Forms.CheckBox
End Class
