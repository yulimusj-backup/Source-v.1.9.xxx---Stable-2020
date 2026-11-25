<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTransaksiStockTaking_Print
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
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtPrimaryStockTaking = New System.Windows.Forms.TextBox
        Me.gboBrand = New System.Windows.Forms.GroupBox
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.clbBrand = New System.Windows.Forms.CheckedListBox
        Me.TxtWarehouseCode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtPrimaryWarehouse = New System.Windows.Forms.TextBox
        Me.btnFindWarehouse = New System.Windows.Forms.Button
        Me.cboLocation = New System.Windows.Forms.ComboBox
        Me.TxtPrimaryLocation = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtWarehouseName = New System.Windows.Forms.TextBox
        Me.TxtStockTakingNumber = New System.Windows.Forms.TextBox
        Me.btnFilter = New System.Windows.Forms.Button
        Me.chkUseSpecification = New System.Windows.Forms.CheckBox
        Me.gboBrand.SuspendLayout()
        Me.SuspendLayout()
        '
        'crv
        '
        Me.crv.ActiveViewIndex = -1
        Me.crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv.DisplayGroupTree = False
        Me.crv.Location = New System.Drawing.Point(13, 204)
        Me.crv.Margin = New System.Windows.Forms.Padding(4)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(956, 336)
        Me.crv.TabIndex = 9
        Me.crv.TabStop = False
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(217, 172)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 24)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 17)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Stock Taking"
        '
        'TxtPrimaryStockTaking
        '
        Me.TxtPrimaryStockTaking.Location = New System.Drawing.Point(106, 12)
        Me.TxtPrimaryStockTaking.Name = "TxtPrimaryStockTaking"
        Me.TxtPrimaryStockTaking.Size = New System.Drawing.Size(24, 22)
        Me.TxtPrimaryStockTaking.TabIndex = 7
        Me.TxtPrimaryStockTaking.Text = "-1"
        Me.TxtPrimaryStockTaking.Visible = False
        '
        'gboBrand
        '
        Me.gboBrand.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboBrand.Controls.Add(Me.LinkLabel2)
        Me.gboBrand.Controls.Add(Me.LinkLabel1)
        Me.gboBrand.Controls.Add(Me.clbBrand)
        Me.gboBrand.Location = New System.Drawing.Point(361, 11)
        Me.gboBrand.Margin = New System.Windows.Forms.Padding(4)
        Me.gboBrand.Name = "gboBrand"
        Me.gboBrand.Padding = New System.Windows.Forms.Padding(4)
        Me.gboBrand.Size = New System.Drawing.Size(608, 185)
        Me.gboBrand.TabIndex = 8
        Me.gboBrand.TabStop = False
        Me.gboBrand.Text = "Brand Filter"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(490, 0)
        Me.LinkLabel2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(82, 17)
        Me.LinkLabel2.TabIndex = 2
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Uncheck All"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(415, 0)
        Me.LinkLabel1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(66, 17)
        Me.LinkLabel1.TabIndex = 1
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Check All"
        '
        'clbBrand
        '
        Me.clbBrand.CheckOnClick = True
        Me.clbBrand.ColumnWidth = 300
        Me.clbBrand.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbBrand.FormattingEnabled = True
        Me.clbBrand.Location = New System.Drawing.Point(4, 19)
        Me.clbBrand.Margin = New System.Windows.Forms.Padding(4)
        Me.clbBrand.MultiColumn = True
        Me.clbBrand.Name = "clbBrand"
        Me.clbBrand.Size = New System.Drawing.Size(600, 157)
        Me.clbBrand.TabIndex = 0
        Me.clbBrand.ThreeDCheckBoxes = True
        '
        'TxtWarehouseCode
        '
        Me.TxtWarehouseCode.Location = New System.Drawing.Point(135, 44)
        Me.TxtWarehouseCode.Name = "TxtWarehouseCode"
        Me.TxtWarehouseCode.Size = New System.Drawing.Size(138, 22)
        Me.TxtWarehouseCode.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 17)
        Me.Label2.TabIndex = 314
        Me.Label2.Text = "Warehouse"
        '
        'TxtPrimaryWarehouse
        '
        Me.TxtPrimaryWarehouse.Location = New System.Drawing.Point(105, 44)
        Me.TxtPrimaryWarehouse.Name = "TxtPrimaryWarehouse"
        Me.TxtPrimaryWarehouse.Size = New System.Drawing.Size(24, 22)
        Me.TxtPrimaryWarehouse.TabIndex = 315
        Me.TxtPrimaryWarehouse.Text = "-1"
        Me.TxtPrimaryWarehouse.Visible = False
        '
        'btnFindWarehouse
        '
        Me.btnFindWarehouse.Location = New System.Drawing.Point(279, 44)
        Me.btnFindWarehouse.Name = "btnFindWarehouse"
        Me.btnFindWarehouse.Size = New System.Drawing.Size(75, 23)
        Me.btnFindWarehouse.TabIndex = 2
        Me.btnFindWarehouse.Text = "Find"
        Me.btnFindWarehouse.UseVisualStyleBackColor = True
        '
        'cboLocation
        '
        Me.cboLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLocation.FormattingEnabled = True
        Me.cboLocation.Location = New System.Drawing.Point(135, 108)
        Me.cboLocation.Name = "cboLocation"
        Me.cboLocation.Size = New System.Drawing.Size(219, 24)
        Me.cboLocation.TabIndex = 4
        '
        'TxtPrimaryLocation
        '
        Me.TxtPrimaryLocation.Location = New System.Drawing.Point(105, 108)
        Me.TxtPrimaryLocation.Name = "TxtPrimaryLocation"
        Me.TxtPrimaryLocation.Size = New System.Drawing.Size(24, 22)
        Me.TxtPrimaryLocation.TabIndex = 318
        Me.TxtPrimaryLocation.Text = "-1"
        Me.TxtPrimaryLocation.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 111)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 17)
        Me.Label3.TabIndex = 319
        Me.Label3.Text = "Location"
        '
        'TxtWarehouseName
        '
        Me.TxtWarehouseName.Location = New System.Drawing.Point(135, 76)
        Me.TxtWarehouseName.Name = "TxtWarehouseName"
        Me.TxtWarehouseName.ReadOnly = True
        Me.TxtWarehouseName.Size = New System.Drawing.Size(219, 22)
        Me.TxtWarehouseName.TabIndex = 3
        Me.TxtWarehouseName.TabStop = False
        '
        'TxtStockTakingNumber
        '
        Me.TxtStockTakingNumber.Location = New System.Drawing.Point(136, 12)
        Me.TxtStockTakingNumber.Name = "TxtStockTakingNumber"
        Me.TxtStockTakingNumber.ReadOnly = True
        Me.TxtStockTakingNumber.Size = New System.Drawing.Size(219, 22)
        Me.TxtStockTakingNumber.TabIndex = 0
        Me.TxtStockTakingNumber.TabStop = False
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(135, 172)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(75, 24)
        Me.btnFilter.TabIndex = 6
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'chkUseSpecification
        '
        Me.chkUseSpecification.AutoSize = True
        Me.chkUseSpecification.Checked = True
        Me.chkUseSpecification.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkUseSpecification.Location = New System.Drawing.Point(135, 140)
        Me.chkUseSpecification.Name = "chkUseSpecification"
        Me.chkUseSpecification.Size = New System.Drawing.Size(139, 21)
        Me.chkUseSpecification.TabIndex = 5
        Me.chkUseSpecification.Text = "Use Specification"
        Me.chkUseSpecification.UseVisualStyleBackColor = True
        '
        'frmTransaksiStockTaking_Print
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(982, 553)
        Me.Controls.Add(Me.chkUseSpecification)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.TxtStockTakingNumber)
        Me.Controls.Add(Me.TxtWarehouseName)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TxtPrimaryLocation)
        Me.Controls.Add(Me.cboLocation)
        Me.Controls.Add(Me.btnFindWarehouse)
        Me.Controls.Add(Me.TxtPrimaryWarehouse)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtWarehouseCode)
        Me.Controls.Add(Me.gboBrand)
        Me.Controls.Add(Me.TxtPrimaryStockTaking)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.crv)
        Me.MinimumSize = New System.Drawing.Size(1000, 600)
        Me.Name = "frmTransaksiStockTaking_Print"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Stock Taking Form"
        Me.gboBrand.ResumeLayout(False)
        Me.gboBrand.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryStockTaking As System.Windows.Forms.TextBox
    Friend WithEvents gboBrand As System.Windows.Forms.GroupBox
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents clbBrand As System.Windows.Forms.CheckedListBox
    Friend WithEvents TxtWarehouseCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryWarehouse As System.Windows.Forms.TextBox
    Friend WithEvents btnFindWarehouse As System.Windows.Forms.Button
    Friend WithEvents cboLocation As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPrimaryLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtWarehouseName As System.Windows.Forms.TextBox
    Friend WithEvents TxtStockTakingNumber As System.Windows.Forms.TextBox
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents chkUseSpecification As System.Windows.Forms.CheckBox
End Class
