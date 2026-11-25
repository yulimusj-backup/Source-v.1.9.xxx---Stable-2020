<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonthlyClosing_DetailAkun
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
        Me.chkUseSupplierInvoiceNumber = New System.Windows.Forms.CheckBox
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Button1 = New System.Windows.Forms.Button
        Me.chkShowReferenceParty = New System.Windows.Forms.CheckBox
        Me.cboProductDetail = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboFormat = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnLoad = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'chkUseSupplierInvoiceNumber
        '
        Me.chkUseSupplierInvoiceNumber.AutoSize = True
        Me.chkUseSupplierInvoiceNumber.Location = New System.Drawing.Point(326, 14)
        Me.chkUseSupplierInvoiceNumber.Name = "chkUseSupplierInvoiceNumber"
        Me.chkUseSupplierInvoiceNumber.Size = New System.Drawing.Size(213, 21)
        Me.chkUseSupplierInvoiceNumber.TabIndex = 1
        Me.chkUseSupplierInvoiceNumber.Text = "Use Supplier Invoice Number"
        Me.chkUseSupplierInvoiceNumber.UseVisualStyleBackColor = True
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.Location = New System.Drawing.Point(0, 73)
        Me.crv1.Margin = New System.Windows.Forms.Padding(4)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(992, 554)
        Me.crv1.TabIndex = 5
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(392, 136)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(172, 34)
        Me.Button1.TabIndex = 8
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'chkShowReferenceParty
        '
        Me.chkShowReferenceParty.AutoSize = True
        Me.chkShowReferenceParty.Location = New System.Drawing.Point(563, 14)
        Me.chkShowReferenceParty.Name = "chkShowReferenceParty"
        Me.chkShowReferenceParty.Size = New System.Drawing.Size(171, 21)
        Me.chkShowReferenceParty.TabIndex = 2
        Me.chkShowReferenceParty.Text = "Show Reference Party"
        Me.chkShowReferenceParty.UseVisualStyleBackColor = True
        '
        'cboProductDetail
        '
        Me.cboProductDetail.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProductDetail.FormattingEnabled = True
        Me.cboProductDetail.Items.AddRange(New Object() {"Code No", "Description", "Code No && Description"})
        Me.cboProductDetail.Location = New System.Drawing.Point(115, 42)
        Me.cboProductDetail.Name = "cboProductDetail"
        Me.cboProductDetail.Size = New System.Drawing.Size(205, 24)
        Me.cboProductDetail.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Format"
        '
        'cboFormat
        '
        Me.cboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat.FormattingEnabled = True
        Me.cboFormat.Items.AddRange(New Object() {"Recap", "Detail"})
        Me.cboFormat.Location = New System.Drawing.Point(115, 12)
        Me.cboFormat.Name = "cboFormat"
        Me.cboFormat.Size = New System.Drawing.Size(205, 24)
        Me.cboFormat.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(97, 17)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Product Detail"
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(326, 41)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(94, 23)
        Me.btnLoad.TabIndex = 4
        Me.btnLoad.Text = "Refresh"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'FrmTransaksiMonthlyClosing_DetailAkun
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.btnLoad)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboFormat)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboProductDetail)
        Me.Controls.Add(Me.chkShowReferenceParty)
        Me.Controls.Add(Me.crv1)
        Me.Controls.Add(Me.chkUseSupplierInvoiceNumber)
        Me.Controls.Add(Me.Button1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmTransaksiMonthlyClosing_DetailAkun"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Product"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkUseSupplierInvoiceNumber As System.Windows.Forms.CheckBox
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents chkShowReferenceParty As System.Windows.Forms.CheckBox
    Friend WithEvents cboProductDetail As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnLoad As System.Windows.Forms.Button
End Class
