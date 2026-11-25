<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSalesInvoice_KeteranganTambahanEFaktur
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtTaxInvoiceSeries = New System.Windows.Forms.TextBox
        Me.cboKeteranganTambahan = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(224, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Your tax invoice series number is: "
        '
        'TxtTaxInvoiceSeries
        '
        Me.TxtTaxInvoiceSeries.Location = New System.Drawing.Point(242, 6)
        Me.TxtTaxInvoiceSeries.Name = "TxtTaxInvoiceSeries"
        Me.TxtTaxInvoiceSeries.ReadOnly = True
        Me.TxtTaxInvoiceSeries.Size = New System.Drawing.Size(162, 22)
        Me.TxtTaxInvoiceSeries.TabIndex = 1
        '
        'cboKeteranganTambahan
        '
        Me.cboKeteranganTambahan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboKeteranganTambahan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKeteranganTambahan.FormattingEnabled = True
        Me.cboKeteranganTambahan.Location = New System.Drawing.Point(15, 57)
        Me.cboKeteranganTambahan.Name = "cboKeteranganTambahan"
        Me.cboKeteranganTambahan.Size = New System.Drawing.Size(536, 24)
        Me.cboKeteranganTambahan.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 37)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(392, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Please specify additional information for E-Faktur data export"
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(479, 92)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmTransaksiSalesInvoice_KeteranganTambahanEFaktur
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(566, 127)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.cboKeteranganTambahan)
        Me.Controls.Add(Me.TxtTaxInvoiceSeries)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(427, 164)
        Me.Name = "FrmTransaksiSalesInvoice_KeteranganTambahanEFaktur"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Additional E-Faktur Data"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Public WithEvents TxtTaxInvoiceSeries As System.Windows.Forms.TextBox
    Public WithEvents cboKeteranganTambahan As System.Windows.Forms.ComboBox
End Class
