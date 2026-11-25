<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSalesInvoice_EditNo2
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.mtb_Number = New System.Windows.Forms.MaskedTextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.MaskedTextBox1 = New System.Windows.Forms.MaskedTextBox
        Me.cboJenisFaktur = New System.Windows.Forms.ComboBox
        Me.TxtPrimaryTaxInvoice = New System.Windows.Forms.TextBox
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel
        Me.Label34 = New System.Windows.Forms.Label
        Me.Label33 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_NoTelepon = New System.Windows.Forms.TextBox
        Me.Txt_E_Faktur_KodePos = New System.Windows.Forms.TextBox
        Me.Txt_E_Faktur_Propinsi = New System.Windows.Forms.TextBox
        Me.Label32 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_Kabupaten = New System.Windows.Forms.TextBox
        Me.Label31 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_Kecamatan = New System.Windows.Forms.TextBox
        Me.Label28 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_Kelurahan = New System.Windows.Forms.TextBox
        Me.Label27 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_RW = New System.Windows.Forms.TextBox
        Me.Txt_E_Faktur_RT = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_No = New System.Windows.Forms.TextBox
        Me.Txt_E_Faktur_Blok = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label35 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_Jalan = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(418, 220)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 4
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'mtb_Number
        '
        Me.mtb_Number.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.mtb_Number.Location = New System.Drawing.Point(395, 8)
        Me.mtb_Number.Mask = "00000000"
        Me.mtb_Number.Name = "mtb_Number"
        Me.mtb_Number.Size = New System.Drawing.Size(168, 20)
        Me.mtb_Number.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(305, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "No Faktur Pajak"
        '
        'MaskedTextBox1
        '
        Me.MaskedTextBox1.Location = New System.Drawing.Point(56, 49)
        Me.MaskedTextBox1.Mask = "00000000"
        Me.MaskedTextBox1.Name = "MaskedTextBox1"
        Me.MaskedTextBox1.Size = New System.Drawing.Size(63, 20)
        Me.MaskedTextBox1.TabIndex = 7
        Me.MaskedTextBox1.Visible = False
        '
        'cboJenisFaktur
        '
        Me.cboJenisFaktur.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboJenisFaktur.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboJenisFaktur.FormattingEnabled = True
        Me.cboJenisFaktur.Location = New System.Drawing.Point(12, 7)
        Me.cboJenisFaktur.Name = "cboJenisFaktur"
        Me.cboJenisFaktur.Size = New System.Drawing.Size(287, 21)
        Me.cboJenisFaktur.TabIndex = 0
        '
        'TxtPrimaryTaxInvoice
        '
        Me.TxtPrimaryTaxInvoice.Location = New System.Drawing.Point(162, 34)
        Me.TxtPrimaryTaxInvoice.Name = "TxtPrimaryTaxInvoice"
        Me.TxtPrimaryTaxInvoice.Size = New System.Drawing.Size(100, 20)
        Me.TxtPrimaryTaxInvoice.TabIndex = 9
        Me.TxtPrimaryTaxInvoice.TabStop = False
        Me.TxtPrimaryTaxInvoice.Visible = False
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel2.ColumnCount = 4
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 75.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 182.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 75.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 219.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.Label34, 2, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.Label33, 0, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.Txt_E_Faktur_NoTelepon, 3, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.Txt_E_Faktur_KodePos, 1, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.Txt_E_Faktur_Propinsi, 4, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.Label32, 2, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.Txt_E_Faktur_Kabupaten, 1, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.Label31, 0, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.Txt_E_Faktur_Kecamatan, 3, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Label28, 2, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Txt_E_Faktur_Kelurahan, 1, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Label27, 0, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Txt_E_Faktur_RW, 3, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Txt_E_Faktur_RT, 1, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Label26, 2, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Label18, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Label17, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Txt_E_Faktur_No, 3, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Txt_E_Faktur_Blok, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Label11, 0, 0)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(14, 58)
        Me.TableLayoutPanel2.Margin = New System.Windows.Forms.Padding(2)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 5
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(551, 122)
        Me.TableLayoutPanel2.TabIndex = 3
        '
        'Label34
        '
        Me.Label34.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(260, 102)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(63, 13)
        Me.Label34.TabIndex = 94
        Me.Label34.Text = "No Telepon"
        '
        'Label33
        '
        Me.Label33.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(3, 102)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(53, 13)
        Me.Label33.TabIndex = 93
        Me.Label33.Text = "Kode Pos"
        '
        'Txt_E_Faktur_NoTelepon
        '
        Me.Txt_E_Faktur_NoTelepon.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_NoTelepon.Location = New System.Drawing.Point(335, 99)
        Me.Txt_E_Faktur_NoTelepon.MaxLength = 250
        Me.Txt_E_Faktur_NoTelepon.Name = "Txt_E_Faktur_NoTelepon"
        Me.Txt_E_Faktur_NoTelepon.Size = New System.Drawing.Size(213, 20)
        Me.Txt_E_Faktur_NoTelepon.TabIndex = 9
        '
        'Txt_E_Faktur_KodePos
        '
        Me.Txt_E_Faktur_KodePos.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_KodePos.Location = New System.Drawing.Point(78, 99)
        Me.Txt_E_Faktur_KodePos.MaxLength = 250
        Me.Txt_E_Faktur_KodePos.Name = "Txt_E_Faktur_KodePos"
        Me.Txt_E_Faktur_KodePos.Size = New System.Drawing.Size(176, 20)
        Me.Txt_E_Faktur_KodePos.TabIndex = 8
        '
        'Txt_E_Faktur_Propinsi
        '
        Me.Txt_E_Faktur_Propinsi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_Propinsi.Location = New System.Drawing.Point(335, 75)
        Me.Txt_E_Faktur_Propinsi.MaxLength = 250
        Me.Txt_E_Faktur_Propinsi.Name = "Txt_E_Faktur_Propinsi"
        Me.Txt_E_Faktur_Propinsi.Size = New System.Drawing.Size(213, 20)
        Me.Txt_E_Faktur_Propinsi.TabIndex = 7
        '
        'Label32
        '
        Me.Label32.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(260, 77)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(44, 13)
        Me.Label32.TabIndex = 89
        Me.Label32.Text = "Propinsi"
        '
        'Txt_E_Faktur_Kabupaten
        '
        Me.Txt_E_Faktur_Kabupaten.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_Kabupaten.Location = New System.Drawing.Point(78, 75)
        Me.Txt_E_Faktur_Kabupaten.MaxLength = 250
        Me.Txt_E_Faktur_Kabupaten.Name = "Txt_E_Faktur_Kabupaten"
        Me.Txt_E_Faktur_Kabupaten.Size = New System.Drawing.Size(176, 20)
        Me.Txt_E_Faktur_Kabupaten.TabIndex = 6
        '
        'Label31
        '
        Me.Label31.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(3, 77)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(59, 13)
        Me.Label31.TabIndex = 87
        Me.Label31.Text = "Kabupaten"
        '
        'Txt_E_Faktur_Kecamatan
        '
        Me.Txt_E_Faktur_Kecamatan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_Kecamatan.Location = New System.Drawing.Point(335, 51)
        Me.Txt_E_Faktur_Kecamatan.MaxLength = 250
        Me.Txt_E_Faktur_Kecamatan.Name = "Txt_E_Faktur_Kecamatan"
        Me.Txt_E_Faktur_Kecamatan.Size = New System.Drawing.Size(213, 20)
        Me.Txt_E_Faktur_Kecamatan.TabIndex = 5
        '
        'Label28
        '
        Me.Label28.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(260, 53)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(61, 13)
        Me.Label28.TabIndex = 85
        Me.Label28.Text = "Kecamatan"
        '
        'Txt_E_Faktur_Kelurahan
        '
        Me.Txt_E_Faktur_Kelurahan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_Kelurahan.Location = New System.Drawing.Point(78, 51)
        Me.Txt_E_Faktur_Kelurahan.MaxLength = 250
        Me.Txt_E_Faktur_Kelurahan.Name = "Txt_E_Faktur_Kelurahan"
        Me.Txt_E_Faktur_Kelurahan.Size = New System.Drawing.Size(176, 20)
        Me.Txt_E_Faktur_Kelurahan.TabIndex = 4
        '
        'Label27
        '
        Me.Label27.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(3, 53)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(55, 13)
        Me.Label27.TabIndex = 83
        Me.Label27.Text = "Kelurahan"
        '
        'Txt_E_Faktur_RW
        '
        Me.Txt_E_Faktur_RW.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_RW.Location = New System.Drawing.Point(335, 27)
        Me.Txt_E_Faktur_RW.MaxLength = 250
        Me.Txt_E_Faktur_RW.Name = "Txt_E_Faktur_RW"
        Me.Txt_E_Faktur_RW.Size = New System.Drawing.Size(213, 20)
        Me.Txt_E_Faktur_RW.TabIndex = 3
        '
        'Txt_E_Faktur_RT
        '
        Me.Txt_E_Faktur_RT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_RT.Location = New System.Drawing.Point(78, 27)
        Me.Txt_E_Faktur_RT.MaxLength = 250
        Me.Txt_E_Faktur_RT.Name = "Txt_E_Faktur_RT"
        Me.Txt_E_Faktur_RT.Size = New System.Drawing.Size(176, 20)
        Me.Txt_E_Faktur_RT.TabIndex = 2
        '
        'Label26
        '
        Me.Label26.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(260, 29)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(26, 13)
        Me.Label26.TabIndex = 80
        Me.Label26.Text = "RW"
        '
        'Label18
        '
        Me.Label18.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(3, 29)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(22, 13)
        Me.Label18.TabIndex = 79
        Me.Label18.Text = "RT"
        '
        'Label17
        '
        Me.Label17.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(260, 5)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(21, 13)
        Me.Label17.TabIndex = 78
        Me.Label17.Text = "No"
        '
        'Txt_E_Faktur_No
        '
        Me.Txt_E_Faktur_No.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_No.Location = New System.Drawing.Point(335, 3)
        Me.Txt_E_Faktur_No.MaxLength = 250
        Me.Txt_E_Faktur_No.Name = "Txt_E_Faktur_No"
        Me.Txt_E_Faktur_No.Size = New System.Drawing.Size(213, 20)
        Me.Txt_E_Faktur_No.TabIndex = 1
        '
        'Txt_E_Faktur_Blok
        '
        Me.Txt_E_Faktur_Blok.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_Blok.Location = New System.Drawing.Point(78, 3)
        Me.Txt_E_Faktur_Blok.MaxLength = 250
        Me.Txt_E_Faktur_Blok.Name = "Txt_E_Faktur_Blok"
        Me.Txt_E_Faktur_Blok.Size = New System.Drawing.Size(176, 20)
        Me.Txt_E_Faktur_Blok.TabIndex = 0
        '
        'Label11
        '
        Me.Label11.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 5)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(28, 13)
        Me.Label11.TabIndex = 75
        Me.Label11.Text = "Blok"
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(15, 36)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(32, 13)
        Me.Label35.TabIndex = 80
        Me.Label35.Text = "Jalan"
        '
        'Txt_E_Faktur_Jalan
        '
        Me.Txt_E_Faktur_Jalan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_Jalan.Location = New System.Drawing.Point(90, 33)
        Me.Txt_E_Faktur_Jalan.MaxLength = 250
        Me.Txt_E_Faktur_Jalan.Name = "Txt_E_Faktur_Jalan"
        Me.Txt_E_Faktur_Jalan.Size = New System.Drawing.Size(472, 20)
        Me.Txt_E_Faktur_Jalan.TabIndex = 2
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(90, 185)
        Me.TxtReadOnly.MaxLength = 250
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(120, 20)
        Me.TxtReadOnly.TabIndex = 81
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'FrmTransaksiSalesInvoice_EditNo2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(571, 258)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.Label35)
        Me.Controls.Add(Me.Txt_E_Faktur_Jalan)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.Controls.Add(Me.TxtPrimaryTaxInvoice)
        Me.Controls.Add(Me.cboJenisFaktur)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.mtb_Number)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MaskedTextBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(544, 263)
        Me.Name = "FrmTransaksiSalesInvoice_EditNo2"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents mtb_Number As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents MaskedTextBox1 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents cboJenisFaktur As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPrimaryTaxInvoice As System.Windows.Forms.TextBox
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_NoTelepon As System.Windows.Forms.TextBox
    Friend WithEvents Txt_E_Faktur_KodePos As System.Windows.Forms.TextBox
    Friend WithEvents Txt_E_Faktur_Propinsi As System.Windows.Forms.TextBox
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_Kabupaten As System.Windows.Forms.TextBox
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_Kecamatan As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_Kelurahan As System.Windows.Forms.TextBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_RW As System.Windows.Forms.TextBox
    Friend WithEvents Txt_E_Faktur_RT As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_No As System.Windows.Forms.TextBox
    Friend WithEvents Txt_E_Faktur_Blok As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_Jalan As System.Windows.Forms.TextBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
End Class
