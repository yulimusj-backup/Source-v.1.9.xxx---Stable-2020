<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSalesInvoice_QuickByStockOut_Option
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
        Me.ChkDirectlySave = New System.Windows.Forms.CheckBox
        Me.ChkDirectlyPrint = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ChkUseCustomPrinter = New System.Windows.Forms.CheckBox
        Me.BtnFind = New System.Windows.Forms.Button
        Me.TxtUseCustomPrinter = New System.Windows.Forms.TextBox
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.BtnStart = New System.Windows.Forms.Button
        Me.DtpDate = New System.Windows.Forms.DateTimePicker
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.SuspendLayout()
        '
        'ChkDirectlySave
        '
        Me.ChkDirectlySave.AutoSize = True
        Me.ChkDirectlySave.Location = New System.Drawing.Point(9, 58)
        Me.ChkDirectlySave.Name = "ChkDirectlySave"
        Me.ChkDirectlySave.Size = New System.Drawing.Size(89, 17)
        Me.ChkDirectlySave.TabIndex = 5
        Me.ChkDirectlySave.Text = "Directly Save"
        Me.ChkDirectlySave.UseVisualStyleBackColor = True
        '
        'ChkDirectlyPrint
        '
        Me.ChkDirectlyPrint.AutoSize = True
        Me.ChkDirectlyPrint.Location = New System.Drawing.Point(114, 58)
        Me.ChkDirectlyPrint.Name = "ChkDirectlyPrint"
        Me.ChkDirectlyPrint.Size = New System.Drawing.Size(85, 17)
        Me.ChkDirectlyPrint.TabIndex = 6
        Me.ChkDirectlyPrint.Text = "Directly Print"
        Me.ChkDirectlyPrint.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Invoice Date"
        '
        'ChkUseCustomPrinter
        '
        Me.ChkUseCustomPrinter.AutoSize = True
        Me.ChkUseCustomPrinter.Location = New System.Drawing.Point(9, 32)
        Me.ChkUseCustomPrinter.Name = "ChkUseCustomPrinter"
        Me.ChkUseCustomPrinter.Size = New System.Drawing.Size(116, 17)
        Me.ChkUseCustomPrinter.TabIndex = 2
        Me.ChkUseCustomPrinter.Text = "Use Custom Printer"
        Me.ChkUseCustomPrinter.UseVisualStyleBackColor = True
        '
        'BtnFind
        '
        Me.BtnFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFind.Location = New System.Drawing.Point(304, 31)
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(58, 21)
        Me.BtnFind.TabIndex = 4
        Me.BtnFind.Text = "Find"
        Me.BtnFind.UseVisualStyleBackColor = True
        '
        'TxtUseCustomPrinter
        '
        Me.TxtUseCustomPrinter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUseCustomPrinter.Location = New System.Drawing.Point(131, 32)
        Me.TxtUseCustomPrinter.Name = "TxtUseCustomPrinter"
        Me.TxtUseCustomPrinter.ReadOnly = True
        Me.TxtUseCustomPrinter.Size = New System.Drawing.Size(168, 20)
        Me.TxtUseCustomPrinter.TabIndex = 3
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(304, 56)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(58, 21)
        Me.BtnCancel.TabIndex = 8
        Me.BtnCancel.Text = "Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'BtnStart
        '
        Me.BtnStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnStart.Location = New System.Drawing.Point(240, 56)
        Me.BtnStart.Name = "BtnStart"
        Me.BtnStart.Size = New System.Drawing.Size(58, 21)
        Me.BtnStart.TabIndex = 7
        Me.BtnStart.Text = "Start"
        Me.BtnStart.UseVisualStyleBackColor = True
        '
        'DtpDate
        '
        Me.DtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDate.CustomFormat = "dd MMM yyyy"
        Me.DtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDate.Location = New System.Drawing.Point(131, 8)
        Me.DtpDate.Name = "DtpDate"
        Me.DtpDate.Size = New System.Drawing.Size(231, 20)
        Me.DtpDate.TabIndex = 9
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'FrmTransaksiSalesInvoice_QuickByStockOut_Option
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(367, 84)
        Me.Controls.Add(Me.DtpDate)
        Me.Controls.Add(Me.BtnStart)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnFind)
        Me.Controls.Add(Me.TxtUseCustomPrinter)
        Me.Controls.Add(Me.ChkUseCustomPrinter)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ChkDirectlyPrint)
        Me.Controls.Add(Me.ChkDirectlySave)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(383, 122)
        Me.Name = "FrmTransaksiSalesInvoice_QuickByStockOut_Option"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ChkDirectlySave As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDirectlyPrint As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ChkUseCustomPrinter As System.Windows.Forms.CheckBox
    Friend WithEvents BtnFind As System.Windows.Forms.Button
    Friend WithEvents TxtUseCustomPrinter As System.Windows.Forms.TextBox
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents BtnStart As System.Windows.Forms.Button
    Friend WithEvents DtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
End Class
