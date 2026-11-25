<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanSupplierList
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkBankaccount = New System.Windows.Forms.CheckBox
        Me.chkConPerson = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.crvSupplierList = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.cboGrupSupplier = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkMul_address = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.chkBankaccount)
        Me.GroupBox1.Controls.Add(Me.chkConPerson)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.btnShow)
        Me.GroupBox1.Controls.Add(Me.cboGrupSupplier)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.chkMul_address)
        Me.GroupBox1.Location = New System.Drawing.Point(5, -2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(775, 506)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'chkBankaccount
        '
        Me.chkBankaccount.AutoSize = True
        Me.chkBankaccount.Location = New System.Drawing.Point(246, 56)
        Me.chkBankaccount.Name = "chkBankaccount"
        Me.chkBankaccount.Size = New System.Drawing.Size(163, 17)
        Me.chkBankaccount.TabIndex = 7
        Me.chkBankaccount.Text = "Show Multiple Bank Account"
        Me.chkBankaccount.UseVisualStyleBackColor = True
        '
        'chkConPerson
        '
        Me.chkConPerson.AutoSize = True
        Me.chkConPerson.Location = New System.Drawing.Point(246, 33)
        Me.chkConPerson.Name = "chkConPerson"
        Me.chkConPerson.Size = New System.Drawing.Size(168, 17)
        Me.chkConPerson.TabIndex = 6
        Me.chkConPerson.Text = "Show Multiple Contact Person"
        Me.chkConPerson.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.crvSupplierList)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 66)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(763, 434)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        '
        'crvSupplierList
        '
        Me.crvSupplierList.ActiveViewIndex = -1
        Me.crvSupplierList.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvSupplierList.DisplayGroupTree = False
        Me.crvSupplierList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvSupplierList.EnableDrillDown = False
        Me.crvSupplierList.Location = New System.Drawing.Point(3, 16)
        Me.crvSupplierList.Name = "crvSupplierList"
        Me.crvSupplierList.SelectionFormula = ""
        Me.crvSupplierList.ShowCloseButton = False
        Me.crvSupplierList.ShowGroupTreeButton = False
        Me.crvSupplierList.ShowRefreshButton = False
        Me.crvSupplierList.Size = New System.Drawing.Size(757, 415)
        Me.crvSupplierList.TabIndex = 0
        Me.crvSupplierList.ViewTimeSelectionFormula = ""
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(145, 37)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(64, 37)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(75, 23)
        Me.btnShow.TabIndex = 3
        Me.btnShow.Text = "&Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'cboGrupSupplier
        '
        Me.cboGrupSupplier.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGrupSupplier.FormattingEnabled = True
        Me.cboGrupSupplier.Items.AddRange(New Object() {"Local Only", "Foreign Only", "Both"})
        Me.cboGrupSupplier.Location = New System.Drawing.Point(64, 10)
        Me.cboGrupSupplier.Name = "cboGrupSupplier"
        Me.cboGrupSupplier.Size = New System.Drawing.Size(156, 21)
        Me.cboGrupSupplier.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Supplier"
        '
        'chkMul_address
        '
        Me.chkMul_address.AutoSize = True
        Me.chkMul_address.Location = New System.Drawing.Point(246, 10)
        Me.chkMul_address.Name = "chkMul_address"
        Me.chkMul_address.Size = New System.Drawing.Size(133, 17)
        Me.chkMul_address.TabIndex = 0
        Me.chkMul_address.Text = "Show Multiple Address"
        Me.chkMul_address.UseVisualStyleBackColor = True
        '
        'FrmLaporanSupplierList
        '
        Me.AcceptButton = Me.btnShow
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(784, 511)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "FrmLaporanSupplierList"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Supplier List Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents crvSupplierList As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents cboGrupSupplier As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkMul_address As System.Windows.Forms.CheckBox
    Friend WithEvents chkBankaccount As System.Windows.Forms.CheckBox
    Friend WithEvents chkConPerson As System.Windows.Forms.CheckBox
End Class
