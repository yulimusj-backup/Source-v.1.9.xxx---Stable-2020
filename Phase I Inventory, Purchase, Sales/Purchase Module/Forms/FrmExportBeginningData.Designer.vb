<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmExportBeginningData
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
        Me.Label10 = New System.Windows.Forms.Label
        Me.btnExportDebitNote = New System.Windows.Forms.Button
        Me.prgBarDebitNote = New System.Windows.Forms.ProgressBar
        Me.Label11 = New System.Windows.Forms.Label
        Me.btnExportCreditNote = New System.Windows.Forms.Button
        Me.prgBarCreditNote = New System.Windows.Forms.ProgressBar
        Me.Label8 = New System.Windows.Forms.Label
        Me.btnExportPO = New System.Windows.Forms.Button
        Me.prgBarPO = New System.Windows.Forms.ProgressBar
        Me.Label9 = New System.Windows.Forms.Label
        Me.btnExportSO = New System.Windows.Forms.Button
        Me.prgBarSO = New System.Windows.Forms.ProgressBar
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnExportPurchaseDP = New System.Windows.Forms.Button
        Me.prgBarPurchaseDP = New System.Windows.Forms.ProgressBar
        Me.Label7 = New System.Windows.Forms.Label
        Me.btnExportSalesDP = New System.Windows.Forms.Button
        Me.prgBarSalesDP = New System.Windows.Forms.ProgressBar
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnExportPurchaseReturn = New System.Windows.Forms.Button
        Me.prgBarPurchaseReturn = New System.Windows.Forms.ProgressBar
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnExportSalesReturn = New System.Windows.Forms.Button
        Me.prgBarSalesReturn = New System.Windows.Forms.ProgressBar
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnExportAR = New System.Windows.Forms.Button
        Me.prgBarAR = New System.Windows.Forms.ProgressBar
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnExportAP = New System.Windows.Forms.Button
        Me.prgBarAP = New System.Windows.Forms.ProgressBar
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnExportProduct = New System.Windows.Forms.Button
        Me.prgBarProduct = New System.Windows.Forms.ProgressBar
        Me.lblProgress = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.bgwProduct = New System.ComponentModel.BackgroundWorker
        Me.bgwAP = New System.ComponentModel.BackgroundWorker
        Me.bgwAR = New System.ComponentModel.BackgroundWorker
        Me.bgwSalesReturn = New System.ComponentModel.BackgroundWorker
        Me.bgwPurchaseReturn = New System.ComponentModel.BackgroundWorker
        Me.bgwSalesDP = New System.ComponentModel.BackgroundWorker
        Me.bgwPurchaseDP = New System.ComponentModel.BackgroundWorker
        Me.bgwSO = New System.ComponentModel.BackgroundWorker
        Me.bgwPO = New System.ComponentModel.BackgroundWorker
        Me.bgwCreditNote = New System.ComponentModel.BackgroundWorker
        Me.bgwDebitNote = New System.ComponentModel.BackgroundWorker
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.btnExportDebitNote)
        Me.GroupBox1.Controls.Add(Me.prgBarDebitNote)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.btnExportCreditNote)
        Me.GroupBox1.Controls.Add(Me.prgBarCreditNote)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.btnExportPO)
        Me.GroupBox1.Controls.Add(Me.prgBarPO)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.btnExportSO)
        Me.GroupBox1.Controls.Add(Me.prgBarSO)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.btnExportPurchaseDP)
        Me.GroupBox1.Controls.Add(Me.prgBarPurchaseDP)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.btnExportSalesDP)
        Me.GroupBox1.Controls.Add(Me.prgBarSalesDP)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.btnExportPurchaseReturn)
        Me.GroupBox1.Controls.Add(Me.prgBarPurchaseReturn)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.btnExportSalesReturn)
        Me.GroupBox1.Controls.Add(Me.prgBarSalesReturn)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.btnExportAR)
        Me.GroupBox1.Controls.Add(Me.prgBarAR)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.btnExportAP)
        Me.GroupBox1.Controls.Add(Me.prgBarAP)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.btnExportProduct)
        Me.GroupBox1.Controls.Add(Me.prgBarProduct)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(457, 299)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 434)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(100, 13)
        Me.Label10.TabIndex = 33
        Me.Label10.Text = "Balance Debit Note"
        Me.Label10.Visible = False
        '
        'btnExportDebitNote
        '
        Me.btnExportDebitNote.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportDebitNote.Location = New System.Drawing.Point(376, 429)
        Me.btnExportDebitNote.Name = "btnExportDebitNote"
        Me.btnExportDebitNote.Size = New System.Drawing.Size(75, 23)
        Me.btnExportDebitNote.TabIndex = 31
        Me.btnExportDebitNote.Text = "Export"
        Me.btnExportDebitNote.UseVisualStyleBackColor = True
        Me.btnExportDebitNote.Visible = False
        '
        'prgBarDebitNote
        '
        Me.prgBarDebitNote.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBarDebitNote.Location = New System.Drawing.Point(141, 429)
        Me.prgBarDebitNote.Name = "prgBarDebitNote"
        Me.prgBarDebitNote.Size = New System.Drawing.Size(229, 23)
        Me.prgBarDebitNote.TabIndex = 32
        Me.prgBarDebitNote.Visible = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(6, 392)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(102, 13)
        Me.Label11.TabIndex = 30
        Me.Label11.Text = "Balance Credit Note"
        Me.Label11.Visible = False
        '
        'btnExportCreditNote
        '
        Me.btnExportCreditNote.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportCreditNote.Location = New System.Drawing.Point(376, 387)
        Me.btnExportCreditNote.Name = "btnExportCreditNote"
        Me.btnExportCreditNote.Size = New System.Drawing.Size(75, 23)
        Me.btnExportCreditNote.TabIndex = 28
        Me.btnExportCreditNote.Text = "Export"
        Me.btnExportCreditNote.UseVisualStyleBackColor = True
        Me.btnExportCreditNote.Visible = False
        '
        'prgBarCreditNote
        '
        Me.prgBarCreditNote.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBarCreditNote.Location = New System.Drawing.Point(141, 387)
        Me.prgBarCreditNote.Name = "prgBarCreditNote"
        Me.prgBarCreditNote.Size = New System.Drawing.Size(229, 23)
        Me.prgBarCreditNote.TabIndex = 29
        Me.prgBarCreditNote.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 104)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(64, 13)
        Me.Label8.TabIndex = 27
        Me.Label8.Text = "Balance PO"
        '
        'btnExportPO
        '
        Me.btnExportPO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportPO.Location = New System.Drawing.Point(376, 99)
        Me.btnExportPO.Name = "btnExportPO"
        Me.btnExportPO.Size = New System.Drawing.Size(75, 23)
        Me.btnExportPO.TabIndex = 25
        Me.btnExportPO.Text = "Export"
        Me.btnExportPO.UseVisualStyleBackColor = True
        '
        'prgBarPO
        '
        Me.prgBarPO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBarPO.Location = New System.Drawing.Point(141, 99)
        Me.prgBarPO.Name = "prgBarPO"
        Me.prgBarPO.Size = New System.Drawing.Size(229, 23)
        Me.prgBarPO.TabIndex = 26
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 64)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(64, 13)
        Me.Label9.TabIndex = 24
        Me.Label9.Text = "Balance SO"
        '
        'btnExportSO
        '
        Me.btnExportSO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportSO.Location = New System.Drawing.Point(376, 59)
        Me.btnExportSO.Name = "btnExportSO"
        Me.btnExportSO.Size = New System.Drawing.Size(75, 23)
        Me.btnExportSO.TabIndex = 22
        Me.btnExportSO.Text = "Export"
        Me.btnExportSO.UseVisualStyleBackColor = True
        '
        'prgBarSO
        '
        Me.prgBarSO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBarSO.Location = New System.Drawing.Point(141, 59)
        Me.prgBarSO.Name = "prgBarSO"
        Me.prgBarSO.Size = New System.Drawing.Size(229, 23)
        Me.prgBarSO.TabIndex = 23
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 352)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 13)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Balance Purchase DP"
        Me.Label6.Visible = False
        '
        'btnExportPurchaseDP
        '
        Me.btnExportPurchaseDP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportPurchaseDP.Location = New System.Drawing.Point(376, 347)
        Me.btnExportPurchaseDP.Name = "btnExportPurchaseDP"
        Me.btnExportPurchaseDP.Size = New System.Drawing.Size(75, 23)
        Me.btnExportPurchaseDP.TabIndex = 19
        Me.btnExportPurchaseDP.Text = "Export"
        Me.btnExportPurchaseDP.UseVisualStyleBackColor = True
        Me.btnExportPurchaseDP.Visible = False
        '
        'prgBarPurchaseDP
        '
        Me.prgBarPurchaseDP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBarPurchaseDP.Location = New System.Drawing.Point(141, 347)
        Me.prgBarPurchaseDP.Name = "prgBarPurchaseDP"
        Me.prgBarPurchaseDP.Size = New System.Drawing.Size(229, 23)
        Me.prgBarPurchaseDP.TabIndex = 20
        Me.prgBarPurchaseDP.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 310)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(93, 13)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Balance Sales DP"
        Me.Label7.Visible = False
        '
        'btnExportSalesDP
        '
        Me.btnExportSalesDP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportSalesDP.Location = New System.Drawing.Point(376, 305)
        Me.btnExportSalesDP.Name = "btnExportSalesDP"
        Me.btnExportSalesDP.Size = New System.Drawing.Size(75, 23)
        Me.btnExportSalesDP.TabIndex = 16
        Me.btnExportSalesDP.Text = "Export"
        Me.btnExportSalesDP.UseVisualStyleBackColor = True
        Me.btnExportSalesDP.Visible = False
        '
        'prgBarSalesDP
        '
        Me.prgBarSalesDP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBarSalesDP.Location = New System.Drawing.Point(141, 305)
        Me.prgBarSalesDP.Name = "prgBarSalesDP"
        Me.prgBarSalesDP.Size = New System.Drawing.Size(229, 23)
        Me.prgBarSalesDP.TabIndex = 17
        Me.prgBarSalesDP.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 264)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(129, 13)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Balance Purchase Return"
        '
        'btnExportPurchaseReturn
        '
        Me.btnExportPurchaseReturn.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportPurchaseReturn.Location = New System.Drawing.Point(376, 259)
        Me.btnExportPurchaseReturn.Name = "btnExportPurchaseReturn"
        Me.btnExportPurchaseReturn.Size = New System.Drawing.Size(75, 23)
        Me.btnExportPurchaseReturn.TabIndex = 13
        Me.btnExportPurchaseReturn.Text = "Export"
        Me.btnExportPurchaseReturn.UseVisualStyleBackColor = True
        '
        'prgBarPurchaseReturn
        '
        Me.prgBarPurchaseReturn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBarPurchaseReturn.Location = New System.Drawing.Point(141, 259)
        Me.prgBarPurchaseReturn.Name = "prgBarPurchaseReturn"
        Me.prgBarPurchaseReturn.Size = New System.Drawing.Size(229, 23)
        Me.prgBarPurchaseReturn.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 224)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(110, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Balance Sales Return"
        '
        'btnExportSalesReturn
        '
        Me.btnExportSalesReturn.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportSalesReturn.Location = New System.Drawing.Point(376, 219)
        Me.btnExportSalesReturn.Name = "btnExportSalesReturn"
        Me.btnExportSalesReturn.Size = New System.Drawing.Size(75, 23)
        Me.btnExportSalesReturn.TabIndex = 10
        Me.btnExportSalesReturn.Text = "Export"
        Me.btnExportSalesReturn.UseVisualStyleBackColor = True
        '
        'prgBarSalesReturn
        '
        Me.prgBarSalesReturn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBarSalesReturn.Location = New System.Drawing.Point(141, 219)
        Me.prgBarSalesReturn.Name = "prgBarSalesReturn"
        Me.prgBarSalesReturn.Size = New System.Drawing.Size(229, 23)
        Me.prgBarSalesReturn.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 144)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Balance Sales"
        '
        'btnExportAR
        '
        Me.btnExportAR.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportAR.Location = New System.Drawing.Point(376, 139)
        Me.btnExportAR.Name = "btnExportAR"
        Me.btnExportAR.Size = New System.Drawing.Size(75, 23)
        Me.btnExportAR.TabIndex = 7
        Me.btnExportAR.Text = "Export"
        Me.btnExportAR.UseVisualStyleBackColor = True
        '
        'prgBarAR
        '
        Me.prgBarAR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBarAR.Location = New System.Drawing.Point(141, 139)
        Me.prgBarAR.Name = "prgBarAR"
        Me.prgBarAR.Size = New System.Drawing.Size(229, 23)
        Me.prgBarAR.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 184)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(94, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Balance Purchase"
        '
        'btnExportAP
        '
        Me.btnExportAP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportAP.Location = New System.Drawing.Point(376, 179)
        Me.btnExportAP.Name = "btnExportAP"
        Me.btnExportAP.Size = New System.Drawing.Size(75, 23)
        Me.btnExportAP.TabIndex = 4
        Me.btnExportAP.Text = "Export"
        Me.btnExportAP.UseVisualStyleBackColor = True
        '
        'prgBarAP
        '
        Me.prgBarAP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBarAP.Location = New System.Drawing.Point(141, 179)
        Me.prgBarAP.Name = "prgBarAP"
        Me.prgBarAP.Size = New System.Drawing.Size(229, 23)
        Me.prgBarAP.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Balance Product"
        '
        'btnExportProduct
        '
        Me.btnExportProduct.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportProduct.Location = New System.Drawing.Point(376, 19)
        Me.btnExportProduct.Name = "btnExportProduct"
        Me.btnExportProduct.Size = New System.Drawing.Size(75, 23)
        Me.btnExportProduct.TabIndex = 0
        Me.btnExportProduct.Text = "Export"
        Me.btnExportProduct.UseVisualStyleBackColor = True
        '
        'prgBarProduct
        '
        Me.prgBarProduct.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBarProduct.Location = New System.Drawing.Point(141, 19)
        Me.prgBarProduct.Name = "prgBarProduct"
        Me.prgBarProduct.Size = New System.Drawing.Size(229, 23)
        Me.prgBarProduct.TabIndex = 2
        '
        'lblProgress
        '
        Me.lblProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(9, 322)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(38, 13)
        Me.lblProgress.TabIndex = 1
        Me.lblProgress.Text = "Ready"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(388, 317)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'FrmExportBeginningData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(481, 352)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(489, 383)
        Me.Name = "FrmExportBeginningData"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Export Beginning Data"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents prgBarProduct As System.Windows.Forms.ProgressBar
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents bgwProduct As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnExportProduct As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnExportAP As System.Windows.Forms.Button
    Friend WithEvents prgBarAP As System.Windows.Forms.ProgressBar
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnExportAR As System.Windows.Forms.Button
    Friend WithEvents prgBarAR As System.Windows.Forms.ProgressBar
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnExportPurchaseReturn As System.Windows.Forms.Button
    Friend WithEvents prgBarPurchaseReturn As System.Windows.Forms.ProgressBar
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnExportSalesReturn As System.Windows.Forms.Button
    Friend WithEvents prgBarSalesReturn As System.Windows.Forms.ProgressBar
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnExportPurchaseDP As System.Windows.Forms.Button
    Friend WithEvents prgBarPurchaseDP As System.Windows.Forms.ProgressBar
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnExportSalesDP As System.Windows.Forms.Button
    Friend WithEvents prgBarSalesDP As System.Windows.Forms.ProgressBar
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnExportPO As System.Windows.Forms.Button
    Friend WithEvents prgBarPO As System.Windows.Forms.ProgressBar
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnExportSO As System.Windows.Forms.Button
    Friend WithEvents prgBarSO As System.Windows.Forms.ProgressBar
    Friend WithEvents bgwAP As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwAR As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwSalesReturn As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwPurchaseReturn As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwSalesDP As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwPurchaseDP As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwSO As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwPO As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents btnExportDebitNote As System.Windows.Forms.Button
    Friend WithEvents prgBarDebitNote As System.Windows.Forms.ProgressBar
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents btnExportCreditNote As System.Windows.Forms.Button
    Friend WithEvents prgBarCreditNote As System.Windows.Forms.ProgressBar
    Friend WithEvents bgwCreditNote As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwDebitNote As System.ComponentModel.BackgroundWorker
End Class
