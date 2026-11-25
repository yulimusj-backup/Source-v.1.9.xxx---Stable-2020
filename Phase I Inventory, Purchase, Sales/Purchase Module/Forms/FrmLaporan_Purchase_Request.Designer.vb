<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporan_Purchase_Request
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.btnFindProduct = New System.Windows.Forms.Button
        Me.BtnFindRefParty = New System.Windows.Forms.Button
        Me.ChkAllProduct = New System.Windows.Forms.CheckBox
        Me.ChkOB = New System.Windows.Forms.CheckBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.chkShowEstimateDelivery = New System.Windows.Forms.CheckBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.chkShowPONumber = New System.Windows.Forms.CheckBox
        Me.TxtPrimaryPR = New System.Windows.Forms.TextBox
        Me.TxtPRNo = New System.Windows.Forms.TextBox
        Me.CboPilih = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnFindPR = New System.Windows.Forms.Button
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.Label1 = New System.Windows.Forms.Label
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.TxtDescRefParty = New System.Windows.Forms.TextBox
        Me.TxtPrimaryRefParty = New System.Windows.Forms.TextBox
        Me.TxtPrimaryProduct = New System.Windows.Forms.TextBox
        Me.TxtRequestBy = New System.Windows.Forms.TextBox
        Me.TxtDescProduct = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtProductCode = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TabUserKelompok = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.DgvUser = New System.Windows.Forms.DataGridView
        Me.ColCreatedBy_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BgwUser = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TabUserKelompok.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(370, 28)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(56, 20)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(370, 3)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(56, 20)
        Me.btnFilter.TabIndex = 5
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'crv
        '
        Me.crv.ActiveViewIndex = -1
        Me.crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv.DisplayGroupTree = False
        Me.crv.EnableDrillDown = False
        Me.crv.Location = New System.Drawing.Point(16, 175)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(871, 352)
        Me.crv.TabIndex = 9
        Me.crv.TabStop = False
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'btnFindProduct
        '
        Me.btnFindProduct.Location = New System.Drawing.Point(198, 38)
        Me.btnFindProduct.Name = "btnFindProduct"
        Me.btnFindProduct.Size = New System.Drawing.Size(31, 18)
        Me.btnFindProduct.TabIndex = 2
        Me.btnFindProduct.TabStop = False
        Me.btnFindProduct.Text = "..."
        Me.btnFindProduct.UseVisualStyleBackColor = True
        '
        'BtnFindRefParty
        '
        Me.BtnFindRefParty.Location = New System.Drawing.Point(198, 62)
        Me.BtnFindRefParty.Name = "BtnFindRefParty"
        Me.BtnFindRefParty.Size = New System.Drawing.Size(31, 18)
        Me.BtnFindRefParty.TabIndex = 5
        Me.BtnFindRefParty.TabStop = False
        Me.BtnFindRefParty.Text = "..."
        Me.BtnFindRefParty.UseVisualStyleBackColor = True
        '
        'ChkAllProduct
        '
        Me.ChkAllProduct.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ChkAllProduct.AutoSize = True
        Me.ChkAllProduct.Location = New System.Drawing.Point(296, 87)
        Me.ChkAllProduct.Name = "ChkAllProduct"
        Me.ChkAllProduct.Size = New System.Drawing.Size(168, 17)
        Me.ChkAllProduct.TabIndex = 10
        Me.ChkAllProduct.TabStop = False
        Me.ChkAllProduct.Text = "Always display all product item"
        Me.ChkAllProduct.UseVisualStyleBackColor = True
        Me.ChkAllProduct.Visible = False
        '
        'ChkOB
        '
        Me.ChkOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ChkOB.AutoSize = True
        Me.ChkOB.Location = New System.Drawing.Point(6, 110)
        Me.ChkOB.Name = "ChkOB"
        Me.ChkOB.Size = New System.Drawing.Size(268, 17)
        Me.ChkOB.TabIndex = 12
        Me.ChkOB.Text = "Only show PR with purchase request order balance"
        Me.ChkOB.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(16, 12)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkShowEstimateDelivery)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkShowPONumber)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryPR)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkOB)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPRNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboPilih)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkAllProduct)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindPR)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TableLayoutPanel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDescRefParty)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryRefParty)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryProduct)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRequestBy)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDescProduct)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindRefParty)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtProductCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindProduct)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TabUserKelompok)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFilter)
        Me.SplitContainer1.Size = New System.Drawing.Size(871, 156)
        Me.SplitContainer1.SplitterDistance = 438
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'chkShowEstimateDelivery
        '
        Me.chkShowEstimateDelivery.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkShowEstimateDelivery.AutoSize = True
        Me.chkShowEstimateDelivery.Location = New System.Drawing.Point(182, 129)
        Me.chkShowEstimateDelivery.Name = "chkShowEstimateDelivery"
        Me.chkShowEstimateDelivery.Size = New System.Drawing.Size(206, 17)
        Me.chkShowEstimateDelivery.TabIndex = 14
        Me.chkShowEstimateDelivery.Text = "Show Estimate Delivery On Availability"
        Me.chkShowEstimateDelivery.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(232, 89)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Auth. Status"
        '
        'chkShowPONumber
        '
        Me.chkShowPONumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkShowPONumber.AutoSize = True
        Me.chkShowPONumber.Location = New System.Drawing.Point(6, 129)
        Me.chkShowPONumber.Name = "chkShowPONumber"
        Me.chkShowPONumber.Size = New System.Drawing.Size(180, 17)
        Me.chkShowPONumber.TabIndex = 13
        Me.chkShowPONumber.Text = "Show PO Number On Availability"
        Me.chkShowPONumber.UseVisualStyleBackColor = True
        '
        'TxtPrimaryPR
        '
        Me.TxtPrimaryPR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryPR.Location = New System.Drawing.Point(6, 82)
        Me.TxtPrimaryPR.Name = "TxtPrimaryPR"
        Me.TxtPrimaryPR.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimaryPR.TabIndex = 37
        Me.TxtPrimaryPR.Text = "-1"
        Me.TxtPrimaryPR.Visible = False
        '
        'TxtPRNo
        '
        Me.TxtPRNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPRNo.Location = New System.Drawing.Point(82, 86)
        Me.TxtPRNo.Name = "TxtPRNo"
        Me.TxtPRNo.Size = New System.Drawing.Size(110, 20)
        Me.TxtPRNo.TabIndex = 7
        '
        'CboPilih
        '
        Me.CboPilih.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboPilih.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPilih.FormattingEnabled = True
        Me.CboPilih.Items.AddRange(New Object() {"All Data", "Authorize Only", "Unauthorize Only"})
        Me.CboPilih.Location = New System.Drawing.Point(302, 85)
        Me.CboPilih.Name = "CboPilih"
        Me.CboPilih.Size = New System.Drawing.Size(133, 21)
        Me.CboPilih.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 89)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(75, 13)
        Me.Label2.TabIndex = 36
        Me.Label2.Text = "Purc. Req. No"
        '
        'btnFindPR
        '
        Me.btnFindPR.Location = New System.Drawing.Point(198, 86)
        Me.btnFindPR.Name = "btnFindPR"
        Me.btnFindPR.Size = New System.Drawing.Size(31, 18)
        Me.btnFindPR.TabIndex = 8
        Me.btnFindPR.TabStop = False
        Me.btnFindPR.Text = "..."
        Me.btnFindPR.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.AutoSize = True
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.DtpTo, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.DtpFrom, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(82, 7)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(352, 28)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(166, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Padding = New System.Windows.Forms.Padding(0, 5, 0, 0)
        Me.Label1.Size = New System.Drawing.Size(20, 18)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "To"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DtpTo
        '
        Me.DtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(192, 3)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(157, 20)
        Me.DtpTo.TabIndex = 1
        '
        'DtpFrom
        '
        Me.DtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(3, 3)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(157, 20)
        Me.DtpFrom.TabIndex = 0
        '
        'TxtDescRefParty
        '
        Me.TxtDescRefParty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescRefParty.Location = New System.Drawing.Point(235, 62)
        Me.TxtDescRefParty.Name = "TxtDescRefParty"
        Me.TxtDescRefParty.ReadOnly = True
        Me.TxtDescRefParty.Size = New System.Drawing.Size(200, 20)
        Me.TxtDescRefParty.TabIndex = 6
        Me.TxtDescRefParty.TabStop = False
        '
        'TxtPrimaryRefParty
        '
        Me.TxtPrimaryRefParty.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryRefParty.Location = New System.Drawing.Point(6, 58)
        Me.TxtPrimaryRefParty.Name = "TxtPrimaryRefParty"
        Me.TxtPrimaryRefParty.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimaryRefParty.TabIndex = 32
        Me.TxtPrimaryRefParty.Text = "-1"
        Me.TxtPrimaryRefParty.Visible = False
        '
        'TxtPrimaryProduct
        '
        Me.TxtPrimaryProduct.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryProduct.Location = New System.Drawing.Point(6, 34)
        Me.TxtPrimaryProduct.Name = "TxtPrimaryProduct"
        Me.TxtPrimaryProduct.Size = New System.Drawing.Size(33, 20)
        Me.TxtPrimaryProduct.TabIndex = 31
        Me.TxtPrimaryProduct.Text = "-1"
        Me.TxtPrimaryProduct.Visible = False
        '
        'TxtRequestBy
        '
        Me.TxtRequestBy.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRequestBy.Location = New System.Drawing.Point(82, 62)
        Me.TxtRequestBy.Name = "TxtRequestBy"
        Me.TxtRequestBy.Size = New System.Drawing.Size(110, 20)
        Me.TxtRequestBy.TabIndex = 4
        '
        'TxtDescProduct
        '
        Me.TxtDescProduct.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescProduct.Location = New System.Drawing.Point(235, 38)
        Me.TxtDescProduct.Name = "TxtDescProduct"
        Me.TxtDescProduct.ReadOnly = True
        Me.TxtDescProduct.Size = New System.Drawing.Size(200, 20)
        Me.TxtDescProduct.TabIndex = 3
        Me.TxtDescProduct.TabStop = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(4, 65)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(62, 13)
        Me.Label7.TabIndex = 29
        Me.Label7.Text = "Request By"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(4, 41)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(72, 13)
        Me.Label8.TabIndex = 27
        Me.Label8.Text = "Product Code"
        '
        'TxtProductCode
        '
        Me.TxtProductCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtProductCode.Location = New System.Drawing.Point(82, 38)
        Me.TxtProductCode.Name = "TxtProductCode"
        Me.TxtProductCode.Size = New System.Drawing.Size(110, 20)
        Me.TxtProductCode.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(4, 12)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(30, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "From"
        '
        'TabUserKelompok
        '
        Me.TabUserKelompok.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabUserKelompok.Controls.Add(Me.TabPage1)
        Me.TabUserKelompok.Controls.Add(Me.TabPage2)
        Me.TabUserKelompok.Location = New System.Drawing.Point(-1, 3)
        Me.TabUserKelompok.Name = "TabUserKelompok"
        Me.TabUserKelompok.SelectedIndex = 0
        Me.TabUserKelompok.Size = New System.Drawing.Size(365, 153)
        Me.TabUserKelompok.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DgvUser)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage1.Size = New System.Drawing.Size(357, 127)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Created By"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'DgvUser
        '
        Me.DgvUser.AllowUserToResizeColumns = False
        Me.DgvUser.AllowUserToResizeRows = False
        Me.DgvUser.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvUser.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCreatedBy_Primary, Me.ColCreatedBy, Me.ColCreatedBy_Name})
        Me.DgvUser.Location = New System.Drawing.Point(2, 2)
        Me.DgvUser.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DgvUser.Name = "DgvUser"
        Me.DgvUser.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvUser.RowTemplate.Height = 24
        Me.DgvUser.Size = New System.Drawing.Size(355, 126)
        Me.DgvUser.TabIndex = 0
        '
        'ColCreatedBy_Primary
        '
        Me.ColCreatedBy_Primary.HeaderText = "Primary"
        Me.ColCreatedBy_Primary.Name = "ColCreatedBy_Primary"
        Me.ColCreatedBy_Primary.ReadOnly = True
        Me.ColCreatedBy_Primary.Visible = False
        Me.ColCreatedBy_Primary.Width = 66
        '
        'ColCreatedBy
        '
        Me.ColCreatedBy.HeaderText = "User Code"
        Me.ColCreatedBy.Name = "ColCreatedBy"
        Me.ColCreatedBy.Width = 82
        '
        'ColCreatedBy_Name
        '
        Me.ColCreatedBy_Name.HeaderText = "Name"
        Me.ColCreatedBy_Name.Name = "ColCreatedBy_Name"
        Me.ColCreatedBy_Name.ReadOnly = True
        Me.ColCreatedBy_Name.Width = 60
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvItemProductGroup)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage2.Size = New System.Drawing.Size(351, 127)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Product Group"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeight = 25
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(0, 0)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(359, 130)
        Me.dgvItemProductGroup.TabIndex = 177
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.Frozen = True
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_No.Width = 27
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.Frozen = True
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.Width = 39
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.Frozen = True
        Me.ColProductGroup_CodeNo.HeaderText = "Group Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_CodeNo.Width = 70
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Desc.Width = 66
        '
        'ColProductGroup_Remarks
        '
        Me.ColProductGroup_Remarks.HeaderText = "Remarks"
        Me.ColProductGroup_Remarks.Name = "ColProductGroup_Remarks"
        Me.ColProductGroup_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Remarks.Width = 55
        '
        'ColProductGroup_PrimaryAcuan
        '
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary Acuan"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 81
        '
        'BackgroundWorker3
        '
        '
        'BgwUser
        '
        '
        'FrmLaporan_Purchase_Request
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(890, 538)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.crv)
        Me.MinimumSize = New System.Drawing.Size(906, 576)
        Me.Name = "FrmLaporan_Purchase_Request"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Request Report"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.TabUserKelompok.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnFindProduct As System.Windows.Forms.Button
    Friend WithEvents BtnFindRefParty As System.Windows.Forms.Button
    Friend WithEvents ChkAllProduct As System.Windows.Forms.CheckBox
    Friend WithEvents ChkOB As System.Windows.Forms.CheckBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtRequestBy As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtProductCode As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtDescRefParty As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescProduct As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryRefParty As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryProduct As System.Windows.Forms.TextBox
    Friend WithEvents CboPilih As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkShowEstimateDelivery As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowPONumber As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryPR As System.Windows.Forms.TextBox
    Friend WithEvents TxtPRNo As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFindPR As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabUserKelompok As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents DgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents ColCreatedBy_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BgwUser As System.ComponentModel.BackgroundWorker
End Class
