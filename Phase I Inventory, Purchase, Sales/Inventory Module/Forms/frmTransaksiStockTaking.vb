Option Explicit On
Option Strict Off

Public Class frmTransaksiStockTaking
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Obj As New clsTransaksiStockTaking
    Dim baru As Boolean
    Dim mBuffer As String = ""

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub TambahTabPageDanAccesoriesnya()
        Try
            Dim i As New TabPage

            'lbTeamNo
            Dim lbTeamNo As New Label
            lbTeamNo.AutoSize = True
            lbTeamNo.Location = New System.Drawing.Point(8, 9)
            lbTeamNo.Name = "lbTeamNo" + CStr(Me.TabControl1.TabPages.Count)
            lbTeamNo.Size = New System.Drawing.Size(51, 13)
            lbTeamNo.TabIndex = 47
            lbTeamNo.Text = "Team No"
            'Tambah Text Box
            Dim txtTeamNumber As New clsTxtCustom
            txtTeamNumber.Location = New System.Drawing.Point(64, 6)
            txtTeamNumber.Name = "txtTeamNumber" + CStr(Me.TabControl1.TabPages.Count)
            txtTeamNumber.Size = New System.Drawing.Size(88, 20)
            txtTeamNumber.TabIndex = 48
            'lbJobDesc
            Dim lbJobDesc As New Label
            lbJobDesc.AutoSize = True
            lbJobDesc.Location = New System.Drawing.Point(8, 29)
            lbJobDesc.Name = "lbJobDesc" + CStr(Me.TabControl1.TabPages.Count)
            lbJobDesc.Size = New System.Drawing.Size(80, 13)
            lbJobDesc.TabIndex = 49
            lbJobDesc.Text = "Job Description"
            'txtTeamJobDesc
            Dim txtTeamJobDesc As New clsTxtCustom
            txtTeamJobDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                        Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
            txtTeamJobDesc.Location = New System.Drawing.Point(11, 45)
            txtTeamJobDesc.Multiline = True
            txtTeamJobDesc.Name = "txtTeamJobDesc" + CStr(Me.TabControl1.TabPages.Count)
            txtTeamJobDesc.Size = New System.Drawing.Size(141, 50)
            txtTeamJobDesc.TabIndex = 50
            'Tambah DGV
            Dim DgvTeam As New ClsDGVCustom
            Me.TambahDGV(DgvTeam)
            'gboTeamStaff
            Dim gboTeamStaff As New GroupBox
            gboTeamStaff.Controls.Add(DgvTeam)
            gboTeamStaff.Location = New System.Drawing.Point(158, 6)
            gboTeamStaff.Name = "gboTeamStaff" + CStr(Me.TabControl1.TabPages.Count)
            gboTeamStaff.Size = New System.Drawing.Size(356, 202)
            gboTeamStaff.TabIndex = 51
            gboTeamStaff.TabStop = False
            gboTeamStaff.Text = "Team Staff"


            'Masukan Item pada Tab Page
            i.Controls.Add(gboTeamStaff)
            i.Controls.Add(txtTeamJobDesc)
            i.Controls.Add(lbJobDesc)
            i.Controls.Add(txtTeamNumber)
            i.Controls.Add(lbTeamNo)
            i.Location = New System.Drawing.Point(4, 22)
            i.Name = "TabPage" + CStr(Me.TabControl1.TabPages.Count)
            i.Padding = New System.Windows.Forms.Padding(3)
            i.Size = New System.Drawing.Size(520, 214)
            i.TabIndex = 1
            i.Text = "Team " + CStr(Me.TabControl1.TabPages.Count)
            i.UseVisualStyleBackColor = True

            Me.TabControl1.TabPages.Add(i)
            gboTeamStaff.Height = i.Height - 10
            gboTeamStaff.Width = i.Width - 165
            gboTeamStaff.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahDGV(ByVal DgvTeam As ClsDGVCustom)
        Try
            Dim Column1 As New System.Windows.Forms.DataGridViewTextBoxColumn
            Dim Column2 As New System.Windows.Forms.DataGridViewTextBoxColumn
            Dim Column3 As New System.Windows.Forms.DataGridViewTextBoxColumn
            '
            'Column1
            '
            Column1.HeaderText = "No"
            Column1.Name = "ColNo" + CStr(Me.TabControl1.TabPages.Count - 1)
            Column1.Width = 46
            Column1.SortMode = DataGridViewColumnSortMode.NotSortable
            '
            'Column2
            '
            Column2.HeaderText = "Name"
            Column2.Name = "ColName" + CStr(Me.TabControl1.TabPages.Count - 1)
            Column2.Width = 60
            '
            'Column3
            '
            Column3.HeaderText = "Remark"
            Column3.Name = "ColRemark" + CStr(Me.TabControl1.TabPages.Count - 1)
            Column3.Width = 69
            '
            DgvTeam.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            DgvTeam.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
            DgvTeam.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            DgvTeam.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Column1, Column2, Column3})
            DgvTeam.Dock = System.Windows.Forms.DockStyle.Fill
            DgvTeam.Location = New System.Drawing.Point(3, 16)
            DgvTeam.Name = "DataGridView" + CStr(Me.TabControl1.TabPages.Count)
            DgvTeam.Size = New System.Drawing.Size(350, 183)
            DgvTeam.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            DgvTeam.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            DgvTeam.MultiSelect = False
            DgvTeam.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DgvTeam.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            DgvTeam.AllowUserToOrderColumns = False
            DgvTeam.AllowUserToResizeColumns = False
            DgvTeam.AllowUserToResizeRows = False
            DgvTeam.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            DgvTeam.TabIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTotalTeam_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTotalTeam.Enter
        Try
            Me.mBuffer = Me.txtTotalTeam.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTotalTeam_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTotalTeam.GotFocus
        Try
            If Me.txtTotalTeam.ReadOnly = False Then
                Me.txtTotalTeam.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTotalTeam_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTotalTeam.KeyDown
        Try
            If e.KeyData = Keys.OemPeriod Or e.KeyData = Keys.Decimal Then
                e.SuppressKeyPress = True
                Exit Sub
            End If
            Me.ObjCD.Number_KeyDown(Me.txtTotalTeam, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTotalTeam_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTotalTeam.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtTotalTeam, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTotalTeam_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTotalTeam.Leave
        Try
            If IsNumeric(Me.txtTotalTeam.Text) = False Then
                Me.txtTotalTeam.Text = Me.mBuffer
            Else
                If Me.TabControl1.TabPages.Count - 1 > CDbl(Me.txtTotalTeam.Text) Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Dim ListItem As New List(Of String)
                    Dim TBL As New dseTransaksiStockTaking.tbltrans_stock_taking_team1DataTable
                    '  Dim ListTag As New List(Of String)
                    Dim AdaTag As Boolean = False
                    Dim view As Data.DataView
                    If Me.Obj.IsikanTable(TBL, Me.TxtPrimaryStockTaking.Text) = True Then
                        AdaTag = True
                    End If
                    Dim Identifier As String = ""
                    Dim JmlTabYangSudahDiCek As Integer = Me.TabControl1.TabPages.Count - 1
                    Dim hj As New FrmTransaksiStockTakingUntukListTag
                    hj.lboItems.Items.Clear()
                    Do While JmlTabYangSudahDiCek <> CDbl(Me.txtTotalTeam.Text)
                        Dim adaItemKah As Boolean = False
                        '        AdaTag = False
                        Dim cont As Control
                        Dim txtTeamNumber As TextBox
                        Dim TeamJobDesc As TextBox
                        '     ListTag.Clear()
                        'txtTeamNumber.Text = ""
                        For Each cont In Me.TabControl1.TabPages(JmlTabYangSudahDiCek).Controls
                            If cont.Name.StartsWith("txtTeamNumber") Then
                                txtTeamNumber = cont
                                If txtTeamNumber.Text <> "" Then
                                    adaItemKah = True
                                    Exit For
                                End If
                            End If
                        Next
                        For Each cont In Me.TabControl1.TabPages(JmlTabYangSudahDiCek).Controls
                            'cari no team
                            If cont.Name.StartsWith("txtTeamJobDesc") Then
                                TeamJobDesc = cont
                                If TeamJobDesc.Text <> "" Then
                                    adaItemKah = True
                                    Exit For
                                End If
                            End If
                        Next
                        For Each cont In Me.TabControl1.TabPages(JmlTabYangSudahDiCek).Controls
                            If cont.Name.StartsWith("gbo") Then
                                Identifier = Me.TxtPrimaryStockTaking.Text + "-" + CStr(JmlTabYangSudahDiCek)
                                Dim cont1 As Control
                                For Each cont1 In cont.Controls
                                    If cont1.Name.StartsWith("DataGridView") Then
                                        Dim Dgv As DataGridView = TryCast(cont1, DataGridView)
                                        Dim row As DataGridViewRow
                                        For Each row In Dgv.Rows
                                            If row.Cells(1).Value <> "" And row.Cells(2).Value <> "" Then
                                                adaItemKah = True
                                                Exit For
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                        Next

                        If adaItemKah = True And IsNothing(txtTeamNumber.Text) = False Then
                            ListItem.Add(txtTeamNumber.Text)
                        End If

                        If AdaTag = True Then
                            'Item langsung di pindahkan ke listbox di FrmTransaksiStockTakingUntukListTag
                            Dim NoTag As Data.DataRowView
                            view = TBL.DefaultView
                            Dim cmd As String = "Identifier = '" + Identifier + "' and [Tag No] <> ''"
                            view.RowFilter = cmd
                            If view.Count > 0 Then
                                hj.lboItems.Items.Add(". " + Me.TabControl1.TabPages(JmlTabYangSudahDiCek).Text + "-" + txtTeamNumber.Text + ", the inventory check tag refering to this team are: ")
                                For Each NoTag In view
                                    hj.lboItems.Items.Add("     " + NoTag("Tag No").ToString)
                                Next
                            End If
                        End If
                        JmlTabYangSudahDiCek -= 1
                    Loop

                    If hj.lboItems.Items.Count > 0 Then
                        Dim Nomor As Int16 = 1
                        Dim NomorAnak As Int16 = 1
                        For i As Integer = 0 To hj.lboItems.Items.Count - 1
                            If CStr(hj.lboItems.Items(i)).StartsWith(".") Then
                                hj.lboItems.Items(i) = CStr(Nomor) + hj.lboItems.Items(i).ToString
                                Nomor += 1
                                NomorAnak = 1
                            ElseIf CStr(hj.lboItems.Items(i)).StartsWith("     ") Then
                                hj.lboItems.Items(i) = "         " + CStr(Nomor - 1) + "." + CStr(NomorAnak) + ".   " + Trim(hj.lboItems.Items(i).ToString)
                                NomorAnak += 1
                            End If
                        Next
                        hj.ShowDialog()
                        hj.Dispose()
                        Me.txtTotalTeam.Text = Me.TabControl1.TabPages.Count - 1
                    ElseIf ListItem.Count > 0 Then
                        Dim kumpulanNumber As String = ""
                        Dim No As String
                        For Each No In ListItem
                            If String.IsNullOrEmpty(kumpulanNumber) = True Then
                                kumpulanNumber = No
                            Else
                                kumpulanNumber += ", " + No
                            End If
                        Next
                        'Hapus semua
                        Windows.Forms.Cursor.Current = Cursors.Arrow
                        If MsgBox(kumpulanNumber + " has value. Do you want to remove?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Do While Me.TabControl1.TabPages.Count - 1 <> CDbl(Me.txtTotalTeam.Text)
                                Me.TabControl1.TabPages.RemoveAt(Me.TabControl1.TabPages.Count - 1)
                            Loop
                        Else
                            Me.txtTotalTeam.Text = Me.TabControl1.TabPages.Count - 1
                        End If
                    Else
                        Do While Me.TabControl1.TabPages.Count - 1 <> CDbl(Me.txtTotalTeam.Text)
                            Me.TabControl1.TabPages.RemoveAt(Me.TabControl1.TabPages.Count - 1)
                        Loop
                    End If
                    ListItem = Nothing
                    '  ListTag = Nothing
                    TBL.Rows.Clear()
                    Windows.Forms.Cursor.Current = Cursors.Arrow
                ElseIf Me.TabControl1.TabPages.Count - 1 < CDbl(Me.txtTotalTeam.Text) Then
                    For i As Integer = Me.TabControl1.TabPages.Count To CDbl(Me.txtTotalTeam.Text)
                        Me.TambahTabPageDanAccesoriesnya()
                    Next
                End If
            End If
            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.txtTotalTeam.Text = Me.TabControl1.TabPages.Count - 1
        End Try
    End Sub

    Private Sub txtTotalTeam_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTotalTeam.LostFocus
        Try
            If Me.txtTotalTeam.ReadOnly = False Then
                Me.txtTotalTeam.BackColor = Me.TxtPrimaryStockTaking.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimaryStockTaking.Text = "-1"
            Me.txtTransactionCode.Text = ""
            Me.txtDescription.Text = ""
            Me.txtTotalTeam.Text = "0"
            Me.dtpStartDate.Value = Now
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Do While Me.TabControl1.TabPages.Count > 1
                Me.TabControl1.TabPages.RemoveAt(Me.TabControl1.TabPages.Count - 1)
            Loop
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.Obj
                .Primary = Me.TxtPrimaryStockTaking.Text
                .NoTransaksi = Me.txtTransactionCode.Text
                .Desc = Me.txtDescription.Text
                .Status = clsTransaksiStockTaking.EnumStatusStockTaking.Active
                .TotalTeam = Me.txtTotalTeam.Text
                .StartDate = Me.dtpStartDate.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                '          Me.KosongkanItem()
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.TabControl1) Then
                            Windows.Forms.Cursor.Current = Cursors.Arrow
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            Me.Obj.CariDetail(Me.TxtPrimaryStockTaking.Text, Me.TabControl1)
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                            Windows.Forms.Cursor.Current = Cursors.Arrow
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.TabControl1) Then
                            Windows.Forms.Cursor.Current = Cursors.Arrow
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            Me.Obj.CariDetail(Me.TxtPrimaryStockTaking.Text, Me.TabControl1)
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                            Windows.Forms.Cursor.Current = Cursors.Arrow
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimaryStockTaking.Text = .Primary
                Me.txtTransactionCode.Text = .NoTransaksi
                Me.txtDescription.Text = .Desc
                Me.txtTotalTeam.Text = .TotalTeam
                Me.dtpStartDate.Value = .StartDate
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, False, False)
            ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            ObjCD.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, True, False, False, False)
            Me.CancelButton = Me.btnExit
            Dim cont As Control
            For Each cont In Me.TabControl1.Controls
                If cont.Name.StartsWith("gbo") Then
                    Dim cont1 As Control
                    For Each cont1 In cont.Controls
                        If cont1.Name.StartsWith("DataGridView") Then
                            Dim Dgv As DataGridView = TryCast(cont1, DataGridView)
                            Dgv.AllowUserToDeleteRows = False
                            Dgv.ReadOnly = True
                        End If
                    Next
                End If
            Next
            Me.NonAktifkanTabPage()
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTeamNo(ByVal TeamName As String) As String
        Dim TimNo_s As New Queue(Of String)
        Dim cont As Control
        Dim Ans As String = ""

        Try
            For i As Integer = 1 To Me.TabControl1.TabPages.Count - 1
                For Each cont In Me.TabControl1.TabPages(i).Controls
                    'cari team jobdesc
                    If cont.Name.StartsWith("txtTeamNumber") And cont.Text = TeamName Then
                        'masukkan ke dalam list
                        TimNo_s.Enqueue(Me.TabControl1.TabPages(i).Text)
                        Exit For
                    End If
                Next
            Next
            While TimNo_s.Count > 0
                If Ans = "" Then
                    Ans = Ans + TimNo_s.Dequeue
                Else
                    If TimNo_s.Count = 1 Then
                        Ans = Ans + " and " + TimNo_s.Dequeue
                    Else
                        Ans = Ans + ", " + TimNo_s.Dequeue
                    End If
                End If
            End While

            Return Ans
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            TimNo_s = Nothing
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim txtTeamNumber As TextBox
        Dim dgv As DataGridView
        Dim TeamDesc As New List(Of String)
        Dim TeamName As String = ""

        Try
            'data deskripsi harus diisikan
            'no transaksi juga harus diisikan
            KoleksiNamaData.Add("stock taking transaction number")
            KoleksiTextBox.Add(Me.txtTransactionCode)
            KoleksiNamaData.Add("stock taking description")
            KoleksiTextBox.Add(Me.txtDescription)
            Me.TabControl1.SelectedIndex = 0
            If ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'tanggal start date harus lebih besar dari tanggal start date terakhir
            If Me.baru = True Then
                Dim i As Double = Me.Obj.GetPrimaryMax
                If i > 0 Then
                    Dim PrevStartDate As Date = Me.Obj.GetStartDate(i)
                    If Me.dtpStartDate.Value <= PrevStartDate Then
                        MsgBox("The previous stock taking date is " + Format(PrevStartDate, "dd MMM yyyy") + _
                        " this stock taking must take place after that date.", MsgBoxStyle.Information, _
                        "Information")
                        Return False
                    End If
                End If
            Else
                Dim i As Double = CDbl(Me.TxtPrimaryStockTaking.Text) - 1
                If i > 0 Then
                    Dim PrevStartDate As Date = Me.Obj.GetStartDate(i)
                    If Me.dtpStartDate.Value <= PrevStartDate Then
                        MsgBox("The previous stock taking date is " + Format(PrevStartDate, "dd MMM yyyy") + _
                        " this stock taking must take place after that date.", MsgBoxStyle.Information, _
                        "Information")
                        Return False
                    End If
                End If
            End If
            If Me.txtTotalTeam.Text = "0" Then
                MsgBox("Please enter at least 1 team for total team.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 0
                Me.txtTotalTeam.SelectAll()
                Me.txtTotalTeam.Focus()
                Return False
            End If
            Dim cont As Control
            For i As Integer = 1 To Me.TabControl1.TabPages.Count - 1
                For Each cont In Me.TabControl1.TabPages(i).Controls
                    'cari team jobdesc
                    If cont.Name.StartsWith("txtTeamNumber") Then
                        txtTeamNumber = cont
                        'cek data yang harus diisi
                        KoleksiNamaData.Clear()
                        KoleksiTextBox.Clear()
                        KoleksiNamaData.Add("Team Number")
                        KoleksiTextBox.Add(txtTeamNumber)
                        If ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                            Me.TabControl1.SelectedIndex = i
                            txtTeamNumber.Focus()
                            Return False
                        End If
                        'masukkan ke dalam list
                        For Each TeamName In TeamDesc
                            If TeamName = txtTeamNumber.Text Then
                                Dim TeamNo_s As String = Me.GetTeamNo(TeamName)
                                MsgBox("There are 2 teams or more using '" + txtTeamNumber.Text + _
                                "' as team name. They are: " + TeamNo_s, MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        Next
                        'kalo sampe sini tambahkan ke dalam list
                        TeamDesc.Add(txtTeamNumber.Text)
                        Exit For
                    End If
                Next
                For Each cont In Me.TabControl1.TabPages(i).Controls
                    If cont.Name.StartsWith("gbo") Then
                        Dim cont1 As Control
                        For Each cont1 In cont.Controls
                            If cont1.Name.StartsWith("DataGridView") Then
                                dgv = cont1
                                'minimal harus punya 1 data item
                                If dgv.RowCount = 1 Then
                                    MsgBox("Please enter at least 1 team member for " + Me.TabControl1.TabPages(i).Text + ".", MsgBoxStyle.Information, "Information")
                                    Me.TabControl1.SelectedIndex = i
                                    Return False
                                End If
                                'setiap baris harus diisikan namanya
                                Dim row As DataGridViewRow
                                For Each row In dgv.Rows
                                    If row.Index <> dgv.RowCount - 1 Then
                                        If IsNothing(row.Cells(1).Value) = True Then
                                            MsgBox("Please enter staff name on row " + CStr(row.Cells(0).Value), MsgBoxStyle.Information, "Information")
                                            Me.TabControl1.SelectedIndex = i
                                            Return False
                                        Else
                                            If CStr(row.Cells(1).Value) = "" Then
                                                MsgBox("Please enter staff name on row " + CStr(row.Cells(0).Value), MsgBoxStyle.Information, "Information")
                                                Me.TabControl1.SelectedIndex = i
                                                Return False
                                            End If
                                        End If
                                        If IsNothing(row.Cells(2).Value) = True Then
                                            row.Cells(2).Value = ""
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            Next

            'tidak boleh ada tim yang double namanya
            'MsgBox("belum")
            If baru = True Then
                If Me.Obj.getJmlByNoTransaksi(Me.txtTransactionCode.Text) > 0 Then
                    MsgBox("Transaction code " + Me.txtTransactionCode.Text + " has been used. Change the transaction code.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf baru = False Then
                Dim NoTranssaksi As String = Me.Obj.GetNoTransaksi(Me.TxtPrimaryStockTaking.Text)
                If NoTranssaksi <> Me.txtTransactionCode.Text Then
                    If Me.Obj.getJmlByNoTransaksi(Me.txtTransactionCode.Text) > 0 Then
                        MsgBox("Transaction code " + Me.txtTransactionCode.Text + " has been used. Change the transaction code.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            TeamDesc = Nothing
        End Try
    End Function

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Dim cont As Control
        Dim dgv As DataGridView
        Try
            Dim KetemuDGV As Boolean = False
            For i As Integer = 1 To Me.TabControl1.TabPages.Count - 1
                For Each cont In Me.TabControl1.TabPages(i).Controls
                    If cont.Name.StartsWith("gbo") Then
                        Dim cont1 As Control
                        For Each cont1 In cont.Controls
                            If cont1.Name.StartsWith("DataGridView") Then
                                dgv = cont1
                                KetemuDGV = True
                                Exit For
                            End If
                        Next
                    End If
                    If KetemuDGV = True Then
                        Exit For
                    End If
                Next
                If KetemuDGV = True Then
                    Exit For
                End If
            Next
            If Me.txtTransactionCode.Text = "" And Me.txtDescription.Text = "" And Me.dtpStartDate.Value.Date = Today And _
            Me.txtTotalTeam.Text = "1" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanItem()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanItem()
                    Else
                        Do While Me.TabControl1.TabPages.Count <> 1
                            Me.TabControl1.TabPages.RemoveAt(Me.TabControl1.TabPages.Count - 1)
                        Loop
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        Me.Obj.CariDetail(Me.TxtPrimaryStockTaking.Text, Me.TabControl1)
                        Me.SalinKeField()
                        Windows.Forms.Cursor.Current = Cursors.Arrow
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanTabPage()
        Try
            Dim j As Int16
            Dim cont As Control
            Dim i As TabPage
            For j = 1 To Me.TabControl1.TabPages.Count - 1
                i = Me.TabControl1.TabPages(j)
                For Each cont In i.Controls
                    'cari team jobdesc
                    If cont.Name.StartsWith("txtTeamJobDesc") Then
                        cont.Enabled = False
                        Exit For
                    End If
                Next
                For Each cont In i.Controls
                    'cari no team
                    If cont.Name.StartsWith("txtTeamNumber") Then
                        cont.Enabled = False
                        Exit For
                    End If
                Next
                'simpan team
                For Each cont In i.Controls
                    If cont.Name.StartsWith("gbo") Then
                        Dim cont1 As Control
                        For Each cont1 In cont.Controls
                            If cont1.Name.StartsWith("DataGridView") Then
                                Dim Dgv As DataGridView = TryCast(cont1, DataGridView)
                                Dgv.AllowUserToDeleteRows = False
                                Dgv.ReadOnly = True
                            End If
                        Next
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanTabPage()
        Try
            Dim j As Int16
            Dim cont As Control
            Dim i As TabPage
            For j = 1 To Me.TabControl1.TabPages.Count - 1
                i = Me.TabControl1.TabPages(j)
                For Each cont In i.Controls
                    'cari team jobdesc
                    If cont.Name.StartsWith("txtTeamJobDesc") Then
                        cont.Enabled = True
                        Exit For
                    End If
                Next
                For Each cont In i.Controls
                    'cari no team
                    If cont.Name.StartsWith("txtTeamNumber") Then
                        cont.Enabled = True
                        Exit For
                    End If
                Next
                'simpan team
                For Each cont In i.Controls
                    If cont.Name.StartsWith("gbo") Then
                        Dim cont1 As Control
                        For Each cont1 In cont.Controls
                            If cont1.Name.StartsWith("DataGridView") Then
                                Dim Dgv As DataGridView = TryCast(cont1, DataGridView)
                                Dgv.AllowUserToDeleteRows = True
                                Dgv.ReadOnly = False
                            End If
                        Next
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariTransaksiStockTaking
        Try
            hj.ShowDialog()
            If Not hj.PrimaryTransaksiTaking = -1 Then
                Do While Me.TabControl1.TabPages.Count > 1
                    Me.TabControl1.TabPages.RemoveAt(Me.TabControl1.TabPages.Count - 1)
                Loop
                If Me.Obj.CariDetail(hj.PrimaryTransaksiTaking, Me.TabControl1) = True Then
                    Me.SalinKeField()
                    Me.NonAktifkanTabPage()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockTaking.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtPrimaryStockTaking.BackColor, False)
            ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, False, False)
            ObjCD.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, False, False)
            ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Dim cont As Control
            Dim dgv As DataGridView
            For i As Integer = 1 To Me.TabControl1.TabPages.Count - 1
                For Each cont In Me.TabControl1.TabPages(i).Controls
                    If cont.Name.StartsWith("gbo") Then
                        Dim cont1 As Control
                        For Each cont1 In cont.Controls
                            If cont1.Name.StartsWith("DataGridView") Then
                                dgv = cont1
                                dgv.AllowUserToDeleteRows = True
                                dgv.ReadOnly = False
                                Exit For
                            End If
                        Next
                    End If
                Next
            Next
            Me.AktifkanTabPage()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Obj.CekApakahAdaYangMasihAktif() = False Or Me.Obj.GetPrimaryMax = -1 Then
                Me.baru = True
                Me.AktifkanDetail()
                Me.KosongkanItem()
                Me.txtTransactionCode.Focus()
                Me.Obj.ResetObject()
                Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
                Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
                Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtPrimaryStockTaking.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPrimaryStockTaking.Text <> "-1" Then
                If Obj.ApakahBolehEdit(Me.TxtPrimaryStockTaking.Text) Then
                    Me.baru = False
                    Me.txtTransactionCode.Focus()
                    Me.AktifkanDetail()
                    Me.AktifkanTabPage()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtPrimaryStockTaking.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPrimaryStockTaking.Text <> "-1" Then
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimaryStockTaking.Text) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.Hapus(Me.TxtPrimaryStockTaking.Text) = True Then
                            Me.KosongkanItem()
                            Do While Me.TabControl1.TabPages.Count > 1
                                Me.TabControl1.TabPages.RemoveAt(Me.TabControl1.TabPages.Count - 1)
                            Loop
                            Me.txtTotalTeam.Text = "0"
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub frmTransaksiStockTaking_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtDescription.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.TabControl1) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.TabControl1) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmTransaksiStockTaking_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDescription_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.GotFocus
        Try
            If Me.txtDescription.ReadOnly = False Then
                Me.txtDescription.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDescription_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.LostFocus
        Try
            If Me.txtDescription.ReadOnly = False Then
                Me.txtDescription.BackColor = Me.TxtPrimaryStockTaking.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTransactionCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTransactionCode.GotFocus
        Try
            Me.txtTransactionCode.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTransactionCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTransactionCode.LostFocus
        Try
            Me.txtTransactionCode.BackColor = Me.TxtPrimaryStockTaking.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If Me.TxtPrimaryStockTaking.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimaryStockTaking.Text <> "-1" Then
                Dim hj As New frmLaporanStockTakingTeam
                hj.Primary = Me.TxtPrimaryStockTaking.Text
                hj.StartDate = Me.dtpStartDate.Value
                hj.StockTakingNo = Me.txtTransactionCode.Text
                hj.TotalTeam = Me.txtTotalTeam.Text
                hj.Description = Me.txtDescription.Text
                hj.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_StockTaking_Create(ByVal vPrimaryTransaksi As Double)
        Try
            Me.baru = False
            Me.NonAktifkanDetail()
            Me.NonAktifkanTabPage()
            Me.KosongkanItem()
            Me.Obj.ResetObject()

            Me.Obj.CariDetail(vPrimaryTransaksi, Me.TabControl1)
            Me.SalinKeField()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_StockTaking_Save() As Double
        Dim Status_ As clsTransaksiStockTaking.EnumStatusStockTaking = Me.Obj.Status

        Try
            Me.baru = True
            Me.SalinKeObjek()
            Me.Obj.Status = Status_
            Me.Obj.ProcessDate = Me.Obj.ModifiedDate
            Me.Obj.SimpanBaru(Me.UserCode, Me.TabControl1, False, True)
            Return Me.Obj.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function
End Class