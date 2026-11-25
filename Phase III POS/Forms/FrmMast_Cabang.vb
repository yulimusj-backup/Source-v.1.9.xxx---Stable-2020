Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmMast_Cabang
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris

    Dim baru As Boolean
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData
    Dim Obj As New ClsMast_Cabang
    Dim ObjLocation As New clsMasterStorageArea
    Dim EraseRowIndex_Branch As Integer = 0

    Private Sub dgvLocation_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvLocation.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_RowsAdded(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvLocation.RowsAdded
        Try
            Me.dgvLocation.TabStop = True
            If Me.TxtRemarks.Enabled = True Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_RowsRemoved(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvLocation.RowsRemoved
        Try
            If Me.dgvLocation.Rows.Count = 0 Then
                Me.dgvLocation.TabStop = False
                If Me.TxtRemarks.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_UserDeletedRow(ByVal sender As Object, ByVal e _
    As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvLocation.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles dgvLocation.SelectionChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvLocation.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvLocation.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvLocation.FirstDisplayedScrollingRowIndex = Me.dgvLocation.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvLocation.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvLocation.Rows(row.Index).Displayed = False
                    Me.dgvLocation.FirstDisplayedScrollingRowIndex = Me.dgvLocation.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvLocation.DisplayedRowCount(True) <> Me.dgvLocation.DisplayedRowCount(False) Then
                    Me.dgvLocation.FirstDisplayedScrollingRowIndex = Me.dgvLocation.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            Me.dgvLocation.AllowUserToAddRows = False
            If Me.dgvLocation.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvLocation.SelectedRows(0)

                If row.Index <> 0 Then
                    Dim indexbaru = row.Index - 1
                    Me.dgvLocation.Rows.Remove(row)
                    Me.dgvLocation.Rows.InsertRange(indexbaru, row)
                    Me.dgvLocation.Focus()
                    Me.dgvLocation.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvLocation)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True)
                    End If
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
            Me.dgvLocation.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            Me.dgvLocation.AllowUserToAddRows = False
            If Me.dgvLocation.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False

                Dim row As DataGridViewRow = Me.dgvLocation.SelectedRows(0)
                If row.Index <> Me.dgvLocation.Rows.Count - 1 Then
                    Dim indexbaru = row.Index + 1
                    Me.dgvLocation.Rows.Remove(row)
                    Me.dgvLocation.Rows.InsertRange(indexbaru, row)
                    Me.dgvLocation.Focus()
                    Me.dgvLocation.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvLocation)
                    Me.ShowRowProcess(row, False)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
            Me.dgvLocation.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAirport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Obj = New ClsMast_Cabang
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
            Me.dgvLocation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.GroupBox1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()
            Me.ChkMainBranch.Enabled = True

            'set dgv
            Me.dgvLocation.ReadOnly = False
            Me.dgvLocation.AllowUserToDeleteRows = True

            If Me.dgvLocation.SelectedRows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.GroupBox1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.BackColorDefault()
            Me.ChkMainBranch.Enabled = False
            Me.dgvLocation.ReadOnly = True
            Me.dgvLocation.AllowUserToDeleteRows = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.GroupBox1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.GroupBox1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.GroupBox1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)

            Me.TxtPrimary.Text = "-1"
            Me.ChkMainBranch.Checked = False
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvLocation.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Code")
            KoleksiTextBox.Add(Me.TxtCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtDesc)
            KoleksiNamaData.Add("Server Name")
            KoleksiTextBox.Add(Me.TxtServerName)
            KoleksiNamaData.Add("User ID")
            KoleksiTextBox.Add(Me.TxtUserId)
            KoleksiNamaData.Add("Password")
            KoleksiTextBox.Add(Me.TxtPassword)
            KoleksiNamaData.Add("Database Name")
            KoleksiTextBox.Add(Me.TxtDatabaseName)
            KoleksiNamaData.Add("Port")
            KoleksiTextBox.Add(Me.TxtPort)

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Duplication in code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Duplikasi code, silahkan ubah code", _
                        MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtCode.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetCode(Me.TxtPrimary.Text)
                If TipeCodeLama <> Me.TxtCode.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtCode.Text) = True Then
                        If Bahasa = EnumBahasa.Inggris Then
                            MsgBox("Duplication in code, please change your code", _
                            MsgBoxStyle.Information, "Information")
                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                            MsgBox("Duplikasi code, silahkan ubah code", _
                            MsgBoxStyle.Information, "Informasi")
                        End If
                        Me.TxtCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If Bahasa = EnumBahasa.Inggris Then
                        If MsgBox("Other user have modified this data, reload latest data?", _
                        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Me.Obj.CariDetail(Me.Obj.Primary, Me.ObjCD, Me)
                            Me.SalinKeField()
                        End If
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        If MsgBox("Beberapa data telah dimodifikasi, apakah mau mengambil kembali data terakhir?", _
                        MsgBoxStyle.YesNo, "Konfirm") = MsgBoxResult.Yes Then
                            Me.Obj.CariDetail(Me.Obj.Primary, Me.ObjCD, Me)
                            Me.SalinKeField()
                        End If
                    End If
                    Return False
                End If
            End If

            'Desc harus diisi
            If String.IsNullOrEmpty(Me.TxtDesc.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input description", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan description", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtDesc.Focus()
                Return False
            End If

            If String.IsNullOrEmpty(Me.TxtServerName.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input server name", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan server name", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtServerName.Focus()
                Return False
            End If

            If String.IsNullOrEmpty(Me.TxtUserId.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input user id", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan user id", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtUserId.Focus()
                Return False
            End If
            If String.IsNullOrEmpty(Me.TxtPassword.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input password", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan password", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtPassword.Focus()
                Return False
            End If
            If String.IsNullOrEmpty(Me.TxtDatabaseName.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input database name", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan database name", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtDatabaseName.Focus()
                Return False
            End If
            If String.IsNullOrEmpty(Me.TxtPort.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input port", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan port", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtPort.Focus()
                Return False
            End If

            If String.IsNullOrEmpty(Me.TxtRemarks.Text) = False Then
                If ObjCD.CekDataKosong(Me.TxtRemarks.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtRemarks.Focus()
                    Return False
                End If
            End If

            'Cek Duplikasi ServerName dan DatabaseName
            If Me.baru = True Then
                If Obj.GetJmlByServerDBName(Me.TxtServerName.Text, Me.TxtDatabaseName.Text) > 0 Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Duplication server name or database name is already in use in another branch", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Duplikasi server name or database name sudah digunakan pada branch yang lain", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtServerName.Focus()
                    Return False
                End If
            Else
                If Obj.GetJmlByServerDBNamePosisiEdit(Me.TxtServerName.Text, Me.TxtDatabaseName.Text, Me.TxtPrimary.Text) > 0 Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Duplication server name or database name is already in use in another branch", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Duplikasi server name or database name sudah digunakan pada branch yang lain", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtServerName.Focus()
                    Return False
                End If
            End If

            'satu server satu main branch
            If Me.baru = True Then
                If Me.ChkMainBranch.Checked = True Then
                    If Obj.GetJmlByMainBranch() >= 1 Then
                        If Bahasa = EnumBahasa.Inggris Then
                            MsgBox("Please uncheck main branch because there is another branch that already use", MsgBoxStyle.Information, "Information")
                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                            MsgBox("Uncheck main branch karena ada cabang lain yang sudah pergunakan", MsgBoxStyle.Information, "Informasi")
                        End If
                        Me.ChkMainBranch.Focus()
                        Return False
                    End If
                End If
            Else
                If Me.ChkMainBranch.Checked = True Then
                    If Obj.GetJmlByMainBranchPosisiEdit(Me.TxtPrimary.Text) >= 1 Then
                        If Bahasa = EnumBahasa.Inggris Then
                            MsgBox("Please uncheck main branch because there is another branch that already use", MsgBoxStyle.Information, "Information")
                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                            MsgBox("Uncheck main branch karena ada cabang lain yang sudah pergunakan", MsgBoxStyle.Information, "Informasi")
                        End If
                        Me.ChkMainBranch.Focus()
                        Return False
                    End If
                End If
            End If

            Me.dgvLocation.AllowUserToAddRows = False
            'If Me.dgvLocation.Rows.Count <= 0 Then
            'If Bahasa = EnumBahasa.Inggris Then
            'MsgBox("Please input location detail", MsgBoxStyle.Information, "Information")
            'ElseIf Bahasa = EnumBahasa.Indonesia Then
            'MsgBox("Masukan location detail", MsgBoxStyle.Information, "Informasi")
            'End If
            'Me.dgvLocation.AllowUserToAddRows = True
            'Return False
            'End If
            Me.dgvLocation.AllowUserToAddRows = True

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .Code = Me.TxtCode.Text
                .Desc = Me.TxtDesc.Text
                .Remarks = Me.TxtRemarks.Text
                If Me.ChkMainBranch.Checked = True Then
                    .MainBranch = 1
                Else
                    .MainBranch = 0
                End If
                .ServerName = Me.TxtServerName.Text
                .UserID = Me.TxtUserId.Text
                .Password = Me.TxtPassword.Text
                .DatabaseName = Me.TxtDatabaseName.Text
                .Port = Me.TxtPort.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtCode.Text = .Code
                Me.TxtDesc.Text = .Desc
                Me.TxtRemarks.Text = .Remarks
                If .MainBranch = 0 Then
                    Me.ChkMainBranch.Checked = False
                Else
                    Me.ChkMainBranch.Checked = True
                End If
                Me.TxtServerName.Text = .ServerName
                Me.TxtUserId.Text = .UserID
                Me.TxtPassword.Text = .Password
                Me.TxtDatabaseName.Text = .DatabaseName
                Me.TxtPort.Text = .Port

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Kontrol Button
    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, _
                        Me.TxtPrimary.Text, Bahasa, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, _
                        Me.TxtPrimary.Text, Bahasa, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtCode.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                        Me.TxtCode.Clear()
                    Else
                        Me.Obj.CariDetail(Me.TxtPrimary.Text, Me.ObjCD, Me)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Obj.ResetObject(Me)
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.TxtCode.Focus()
            Me.TxtCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ijk As New FrmMast_Cabang_Cari
        Try
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            ijk.ShowDialog()

            If Not ijk.mPrimary = -1 Then
                If Me.Obj.CariDetail(ijk.mPrimary, Me.ObjCD, Me) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCode.Text <> "" Then
                If Me.Obj.ApakahBolehEdit(Me.TxtPrimary.Text, Me.UserCode, True, True, Bahasa) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.TxtCode.Focus()
                    Me.TxtCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCode.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimary.Text, Me.UserCode, True, True, _
                 Bahasa, Me) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.TxtPrimary.Text, Me.UserCode, True, True, _
                        Bahasa, Me) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, _
                            Me.TxtPrimary.Text, Bahasa, Me) = True Then
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
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, _
                            Me.TxtPrimary.Text, Bahasa, Me) = True Then
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

    'Coding BackColor
    Private Sub TxtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCode.Enabled = True Then
                    If Me.TxtCode.ReadOnly = False Then
                        TxtCode.BackColor = Me.BackColorAll
                    Else
                        TxtCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCode.Enabled = True Then
                    If Me.TxtCode.ReadOnly = False Then
                        TxtCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNama_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDesc.Enabled = True Then
                    If Me.TxtDesc.ReadOnly = False Then
                        TxtDesc.BackColor = Me.BackColorAll
                    Else
                        TxtDesc.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDesc.Enabled = True Then
                    If Me.TxtDesc.ReadOnly = False Then
                        TxtDesc.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtDesc.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.BackColorAll
                Else
                    TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCity_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtServerName.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtServerName.Enabled = True Then
                    If Me.TxtServerName.ReadOnly = False Then
                        Me.TxtServerName.BackColor = Me.BackColorAll
                    Else
                        Me.TxtServerName.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCity_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtServerName.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtServerName.Enabled = True Then
                    If Me.TxtServerName.ReadOnly = False Then
                        Me.TxtServerName.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtServerName.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDescItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUserId.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtUserId.Enabled = True Then
                    If Me.TxtUserId.ReadOnly = False Then
                        Me.TxtUserId.BackColor = Me.BackColorAll
                    Else
                        Me.TxtUserId.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDescItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUserId.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtUserId.Enabled = True Then
                    If Me.TxtUserId.ReadOnly = False Then
                        Me.TxtUserId.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtUserId.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarksItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPassword.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPassword.Enabled = True Then
                    If Me.TxtPassword.ReadOnly = False Then
                        Me.TxtPassword.BackColor = Me.BackColorAll
                    Else
                        Me.TxtPassword.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarksItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPassword.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPassword.Enabled = True Then
                    If Me.TxtPassword.ReadOnly = False Then
                        Me.TxtPassword.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtPassword.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTerminalCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDatabaseName.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDatabaseName.Enabled = True Then
                    If Me.TxtDatabaseName.ReadOnly = False Then
                        Me.TxtDatabaseName.BackColor = Me.BackColorAll
                    Else
                        Me.TxtDatabaseName.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTerminalCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDatabaseName.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDatabaseName.Enabled = True Then
                    If Me.TxtDatabaseName.ReadOnly = False Then
                        Me.TxtDatabaseName.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtDatabaseName.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPort_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPort.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPort.Enabled = True Then
                    If Me.TxtPort.ReadOnly = False Then
                        Me.TxtPort.BackColor = Me.BackColorAll
                    Else
                        Me.TxtPort.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPort_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPort.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPort.Enabled = True Then
                    If Me.TxtPort.ReadOnly = False Then
                        Me.TxtPort.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtPort.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vPrimaryLocation As String) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvLocation.Rows
                If rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvLocation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvLocation.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColLocationCode.Index Or _
            e.ColumnIndex = Me.ColRemarksItem.Index Then
                Me.buffer = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvLocation.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvLocation.Rows(e.RowIndex)
        Dim Index As Integer = -1
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                    'Cek string keterangan
                    If e.ColumnIndex = Me.ColRemarksItem.Index Then
                        If ObjCD.CekDataKosong(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            MsgBox("Data keterangan" _
                            + " tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Information")
                            Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                    ''Cell kode 
                    If e.ColumnIndex = Me.ColLocationCode.Index Then
                        Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjLocation.ApakahSudahAda(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Index = Me.GetIndexItem(Me.ObjLocation.GetPrimary(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value))
                            If Index < 0 Then
                                Me.dgvLocation.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value = Me.ObjLocation.GetPrimary(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvLocation.Item(Me.ColLocationDesc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(Me.dgvLocation.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> Index Then
                                    If Bahasa = EnumBahasa.Inggris Then
                                        MsgBox("You cannot input the same code", MsgBoxStyle.Information, "Information")
                                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                                        MsgBox("Anda tidak dapat menginput code yang sama", MsgBoxStyle.Information, "Informasi")
                                    End If
                                    Me.dgvLocation.Item(Me.ColLocationCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex_Branch = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvLocation)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New FrmCariMasterStorageArea
                            ij.txtCari.Text = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            If ij.PrimaryStorageArea <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Index = Me.GetIndexItem(ij.PrimaryStorageArea)
                                If Index < 0 Then
                                    Me.dgvLocation.Item(Me.ColLocationCode.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(ij.PrimaryStorageArea)
                                    Me.dgvLocation.Item(Me.ColLocationDesc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(ij.PrimaryStorageArea)
                                    Me.dgvLocation.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value = ij.PrimaryStorageArea
                                Else
                                    If e.RowIndex <> Index Then
                                        If Bahasa = EnumBahasa.Inggris Then
                                            MsgBox("You cannot input the same code", MsgBoxStyle.Information, "Information")
                                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                                            MsgBox("Anda tidak dapat menginput code yang sama", MsgBoxStyle.Information, "Informasi")
                                        End If
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex_Branch = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    Else
                                        Me.dgvLocation.Item(Me.ColLocationCode.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(ij.PrimaryStorageArea)
                                        Me.dgvLocation.Item(Me.ColLocationDesc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(ij.PrimaryStorageArea)
                                        Me.dgvLocation.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value = ij.PrimaryStorageArea
                                    End If
                                End If
                            ElseIf ij.PrimaryStorageArea = -1 Then
                                If Me.dgvLocation.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_Branch = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvLocation)
                    End If
                    'Hapus Baris Jika Code No DownGrade Kosong
                    If e.ColumnIndex = Me.ColRemarksItem.Index Then
                        If String.IsNullOrEmpty(Me.dgvLocation.Item(Me.ColLocationCode.Index, e.RowIndex).Value) = True Then
                            If Me.dgvLocation.Rows.Count - 1 <> e.RowIndex Then
                                Me.EraseRowIndex_Branch = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarksItem.Index Then
                        Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        If e.ColumnIndex = Me.ColLocationCode.Index Then
                            If String.IsNullOrEmpty(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                If Me.dgvLocation.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_Branch = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                End If
                            End If
                        Else
                            Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.dgvLocation.Rows.RemoveAt(Me.EraseRowIndex_Branch)
            Me.ObjCD.SetNumber(Me.dgvLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class