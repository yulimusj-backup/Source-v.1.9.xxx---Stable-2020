Option Explicit On
Option Strict Off

Public Class FrmMasterUser
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim objuser As New ClsUser
    Dim baru As Boolean
    Dim objCd As New ClsCekData

    Private Sub cmdKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKeluar.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cmdSimpan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSimpan.Click
        Dim hasil As MsgBoxResult
        Try
            'konfirmasi
            If Me.ApakahDataBenar = True Then
                hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")

                If hasil = MsgBoxResult.Yes Then
                    If CekData() = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        SalinKeObjek()
                        If baru = True Then
                            SimpanBaru()
                        Else
                            SimpanPerubahan()
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

    Private Sub SimpanBaru()
        Dim Obj As New ClsUser
        Try
            If ModUser.Simpan(Me.objuser, Me.tbKodeUser, Me.tbNamaUser, Me.tbPassword, Me.tbKeterangan, CDbl(Me.TxtPrimaryUserGroup.Text)) = True Then
                'konfirmasi simpan
                MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                ''Obj.Simpan(Me.tbKodeUser, tbNamaUser, Me.tbPassword, Me.tbKeterangan.Text, CDbl(Me.TxtPrimaryUserGroup.Text))
                Me.SalinKeField()
                'kembalikan ke kondisi awal
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Sub

    Private Sub SimpanPerubahan()
        Dim Obj As New ClsUser
        Try
            If ModUser.SimpanPerubahan(Me.objuser) = True Then
                'konfirmasi simpan
                MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                'Obj.SimpanPerubahan(Me.tbKodeUser.Text, tbNamaUser.Text, Me.tbPassword.Text, Me.tbKeterangan.Text, CDbl(Me.TxtPrimaryUserGroup.Text))
                Me.SalinKeField()
                'kembalikan ke kondisi awal
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Sub

    Private Sub FrmTransaksiPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            '' If Me.tbKeterangan.Enabled = True Then
            ''Me.BringToFront()
            ''Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
            ''If hasil = MsgBoxResult.Yes Then
            ''If Me.ApakahDataBenar = True Then
            ''hasil = MsgBox("Anda yakin akan menyimpan data ini", MsgBoxStyle.YesNo, "Konfirmasi")
            ''If CekData() = True Then
            ''SalinKeObjek()
            ''If baru = True Then
            ''SimpanBaru()
            ''Else
            '' SimpanPerubahan()
            '' End If
            '' End If
            ''End If
            ''ElseIf hasil = MsgBoxResult.Cancel Then
            ''e.Cancel = True
            ''End If
            ''End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Group Code")
            KoleksiTextBox.Add(Me.txtUserGroupCode)
            KoleksiNamaData.Add("Kode User")
            KoleksiTextBox.Add(Me.tbKodeUser)

            If objCd.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub cmdBatal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBatal.Click
        Try
            If MsgBox("Cancel save?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                'kembalikan ke data awal
                If baru = True Then
                    Me.KosongkanField()
                Else
                    Me.SalinKeField()
                End If
                'kembalikan ke kondisi awal
                Me.NonAktifkanDetail()
                Me.cmdNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        'text box
        Me.tbKeterangan.Clear()
        Me.tbKodeUser.Clear()
        Me.tbNamaUser.Clear()
        Me.tbPassword.Clear()
        Me.txtUserGroupCode.Clear()
        Me.tbKonfimasiPassword.Clear()
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ''default
            Me.tbNamaUser.BackColor = Me.txtReadOnly.BackColor
            Me.tbKodeUser.BackColor = Me.txtReadOnly.BackColor
            Me.txtUserGroupCode.BackColor = Me.txtReadOnly.BackColor
            Me.tbKeterangan.BackColor = Me.txtReadOnly.BackColor
            Me.tbPassword.BackColor = Me.txtReadOnly.BackColor
            Me.tbKonfimasiPassword.BackColor = Me.txtReadOnly.BackColor
            'text box
            Me.tbNamaUser.ReadOnly = True
            Me.tbKodeUser.ReadOnly = True
            Me.txtUserGroupCode.ReadOnly = True
            Me.tbKeterangan.ReadOnly = True
            Me.tbPassword.ReadOnly = True
            Me.tbKonfimasiPassword.ReadOnly = True
            'button
            Me.btnFindUserGroup.Enabled = False
            Me.cmdBatal.Enabled = False
            Me.cmdSimpan.Enabled = False
            Me.btnHapus.Enabled = True
            Me.cmdNew.Enabled = True
            Me.cmdEdit.Enabled = True
            Me.cmdCari.Enabled = True
            Me.cmdKeluar.Enabled = True
            Me.chkTempPass.Enabled = False
            Me.objCd.SetBackColor(Me.Controls, Me.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            ''default back color
            Me.tbNamaUser.BackColor = Color.White
            If Me.baru = True Then
                Me.tbKodeUser.BackColor = Color.White
                Me.txtUserGroupCode.BackColor = Color.White
            End If
            Me.txtUserGroupCode.BackColor = Color.White
            Me.tbKeterangan.BackColor = Color.White
            Me.tbPassword.BackColor = Color.White
            Me.tbKonfimasiPassword.BackColor = Color.White
            Me.btnFindUserGroup.BackColor = Color.White

            'text box
            Me.tbNamaUser.ReadOnly = False
            If Me.baru = True Then
                Me.tbKodeUser.ReadOnly = False
                Me.txtUserGroupCode.ReadOnly = False
            End If
            Me.txtUserGroupCode.ReadOnly = False
            Me.tbKeterangan.ReadOnly = False
            Me.tbPassword.ReadOnly = False
            Me.tbKonfimasiPassword.ReadOnly = False
            Me.btnFindUserGroup.Enabled = True
            'button
            Me.cmdBatal.Enabled = True
            Me.cmdSimpan.Enabled = True
            Me.btnHapus.Enabled = False
            Me.cmdNew.Enabled = False
            Me.cmdEdit.Enabled = False
            Me.cmdCari.Enabled = False
            Me.cmdKeluar.Enabled = False
            Me.chkTempPass.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim obj As New ClsMasterUserGroup
        Try
            With Me.objuser
                Me.tbKeterangan.Text = .Keterangan
                Me.tbKodeUser.Text = .KodeUser
                Me.tbNamaUser.Text = .NamaUser
                Me.tbKonfimasiPassword.Text = .Password
                Me.tbPassword.Text = .Password
                Me.TxtPrimaryUserGroup.Text = objuser.GetPrimaryGroup(Me.tbKodeUser.Text)
                Me.txtUserGroupCode.Text = obj.GetGroupCode(Me.TxtPrimaryUserGroup.Text)
                Me.chkTempPass.Checked = .UseTempPass
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub
    Private Sub SalinKeObjek()
        Try
            With Me.objuser
                .Keterangan = Me.tbKeterangan.Text
                .KodeUser = Me.tbKodeUser.Text
                .NamaUser = Me.tbNamaUser.Text
                .Password = Me.tbKonfimasiPassword.Text
                .Password = Me.tbPassword.Text
                .PrimaryGroup = Me.TxtPrimaryUserGroup.Text
                .UseTempPass = Me.chkTempPass.Checked
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cmdCari_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCari.Click
        Dim ij As New FrmCariMasterUser
        Try
            ij.Show()
            ij.TxtCari.Focus()
            ij.Hide()
            ij.ShowDialog()
            If Not ij.Kodeuser = "" Then
                Me.tbKodeUser.Text = ij.Kodeuser
                If ModUser.CariDetail(Me.tbKodeUser.Text, Me.objuser) = True Then
                    Me.SalinKeField()
                Else
                    Me.KosongkanField()
                End If
            Else
                Me.KosongkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
        Me.baru = True
        Me.AktifkanDetail()
        Me.KosongkanField()
        Me.tbKodeUser.Focus()
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        If Not Me.tbKodeUser.Text = "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.tbNamaUser.Focus()
            Me.tbPassword.ReadOnly = True
            Me.tbKonfimasiPassword.ReadOnly = True
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHapus.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.tbKodeUser.Text = "" Then
                Me.cmdCari.PerformClick()
            End If
            If Me.tbKodeUser.Text = "" Then
                'MsgBox("Please Search User Code", MsgBoxStyle.Information, "Information")
            Else
                hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Not Me.tbKodeUser.Text = "" Then
                        ModUser.Hapus(Me.tbKodeUser.Text)

                        'kembalikan ke keadaan semula
                        Me.KosongkanField()
                        MsgBox("Data deleted", MsgBoxStyle.Information, "information")
                        Me.baru = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekData() As Boolean
        Dim KoleksiData, KoleksiNamaData As Collection
        KoleksiData = New Collection
        KoleksiNamaData = New Collection
        Try
            'salin data, untuk cek data kosong
            IsiKoleksiCekData(KoleksiData, KoleksiNamaData)

            'cek data kosong
            If ModCekData.CekDataKosong(KoleksiData, KoleksiNamaData) = False Then
                Return False
            End If

            'cek konsistensi pasword yang diberika
            If Me.tbPassword.Text.Equals(Me.tbKonfimasiPassword.Text) = False Then
                MsgBox("Inconsisten password, please fix.", MsgBoxStyle.Critical, "Error")
                Me.tbKonfimasiPassword.Clear()
                Me.tbPassword.Clear()
                Me.tbPassword.Focus()
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            KoleksiData = Nothing
            KoleksiNamaData = Nothing
        End Try
    End Function

    Private Sub IsiKoleksiCekData(ByVal Koleksidata As Collection, ByVal KoleksiNamaData As Collection)
        'kosongkan
        Koleksidata.Clear()
        KoleksiNamaData.Clear()

        'isi koleksi data yang akan dicek
        Koleksidata.Add(Me.tbKodeUser)
        KoleksiNamaData.Add("User code")
        Koleksidata.Add(Me.tbNamaUser)
        KoleksiNamaData.Add("User name")
        Koleksidata.Add(Me.tbPassword)
        KoleksiNamaData.Add("Password")
        Koleksidata.Add(Me.tbKonfimasiPassword)
        KoleksiNamaData.Add("Password confirm")
    End Sub

    Private Sub tbKodeUser_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tbKodeUser.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.tbNamaUser.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tbNamaUser_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tbNamaUser.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.tbPassword.Enabled = False Then
                    Me.tbPassword.Focus()
                Else
                    Me.tbKeterangan.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tbPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tbPassword.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.tbKonfimasiPassword.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tbKonfimasiPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tbKonfimasiPassword.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.tbKeterangan.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tbKeterangan_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tbKeterangan.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.cmdSimpan.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tbKodeUser_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbKodeUser.GotFocus
        Try
            If Me.tbKeterangan.Enabled = True Then
                If Me.tbKodeUser.Enabled = True Then
                    If Me.tbKodeUser.ReadOnly = False Then
                        tbKodeUser.BackColor = BackColorAll
                    Else
                        tbKodeUser.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tbKodeUser_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbKodeUser.LostFocus
        Try
            If Me.tbKeterangan.Enabled = True Then
                If Me.tbKeterangan.ReadOnly = False Then
                    Me.tbKodeUser.BackColor = Me.TxtPrimaryUserGroup.BackColor
                Else
                    tbKodeUser.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tbNamaUser_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbNamaUser.GotFocus
        If Me.tbKeterangan.Enabled = True Then
            If Me.tbNamaUser.Enabled = True Then
                If Me.tbNamaUser.ReadOnly = False Then
                    tbNamaUser.BackColor = BackColorAll
                Else
                    tbNamaUser.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        End If
    End Sub

    Private Sub tbNamaUser_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbNamaUser.LostFocus
        If Me.tbKeterangan.Enabled = True Then
            If Me.tbKeterangan.ReadOnly = False Then
                Me.tbNamaUser.BackColor = Me.TxtPrimaryUserGroup.BackColor
            Else
                tbNamaUser.BackColor = Me.txtReadOnly.BackColor
            End If
        End If
    End Sub

    Private Sub tbPassword_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbPassword.GotFocus
        If Me.tbKeterangan.Enabled = True Then
            If Me.tbPassword.Enabled = True Then
                If Me.tbPassword.ReadOnly = False Then
                    tbPassword.BackColor = BackColorAll
                Else
                    tbPassword.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        End If
    End Sub

    Private Sub tbPassword_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbPassword.LostFocus
        If Me.tbKeterangan.Enabled = True Then
            If Me.tbKeterangan.ReadOnly = False Then
                Me.tbPassword.BackColor = Me.TxtPrimaryUserGroup.BackColor
            Else
                tbPassword.BackColor = Me.txtReadOnly.BackColor
            End If
        End If
    End Sub

    Private Sub tbKonfimasiPassword_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbKonfimasiPassword.GotFocus
        If Me.tbKeterangan.Enabled = True Then
            If Me.tbKonfimasiPassword.Enabled = True Then
                If Me.tbKonfimasiPassword.ReadOnly = False Then
                    tbKonfimasiPassword.BackColor = BackColorAll
                Else
                    tbKonfimasiPassword.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        End If
    End Sub

    Private Sub tbKonfimasiPassword_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbKonfimasiPassword.LostFocus
        If Me.tbKeterangan.Enabled = True Then
            If Me.tbKeterangan.ReadOnly = False Then
                Me.tbKonfimasiPassword.BackColor = Me.TxtPrimaryUserGroup.BackColor
            Else
                tbKonfimasiPassword.BackColor = Me.txtReadOnly.BackColor
            End If
        End If
    End Sub

    Private Sub tbKeterangan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbKeterangan.GotFocus
        If Me.tbKeterangan.Enabled = True Then
            If Me.tbKeterangan.Enabled = True Then
                If Me.tbKeterangan.ReadOnly = False Then
                    tbKeterangan.BackColor = BackColorAll
                Else
                    tbKeterangan.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        End If
    End Sub

    Private Sub tbKeterangan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbKeterangan.LostFocus
        If Me.tbKeterangan.Enabled = True Then
            If Me.tbKeterangan.ReadOnly = False Then
                Me.tbKeterangan.BackColor = Me.TxtPrimaryUserGroup.BackColor
            Else
                tbKeterangan.BackColor = Me.txtReadOnly.BackColor
            End If
        End If
    End Sub

    Private Sub btnFindUserGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindUserGroup.Click
        Dim i As New FrmCariMasterHakAkses2
        i.BackColorAll = Me.BackColorAll
        Dim objGroup As New ClsMasterUserGroup
        Try
            If Me.txtUserGroupCode.Text <> "" Then
                i.txtCari.Text = Me.txtUserGroupCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryLPB > 0 Then
                Me.TxtPrimaryUserGroup.Text = i.PrimaryLPB
                Me.txtUserGroupCode.Text = objGroup.GetGroupCode(Me.TxtPrimaryUserGroup.Text)
                Me.tbKeterangan.Focus()
            Else
                Me.TxtPrimaryUserGroup.Text = "-1"
                Me.txtUserGroupCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objGroup = Nothing
        End Try
    End Sub

    Private Sub txtUserGroupCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserGroupCode.GotFocus
        Try
            If Me.tbKeterangan.Enabled = True Then
                If Me.tbKeterangan.Enabled = True Then
                    If Me.txtUserGroupCode.ReadOnly = False Then
                        txtUserGroupCode.BackColor = BackColorAll
                    Else
                        txtUserGroupCode.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtUserGroupCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserGroupCode.LostFocus
        Try
            If Me.tbKeterangan.Enabled = True Then
                If Me.tbKeterangan.ReadOnly = False Then
                    Me.txtUserGroupCode.BackColor = Me.TxtPrimaryUserGroup.BackColor
                Else
                    txtUserGroupCode.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtUserGroupCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUserGroupCode.KeyDown
        Dim objGroup As New ClsMasterUserGroup
        Dim i As New FrmCariMasterHakAkses2
        i.BackColorAll = Me.BackColorAll
        Try
            'If e.KeyValue = 13 Then
            'If Me.txtUserGroupCode.Text <> "" Then
            'i.txtCari.Text = Me.txtUserGroupCode.Text
            'End If
            'i.ShowDialog()
            'If i.PrimaryLPB > 0 Then
            'Me.TxtPrimaryUserGroup.Text = i.PrimaryLPB
            'Me.txtUserGroupCode.Text = objGroup.GetGroupCode(Me.TxtPrimaryUserGroup.Text)
            'Me.tbKeterangan.Focus()
            'Else
            'Me.TxtPrimaryUserGroup.Text = "-1"
            'Me.txtUserGroupCode.Clear()
            'End If
            'End If
            If e.KeyValue = 13 Then
                If Me.txtUserGroupCode.Text <> "" Then
                    If objGroup.ApakahSudahAda(Me.txtUserGroupCode.Text) = True Then
                        Me.TxtPrimaryUserGroup.Text = objGroup.GetPrimary(Me.txtUserGroupCode.Text)
                        Me.txtUserGroupCode.Text = objGroup.GetGroupCode(Me.TxtPrimaryUserGroup.Text)
                        Me.tbKeterangan.Focus()
                    Else
                        Me.btnFindUserGroup.PerformClick()
                    End If
                Else
                    Me.btnFindUserGroup.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objGroup = Nothing
        End Try
    End Sub

    Private Sub txtUserGroupCode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUserGroupCode.Leave
        Dim objGroup As New ClsMasterUserGroup
        Dim i As New FrmCariMasterHakAkses2
        i.BackColorAll = Me.BackColorAll
        Try
            'If Me.txtUserGroupCode.Text <> "" Then
            'Dim Primary As Double = objGroup.GetPrimary(Me.txtUserGroupCode.Text)
            'If Primary <> -1 Then
            'Me.TxtPrimaryUserGroup.Text = Primary
            'Me.txtUserGroupCode.Text = objGroup.GetGroupCode(Me.TxtPrimaryUserGroup.Text)
            'Me.tbKeterangan.Focus()
            'Else
            'Me.btnFindUserGroup.PerformClick()
            'End If
            'Else
            'Me.TxtPrimaryUserGroup.Text = "-1"
            'Me.txtUserGroupCode.Clear()
            'End If

            ''standar coding
            If Me.txtUserGroupCode.Text <> "" Then
                If objGroup.ApakahSudahAda(Me.txtUserGroupCode.Text) = True Then
                    Me.TxtPrimaryUserGroup.Text = objGroup.GetPrimary(Me.txtUserGroupCode.Text)
                    Me.txtUserGroupCode.Text = objGroup.GetGroupCode(Me.TxtPrimaryUserGroup.Text)
                Else
                    If Me.txtUserGroupCode.Text <> "" Then
                        i.txtCari.Text = Me.txtUserGroupCode.Text
                    End If
                    i.ShowDialog()
                    If i.PrimaryLPB > 0 Then
                        Me.TxtPrimaryUserGroup.Text = i.PrimaryLPB
                        Me.txtUserGroupCode.Text = objGroup.GetGroupCode(Me.TxtPrimaryUserGroup.Text)
                        Me.tbKeterangan.Focus()
                    Else
                        Me.TxtPrimaryUserGroup.Text = "-1"
                        Me.txtUserGroupCode.Clear()
                        Me.txtUserGroupCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryUserGroup.Text = "-1"
                Me.txtUserGroupCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objGroup = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub FrmMasterUser_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub
End Class