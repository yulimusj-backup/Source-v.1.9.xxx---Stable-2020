Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiReminder_New
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightSeaGreen
    Public Baru As Boolean
    Public ObjRemain As New ClsTransaksiReminder
    Public TBLUSer As New Core_Module.DSUser.tbluser1DataTable
    Public PrimaryReminder As Double = -1

    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Ok As Boolean
    Dim Cancel_No As Boolean = False
    '''Dim FrmReminder As FrmTransaksiReminder

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Subject")
            KoleksiTextBox.Add(Me.TxtSubject)
            KoleksiNamaData.Add("Memo")
            KoleksiTextBox.Add(Me.TxtMemo)


            If ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            ' data salah kalau tidak ada yang di cek
            If Me.clbUser.CheckedItems.Count <= 0 Then
                MsgBox("Please choose remind who", MsgBoxStyle.Information, "Information")
                Return False
            End If


            'cek data yang salah, double dalam kode
            'Dim AkunCode As String = ""
            'AkunCode = Me.TxtSubject.Text
            If Me.Baru = True Then
                'If Objakun.ApakahSudahAda(AkunCode) = True Then
                'MsgBox("Duplication in account code, please change your code", _
                ' MsgBoxStyle.Information, "Information")
                'Me.TxtAccountNo.Focus()
                'Return False
                'End If
            Else
                'Dim TipeCodeLama As String = Me.Objakun.GetAccountCode(Me.TxtPrimaryAccount.Text)
                'If TipeCodeLama <> AkunCode Then
                'If Me.Objakun.ApakahSudahAda(AkunCode) = True Then
                'MsgBox("Duplication in account code, please change your code", _
                'MsgBoxStyle.Information, "Information")
                'Me.TxtAccountNo.Focus()
                'Return False
                'End If
                'End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit

                Dim OldTimeStamp As Date = Me.ObjRemain.GetTimestamp(Me.TxtPrimary.Text)
                If Me.ObjRemain.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjRemain.CariDetail(Me.clbUser, Me.TBLUSer, Me.ObjRemain.Primary)
                        'Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'pastikan text yang dapat di inputkan pada memo maksimal 1000 karakter
            If Me.TxtMemo.Text.Length > 1000 Then
                MsgBox("Max length for memo column is 1000 characters", MsgBoxStyle.Information, "Information")
                Me.TxtMemo.Focus()
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

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'kalo simpan data berhasil maka set dialog result jadi ok --> digunakan untuk menandai apakah harus reload data ulang
            'pada frmTransaksiReminder
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjRemain.SimpanBaru(Me.clbUser, Me.TBLUSer, Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            ''Me.SalinKeField()
                            ''Me.NonAktifkanDetail()
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                            Me.Ok = True
                            Me.Cancel_No = False
                            Me.Close()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjRemain.SimpanPerubahan(Me.clbUser, Me.TBLUSer, Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            ''Me.SalinKeField()
                            ''Me.NonAktifkanDetail()
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                            Me.Ok = True
                            Me.Cancel_No = False
                            Me.Close()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiReminder_New_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.TxtMemo.MaxLength = 1000
            Me.ObjUser.IsiListBoxUser(Me.clbUser, Me.TBLUSer, "")

            If Me.Baru = False Then
                Me.ObjRemain.CariDetail(Me.clbUser, Me.TBLUSer, Me.PrimaryReminder)
                With Me.ObjRemain
                    Me.TxtPrimary.Text = .Primary
                    Me.TxtSubject.Text = .Subject
                    Me.TxtMemo.Text = .Memo
                    Me.dtpDate.Value = .Tanggal
                End With
            End If
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            With Me.ObjRemain
                .Primary = Me.TxtPrimary.Text
                .Subject = Me.TxtSubject.Text
                .Tanggal = Me.dtpDate.Value
                .Memo = Me.TxtMemo.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtSubject.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                        Me.Cancel_No = False
                    Else
                        ''Me.SalinKeField()
                        Me.Cancel_No = False
                    End If
                Else
                    Me.Ok = True
                    Me.Cancel_No = True
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnCancel.Focus()
            Me.Ok = True
            Me.Cancel_No = False
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAkunDetail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.Cancel_No = False Then
                If Me.Ok = False Then
                    If Me.TxtSubject.Enabled = True Then
                        Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ApakahDataBenar() = True Then
                                Me.SalinKeObjek()
                                If Baru = True Then
                                    'simpan baru
                                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                    If Me.ObjRemain.SimpanBaru(Me.clbUser, Me.TBLUSer, Me.UserCode) = True Then
                                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                        Me.DialogResult = Windows.Forms.DialogResult.OK
                                    Else
                                        e.Cancel = True
                                    End If
                                Else
                                    'simpan perubahan
                                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                    If Me.ObjRemain.SimpanPerubahan(Me.clbUser, Me.TBLUSer, Me.UserCode) = True Then
                                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                        Me.DialogResult = Windows.Forms.DialogResult.OK
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
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.Controls, True, False, True, False, False, False, True, False)
            ''ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.ObjCD.SetBackColor(Me.Controls, Me.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KosongkanField()
        Try
            ObjCD.KosongkanControl(Me.Controls, True, False, True)
            Me.TxtPrimary.Text = "-1"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMemo.GotFocus
        Try
            If Me.TxtMemo.Enabled = True Then
                If Me.TxtMemo.ReadOnly = False Then
                    Me.TxtMemo.BackColor = Me.BackColorAll
                Else
                    TxtMemo.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMemo.LostFocus
        Try
            If Me.TxtMemo.Enabled = True Then
                If Me.TxtMemo.ReadOnly = False Then
                    Me.TxtMemo.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtMemo.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtScheduleNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSubject.GotFocus
        Try
            If Me.TxtSubject.Enabled = True Then
                If Me.TxtSubject.ReadOnly = False Then
                    Me.TxtSubject.BackColor = Me.BackColorAll
                Else
                    TxtSubject.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtScheduleNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSubject.LostFocus
        Try
            If Me.TxtSubject.Enabled = True Then
                If Me.TxtSubject.ReadOnly = False Then
                    Me.TxtSubject.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtSubject.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class