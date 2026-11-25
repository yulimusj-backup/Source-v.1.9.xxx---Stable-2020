Option Explicit On
Option Strict Off
Public Class FrmMasterRefParty

    Public UserCode As String = ""
    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim baru As Boolean

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjClsMasRefParty As New ClsMasterRefParty
    Dim DefPrec As Double = 0

    Private Sub AktifkanDetail()
        Try
            ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryRefParty.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)


            ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.ChkAktif.Enabled = True
            'If Me.TxtPrimaryRefParty.Text <> "-1" Then
            '    If ObjClsMasRefParty.ApakahSudahAda3(Me.TxtPrimaryRefParty.Text) = True Then
            '        Me.ChkAktif.Enabled = False
            '    Else
            '        Me.ChkAktif.Enabled = True
            '    End If
            'End If
            If baru = True Then
                Me.ChkAktif.Checked = True
            End If
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ChkAktif.Enabled = False
            ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)

            Me.CancelButton = Me.btnExit
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ObjCD.KosongkanControl(Me.gboDetail.Controls, True, False, True)

            Me.TxtPrimaryRefParty.Text = "-1"
            Me.ChkAktif.Checked = False
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection

        Dim KoleksiNamaData1 As New Collection
        Dim KoleksiTextBox1 As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Code")
            KoleksiNamaData.Add("Description")
            KoleksiNamaData.Add("Email")
            KoleksiTextBox.Add(Me.TxtCode)
            KoleksiTextBox.Add(Me.TxtDescription)
            KoleksiTextBox.Add(Me.TxtEmail)

            If ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If


            'khusus untuk remarks saja
            KoleksiNamaData1.Add("Remarks")
            KoleksiTextBox1.Add(Me.TxtRemarks)

            If Me.CekDataRemarks(KoleksiTextBox1, KoleksiNamaData1) = False Then
                Return False
            End If


            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.ObjClsMasRefParty.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    MsgBox("Duplication in Code, please change your Code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.ObjClsMasRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                If ProductGroupLama <> Me.TxtCode.Text Then
                    If Me.ObjClsMasRefParty.ApakahSudahAda(Me.TxtCode.Text) = True Then
                        MsgBox("Duplication in Code, please change your Code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai timestamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjClsMasRefParty.GetTimestamp(Me.TxtPrimaryRefParty.Text)
                If Me.ObjClsMasRefParty.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjClsMasRefParty.CariDetail(Me.ObjClsMasRefParty.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'cekEmail
            If Me.TxtEmail.Text <> "" Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(Me.TxtEmail.Text) = False Then
                    MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                    Me.TxtEmail.Focus()
                    Me.TxtEmail.SelectAll()
                    Return False
                End If
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

    Private Sub SalinKeObjek()
        Try
            With Me.ObjClsMasRefParty
                .Primary = Me.TxtPrimaryRefParty.Text
                .Code = Me.TxtCode.Text
                .Desc = Me.TxtDescription.Text
                .Email = Me.TxtEmail.Text
                .Remarks = Me.TxtRemarks.Text

                If Me.ChkAktif.Checked = False Then
                    .Aktif = ClsMasterRefParty.EnumAktif.No
                Else
                    .Aktif = ClsMasterRefParty.EnumAktif.Yes
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjClsMasRefParty
                Me.TxtPrimaryRefParty.Text = .Primary
                Me.TxtCode.Text = .Code
                Me.TxtDescription.Text = .Desc
                Me.TxtEmail.Text = .Email
                Me.TxtRemarks.Text = .Remarks

                If .Aktif = False Then
                    Me.ChkAktif.Checked = False
                Else
                    Me.ChkAktif.Checked = True
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjClsMasRefParty.SimpanBaru(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'disimpan atau tidak
                    'If Me.ObjClsMasRefParty.ApakahBolehEdit(Me.TxtPrimaryRefParty.Text, UserCode, True, True) = True Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjClsMasRefParty.SimpanPerubahan(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtCode.Text = "" And Me.TxtCode.Text <> "" Then
                'Tipe batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
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

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterRefParty_Cari
        Try
            hj.BackColorAll = Me.BAckColorAll
            hj.ShowDialog()

            If Not hj.PrimaryRefParty = -1 Then
                If Me.ObjClsMasRefParty.CariDetail(hj.PrimaryRefParty) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()

                End If
                Me.btnFind.Focus()
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TxtCode.Focus()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtPrimaryRefParty.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimaryRefParty.Text <> "-1" Then
                'If Me.ObjClsMasRefParty.ApakahBolehEdit(Me.ObjClsMasRefParty.Primary, UserCode, True) = True Then
                Me.baru = False
                Me.ObjClsMasRefParty.CariDetail(Me.ObjClsMasRefParty.Primary)
                Me.AktifkanDetail()
                Me.SalinKeField()
                'End If
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
                If Me.ObjClsMasRefParty.ApakahBolehHapus(Me.TxtCode.Text, UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjClsMasRefParty.Hapus(Me.TxtPrimaryRefParty.Text, UserCode) = True Then
                            Me.KosongkanField()
                            Me.NonAktifkanDetail()
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
        Me.Close()
    End Sub

    Private Sub FrmMasterLoad_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjClsMasRefParty.SimpanBaru(Me.UserCode) = True Then
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
                            If Me.ObjClsMasRefParty.SimpanPerubahan(Me.UserCode) = True Then
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

    Private Sub txtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.GotFocus, TxtDescription.GotFocus, _
    TxtEmail.GotFocus, TxtRemarks.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BAckColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtcode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.LostFocus, TxtDescription.LostFocus, _
    TxtEmail.LostFocus, TxtRemarks.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimaryRefParty.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.btnFind.Select()
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.KosongkanField()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function PengecekanEmail() As Boolean
        Try
            If Me.TxtEmail.Text <> "" Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                If Expression.IsMatch(Me.TxtEmail.Text) = False Then
                    MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                    Me.TxtEmail.Focus()
                    Me.TxtEmail.SelectAll()
                    Return False
                End If
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Private Function HapusSpasi(ByVal Data As String) As String
        Try
            Do While Data.EndsWith(" ") And Data.Length > 0
                Data = Data.Remove(Data.Length - 1, 1)
            Loop
            Return Data
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function CekDataRemarks(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim i As Int16
        Dim data As String
        Dim Txt As Windows.Forms.TextBox

        Try
            Do While i < KoleksiTextBox.Count
                'trim kanan kiri
                Txt = TryCast(KoleksiTextBox.Item(i + 1), Windows.Forms.TextBox)
                If IsNothing(Txt) = False Then
                    Txt.Text = Txt.Text.Trim
                End If
                'salin data
                data = KoleksiTextBox.Item(i + 1).text
                data = HapusSpasi(data)

                ''cek data
                'If data.Length = 0 Then
                '    MsgBox(Convert.ToString(KoleksiNamaData.Item(i + 1)) + " cannot be an empty string.", MsgBoxStyle.Critical, "Error")
                '    KoleksiTextBox.Item(i + 1).focus()
                '    Return False
                'End If

                'cek berkaitan dengan karakter yang tidak diingikan --> ', *, %
                If data.Contains("'") = True Or data.Contains("*") = True Or data.Contains("%") = True Then
                    MsgBox("" + KoleksiNamaData.Item(i + 1) _
                        + " data can not contain the following characters: ' , % , or * .", MsgBoxStyle.Critical, "Error")
                    Return False
                End If
                i += 1
            Loop

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

End Class