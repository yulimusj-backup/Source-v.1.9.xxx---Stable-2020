Option Explicit On
Option Strict Off

Public Class FrmMasterProses
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim Obj As New ClsMasterProses
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.BackColorDefault()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.TxtCode.BackColor = Me.TxtPrimary.BackColor
            Me.TxtDesc.BackColor = Me.TxtPrimary.BackColor
            Me.TxtRemarks.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.TxtCode.BackColor = Me.TxtReadOnly.BackColor
            Me.TxtDesc.BackColor = Me.TxtReadOnly.BackColor
            Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            Me.TxtPrimary.Text = "-1"
            Me.cboStatus.SelectedIndex = 0
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
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Kode")
            KoleksiTextBox.Add(Me.TxtCode)
            KoleksiNamaData.Add("Desc")
            KoleksiTextBox.Add(Me.TxtDesc)

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    MsgBox("Code duplication, please change code.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtCode.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetCode(Me.TxtPrimary.Text)
                If TipeCodeLama <> Me.TxtCode.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtCode.Text) = True Then
                        MsgBox("Code duplication, please change code.", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user changed this data. Reload from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.Obj.Primary)
                        Me.SalinKeField()
                    End If
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
            With Me.Obj
                .Code = Me.TxtCode.Text
                .Desc = Me.TxtDesc.Text
                .Remarks = Me.TxtRemarks.Text
                If Me.cboStatus.SelectedIndex = 0 Then
                    .Aktif = ClsMasterProses.Enum_Aktif.Aktif
                Else
                    .Aktif = ClsMasterProses.Enum_Aktif.NonAktif
                End If
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
                If .Aktif = ClsMasterProses.Enum_Aktif.Aktif Then
                    Me.cboStatus.SelectedIndex = 0
                Else
                    Me.cboStatus.SelectedIndex = 1
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

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ShowIcon = False
            Me.Obj = New ClsMasterProses
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
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
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode) Then
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
                        If Me.Obj.SimpanPerubahan(Me.UserCode) Then
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
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Konfirmasi")
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

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtCode.Focus()
            Me.TxtPrimary.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ijk As New FrmMasterProsesCari

        Try
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            ijk.mStatusAktif = True
            ijk.mSearchAll = True
            ijk.ShowDialog()

            If Not ijk.mPrimary = -1 Then
                If Me.Obj.CariDetail(ijk.mPrimary) = True Then
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
                Me.baru = False
                Me.AktifkanDetail()
                Me.TxtCode.Focus()
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
                'khusus untuk IDR, tidak bisa dihapus, karena digunakan oleh sistem
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimary.Text, Me.UserCode, True, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.TxtPrimary.Text, Me.UserCode, True, True) = True Then
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
        Me.Close()
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
                            If Me.Obj.SimpanBaru(Me.UserCode) = True Then
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
                            If Me.Obj.SimpanPerubahan(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Infomation")
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
End Class