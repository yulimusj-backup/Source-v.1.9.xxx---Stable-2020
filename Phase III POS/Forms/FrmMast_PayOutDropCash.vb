Option Explicit On
Option Strict Off
Imports Core_Module, AccModule

Public Class FrmMast_PayOutDropCash
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris

    Dim baru As Boolean
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData
    Dim Obj As New ClsMast_PayoutDropCash
    Dim ObjAkun As New ClsMasterAkun

    Private Sub FrmMast_PayOutDropCash_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.CboTipe.SelectedIndex = 0
            Me.Obj = New ClsMast_PayoutDropCash
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboTipe_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboTipe.SelectedIndexChanged
        Try
            If Me.CboTipe.SelectedIndex = 0 Then
                Me.LblAccount.Text = "Cost Account"
            Else
                Me.LblAccount.Text = "Income Account"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccountCostCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtAccountCostCode.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAccountCostCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtAccountCostCode.Text) = True Then
                        Me.TxtPrimaryAccountCost.Text = objCode.GetPrimary(Me.TxtAccountCostCode.Text)
                        Me.TxtAccountCostCode.Text = objCode.GetAccountCode(Me.TxtPrimaryAccountCost.Text)
                        Me.TxtRemarks.Focus()
                    Else
                        Me.btnFindAccount.PerformClick()
                    End If
                Else
                    Me.btnFindAccount.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtAccountCostCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtAccountCostCode.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccountCostCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtAccountCostCode.Text) = True Then
                    Me.TxtPrimaryAccountCost.Text = objCode.GetPrimary(Me.TxtAccountCostCode.Text)
                    Me.TxtAccountCostCode.Text = objCode.GetAccountCode(Me.TxtPrimaryAccountCost.Text)
                Else
                    If Me.TxtAccountCostCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtAccountCostCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccountCost.Text = i.mPrimaryAkun
                        Me.TxtAccountCostCode.Text = objCode.GetAccountCode(Me.TxtPrimaryAccountCost.Text)
                    Else
                        Me.TxtAccountCostCode.Clear()
                        Me.TxtPrimaryAccountCost.Text = "-1"
                        Me.TxtAccountCostCode.Focus()
                    End If
                End If
            Else
                Me.TxtAccountCostCode.Clear()
                Me.TxtPrimaryAccountCost.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFindAccount.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccountCostCode.Text <> "" Then
                i.txtCari.Text = Me.TxtAccountCostCode.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccountCost.Text = i.mPrimaryAkun
                Me.TxtAccountCostCode.Text = ObjCode.GetAccountCode(Me.TxtPrimaryAccountCost.Text)
                Me.TxtRemarks.Focus()
            Else
                Me.TxtAccountCostCode.Clear()
                Me.TxtPrimaryAccountCost.Text = "-1"
                Me.TxtAccountCostCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

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
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.TxtPrimary.Text = "-1"
            Me.CboTipe.SelectedIndex = 0
            Me.TxtPrimaryAccountCost.Text = "-1"

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
            KoleksiNamaData.Add("Code")
            KoleksiTextBox.Add(Me.TxtCode)
            KoleksiNamaData.Add("Description")
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

            If String.IsNullOrEmpty(Me.TxtDesc.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input description", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan description", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtDesc.Focus()
                Return False
            End If

            If Me.TxtPrimaryAccountCost.Text = "-1" Then
                If Me.CboTipe.SelectedIndex = 0 Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input cost account", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan cost account", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtAccountCostCode.Focus()
                    Return False
                Else
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input income account", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan income account", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtAccountCostCode.Focus()
                    Return False
                End If
            End If

            If String.IsNullOrEmpty(Me.TxtRemarks.Text) = False Then
                If ObjCD.CekDataKosong(Me.TxtRemarks.Text, False) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtRemarks.Focus()
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
                If Me.CboTipe.SelectedIndex = 0 Then
                    .Tipe = ClsMast_PayoutDropCash.EnumTipe.Payout
                Else
                    .Tipe = ClsMast_PayoutDropCash.EnumTipe.Drop_Cash
                End If
                .PrimaryCostAccount = Me.TxtPrimaryAccountCost.Text
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
                If .Tipe = ClsMast_PayoutDropCash.EnumTipe.Payout Then
                    Me.CboTipe.SelectedIndex = 0
                Else
                    Me.CboTipe.SelectedIndex = 1
                End If
                Me.TxtPrimaryAccountCost.Text = .PrimaryCostAccount
                Me.TxtAccountCostCode.Text = ObjAkun.GetAccountCode(Me.TxtPrimaryAccountCost.Text)

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
                If Baru = True Then
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
            Me.Baru = True
            Obj.ResetObject()
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
        Dim ijk As New FrmMast_PayOutDropCash_Cari
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

    Private Sub TxtDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.GotFocus
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

    Private Sub TxtAccountCostCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccountCostCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtAccountCostCode.Enabled = True Then
                    If Me.TxtAccountCostCode.ReadOnly = False Then
                        Me.TxtAccountCostCode.BackColor = Me.BackColorAll
                    Else
                        Me.TxtAccountCostCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccountCostCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccountCostCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtAccountCostCode.Enabled = True Then
                    If Me.TxtAccountCostCode.ReadOnly = False Then
                        Me.TxtAccountCostCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtAccountCostCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class