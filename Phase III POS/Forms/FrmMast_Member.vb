Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmMast_Member
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris

    Dim baru As Boolean
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData
    Dim Obj As New ClsMast_Member
    Dim ObjCust As New clsMasterCustomer

    Private Sub FrmMasterMember_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Obj = New ClsMast_Member
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()
            Me.dtpExpiry.Enabled = True
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
            Me.dtpExpiry.Enabled = False
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
            Me.TxtPrimaryCust.Text = "-1"
            Me.dtpExpiry.Value = Now
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
                        MsgBox("Duplication in reference name, please change your code", _
                        MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Duplikasi code, silahkan ubah reference name", _
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
                            MsgBox("Duplication in reference name, please change your code", _
                            MsgBoxStyle.Information, "Information")
                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                            MsgBox("Duplikasi code, silahkan ubah reference name", _
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

            If Me.TxtPrimaryCust.Text = "-1" Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input customer code", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan customer code", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtCustCode.Focus()
                Return False
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

            'Item tidak boleh kosong 
            'If Me.DGVItem.Rows.Count <= 0 Then
            'If Bahasa = EnumBahasa.Inggris Then
            'MsgBox("Please input item data", MsgBoxStyle.Information, "Information")
            'ElseIf Bahasa = EnumBahasa.Indonesia Then
            'MsgBox("Masukan data item", MsgBoxStyle.Information, "Informasi")
            'End If
            'Return False
            'End If

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
                .PrimaryCustomer = Me.TxtPrimaryCust.Text
                .ExpiryMember = Me.dtpExpiry.Value
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
                Me.TxtPrimaryCust.Text = .PrimaryCustomer
                Me.TxtCustCode.Text = ObjCust.GetCustomerCode(Me.TxtPrimaryCust.Text)
                Me.TxtCustomerName.Text = ObjCust.GetCustName(Me.TxtPrimaryCust.Text)
                Me.dtpExpiry.Value = .ExpiryMember

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
                        Me.Obj.CariDetail(Me.TxtPrimary.Text, _
                        Me.ObjCD, Me)
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
        Dim ijk As New FrmMast_Member_Cari
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

    Private Sub TxtCustCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCustCode.Enabled = True Then
                    If Me.TxtCustCode.ReadOnly = False Then
                        TxtCustCode.BackColor = Me.BackColorAll
                    Else
                        TxtCustCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCustCode.Enabled = True Then
                    If Me.TxtCustCode.ReadOnly = False Then
                        TxtCustCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtCustCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCustCode.KeyDown

        Dim objCode As New clsMasterCustomer
        Dim i As New FrmCariMasterCustomer
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCustCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtCustCode.Text) = True Then
                        Me.TxtPrimaryCust.Text = objCode.GetPrimary(Me.TxtCustCode.Text)
                        Me.TxtCustCode.Text = objCode.GetCustomerCode(Me.TxtPrimaryCust.Text)
                        Me.TxtCustomerName.Text = objCode.GetCustName(Me.TxtPrimaryCust.Text)
                        Me.dtpExpiry.Focus()
                    Else
                        Me.btnFindCust.PerformClick()
                    End If
                Else
                    Me.btnFindCust.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCustCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustCode.Leave

        Dim objCode As New clsMasterCustomer
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCustCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtCustCode.Text) = True Then
                    Me.TxtPrimaryCust.Text = objCode.GetPrimary(Me.TxtCustCode.Text)
                    Me.TxtCustCode.Text = objCode.GetCustomerCode(Me.TxtPrimaryCust.Text)
                    Me.TxtCustomerName.Text = objCode.GetCustName(Me.TxtPrimaryCust.Text)
                Else
                    If Me.TxtCustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtPrimaryCust.Text = i.PrimaryCustomer
                        Me.TxtCustCode.Text = objCode.GetCustomerCode(Me.TxtPrimaryCust.Text)
                        Me.TxtCustomerName.Text = objCode.GetCustName(Me.TxtPrimaryCust.Text)
                    Else
                        Me.TxtCustCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.TxtPrimaryCust.Text = "-1"
                        Me.TxtCustCode.Focus()
                    End If
                End If
            Else
                Me.TxtCustCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.TxtPrimaryCust.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnFindCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCust.Click
        Dim i As New FrmCariMasterCustomer
        Dim ObjCode As New clsMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCustCode.Text <> "" Then
                i.txtCari.Text = Me.TxtCustCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtPrimaryCust.Text = i.PrimaryCustomer
                Me.TxtCustCode.Text = ObjCode.GetCustomerCode(Me.TxtPrimaryCust.Text)
                Me.TxtCustomerName.Text = ObjCode.GetCustName(Me.TxtPrimaryCust.Text)
                Me.dtpExpiry.Focus()
            Else
                Me.TxtCustCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.TxtPrimaryCust.Text = "-1"
                Me.TxtCustCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub
End Class