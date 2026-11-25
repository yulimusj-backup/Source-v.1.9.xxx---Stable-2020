Option Explicit On
Option Strict Off

Public Class FrmManajemenHakKhusus
    Dim objUser As ClsUser
    Dim objhakKhusus As ClsHakKhusus

    Dim baru As Boolean
    Private Sub ListSemuaForm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListSemuaForm.SelectedIndexChanged
        If Me.ListSemuaForm.SelectedIndex <> -1 Then
            Me.AturListBox(Me.ListSemuaForm)
        End If
    End Sub

    Private Sub AturListBox(ByVal LB As ListBox)
        'atur agar yang terpilih hanya 1
        If LB.Equals(Me.ListFormAktif) Then
            Me.ListSemuaForm.SelectedIndex = -1
        Else
            Me.ListFormAktif.SelectedIndex = -1
        End If
    End Sub

    Private Sub ListFormAktif_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListFormAktif.SelectedIndexChanged
        If Me.ListFormAktif.SelectedIndex <> -1 Then
            Me.AturListBox(Me.ListFormAktif)
        End If
    End Sub

    Private Sub btnKanan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKanan.Click
        If Me.ListSemuaForm.SelectedItems.Count > 0 Then
            Me.ListFormAktif.Items.Add(Me.ListSemuaForm.SelectedItem)
            Me.ListSemuaForm.Items.RemoveAt(Me.ListSemuaForm.SelectedIndex)
        End If
    End Sub

    Private Sub btnKiri_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKiri.Click
        If Me.ListFormAktif.SelectedItems.Count > 0 Then
            Me.ListSemuaForm.Items.Add(Me.ListFormAktif.SelectedItem)
            Me.ListFormAktif.Items.RemoveAt(Me.ListFormAktif.SelectedIndex)
        End If
    End Sub

    Private Sub btnKananSemua_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKananSemua.Click
        Dim i As String
        If Me.ListSemuaForm.Items.Count > 0 Then
            For Each i In Me.ListSemuaForm.Items
                Me.ListFormAktif.Items.Add(i)
            Next
            Me.ListSemuaForm.Items.Clear()
        End If
    End Sub

    Private Sub btnKiriSemua_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKiriSemua.Click
        Dim i As String
        If Me.ListFormAktif.Items.Count > 0 Then
            For Each i In Me.ListFormAktif.Items
                Me.ListSemuaForm.Items.Add(i)
            Next
            Me.ListFormAktif.Items.Clear()
        End If
    End Sub

    Private Sub CBOkodeuser_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBOkodeuser.SelectedIndexChanged
        Dim obj As New ClsHakKhusus
        Dim item1, item2 As String

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'cari data ke basis data
            ModHakKhusus.IsiListBox(Me.CBOkodeuser.SelectedItem, Me.ListFormAktif)
            ModHakKhusus.IsiListBox(Me.ListSemuaForm)
            For Each item1 In Me.ListFormAktif.Items
                For Each item2 In Me.ListSemuaForm.Items
                    If item1 = item2 Then
                        Me.ListSemuaForm.Items.Remove(item2)
                        Exit For
                    End If
                Next
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            obj = Nothing
        End Try
    End Sub

    Private Sub BtnKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnKeluar.Click
        Me.Close()
    End Sub

    Private Sub BtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit.Click
        Me.AktifkanField()
    End Sub
    Private Sub AktifkanField()
        Try
            Me.ListFormAktif.Enabled = True
            Me.ListSemuaForm.Enabled = True
            Me.BtnEdit.Enabled = False
            Me.BtnKeluar.Enabled = False
            Me.CBOkodeuser.Enabled = False
            Me.TxtNamaUser.ReadOnly = True
            Me.btnKanan.Enabled = True
            Me.btnKiri.Enabled = True
            Me.btnKiriSemua.Enabled = True
            Me.btnKananSemua.Enabled = True
            Me.BtnSimpan.Enabled = True
            Me.BtnBatal.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmManagementhakKhusus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'isi combo box
            ModUser.IsiComboBox(Me.CBOkodeuser)
            If Me.CBOkodeuser.Items.Count > 0 Then
                Me.CBOkodeuser.SelectedIndex = 0
                Me.BtnEdit.Enabled = True
            Else
                Me.BtnEdit.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnSimpan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSimpan.Click
        Try
            If MsgBox("Anda yakin akan menyimpan data ini?", MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.Yes Then
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

                Me.objUser = New ClsUser
                Me.objUser.KodeUser = Me.CBOkodeuser.SelectedItem
                Me.SalinKeObjek()
                ModUser.UpdateHakKhusus(Me.objUser, Me.objUser.KoleksiHakKhusus)

                'hanya untuk testing
                ModUser.CariDetail(Me.CBOkodeuser.SelectedItem, Me.objUser)
                Me.SalinKeField()

                System.Windows.Forms.Cursor.Current = Cursors.Arrow
                Me.NonAktifkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.objUser = Nothing
        End Try
    End Sub
    Private Sub SalinKeField()
        Try
            With Me.objUser
                Me.CBOkodeuser.SelectedItem = .KodeUser

                'isi data list box
                ModHakKhusus.IsiListBox(Me.CBOkodeuser.SelectedItem, Me.ListFormAktif)

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub SalinKeObjek()
        Dim i As String
        Try
            Me.objUser.KodeUser = Me.CBOkodeuser.SelectedItem
            Me.objUser.KoleksiHakKhusus.Clear()

            For Each i In Me.ListFormAktif.Items
                Me.objhakKhusus = New ClsHakKhusus
                Me.objhakKhusus.Description = i
                Me.objhakKhusus.KodehakKhusus = Me.objhakKhusus.GetKode1(Me.objhakKhusus.Description)
                Me.objUser.KoleksiHakKhusus.Add(Me.objhakKhusus)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub NonAktifkanField()
        'text box
        Me.ListFormAktif.Enabled = False
        Me.ListSemuaForm.Enabled = False
        Me.BtnEdit.Enabled = True
        Me.BtnKeluar.Enabled = True
        Me.CBOkodeuser.Enabled = True
        Me.TxtNamaUser.ReadOnly = False
        Me.btnKanan.Enabled = False
        Me.btnKiri.Enabled = False
        Me.btnKiriSemua.Enabled = False
        Me.btnKananSemua.Enabled = False
        Me.BtnSimpan.Enabled = False
        Me.BtnBatal.Enabled = False
    End Sub


    Private Sub BtnBatal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBatal.Click
        If MsgBox("Anda yakin batal menyimpan data ini?", MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.Yes Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Dim i As String = Me.CBOkodeuser.SelectedItem
            Me.NonAktifkanField()
            Me.CBOkodeuser.SelectedIndex = -1
            Me.CBOkodeuser.SelectedItem = i

            Windows.Forms.Cursor.Current = Cursors.Arrow
        End If
    End Sub
End Class