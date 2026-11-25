Option Explicit On
Option Strict Off

Public Class FrmMasterHakAkses

    Dim objhakakses As ClsHakAkses
    Dim baru As Boolean

    Private Sub cmdCari_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCari.Click
        Dim hj As New FrmCariHakAkses
        Try
            hj.ShowDialog()
            If Not hj.KodeHakAkses = "" Then
                Me.objhakakses.KodeHakAkses = hj.KodeHakAkses
                If ModHakAkses.CariDetail(Me.objhakakses) = True Then
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
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHapus.Click
        'hapus hanya bisa dilakukan jika hakakses tidak pernah digunakan oleh finish goods yang ada
        'cek ke tabel finish gooods, jika ada ID hakakses yang bersangkutan maka data tidak bisa dihapus
        If Not Me.TxtKodeHakAkses.Text = "" Then
            If MsgBox("Anda yakin akan menghapus data ini?", MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.Yes Then
                ModHakAkses.Hapus(Me.TxtKodeHakAkses.Text)

                'kembalikan ke keadaan semula
                Me.KosongkanField()
                Me.baru = True
            End If
        End If
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Dim hasil As MsgBoxResult
        Try
            'konfirmasi
            hasil = MsgBox("Anda yakin akan menyimpan data ini?", MsgBoxStyle.YesNo, "Konfirmasi")

            If hasil = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor

                If CekData() = True Then
                    SalinKeObjek()
                    If baru = True Then
                        SimpanBaru()
                    Else
                        SimpanPerubahan()
                    End If
                End If

                Windows.Forms.Cursor.Current = Cursors.Arrow
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanBaru()
        If ModHakAkses.Simpan(Me.objhakakses, Me.TxtKodeHakAkses, Me.TxtNamaHakAkses, Me.txtDeskripsi) = True Then
            'kembalikan ke kondisi awal
            Me.NonAktifkanDetail()
        End If
    End Sub

    Private Sub SimpanPerubahan()
        If ModHakAkses.SimpanPerubahan(Me.objhakakses, Me.TxtKodeHakAkses, Me.TxtNamaHakAkses, Me.txtDeskripsi) = True Then
            'kembalikan ke kondisi awal
            Me.NonAktifkanDetail()
        End If
    End Sub

    Private Sub FrmMasterhakakses_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objhakakses = New ClsHakAkses
        Me.baru = True
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objhakakses
                Me.TxtKodeHakAkses.Text = .KodeHakAkses
                Me.TxtNamaHakAkses.Text = .NamaHakAkses
                Me.txtDeskripsi.Text = .Description
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            With Me.objhakakses
                .KodeHakAkses = Me.TxtKodeHakAkses.Text
                .NamaHakAkses = Me.TxtNamaHakAkses.Text
                .Description = Me.txtDeskripsi.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        'text box
        Me.TxtNamaHakAkses.Enabled = True
        If Me.baru = True Then
            Me.TxtKodeHakAkses.Enabled = True
        Else
            Me.TxtKodeHakAkses.Enabled = False
        End If
        Me.txtDeskripsi.Enabled = True

        'button
        Me.cmdCancel.Enabled = True
        Me.cmdOK.Enabled = True
        Me.btnHapus.Enabled = False
        Me.cmdNew.Enabled = False
        Me.cmdEdt.Enabled = False
        Me.cmdCari.Enabled = False
        Me.cmdKeluar.Enabled = False
    End Sub

    Private Sub NonAktifkanDetail()
        'text box
        Me.TxtNamaHakAkses.Enabled = False
        Me.TxtKodeHakAkses.Enabled = False
        Me.txtDeskripsi.Enabled = False

        'button
        Me.cmdCancel.Enabled = False
        Me.cmdOK.Enabled = False
        Me.btnHapus.Enabled = True
        Me.cmdNew.Enabled = True
        Me.cmdEdt.Enabled = True
        Me.cmdCari.Enabled = True
        Me.cmdKeluar.Enabled = True
    End Sub

    Private Sub KosongkanField()
        'text box
        Me.TxtNamaHakAkses.Clear()
        Me.txtDeskripsi.Clear()
    End Sub

    Private Function CekData() As Boolean
        Dim KoleksiData, KoleksiNamaData As Collection
        KoleksiData = New Collection
        KoleksiNamaData = New Collection
        Try
            'salin data, untuk cek data kosong
            KoleksiData.Clear()
            KoleksiNamaData.Clear()
            KoleksiData.Add(Me.TxtNamaHakAkses)
            KoleksiNamaData.Add("Nama hakakses")
            KoleksiData.Add(Me.TxtKodeHakAkses)
            KoleksiNamaData.Add("Kode hakakses")

            'cek data kosong
            If ModCekData.CekDataKosong(KoleksiData, KoleksiNamaData) = False Then
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

    Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
        'setting agar siap memasukkan data
        Me.baru = True
        Me.AktifkanDetail()
        Me.KosongkanField()
        Me.TxtNamaHakAkses.Focus()
    End Sub

    Private Sub cmdEdt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdt.Click
        'setting agar siap memasukan data
        If Not Me.TxtKodeHakAkses.Text = "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.TxtNamaHakAkses.Focus()
        End If
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Try
            If MsgBox("Apa Anda yakin batal batal menyimpan data ini?", MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.Yes Then
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
    Private Sub cmdKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKeluar.Click
        Me.Hide()
    End Sub
End Class