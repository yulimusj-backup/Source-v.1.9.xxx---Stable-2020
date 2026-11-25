Public Class ClsUser

    Private mKodeUser As String
    Private mNamaUser As String
    Private mPassword As String
    Private mKeterangan As String
    Private mKoleksiHakAkses As New Collection
    Private mKoleksiHakKhusus As New Collection
    Private mPrimaryGroup As Double = -1

    Private DS As DSUser
    Private DTA As DSUserTableAdapters.TblUserTableAdapter

    Public Property PrimaryGroup() As Double
        Get
            Return mPrimaryGroup
        End Get
        Set(ByVal value As Double)
            mPrimaryGroup = value
        End Set
    End Property

    Public Property KodeUser() As String
        Get
            Return mKodeUser
        End Get
        Set(ByVal value As String)
            mKodeUser = value
        End Set
    End Property

    Public Property NamaUser() As String
        Get
            Return mNamaUser
        End Get
        Set(ByVal value As String)
            mNamaUser = value
        End Set
    End Property

    Public Property Keterangan() As String
        Get
            Return mKeterangan
        End Get
        Set(ByVal value As String)
            mKeterangan = value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return mPassword
        End Get
        Set(ByVal value As String)
            mPassword = value
        End Set
    End Property
    Public Property KoleksiHakAkses() As Collection
        Get
            Return mKoleksiHakAkses
        End Get
        Set(ByVal value As Collection)
            mKoleksiHakAkses = value
        End Set
    End Property
    Public Property KoleksiHakKhusus() As Collection
        Get
            Return mKoleksiHakKhusus
        End Get
        Set(ByVal value As Collection)
            mKoleksiHakKhusus = value
        End Set
    End Property

    ''' <summary>
    ''' mengembalikan nilai true jika user memiliki hak akses atas form tertentu.
    ''' </summary>
    ''' <param name="vUserCode"></param>
    ''' <param name="vNamaHakAkses"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ApakahPunyaHakAkses(ByVal vUserCode As String, ByVal vNamaHakAkses As String) As Boolean
        Dim DTA As New DSUserTableAdapters.DataTable1TableAdapter
        Dim TBL As New DSUser.DataTable1DataTable

        Try
            TBL.Rows.Clear()
            DTA.Fill(TBL, vUserCode, vNamaHakAkses)
            If TBL.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Private Sub LoadDataBase()
        Try
            DS = New DSUser
            DTA = New DSUserTableAdapters.TblUserTableAdapter
            DTA.Fill(DS.TblUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' Fungsi ini mengembalikan nilai true jika user yang diberikan kodenya memiliki hak akses yang diberikan namanya
    ''' </summary>
    ''' <param name="KodeUser">Kode user yang dipertanyakan hak aksesnya</param>
    ''' <param name="NamaHakAkses">Nama hak akses yang dipertanyakan</param>
    Public Function GetHakAkses(ByVal KodeUser As String, ByVal NamaHakAkses As String) As Boolean

    End Function

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DSUserTableAdapters.tbluserTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByKodeUser(vCode).Value > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vKodeUser As String) As Double
        Dim DTA As New DSUserTableAdapters.tbluserTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimary(vKodeUser).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vKodeUser).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini mengembalikan nilai true jika user yang diberikan kodenya memiliki hak khusus yang diberikan namanya
    ''' </summary>
    ''' <param name="KodeUser">Kode user yang dipertanyakan hak khususnya</param>
    ''' <param name="NamaHakKhusus">Nama hak khusus yang dipertanyakan</param>
    Public Function GetHakKhusus(ByVal KodeUser As String, ByVal NamaHakKhusus As String) As Boolean
        Dim DTA As New DSUserTableAdapters.TblAntaraUserHakKhususTableAdapter

        Try
            If DTA.ScalarQuery_HitungHakKhususPerUser(NamaHakKhusus, KodeUser).HasValue Then
                If DTA.ScalarQuery_HitungHakKhususPerUser(NamaHakKhusus, KodeUser).Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetNamaByPrimary(ByVal vPrimaryUser As Double) As String
        DTA = New DSUserTableAdapters.tbluserTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNamaByPrimary(vPrimaryUser)) = False Then
                Return DTA.ScalarQuery_GetNamaByPrimary(vPrimaryUser)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNama(ByVal KodeUser As String) As String
        DTA = New DSUserTableAdapters.TblUserTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNama(KodeUser)) = False Then
                Return DTA.ScalarQuery_GetNama(KodeUser)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPassword(ByVal KodeUser As String) As String
        DTA = New DSUserTableAdapters.TblUserTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPassword(KodeUser)) = False Then
                Return DTA.ScalarQuery_GetPassword(KodeUser)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Login(ByVal KodeUser As String, ByVal Password As String, ByVal txtkodeUser As TextBox, _
    ByVal txtPassword As TextBox) As Boolean
        Try
            LoadDataBase()

            'cek data yang diberikan
            'cek keberadaan user
            If DTA.ScalarQuery_CekKode(KodeUser).Value = 0 Then
                MsgBox("Unknown user name", MsgBoxStyle.Critical, "Error")
                txtkodeUser.SelectAll()
                txtkodeUser.Focus()
                Return False
            End If
            'cek kecocokan password dengan user
            Password = Me.EnkripsiPassword(Password)
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPassword(KodeUser)) = False Then
                If Password.Equals(DTA.ScalarQuery_GetPassword(KodeUser)) = False Then
                    MsgBox("Wrong Password.", MsgBoxStyle.Critical, "Error")
                    txtPassword.SelectAll()
                    txtPassword.Focus()
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function Login(ByVal txtKodeUser As TextBox, ByVal txtPassword As TextBox, _
    ByVal txtKonfirmasiPassword As TextBox) As Boolean
        Try
            LoadDataBase()

            'cek data yang diberikan
            'cek keberadaan user
            If DTA.ScalarQuery_CekKode(KodeUser).Value = 0 Then
                MsgBox("User Anda berikan tidak ada dalam database", MsgBoxStyle.Critical, "Error")
                txtKodeUser.SelectAll()
                txtKodeUser.Focus()
                Return False
            End If
            'cek kecocokan password dengan user
            Password = Me.EnkripsiPassword(Password)
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPassword(KodeUser)) = False Then
                If Password.Equals(DTA.ScalarQuery_GetPassword(KodeUser)) = False Then
                    MsgBox("Password yang Anda berikan salah.", MsgBoxStyle.Critical, "Error")
                    txtPassword.SelectAll()
                    txtPassword.Focus()
                    Return False
                End If
            End If

            'sampe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini mengubah password user yang diberikan kode user dan password lamanya
    ''' </summary>
    Public Function SetPassword(ByVal KodeUser As String, ByVal txtPasswordLama As TextBox, _
    ByVal txtpasswordbaru As TextBox, ByVal txtKonfirmasiPassword As TextBox) As Boolean
        Dim i As String = ""

        Try
            LoadDataBase()

            'cek kecocokan password dengan user
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPassword(KodeUser)) = False Then
                i = txtPasswordLama.Text
                txtPasswordLama.Text = Me.EnkripsiPassword(txtPasswordLama.Text)
                If txtPasswordLama.Text.Equals(DTA.ScalarQuery_GetPassword(KodeUser)) = False Then
                    MsgBox("Password lama yang Anda berikan salah. Perbaiki data password lama.", MsgBoxStyle.Critical, "Error")
                    txtPasswordLama.SelectAll()
                    txtPasswordLama.Focus()
                    txtPasswordLama.Clear()
                    Return False
                Else
                    'kembalikan tulisan password
                    txtPasswordLama.Text = i
                End If
            End If

            'cek kecocokan password baru dan konfirmasi password baru
            If txtpasswordbaru.Text.Equals(txtKonfirmasiPassword.Text) = False Then
                MsgBox("Password yang Anda berikan tidak konsisten. Perbaiki password.", MsgBoxStyle.Critical, "Error")
                txtpasswordbaru.Clear()
                txtKonfirmasiPassword.Clear()
                txtpasswordbaru.Focus()
                Return False
            End If

            'cek jangan sampai password baru kosong
            If txtpasswordbaru.Text = "" Then
                MsgBox("Password tidak boleh kosong", MsgBoxStyle.Information, "Informasi")
                Return False
            End If

            'cek apakah password yang diberikan kuat 
            If Me.ValidatePassword(txtpasswordbaru.Text) = False Then
                If MsgBox("Password yang Anda berikan kurang kompleks. Gunakan password minimal 8 karakter dengan minimal 2 karakter angka, 2 karakter dengan huruf besar, 2 karakter dengan huruf kecil, dan 2 karakter non huruf; contoh: 'P@ssWerd123'. Lanjutkan proses simpan?", _
                MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.No Then
                    Return False
                End If
            End If

            'simpan perubahan
            txtpasswordbaru.Text = Me.EnkripsiPassword(txtpasswordbaru.Text)
            DTA.UpdateQuery_SetPassword(txtpasswordbaru.Text, KodeUser)

            'sampe sini proses ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Function ValidatePassword(ByVal pwd As String, _
    Optional ByVal minLength As Integer = 8, _
    Optional ByVal numUpper As Integer = 2, _
    Optional ByVal numLower As Integer = 2, _
    Optional ByVal numNumbers As Integer = 2, _
    Optional ByVal numSpecial As Integer = 2) _
    As Boolean

        ' Replace [A-Z] with \p{Lu}, to allow for Unicode uppercase letters.
        Dim upper As New System.Text.RegularExpressions.Regex("[A-Z]")
        Dim lower As New System.Text.RegularExpressions.Regex("[a-z]")
        Dim number As New System.Text.RegularExpressions.Regex("[0-9]")
        ' Special is "none of the above".
        Dim special As New System.Text.RegularExpressions.Regex("[^a-zA-Z0-9]")

        ' Check the length.
        If Len(pwd) < minLength Then Return False
        ' Check for minimum number of occurrences.
        If upper.Matches(pwd).Count < numUpper Then Return False
        If lower.Matches(pwd).Count < numLower Then Return False
        If number.Matches(pwd).Count < numNumbers Then Return False
        If special.Matches(pwd).Count < numSpecial Then Return False

        ' Passed all checks.
        Return True
    End Function

    ''' <summary>
    ''' Fungsi ini mengubah password user yang diberikan kodenya
    ''' </summary>
    ''' <param name="KodeUser">Kode user yang akan diubah passwordnya</param>
    ''' <param name="txtpasswordbaru">Textbox yang berisikan password baru</param>
    ''' <param name="txtKonfirmasiPassword">Textbox yang berisikan konfimasri password baru</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetPassword(ByVal KodeUser As String, _
    ByVal txtpasswordbaru As TextBox, ByVal txtKonfirmasiPassword As TextBox) As Boolean
        Try
            Me.DTA = New DSUserTableAdapters.TblUserTableAdapter

            'cek kecocokan password baru dan konfirmasi password baru
            If txtpasswordbaru.Text.Equals(txtKonfirmasiPassword.Text) = False Then
                MsgBox("Password yang Anda berikan tidak konsisten. Perbaiki password.", MsgBoxStyle.Critical, "Error")
                txtpasswordbaru.Clear()
                txtKonfirmasiPassword.Clear()
                txtpasswordbaru.Focus()
                Return False
            End If

            'cek jangan sampai password baru kosong
            If txtpasswordbaru.Text = "" Then
                MsgBox("Password tidak boleh kosong", MsgBoxStyle.Information, "Informasi")
                Return False
            End If

            'cek apakah password yang diberikan kuat 
            If Me.ValidatePassword(txtpasswordbaru.Text) = False Then
                If MsgBox("Password yang Anda berikan kurang kompleks. Gunakan password minimal 8 karakter dengan minimal 2 karakter angka, 2 karakter dengan huruf besar, 2 karakter dengan huruf kecil, dan 2 karakter non huruf; contoh: 'P@ssWerd123'. Lanjutkan proses simpan?", _
                MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.No Then
                    Return False
                End If
            End If

            'simpan perubahan
            txtpasswordbaru.Text = Me.EnkripsiPassword(txtpasswordbaru.Text)
            txtKonfirmasiPassword.Text = txtpasswordbaru.Text
            DTA.UpdateQuery_SetPassword(txtpasswordbaru.Text, KodeUser)

            'konfirmasi
            MsgBox("Data password untuk user: " + KodeUser + " sudah diubah.", MsgBoxStyle.Information, "Informasi")

            'sampe sini proses ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function Simpan(ByVal txtKodeUser As TextBox, ByVal txtNamaUser As TextBox, _
    ByVal txtpassword As TextBox, ByVal keterangan As String) As Boolean
        DTA = New DSUserTableAdapters.TblUserTableAdapter
        Try
            'cek kode
            If DTA.ScalarQuery_CekKode(txtKodeUser.Text).HasValue = True Then
                If DTA.ScalarQuery_CekKode(txtKodeUser.Text).Value > 0 Then
                    MsgBox("Sudah ada Customer dengan kode: " + txtKodeUser.Text + "." + vbCrLf + _
                    "Ganti kode Customer.", MsgBoxStyle.Critical, "Error")
                    txtKodeUser.SelectAll()
                    txtKodeUser.Focus()
                    Return False
                End If
            End If
            'simpan
            Me.Password = Me.EnkripsiPassword(Me.Password)
            DTA.Insert(Me.KodeUser, Me.NamaUser, Me.Password, Me.Keterangan, Me.PrimaryGroup)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Simpan(ByVal txtKodeUser As TextBox, ByVal txtNamaUser As TextBox, _
      ByVal txtpassword As TextBox, ByVal keterangan As String, ByVal PrimaryGroup As Double) As Boolean
        DTA = New DSUserTableAdapters.tbluserTableAdapter
        Dim dta2 As New DSUserTableAdapters.tbluser2TableAdapter
        Try

            'cek kode
            If DTA.ScalarQuery_CekKode(txtKodeUser.Text).HasValue = True Then
                If DTA.ScalarQuery_CekKode(txtKodeUser.Text).Value > 0 Then
                    MsgBox("Sudah ada user dengan kode: " + txtKodeUser.Text + "." + vbCrLf + _
                    "Ganti kode user.", MsgBoxStyle.Critical, "Error")
                    txtKodeUser.SelectAll()
                    txtKodeUser.Focus()
                    Return False
                End If
            End If

            'simpan
            Me.Password = Me.EnkripsiPassword(Me.Password)
            dta2.Insert(KodeUser, NamaUser, Password, keterangan, PrimaryGroup)

            ''Simpan Tabel Antara Hak Akses
            Me.SimpanTBLAntara(KodeUser, PrimaryGroup)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            dta2.Dispose()
        End Try
    End Function

    Public Function SimpanTBLAntara(ByVal VKodeUser As String, ByVal VPrimaryGroup As Double) As Boolean
        Dim DTA As New DSUserTableAdapters.TblPerbandinganTableAdapter
        Dim DTA1 As New DSUserTableAdapters.tblantarauserhakakses1TableAdapter
        Dim row As DSUser.TblPerbandinganRow
        Dim DS As New DSUser
        Dim row1 As DSUser.Tbl_DataGroupRow
        Try
            'Bersihkan
            DS.Tbl_DataGroup.Clear()

            'Kumpulkan Data
            For Each row In DTA.GetDataBy_PrimaryGroup(VPrimaryGroup)
                row1 = DS.Tbl_DataGroup.NewTbl_DataGroupRow
                row1.PrimaryGroup = row.Primary
                row1.KodeHakAkses = row.Kode_Hak_Akses
                row1.Check = row.Checked
                DS.Tbl_DataGroup.Rows.Add(row1)
            Next

            'hapus data antara untuk user ybs
            DTA1.Delete(VKodeUser)

            'Simpan Table Antara
            For Each row1 In DS.Tbl_DataGroup
                DTA1.Insert(VKodeUser, row1.KodeHakAkses)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DS.Dispose()
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal KodeUser As String, ByVal NamaUser As String, _
   ByVal password As String, ByVal keterangan As String, ByVal PrimaryGroup As Double) As Boolean
        DTA = New DSUserTableAdapters.tbluserTableAdapter
        Dim DTA2 As New DSUserTableAdapters.tbluser2TableAdapter
        Try

            ''hapus TBLAntaraHakAkses
            Me.HapusTBLAntara(KodeUser)

            'simpan perubahan data ekspedisi
            DTA2.Update(KodeUser, NamaUser, password, keterangan, PrimaryGroup, KodeUser)

            ''Simpan Tabel Antara Hak Akses
            Me.SimpanTBLAntara(KodeUser, PrimaryGroup)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA2.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal KodeUser As String, ByVal NamaUser As String, _
    ByVal password As String, ByVal keterangan As String) As Boolean
        DTA = New DSUserTableAdapters.TblUserTableAdapter
        Try
            'simpan perubahan data ekspedisi
            DTA.Update(Me.KodeUser, Me.NamaUser, Me.Password, Me.Keterangan, Me.KodeUser)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal kode_user As String) As Boolean
        DTA = New DSUserTableAdapters.tbluserTableAdapter
        Try
            DTA.Delete(kode_user)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Function CariDetail(ByVal Kode_User As String, ByVal ObjUser As ClsUser) As Boolean 'DSFinishGoods.TblUserDataTable
        Dim tabel As DSUser.TblUserDataTable
        Try
            LoadDataBase()

            'get data
            tabel = DTA.GetDataBy5_KodeUser(Kode_User)

            'salin ke objek
            If tabel.Rows.Count > 0 Then
                ObjUser.KodeUser = tabel.Rows.Item(0).Item("Kode User")
                ObjUser.NamaUser = tabel.Rows.Item(0).Item("Nama User")
                ObjUser.Password = tabel.Rows.Item(0).Item("Password")
                ObjUser.Keterangan = tabel.Rows.Item(0).Item("Keterangan")
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function IsiListBoxUser(ByVal lbo As CheckedListBox, ByVal TBL As DSUser.tbluser1DataTable, _
    ByVal UserCode As String) As Boolean
        Dim DTA As New DSUserTableAdapters.tbluser1TableAdapter
        Dim row As DSUser.tbluser1Row

        Try
            DTA.Fill(TBL)
            For Each row In TBL.Rows
                If row.Kode_User = UserCode Then
                    TBL.Rows.Remove(row)
                    Exit For
                End If
            Next
            lbo.DataSource = TBL.DefaultView
            For Each row In TBL.Rows
                row.Kombinasi = row.Kode_User + " - " + row.Nama_User
            Next
            lbo.DisplayMember = "Kombinasi"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal DGV As DataGridView, ByVal View As DataView, ByVal TextCari As String, _
    Optional ByVal DenganKode As Boolean = True) As Boolean
        Dim cmd As String
        Try
            LoadDataBase()


            'hilangkan kolom password
            Me.DS.TblUser.Columns.RemoveAt(2)

            'set view
            View = DS.TblUser.DefaultView

            'set filter
            If DenganKode = True Then
                cmd = "[Kode User] LIKE '" + TextCari + "*'"
            Else
                cmd = "[Nama User] LIKE '" + TextCari + "*'"
            End If
            View.RowFilter = cmd

            'set data source
            DGV.DataSource = View

            'kembalikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function IsiComboBox(ByVal CBO As ComboBox) As Boolean
        Dim row As DSUser.TblUserRow

        Try
            CBO.Items.Clear()

            LoadDataBase()

            For Each row In Me.DS.TblUser.Rows
                CBO.Items.Add(row.Kode_User)
            Next

            CBO.SelectedIndex = 0
            'kembalikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini menyimpan hak khusus untuk user tertentu
    ''' </summary>
    Public Function SimpanHakKhusus(ByVal KodeUser As String, ByVal KoleksiHakKhusus As Collection) As Boolean
        Dim DTS As New DSUserTableAdapters.TblAntaraUserHakKhususTableAdapter
        Dim TBL1 As New DSUser.TblAntaraUserHakKhususDataTable
        Dim objHakKhusus As ClsHakKhusus
        Try
            'hapus hak akses untuk user ybs
            DTS.Delete(KodeUser)
            'simpan
            For Each objHakKhusus In KoleksiHakKhusus
                DTS.Insert(KodeUser, objHakKhusus.KodehakKhusus)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTS.Dispose()
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Fungsi ini menyimpan hak akses untuk user tertentu
    ''' </summary>
    Public Function SimpanHakAkses(ByVal KodeUser As String, ByVal KoleksiHakAkses As Collection) As Boolean
        Dim DTS As New DSUserTableAdapters.TblAntaraUserHakAksesTableAdapter
        Dim TBL1 As New DSUser.TblAntaraUserHakAksesDataTable
        Dim row As DSUser.TblAntaraUserHakAksesRow

        Dim objHakAkses As ClsHakAkses

        Try
            'hapus hak akses untuk user ybs
            DTS.FillBy_KodeUser(TBL1, KodeUser)
            For Each row In TBL1.Rows
                row.Delete()
            Next
            DTS.Update(TBL1)

            'simpan
            For Each objHakAkses In KoleksiHakAkses
                DTS.Insert(KodeUser, objHakAkses.GetKode1(objHakAkses.Description))
            Next

            'konfirmasi
            MsgBox("Hak akses tersimpan ke basis data", MsgBoxStyle.Information, "Informasi")

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTS.Dispose()
        End Try
    End Function

    Private Function EnkripsiPassword(ByVal PasswordUntukDiEnkripsi As String) As String
        Try
            Dim SHA1 As System.Security.Cryptography.SHA1CryptoServiceProvider
            Dim bytValue() As Byte
            Dim bytHash() As Byte

            ' Create New Crypto Service Provider Object
            SHA1 = New System.Security.Cryptography.SHA1CryptoServiceProvider

            ' Convert the original string to array of Bytes
            bytValue = _
             System.Text.Encoding.UTF8.GetBytes(PasswordUntukDiEnkripsi)

            ' Compute the Hash, returns an array of Bytes
            bytHash = SHA1.ComputeHash(bytValue)

            SHA1.Clear()

            ' Return a base 64 encoded string of the Hash value
            Return Convert.ToBase64String(bytHash)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini mengembalikan nilai true jika user ybs boleh akses form tsb
    ''' </summary>
    ''' <param name="KodeUser"></param>
    ''' <param name="NamaForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function HakAksesForm(ByVal KodeUser As String, ByVal NamaForm As String) As Boolean
        Dim DTA As New DSUserTableAdapters.TblAntaraUserHakAksesTableAdapter
        Dim obj As New ClsHakAkses

        Try
            If DTA.ScalarQuery(KodeUser, obj.GetKode(NamaForm)).HasValue Then
                If DTA.ScalarQuery(KodeUser, obj.GetKode(NamaForm)).Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryGroup(ByVal kodeuser As String) As Double
        Dim dta As New DSUserTableAdapters.tbluser2TableAdapter
        Try
            If dta.ScalarQuery_GetPrimaryGroupCode(kodeuser).HasValue = True Then
                Return dta.ScalarQuery_GetPrimaryGroupCode(kodeuser).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJmlByKodeUser(ByVal KodeUser As String) As Double
        Dim DTA As New DSUserTableAdapters.tblantarauserhakaksesTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByKodeUser(KodeUser).HasValue Then
                Return DTA.ScalarQuery_GetJmlByKodeUser(KodeUser).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetKodeUser(ByVal vPrimary As Double) As String
        Dim DTA As New DSUserTableAdapters.tbluserTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetKodeUser(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetKodeUser(vPrimary)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function HapusTBLAntara(ByVal VKodeUser As String) As Boolean
        Dim DTA As New DSUserTableAdapters.tblantarauserhakakses1TableAdapter
        Try
            DTA.Delete(VKodeUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
