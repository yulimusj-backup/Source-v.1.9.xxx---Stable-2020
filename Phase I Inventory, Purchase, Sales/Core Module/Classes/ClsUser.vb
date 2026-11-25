Option Explicit On
Option Strict Off

Public Class ClsUser

    Private mKodeUser As String
    Private mNamaUser As String
    Private mPassword As String
    Private mKeterangan As String
    Private mKoleksiHakAkses As New Collection
    Private mKoleksiHakKhusus As New Collection
    Private mPrimaryGroup As Double = -1
    Private mUseTempPass As Integer = 0

    Private DS As DSUser
    Private DTA As DSUserTableAdapters.tbluserTableAdapter

    'ini juga ada di clstransaksi journal
    Public Enum EnumJournalAccesType As Integer
        FullAccess
        Selected_Users_Only
        Respected_User_Only
    End Enum

    'ini ada kembarannya di 'FrmSettingAP_AR_LocationAccess'
    Public Enum EnumLocationAccessType As Integer
        STOCK_IN_ADJUST '
        STOCK_IN_PURCAHSE '
        STOCK_IN_SALES_RETURN '
        STOCK_TRANSFER_IN
        STOCK_BUNDLING_AFTER
        STOCK_OUT_ADJUST
        STOCK_OUT_SALES '
        STOCK_OUT_PURCHASE_RETURN '
        STOCK_TANSFER_OUT
        STOCK_BUNDLING_BEFORE
        PRODUCTION_LOG_BEFORE
        PRODUCTION_LOG_AFTER
    End Enum

    Public Property UseTempPass() As Boolean
        Get
            If Me.mUseTempPass = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mUseTempPass = 1
            Else
                Me.mUseTempPass = 0
            End If
        End Set
    End Property

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

    Private Sub LoadDataBase()
        Try
            DS = New DSUser
            DTA = New DSUserTableAdapters.tbluserTableAdapter
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(DS.tbluser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetUserCode(ByVal vPrimaryUser As Double) As String
        Dim DTA As New DSUserTableAdapters.tbluserTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetKodeUser(vPrimaryUser)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetKodeUser(vPrimaryUser)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DSUserTableAdapters.tbluserTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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

    ''' diijoin karena error
    ''' <summary>
    ''' mengembalikan nilai true jika user memiliki hak akses atas form tertentu.
    ''' </summary>
    ''' <param name="vUserCode"></param>
    ''' <param name="vNamaHakAkses"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    'Public Function ApakahMemilikiHakAkses(ByVal vUserCode As String, ByVal vNamaHakAkses As String) As Boolean
    '    Dim DTA As New DSUserTableAdapters.DataTable1aTableAdapter
    '    Dim TBL As New DSUser.DataTable1aDataTable

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        TBL.Rows.Clear()
    '        DTA.Fill(TBL, vUserCode, vNamaHakAkses)
    '        If TBL.Rows.Count > 0 Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        TBL.Dispose()
    '    End Try
    'End Function

    Public Function GetPrimary(ByVal vKodeUser As String) As Double
        Dim DTA As New DSUserTableAdapters.tbluserTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
        Dim DTA As New DSUserTableAdapters.tblantarauserhakkhususTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
        DTA = New DSUserTableAdapters.tbluserTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
        DTA = New DSUserTableAdapters.tbluserTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
        Dim RetVal As Nullable(Of Integer)

        Try
            Me.LoadDataBase()
            DTA.Connection.ConnectionString = Core_Module.GetConstring

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
                    RetVal = DTA.ScalarQuery_GetAllowUseTempPass(KodeUser)
                    If RetVal.HasValue = True Then
                        If RetVal.Value = 1 Then
                            Dim ObjTempPass As New ClsUser_TempPass
                            If ObjTempPass.Login(KodeUser, Password) = False Then
                                MsgBox("Wrong Password.", MsgBoxStyle.Critical, "Error")
                                Return False
                            End If
                            ObjTempPass = Nothing
                        Else
                            MsgBox("Wrong Password.", MsgBoxStyle.Critical, "Error")
                            Return False
                        End If
                    Else
                        MsgBox("Wrong Password.", MsgBoxStyle.Critical, "Error")
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function Login(ByVal KodeUser As String, ByVal Password As String) As Boolean
        Dim RetVal As Nullable(Of Integer)

        Try
            Me.LoadDataBase()
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cek data yang diberikan
            'cek keberadaan user
            If DTA.ScalarQuery_CekKode(KodeUser).Value = 0 Then
                MsgBox("Unknown user name", MsgBoxStyle.Critical, "Error")
                Return False
            End If
            'cek kecocokan password dengan user
            Password = Me.EnkripsiPassword(Password)
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPassword(KodeUser)) = False Then
                If Password.Equals(DTA.ScalarQuery_GetPassword(KodeUser)) = False Then
                    RetVal = DTA.ScalarQuery_GetAllowUseTempPass(KodeUser)
                    If RetVal.HasValue = True Then
                        If RetVal.Value = 1 Then
                            Dim ObjTempPass As New ClsUser_TempPass
                            If ObjTempPass.Login(KodeUser, Password) = False Then
                                MsgBox("Wrong Password.", MsgBoxStyle.Critical, "Error")
                                Return False
                            End If
                            ObjTempPass = Nothing
                        Else
                            MsgBox("Wrong Password.", MsgBoxStyle.Critical, "Error")
                            Return False
                        End If
                    Else
                        MsgBox("Wrong Password.", MsgBoxStyle.Critical, "Error")
                        Return False
                    End If
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
        Dim RetVal As Nullable(Of Integer)

        Try
            Me.LoadDataBase()
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cek data yang diberikan
            'cek keberadaan user
            If DTA.ScalarQuery_CekKode(KodeUser).Value = 0 Then
                MsgBox("Unknown user name.", MsgBoxStyle.Critical, "Error")
                txtKodeUser.SelectAll()
                txtKodeUser.Focus()
                Return False
            End If

            'cek kecocokan password dengan user
            Password = Me.EnkripsiPassword(Password)
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPassword(KodeUser)) = False Then
                If Password.Equals(DTA.ScalarQuery_GetPassword(KodeUser)) = False Then
                    RetVal = DTA.ScalarQuery_GetAllowUseTempPass(KodeUser)
                    If RetVal.HasValue = True Then
                        If RetVal.Value = 1 Then
                            Dim ObjTempPass As New ClsUser_TempPass
                            If ObjTempPass.Login(KodeUser, Password) = False Then
                                MsgBox("Wrong Password.", MsgBoxStyle.Critical, "Error")
                                Return False
                            End If
                            ObjTempPass = Nothing
                        Else
                            MsgBox("Wrong Password.", MsgBoxStyle.Critical, "Error")
                            Return False
                        End If
                    Else
                        MsgBox("Wrong Password.", MsgBoxStyle.Critical, "Error")
                        Return False
                    End If
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
                    MsgBox("Old password is wrong, please fix.", MsgBoxStyle.Critical, "Error")
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
                MsgBox("Inconsistent password please fix.", MsgBoxStyle.Critical, "Error")
                txtpasswordbaru.Clear()
                txtKonfirmasiPassword.Clear()
                txtpasswordbaru.Focus()
                Return False
            End If

            'cek jangan sampai password baru kosong
            If txtpasswordbaru.Text = "" Then
                MsgBox("Please specify password. A good password is a word with at least 8 characters of combination of symbol, number, uppercase and lowercase letters.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek apakah password yang diberikan kuat 
            If Me.ValidatePassword(txtpasswordbaru.Text) = False Then
                If MsgBox("Your password is too weak. Stong password uses at least 8 character with combination of symbol, number, uppercase and lowercase letters; i.e: 'P@ssWerd123'. Continue save?", _
                MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTA = New DSUserTableAdapters.tbluserTableAdapter

            'cek kecocokan password baru dan konfirmasi password baru
            If txtpasswordbaru.Text.Equals(txtKonfirmasiPassword.Text) = False Then
                MsgBox("Inconsistent password.", MsgBoxStyle.Critical, "Error")
                txtpasswordbaru.Clear()
                txtKonfirmasiPassword.Clear()
                txtpasswordbaru.Focus()
                Return False
            End If

            'cek jangan sampai password baru kosong
            If txtpasswordbaru.Text = "" Then
                MsgBox("Please specify password. A good password is a word with at least 8 characters of combination of symbol, number, uppercase and lowercase letters.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek apakah password yang diberikan kuat 
            If Me.ValidatePassword(txtpasswordbaru.Text) = False Then
                If MsgBox("Your password is too weak. Stong password uses at least 8 character with combination of symbol, number, uppercase and lowercase letters; i.e: 'P@ssWerd123'. Continue save?", _
                MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                    Return False
                End If
            End If

            'simpan perubahan
            txtpasswordbaru.Text = Me.EnkripsiPassword(txtpasswordbaru.Text)
            txtKonfirmasiPassword.Text = txtpasswordbaru.Text
            DTA.UpdateQuery_SetPassword(txtpasswordbaru.Text, KodeUser)

            'konfirmasi
            MsgBox("User " + KodeUser + " password's changed.", MsgBoxStyle.Information, "Information")

            'sampe sini proses ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    'Public Function SimpanPerubahan(ByVal KodeUser As String, ByVal NamaUser As String, _
    'ByVal password As String, ByVal keterangan As String, ByVal vPrimaryGroup As String) As Boolean
    '    DTA = New DSUserTableAdapters.tbluserTableAdapter
    '    Try
    '        'simpan perubahan data ekspedisi
    '        DTA.Update(KodeUser, NamaUser, password, keterangan, vPrimaryGroup, KodeUser)
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function CariDetail(ByVal Kode_User As String, ByVal ObjUser As ClsUser) As Boolean 'DSFinishGoods.TblUserDataTable
        Dim tabel As New DSUser.tbluserDataTable
        Try
            LoadDataBase()

            'get data
            DTA.FillBy_KodeUser(tabel, Kode_User)

            'salin ke objek
            If tabel.Rows.Count > 0 Then
                ObjUser.KodeUser = tabel.Rows.Item(0).Item("Kode User")
                ObjUser.NamaUser = tabel.Rows.Item(0).Item("Nama User")
                ObjUser.Password = tabel.Rows.Item(0).Item("Password")
                ObjUser.Keterangan = tabel.Rows.Item(0).Item("Keterangan")
                ObjUser.mUseTempPass = tabel.Rows.Item(0).Item("Allow Use Temp Pass")
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            tabel.Dispose()
        End Try
    End Function

    Public Function IsiListBoxUser(ByVal lbo As CheckedListBox, ByVal TBL As DSUser.tbluser1DataTable, _
    ByVal UserCode As String) As Boolean
        Dim DTA As New DSUserTableAdapters.tbluser1TableAdapter
        Dim row As DSUser.tbluser1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            Me.DS.tbluser.Columns.RemoveAt(2)
            Me.DS.tbluser.Columns.RemoveAt(3)

            'set view
            View = DS.tbluser.DefaultView

            'set filter
            If DenganKode = True Then
                cmd = "[Kode User] LIKE '" + TextCari + "*'"
            Else
                cmd = "[Nama User] LIKE '" + TextCari + "*'"
            End If
            View.RowFilter = cmd

            'set data source
            DGV.DataSource = View
            DGV.Columns(0).HeaderText = "User Code"
            DGV.Columns(1).HeaderText = "User Name"
            'DGV.Columns(2).Visible = False
            DGV.Columns(2).HeaderText = "Remarks"

            'kembalikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function IsiComboBox_ByHakAkses(ByVal CBO As ComboBox, ByVal vNamaHakAkses As String) As Boolean
        Dim DTA As New DSUserTableAdapters.tblantarauserhakaksesTableAdapter
        Dim TBL As New DSUser.tblantarauserhakaksesDataTable
        Dim row As DSUser.tblantarauserhakaksesRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_NamaHakAkses(TBL, vNamaHakAkses)
            CBO.Items.Clear()
            For Each row In TBL.Rows
                CBO.Items.Add(row.Kode_User)
            Next
            If CBO.Items.Count > 0 Then
                CBO.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsiComboBox(ByVal CBO As ComboBox) As Boolean
        Dim row As DSUser.tbluserRow

        Try
            CBO.Items.Clear()

            LoadDataBase()

            For Each row In Me.DS.tbluser.Rows
                CBO.Items.Add(row.Kode_User)
            Next

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
        Dim DTS As New DSUserTableAdapters.tblantarauserhakkhususTableAdapter
        Dim TBL1 As New DSUser.tblantarauserhakkhususDataTable
        Dim objHakKhusus As ClsHakKhusus
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
        Dim DTS As New DSUserTableAdapters.tblantarauserhakaksesTableAdapter
        Dim TBL1 As New DSUser.tblantarauserhakaksesDataTable
        Dim row As DSUser.tblantarauserhakaksesRow

        Dim objHakAkses As ClsHakAkses

        Try
            DTS.Connection.ConnectionString = Core_Module.GetConstring
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
            MsgBox("Hak akses tersimpan ke basis data", MsgBoxStyle.Information, "Information")

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTS.Dispose()
        End Try
    End Function

    Public Function EnkripsiPassword(ByVal PasswordUntukDiEnkripsi As String) As String
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

    Public Sub ShowHakAksesLokasi(ByVal vUserCode As String)
        Dim DTAUser As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilegeTableAdapter
        Dim TBLUser As New DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilegeDataTable
        Dim rowUser As DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilegeRow
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Dummy As String = ""
        Dim FrmRef As New FrmUserLocationPrivilegeSetting2
        Dim DTALocation As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilege_itemTableAdapter
        Dim tblLocation As New DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilege_itemDataTable
        Dim rowLocation As DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilege_itemRow
        Dim rowDGV As DataGridViewRow
        Dim ObjLoc As New Core_Module.clsMasterStorageArea
        Dim ObjWarehouse As New Core_Module.clsMasterWarehouse

        Try
            'cek apakah user ybs ada di dalam daftar user
            DTALocation.Connection.ConnectionString = Core_Module.GetConstring
            DTAUser.Connection.ConnectionString = Core_Module.GetConstring
            DTAUser.FillBy_UserCode(TBLUser, vUserCode)
            If TBLUser.Rows.Count = 0 Then
                Dummy = ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Location_Privileges, _
                ClsTransaksiSetting.EnumNamaParameter.Default_User_Not_IncludeInList)
                If Dummy = "1" Then
                    'semua lokasi ga boleh
                    MsgBox("The user: " + vUserCode + " can't access any location.", MsgBoxStyle.Information, "Information")
                Else
                    'semua lokasi boleh
                    MsgBox("The user: " + vUserCode + " can access every location.", MsgBoxStyle.Information, "Information")
                End If
            Else
                rowUser = TBLUser.Rows(0)
                If rowUser.default_access = 1 Then
                    'secara default tidak boleh
                    FrmRef.Label1.Text = "User with ID: " + vUserCode + " has the following access detail. User can't access all other location not specified in the table below."
                Else
                    'secara default boleh
                    FrmRef.Label1.Text = "User with ID: " + vUserCode + " has the following access detail. User can access all other location not specified in the table below."
                End If
                DTALocation.FillBy_UserCode(tblLocation, vUserCode)
                For Each rowLocation In tblLocation.Rows
                    rowDGV = FrmRef.dgvLocation.Rows(FrmRef.dgvLocation.Rows.Add())
                    With rowDGV
                        If rowLocation.location_level = 1 Then
                            .Cells(FrmRef.colLocation_Code.Index).Value = ObjLoc.GetStorageAreaCode(rowLocation.primary_location)
                            .Cells(FrmRef.colLocation_Desc.Index).Value = ObjLoc.GetDescription(rowLocation.primary_location)
                        Else
                            .Cells(FrmRef.colLocation_Code.Index).Value = ObjWarehouse.GetWarehouseCode(rowLocation.primary_location)
                            .Cells(FrmRef.colLocation_Desc.Index).Value = ObjWarehouse.GetDescription(rowLocation.primary_location)
                        End If

                        If rowLocation.location_level = 0 Then
                            .Cells(FrmRef.colLocation_Level.Index).Value = "Warehouse"
                        Else
                            .Cells(FrmRef.colLocation_Level.Index).Value = "Location"
                        End If
                        .Cells(FrmRef.colLocation_No.Index).Value = FrmRef.dgvLocation.Rows.Count
                        .Cells(FrmRef.ColLocation_StockBundlingAfter.Index).Value = rowLocation.stock_bundling_after * -1
                        .Cells(FrmRef.ColLocation_StockBundlingBefore.Index).Value = rowLocation.stock_bundling_before * -1
                        .Cells(FrmRef.ColLocation_StockInAdjust.Index).Value = rowLocation.stock_in_adjust * -1
                        .Cells(FrmRef.colLocation_StockInPurchase.Index).Value = rowLocation.stock_in_purchase * -1
                        .Cells(FrmRef.colLocation_StockInSalesReturn.Index).Value = rowLocation.stock_in_sales_return * -1
                        .Cells(FrmRef.ColLocation_StockOutAdjust.Index).Value = rowLocation.stock_out_adjust * -1
                        .Cells(FrmRef.ColLocation_StockOutPurcahseReturn.Index).Value = rowLocation.stock_out_purchase_return * -1
                        .Cells(FrmRef.ColLocation_StockOutSales.Index).Value = rowLocation.stock_out_sales * -1
                        .Cells(FrmRef.ColLocation_StockTransferIn.Index).Value = rowLocation.stock_transfer_in * -1
                        .Cells(FrmRef.ColLocation_StockTransferOut.Index).Value = rowLocation.stock_transfer_out * -1
                        .Cells(FrmRef.colLocationPrimaryRefParty.Index).Value = rowLocation.primary_location
                        .Cells(FrmRef.ColLocation_ProductionlogBefore.Index).Value = rowLocation.production_log_before * -1
                        .Cells(FrmRef.ColLocation_ProductionLogAfter.Index).Value = rowLocation.production_log_after * -1
                    End With
                Next
                FrmRef.ShowDialog()

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahMemilikiHakAksesLokasi(ByVal vUserCode As String ) As Boolean
        Dim DTAUser As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilegeTableAdapter 
        Dim Jml As Nullable(Of Long)

        Try
            DTAUser.Connection.ConnectionString = Core_Module.GetConstring
            Jml = DTAUser.ScalarQuery_GetJml(vUserCode)
            If Jml.HasValue = False Then
                Return False
            Else
                If Jml.Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahMemilikiHakAksesLokasi(ByVal vUserCode As String, ByVal vTipeTrans As EnumLocationAccessType, _
    ByVal vTBLDetail As DseUserLocationPrivilegeSetting.tblmaster_storage_areaDataTable) As Boolean
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Dummy As String = ""
        Dim DTAUser As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilegeTableAdapter
        Dim TBLUser As New DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilegeDataTable
        Dim rowUser As DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilegeRow
        Dim DTAUserDetail As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilege_itemTableAdapter
        Dim TBLUserDetail As New DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilege_itemDataTable
        Dim rowUserDetail As DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilege_itemRow
        Dim Msg As String = ""
        Dim ColumnName1 As String = ""
        Dim rowDetail As DseUserLocationPrivilegeSetting.tblmaster_storage_areaRow
        Dim Ketemu As Boolean = False
        Dim PunyaHakAkses As Boolean = False
        Dim ObjLoc As New Core_Module.clsMasterStorageArea
        Dim counter As Integer = 0

        Try
            DTAUser.Connection.ConnectionString = Core_Module.GetConstring
            DTAUser.FillBy_UserCode(TBLUser, vUserCode)
            If TBLUser.Rows.Count = 0 Then
                Dummy = ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Location_Privileges, _
                                         ClsTransaksiSetting.EnumNamaParameter.Default_User_Not_IncludeInList)
                If Dummy = "1" Then
                    'ga boleh pake semua lokasi
                    For Each rowDetail In vTBLDetail.Rows
                        counter += 1
                        If Msg = "" Then
                            Msg = counter.ToString + ". " + rowDetail.Storage_Area_Code
                        Else
                            Msg = Msg + vbCrLf + counter.ToString + ". " + rowDetail.Storage_Area_Code
                        End If
                    Next
                    MsgBox("You don't have access to the following location: " + vbCrLf + Msg, MsgBoxStyle.Information, "Information")
                    Return False
                Else
                    'anggap boleh pake semua lokasi
                    Return True
                End If
            Else
                rowUser = TBLUser.Rows(0)
                DTAUserDetail.FillBy_UserCode(TBLUserDetail, vUserCode)

                If vTipeTrans = EnumLocationAccessType.STOCK_BUNDLING_AFTER Or _
                vTipeTrans = EnumLocationAccessType.PRODUCTION_LOG_AFTER Then
                    ColumnName1 = "stock bundling after"
                ElseIf vTipeTrans = EnumLocationAccessType.STOCK_BUNDLING_BEFORE Or _
                vTipeTrans = EnumLocationAccessType.PRODUCTION_LOG_BEFORE Then
                    ColumnName1 = "stock bundling before"
                ElseIf vTipeTrans = EnumLocationAccessType.STOCK_IN_ADJUST Then
                    ColumnName1 = "stock in adjust"
                ElseIf vTipeTrans = EnumLocationAccessType.STOCK_IN_PURCAHSE Then
                    ColumnName1 = "stock in purchase"
                ElseIf vTipeTrans = EnumLocationAccessType.STOCK_IN_SALES_RETURN Then
                    ColumnName1 = "stock in sales return"
                ElseIf vTipeTrans = EnumLocationAccessType.STOCK_OUT_ADJUST Then
                    ColumnName1 = "stock out adjust"
                ElseIf vTipeTrans = EnumLocationAccessType.STOCK_OUT_PURCHASE_RETURN Then
                    ColumnName1 = "stock out purchase return"
                ElseIf vTipeTrans = EnumLocationAccessType.STOCK_OUT_SALES Then
                    ColumnName1 = "stock out sales"
                ElseIf vTipeTrans = EnumLocationAccessType.STOCK_TANSFER_OUT Then
                    ColumnName1 = "stock transfer out"
                ElseIf vTipeTrans = EnumLocationAccessType.STOCK_TRANSFER_IN Then
                    ColumnName1 = "stock transfer in"
                End If

                For Each rowDetail In vTBLDetail.Rows
                    Ketemu = False
                    PunyaHakAkses = False
                    'jika secara primary lokasi, dibilang boleh atau tidak boleh, maka langsung ikuti
                    For Each rowUserDetail In TBLUserDetail.Rows
                        If rowUserDetail.location_level = 1 Then
                            'location
                            If rowDetail.Primary = rowUserDetail.primary_location Then
                                If rowUserDetail.Item(ColumnName1) = 1 Then
                                    PunyaHakAkses = True
                                Else
                                    PunyaHakAkses = False
                                End If
                                Ketemu = True
                                Exit For
                            End If
                        End If
                    Next
                    'jika blm ketemu, maka cek pada level warehouse
                    If Ketemu = False Then
                        For Each rowUserDetail In TBLUserDetail.Rows
                            If rowUserDetail.location_level = 0 Then
                                'location
                                If rowDetail.Primary_Warehouse = rowUserDetail.primary_location Then
                                    If rowUserDetail.Item(ColumnName1) = 1 Then
                                        PunyaHakAkses = True
                                    Else
                                        PunyaHakAkses = False
                                    End If
                                    Ketemu = True
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                    If Ketemu = False Then
                        If rowUser.default_access = 0 Then
                            'secara default boleh
                            PunyaHakAkses = True
                        Else
                            'secara default ga boleh
                            PunyaHakAkses = False
                        End If
                    End If
                    If PunyaHakAkses = False Then
                        counter += 1
                        If Msg = "" Then
                            Msg = counter.ToString + ". " + rowDetail.Storage_Area_Code
                        Else
                            Msg = Msg + vbCrLf + counter.ToString + ". " + rowDetail.Storage_Area_Code
                        End If
                    End If
                Next

                If Msg = "" Then
                    Return True
                Else
                    MsgBox("You don't have access to the following location: " + vbCrLf + Msg, MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Public Function FillLocationAccess(ByVal vUserCode As String, ByVal vTipeTrans As EnumLocationAccessType, _
    'ByVal vTBLDetail As DseUserLocationPrivilegeSetting.tblmaster_storage_areaDataTable) As Boolean
    '    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    '    Dim Dummy As String = ""
    '    Dim DTAUser As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilegeTableAdapter
    '    Dim TBLUser As New DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilegeDataTable
    '    Dim rowUser As DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilegeRow
    '    Dim DTAUserDetail As New DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilege_itemTableAdapter
    '    Dim TBLUserDetail As New DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilege_itemDataTable
    '    Dim rowUserDetail As DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilege_itemRow
    '    Dim DTADetail As New DseUserLocationPrivilegeSettingTableAdapters.tblmaster_storage_areaTableAdapter
    '    Dim ColumnName1 As String = ""

    '    Try
    '        DTAUser.Connection.ConnectionString = Core_Module.GetConstring
    '        DTADetail.Connection.ConnectionString = DTAUser.Connection.ConnectionString
    '        DTAUserDetail.Connection.ConnectionString = DTAUser.Connection.ConnectionString

    '        DTAUser.FillBy_UserCode(TBLUser, vUserCode)
    '        If TBLUser.Rows.Count = 0 Then
    '            'cek nilai default u/ user yang tidak ada dalam setting
    '            Dummy = ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Location_Privileges, _
    '                      ClsTransaksiSetting.EnumNamaParameter.Default_User_Not_IncludeInList)
    '            If Dummy = "1" Then
    '                'ga boleh pake semua lokasi
    '                vTBLDetail.Rows.Clear()
    '            Else
    '                'anggap boleh pake semua lokasi
    '                DTADetail.Fill(vTBLDetail)
    '            End If
    '        Else
    '            rowUser = TBLUser.Rows(0)
    '            DTAUserDetail.FillBy_UserCode(TBLUserDetail, vUserCode)
    '            If vTipeTrans = EnumLocationAccessType.STOCK_BUNDLING_AFTER Then
    '                ColumnName1 = "stock bundling after"
    '            ElseIf vTipeTrans = EnumLocationAccessType.STOCK_BUNDLING_BEFORE Then
    '                ColumnName1 = "stock bundling before"
    '            ElseIf vTipeTrans = EnumLocationAccessType.STOCK_IN_ADJUST Then
    '                ColumnName1 = "stock in adjust"
    '            ElseIf vTipeTrans = EnumLocationAccessType.STOCK_IN_PURCAHSE Then
    '                ColumnName1 = "stock in purchase"
    '            ElseIf vTipeTrans = EnumLocationAccessType.STOCK_IN_SALES_RETURN Then
    '                ColumnName1 = "stock in sales return"
    '            ElseIf vTipeTrans = EnumLocationAccessType.STOCK_OUT_ADJUST Then
    '                ColumnName1 = "stock out adjust"
    '            ElseIf vTipeTrans = EnumLocationAccessType.STOCK_OUT_PURCHASE_RETURN Then
    '                ColumnName1 = "stock out purchase return"
    '            ElseIf vTipeTrans = EnumLocationAccessType.STOCK_OUT_SALES Then
    '                ColumnName1 = "stock out sales"
    '            ElseIf vTipeTrans = EnumLocationAccessType.STOCK_TANSFER_OUT Then
    '                ColumnName1 = "stock transfer out"
    '            ElseIf vTipeTrans = EnumLocationAccessType.STOCK_TRANSFER_IN Then
    '                ColumnName1 = "stock transfer in"
    '            End If

    '            If rowUser.default_access = 0 Then
    '                'secara default ga boleh, maka yang tampil hanya ada di tabel item saja
    '                'load data per warehouse
    '                For Each rowUserDetail In TBLUserDetail.Rows
    '                    If rowUserDetail.location_level = 1 Then
    '                        If rowUserDetail.Item(ColumnName1) = 1 Then
    '                            'tambahkan warehouse ke dalam item detail

    '                        ElseIf rowUserDetail.Item(ColumnName1) = 0 Then
    '                            'buang warehouse

    '                        End If
    '                    End If
    '                Next

    '                'load data per lokasi
    '                For Each rowUserDetail In TBLUserDetail.Rows
    '                    If rowUserDetail.location_level = 0 Then
    '                        If rowUserDetail.Item(ColumnName1) = 1 Then
    '                            'tambahkan lokasi ke dalam item detail

    '                        ElseIf rowUserDetail.Item(ColumnName1) = 0 Then
    '                            'buang lokasi

    '                        End If
    '                    End If
    '                Next
    '            Else
    '                'secara default boleh, maka yang tampil adalah semua, dikurangi yang ga boleh di tabel item

    '            End If
    '        End If

    '        '1. cek apakah ada di tabel user
    '        'jika ada, maka fill tabel lokasi yang diberikan
    '        'jika tidak ada, maka lihar berdasarkan nilai default

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Public Function ApakahMemilikiAksesLokasi(ByVal vUserCode As String, ByVal vListPrimaryLocation As List(Of Double), _
    'ByVal vShowMsg As Boolean, ByVal vTipeTrans As EnumLocationAccessType, ByVal vTBLUser As DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilegeDataTable, _
    'ByVal vTBLDetail As DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilege_itemDataTable) As Boolean
    '    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    '    Dim DefaultBolehPakeLokasi As Boolean = True
    '    Dim Dummy As String = ""
    '    Dim DefaultUserBolehPakeLokasi As Boolean = True

    '    Try
    '        Dummy = ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Location_Privileges, _
    '        ClsTransaksiSetting.EnumNamaParameter.Default_User_Not_IncludeInList)
    '        If Dummy = "1" Then
    '            'anggap ga boleh pake
    '            DefaultBolehPakeLokasi = False
    '        Else
    '            'anggap boleh pake
    '            DefaultBolehPakeLokasi = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Public Function ApakahBolehMemilikiTempPass(ByVal vPrimaryUser As Double) As Boolean
        Dim DTA As New DSUserTableAdapters.tbluser2TableAdapter
        Dim RetVal As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetUseTempPass(vPrimaryUser)
            If RetVal.HasValue = True Then
                If RetVal.Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return False
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini mengembalikan nilai true jika user ybs boleh akses form tsb
    ''' </summary>
    ''' <param name="KodeUser"></param>
    ''' <param name="NamaForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ApakahMemilikiHakAkses(ByVal KodeUser As String, ByVal NamaForm As String) As Boolean
        Dim DTA As New DSUserTableAdapters.tblantarauserhakaksesTableAdapter
        Dim obj As New ClsHakAkses

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(KodeUser, obj.GetKode(NamaForm)).HasValue Then
                If DTA.ScalarQuery_GetJml(KodeUser, obj.GetKode(NamaForm)).Value > 0 Then
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

    Public Function Simpan(ByVal txtKodeUser As TextBox, ByVal txtNamaUser As TextBox, _
        ByVal txtpassword As TextBox, ByVal keterangan As String, ByVal PrimaryGroup As Double) As Boolean
        DTA = New DSUserTableAdapters.tbluserTableAdapter
        Dim dta2 As New DSUserTableAdapters.tbluser2TableAdapter
        Dim ObjUG As New Core_Module.ClsMasterUserGroup
        Dim vTBLUser As New DseMasterUserGroup.Tbl_JmlUserDataTable
        Dim rowUser As DseMasterUserGroup.Tbl_JmlUserRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            dta2.Connection.ConnectionString = DTA.Connection.ConnectionString
            'cek kode
            If DTA.ScalarQuery_CekKode(txtKodeUser.Text).HasValue = True Then
                If DTA.ScalarQuery_CekKode(txtKodeUser.Text).Value > 0 Then
                    MsgBox("User code duplication, please change user code.", MsgBoxStyle.Information, "Information")
                    txtKodeUser.SelectAll()
                    txtKodeUser.Focus()
                    Return False
                End If
            End If

            'simpan
            Me.Password = Me.EnkripsiPassword(Me.Password)
            dta2.Insert(KodeUser, NamaUser, Password, keterangan, PrimaryGroup, Me.mUseTempPass)

            ''Simpan Tabel Antara Hak Akses
            rowUser = vTBLUser.NewTbl_JmlUserRow
            rowUser.KodeUser = KodeUser
            rowUser.Primary = PrimaryGroup
            vTBLUser.Rows.Add(rowUser)
            ObjUG.SimpanPerubahanTBLAntara(vTBLUser, PrimaryGroup)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            dta2.Dispose()
            ObjUG = Nothing
            vTBLUser = Nothing
        End Try
    End Function

    'Public Function SimpanTBLAntara(ByVal VKodeUser As String, ByVal VPrimaryGroup As Double) As Boolean
    '    Dim DTA As New DSUserTableAdapters.TblPerbandinganTableAdapter
    '    Dim DTA1 As New DSUserTableAdapters.tblantarauserhakakses1TableAdapter
    '    Dim row As DSUser.TblPerbandinganRow
    '    Dim DS As New DSUser
    '    Dim row1 As DSUser.Tbl_DataGroupRow
    '    Try
    '        'Bersihkan
    '        DS.Tbl_DataGroup.Clear()

    '        'Kumpulkan Data
    '        For Each row In DTA.GetDataBy_PrimaryGroup(VPrimaryGroup)
    '            row1 = DS.Tbl_DataGroup.NewTbl_DataGroupRow
    '            row1.PrimaryGroup = row.Primary
    '            row1.KodeHakAkses = row.Kode_Hak_Akses
    '            row1.Check = row.Checked
    '            DS.Tbl_DataGroup.Rows.Add(row1)
    '        Next

    '        'hapus data antara untuk user ybs
    '        DTA1.Delete(VKodeUser)

    '        'Simpan Table Antara
    '        For Each row1 In DS.Tbl_DataGroup
    '            DTA1.Insert(VKodeUser, row1.KodeHakAkses)
    '        Next

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return False
    '    Finally
    '        DS.Dispose()
    '        DTA.Dispose()
    '        DTA1.Dispose()
    '    End Try
    'End Function

    Public Function SimpanPerubahan(ByVal KodeUser As String, ByVal NamaUser As String, _
    ByVal password As String, ByVal keterangan As String, ByVal PrimaryGroup As Double, ByVal vTempPass As Integer) As Boolean
        DTA = New DSUserTableAdapters.tbluserTableAdapter
        Dim DTA2 As New DSUserTableAdapters.tbluser2TableAdapter
        Dim vTBLUser As New DseMasterUserGroup.Tbl_JmlUserDataTable
        Dim rowUser As DseMasterUserGroup.Tbl_JmlUserRow
        Dim ObjUG As New Core_Module.ClsMasterUserGroup

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString

            ''hapus TBLAntaraHakAkses
            Me.HapusTBLAntara(KodeUser)

            'simpan perubahan data ekspedisi
            DTA2.Update(KodeUser, NamaUser, password, keterangan, PrimaryGroup, vTempPass, KodeUser)

            ''Simpan Tabel Antara Hak Akses
            rowUser = vTBLUser.NewTbl_JmlUserRow
            rowUser.KodeUser = KodeUser
            rowUser.Primary = PrimaryGroup
            vTBLUser.Rows.Add(rowUser)
            ObjUG.SimpanPerubahanTBLAntara(vTBLUser, PrimaryGroup)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA2.Dispose()
            ObjUG = Nothing
            vTBLUser.Dispose()

        End Try
    End Function

    Public Function Hapus(ByVal kode_user As String) As Boolean
        DTA = New DSUserTableAdapters.tbluserTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            ''Delete TblAntaraHakAkses
            Me.HapusTBLAntara(kode_user)

            ''Delete
            DTA.Delete(kode_user)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function HapusTBLAntara(ByVal VKodeUser As String) As Boolean
        Dim DTA As New DSUserTableAdapters.tblantarauserhakakses1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(VKodeUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryGroup(ByVal kodeuser As String) As Double
        Dim dta As New DSUserTableAdapters.tbluser2TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetPrimaryGroupCode(kodeuser).HasValue = True Then
                Return dta.ScalarQuery_GetPrimaryGroupCode(kodeuser).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetJmlByNamaHakAkses(ByVal vNamaHakAkses As String) As Double
        Dim DTA As New DSUserTableAdapters.tblantarauserhakaksesTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByNamaHakAkses(vNamaHakAkses).HasValue Then
                Return DTA.ScalarQuery_GetJmlByNamaHakAkses(vNamaHakAkses).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByKodeUser(ByVal KodeUser As String) As Double
        Dim DTA As New DSUserTableAdapters.tblantarauserhakaksesTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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

    Public Function GetLedgerAccess(ByVal vKodeUser As String) As EnumJournalAccesType
        Dim DTA As New DSUserTableAdapters.tbl_ledger_accessTableAdapter
        Dim TBL As New DSUser.tbl_ledger_accessDataTable
        Dim row As DSUser.tbl_ledger_accessRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_kodeUser(TBL, vKodeUser)
            If TBL.Rows.Count = 0 Then
                Return EnumJournalAccesType.Respected_User_Only
            Else
                row = TBL.Rows(0)
                If row.tipe_access = CInt(EnumJournalAccesType.FullAccess) Then
                    Return EnumJournalAccesType.FullAccess
                Else
                    Return EnumJournalAccesType.Selected_Users_Only
                End If
            End If

            Return EnumJournalAccesType.Respected_User_Only
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumJournalAccesType.Respected_User_Only
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function
End Class
