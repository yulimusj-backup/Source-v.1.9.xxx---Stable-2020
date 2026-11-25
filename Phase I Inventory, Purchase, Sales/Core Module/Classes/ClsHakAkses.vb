Option Explicit On
Option Strict Off

Public Class ClsHakAkses
    Private mKodehakakses As String
    Private mNamahakakses As String
    Private mDescription As String

    Public Property KodeHakAkses() As String
        Get
            Return mKodehakakses
        End Get
        Set(ByVal value As String)
            mKodehakakses = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return mDescription
        End Get
        Set(ByVal value As String)
            mDescription = value
        End Set
    End Property

    Public Property NamaHakAkses() As String
        Get
            Return mNamahakakses
        End Get
        Set(ByVal value As String)
            mNamahakakses = value
        End Set
    End Property

    ''' <summary>
    ''' Fungsi ini menyimpan data master ke basis data
    ''' </summary>
    Public Function Simpan(ByVal txtKodeHak As TextBox) As Boolean
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cek apakah ada yang sama
            If DTA.ScalarQuery_CekNama(Me.NamaHakAkses).HasValue = False Then
                DTA.Insert(Me.NamaHakAkses, Me.Description, "", 0)
            Else
                If DTA.ScalarQuery_CekNama(Me.NamaHakAkses).Value = 0 Then
                    DTA.Insert(Me.NamaHakAkses, Me.Description, "", 0)
                Else
                    MsgBox("Sudah ada hak akses dengan nama: " + Me.NamaHakAkses + " proses simpan data dihentikan.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            txtKodeHak.Text = Me.GetKode(Me.NamaHakAkses)

            MsgBox("Data tersimpan dibasis data.", MsgBoxStyle.Information, "Information")
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Simpan(ByVal vNama As String, ByVal vDesc As String, Optional ByVal vNamaAcuan As String = "", _
    Optional ByVal vDefaultValue_Turunan As Integer = 0) As Boolean
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cek apakah ada yang sama
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_CekNama(vNama).HasValue = False Then
                DTA.Insert(vNama, vDesc, vNamaAcuan, vDefaultValue_Turunan)
            Else
                If DTA.ScalarQuery_CekNama(vNama).Value = 0 Then
                    DTA.Insert(vNama, vDesc, vNamaAcuan, vDefaultValue_Turunan)
                Else
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub SetConnectionString(ByVal DTA As DSUserTableAdapters.tblmasterhakaksesTableAdapter)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahSudahAda(ByVal vKode As String) As Boolean
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_CekNama(vKode).HasValue = True Then
                If DTA.ScalarQuery_CekNama(vKode).Value > 0 Then
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

    Public Function ApakahSudahAda(ByVal vKode As String, ByVal DTA As DSUserTableAdapters.tblmasterhakaksesTableAdapter) As Boolean
        Dim Retval As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_CekNama(vKode)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
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

    ''' <summary>
    ''' Fungsi ini mencari detail data untuk hak khusus tertentu
    ''' </summary>
    Public Function CariDetail() As Boolean
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter
        Dim TBL As New DSUser.tblmasterhakaksesDataTable
        Dim row As DSUser.tblmasterhakaksesRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fill basis data
            DTA.FillBy_KodeHakAkses(TBL, Me.KodeHakAkses)

            'salin data
            If TBL.Rows.Count > 0 Then
                row = TBL.Item(0)
                With row
                    Me.KodeHakAkses = .Kode
                    Me.NamaHakAkses = .Nama
                    Me.Description = .Description
                End With
            Else
                Me.KodeHakAkses = ""
                Me.Description = ""
                Me.NamaHakAkses = ""
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub HapusByNama(ByVal vNama As String, Optional ByVal vShowMsgbox As Boolean = False)
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery_ByNama(vNama)

            If vShowMsgbox = True Then
                MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''' <param name="Kodehakakses">Kode Hak khusus yang akan dihapus datanya</param>
    Public Function Hapus(ByVal Kodehakakses As String) As Boolean
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter
        Dim DTA1 As New DSUserTableAdapters.tblantarauserhakaksesTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            'cek apakah ada yang sudah menggunakan
            If DTA1.ScalarQuery_HitungHakAkses(Kodehakakses).HasValue Then
                If DTA1.ScalarQuery_HitungHakAkses(Kodehakakses).Value = 0 Then
                    DTA.Delete(Kodehakakses)
                Else
                    Return False
                End If
            Else
                DTA.Delete(Kodehakakses)
            End If

            MsgBox("Data sudah dihapus", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan() As Boolean
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Update(Me.NamaHakAkses, Me.Description, Me.KodeHakAkses)

            'konfimasi
            MsgBox("Perubahan data tersimpan dibasis data", MsgBoxStyle.Information, "Information")

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini mengisikan list box dengan semua nama hak khusus yang ada dibasis data
    ''' </summary>
    Public Function IsiListBox(ByVal KodeUser As String, ByVal LB As ListBox) As Boolean
        Dim DTA As New DSUserTableAdapters.tblantarauserhakaksesTableAdapter
        Dim TBL As New DSUser.tblantarauserhakaksesDataTable
        Dim row As DSUser.tblantarauserhakaksesRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            LB.Items.Clear()

            'fill data
            DTA.FillBy_KodeUser(TBL, KodeUser)

            For Each row In TBL.Rows
                LB.Items.Add(Me.GetDescription(row.Kode_Hak_Akses))
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsiListBox(ByVal LB As ListBox) As Boolean
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter
        Dim TBL As New DSUser.tblmasterhakaksesDataTable
        Dim row As DSUser.tblmasterhakaksesRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            LB.Items.Clear()

            'fill data
            DTA.Fill(TBL)

            For Each row In TBL.Rows
                LB.Items.Add(Me.GetDescription(row.Kode))
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal DGV As DataGridView, ByVal View As DataView, ByVal Text As String, _
    Optional ByVal DenganKode As Boolean = True) As Boolean
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter
        Dim TBL As New DSUser.tblmasterhakaksesDataTable
        Dim cmd As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil data 
            DTA.Fill(TBL)

            'filter
            If DenganKode = True Then
                cmd = "Kode LIKE '" + Text + "*'"
            Else
                cmd = "Nama LIKE '" + Text + "*'"
            End If

            View = TBL.DefaultView
            View.RowFilter = cmd

            'fill DGV
            DGV.DataSource = View
            DGV.Refresh()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL.Dispose()
            DTA.Dispose()
        End Try
    End Function

    Public Function GetKode(ByVal Nama As String) As Integer
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetKode(Nama).HasValue Then
                Return DTA.ScalarQuery_GetKode(Nama).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetKode1(ByVal Description As String) As Integer
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetKode1(Description).HasValue = False Then
                Return "0"
            Else
                Return DTA.ScalarQuery_GetKode1(Description).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "0"
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNama(ByVal vKode As Double) As String
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetNama(vKode)
            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDescription(ByVal Kode As String) As String
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDescription(Kode)) = False Then
                Return DTA.ScalarQuery_GetDescription(Kode)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Class
