Option Explicit On
Option Strict Off

Public Class ClsHakKhusus

    Private mKodehakKhusus As String
    Private mNamahakKhusus As String
    Private mDescription As String

    Public Property KodehakKhusus() As String
        Get
            Return mKodehakKhusus
        End Get
        Set(ByVal value As String)
            mKodehakKhusus = value
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

    Public Property NamahakKhusus() As String
        Get
            Return mNamahakKhusus
        End Get
        Set(ByVal value As String)
            mNamahakKhusus = value
        End Set
    End Property

    ''' <summary>
    ''' Fungsi ini menyimpan data master ke basis data
    ''' </summary>
    Public Function Simpan(ByVal txtKodeHak As TextBox) As Boolean
        Dim DTA As New DSUserTableAdapters.tblmasterhakkhususTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cek apakah ada yang sama
            If DTA.ScalarQuery_CekNama(Me.NamahakKhusus).HasValue = False Then
                DTA.Insert(Me.NamahakKhusus, Me.Description)
            Else
                If DTA.ScalarQuery_CekNama(Me.NamahakKhusus).Value = 0 Then
                    DTA.Insert(Me.NamahakKhusus, Me.Description)
                Else
                    Return False
                End If
            End If

            txtKodeHak.Text = Me.GetKode(Me.NamahakKhusus)

            MsgBox("Data tersimpan dibasis data.", MsgBoxStyle.Information, "Information")
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini mencari detail data untuk hak khusus tertentu
    ''' </summary>
    Public Function CariDetail() As Boolean
        Dim DTA As New DSUserTableAdapters.tblmasterhakkhususTableAdapter
        Dim TBL As New DSUser.tblmasterhakkhususDataTable
        Dim row As DSUser.tblmasterhakkhususRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fill basis data
            DTA.FillBy_KodeHakKhusus(TBL, Me.KodehakKhusus)

            'salin data
            If TBL.Rows.Count > 0 Then
                row = TBL.Item(0)
                With row
                    Me.KodehakKhusus = .Kode
                    Me.NamahakKhusus = .Nama
                    Me.Description = .Descriptionl
                End With
            Else
                Me.KodehakKhusus = ""
                Me.Description = ""
                Me.NamahakKhusus = ""
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    ''' <param name="KodehakKhusus">Kode Hak khusus yang akan dihapus datanya</param>
    Public Function Hapus(ByVal KodehakKhusus As String) As Boolean
        Dim DTA As New DSUserTableAdapters.tblmasterhakkhususTableAdapter
        Dim DTA1 As New DSUserTableAdapters.tblantarauserhakkhususTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            'cek apakah ada yang sudah menggunakan
            If DTA1.ScalarQuery_HitungHakKhusus(KodehakKhusus).HasValue Then
                If DTA1.ScalarQuery_HitungHakKhusus(KodehakKhusus).Value = 0 Then
                    DTA.Delete(KodehakKhusus)
                Else
                    Return False
                End If
            Else
                DTA.Delete(KodehakKhusus)
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
        Dim DTA As New DSUserTableAdapters.tblmasterhakkhususTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Update(Me.NamahakKhusus, Me.Description, Me.KodehakKhusus)
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
        Dim DTA As New DSUserTableAdapters.tblantarauserhakkhususTableAdapter
        Dim TBL As New DSUser.tblantarauserhakkhususDataTable
        Dim row As DSUser.tblantarauserhakkhususRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            LB.Items.Clear()

            'fill data
            DTA.FillBy_KodeUser(TBL, KodeUser)

            For Each row In TBL.Rows
                LB.Items.Add(Me.GetDescription(row.Kode_Hak_Khusus))
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsiListBox(ByVal LB As ListBox) As Boolean
        Dim DTA As New DSUserTableAdapters.tblmasterhakkhususTableAdapter
        Dim TBL As New DSUser.tblmasterhakkhususDataTable
        Dim row As DSUser.tblmasterhakkhususRow

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
        Dim DTA As New DSUserTableAdapters.tblmasterhakkhususTableAdapter
        Dim TBL As New DSUser.tblmasterhakkhususDataTable
        Dim filter As String

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil data 
            DTA.Fill(TBL)

            'filter
            If DenganKode = True Then
                filter = "Kode LIKE '" + Text + "*'"
            Else
                filter = "Nama LIKE '" + Text + "*'"
            End If
            View = TBL.DefaultView
            View.RowFilter = filter

            'fil DGV
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
        Dim DTA As New DSUserTableAdapters.tblmasterhakkhususTableAdapter

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
        Dim DTA As New DSUserTableAdapters.tblmasterhakkhususTableAdapter

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

    Public Function GetDescription(ByVal Kode As String) As String
        Dim DTA As New DSUserTableAdapters.tblmasterhakkhususTableAdapter

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
