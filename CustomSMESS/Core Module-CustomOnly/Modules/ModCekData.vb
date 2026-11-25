Option Explicit On
Option Strict Off

Public Module ModCekData
    ''' <summary>
    ''' Fungsi ini mengembalikan tanggal pertama dari bulan yang sama dengan 
    ''' tanggal pada parameter
    ''' </summary>
    ''' <param name="Tanggal"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFirstDateofTheMonth(ByVal Tanggal As Date) As Date
        Dim obj As New ClsCekData

        Try
            Return obj.GetFirstDateofTheMonth(Tanggal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini mengembalikan tanggal terakhir dari bulan yang sama
    ''' dengan tanggal pada parameter
    ''' </summary>
    ''' <param name="Tanggal"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLasDateOfTheMonth(ByVal Tanggal As Date) As Date
        Dim obj As New ClsCekData

        Try
            Return obj.GetLasDateOfTheMonth(Tanggal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini mengosongkan semua kontrol dalam koleksi
    ''' </summary>
    ''' <param name="KoleksiControl"></param>
    ''' <param name="KosongkanText">Jika true maka text box yang dalam koleksi akan dikosongkan</param>
    ''' <param name="KosongkanDTP">Jika true maka date time picker dalam koleksi akan bernilai tanggal hari ini</param>
    ''' <param name="KosongkanCombo">Jika true maka item pertama yang akan dipilih</param>
    ''' <remarks></remarks>
    Public Sub KosongkanControl(ByVal KoleksiControl As Windows.Forms.Control.ControlCollection, _
    ByVal KosongkanText As Boolean, ByVal KosongkanDTP As Boolean, ByVal KosongkanCombo As _
    Boolean)
        Dim obj As New ClsCekData

        Try
            obj.KosongkanControl(KoleksiControl, KosongkanText, KosongkanDTP, KosongkanCombo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Fungsi ini melakukan format text yang merupakan angka dalam harga rupiah
    ''' </summary>
    ''' <param name="Angka">Angka ingin diformat, harus dalam format #,###</param>
    ''' <param name="AngkaDecimal">Jumlah angka dibelakang koma yang diinginkan</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FormatHarga(ByVal Angka As String, ByVal AngkaDecimal As Integer, _
    Optional ByVal TampilkanSatuan As Boolean = False, Optional ByVal Satuan As String = "Rp. ") As String
        Dim obj As New ClsCekData

        Try
            Return obj.FormatHargaUS(Angka, AngkaDecimal, TampilkanSatuan, Satuan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Rp. 0,00"
        Finally
            obj = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini melakukan format text yang merupakan angka dalam format US, format angka dalam textbox harus 
    ''' #,###
    ''' </summary>
    ''' <param name="KoleksiTextBox">Koleksi yang berisikan text box yang textnya sudah dalam format
    ''' angka semua</param>
    ''' <param name="AngkaDecimal">Jumlah angka dibelakang koma yang diinginkan</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FormatHarga(ByVal KoleksiTextBox As Collection, ByVal AngkaDecimal As Integer, _
    Optional ByVal TampilkanSatuan As Boolean = True, Optional ByVal Satuan As String = "Rp. ") As Boolean
        Dim obj As New ClsCekData

        Try
            Return obj.FormatHargaUS(KoleksiTextBox, AngkaDecimal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-"
        Finally
            obj = Nothing
        End Try
    End Function

    Public Function FormatAngka(ByVal Data As String, ByVal AngkaDecimal As Integer) As String
        Dim obj As New ClsCekData

        Try
            Return obj.FormatAngka(Data, AngkaDecimal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            obj = Nothing
        End Try
    End Function

    Public Sub FormatAngka(ByVal KoleksiTextBox As Collection, ByVal AngkaDecimal As String)
        Dim obj As New ClsCekData

        Try
            obj.FormatAngka(KoleksiTextBox, AngkaDecimal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    Public Function CekNumerik(ByVal Data As String) As Boolean
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.CekNumerik(Data)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objCekData = Nothing
        End Try
    End Function

    Public Function CekNumerik(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.CekNumerik(KoleksiTextBox, KoleksiNamaData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objCekData = Nothing
        End Try
    End Function

    Public Function CekNumerik(ByVal KoleksiData As List(Of String), _
    ByVal KoleksiNamaData As List(Of String), _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.CekNumerik(KoleksiData, KoleksiNamaData, TampilkanPesan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objCekData = Nothing
        End Try
    End Function

    Public Function CekBilanganPositif(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.CekBilanganPositif(KoleksiTextBox, KoleksiNamaData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objCekData = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Cek bilangan positif adalah fungsi untuk mengecek apakah isi dari sekumpulan cek box adalah bilang 
    ''' positif. Sudah termasuk pengecekan numerik.
    ''' </summary>
    Public Function CekBilanganPositif(ByVal KoleksiData As List(Of String), _
    ByVal KoleksiNamaData As List(Of String), _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.CekBilanganPositif(KoleksiData, KoleksiNamaData, TampilkanPesan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objCekData = Nothing
        End Try
    End Function

    Public Function CekBilanganPositif(ByVal Data As String) As Boolean
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.CekBilanganPositif(Data)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objCekData = Nothing
        End Try
    End Function

    Public Function CekNonNegatif(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.CekNonNegatif(KoleksiTextBox, KoleksiNamaData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objCekData = Nothing
        End Try
    End Function

    Public Function CekNonNegatif(ByVal KoleksiData As List(Of String), _
    ByVal KoleksiNamaData As List(Of String), Optional ByVal tampilkanPesan As _
    Boolean = True) As Boolean
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.CekNonNegatif(KoleksiData, KoleksiNamaData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objCekData = Nothing
        End Try
    End Function

    Public Function CekNonNegatif(ByVal Data As String) As Boolean
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.CekNonNegatif(Data)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objCekData = Nothing
        End Try
    End Function

    Public Function CekDataKosong(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objCekData = Nothing
        End Try
    End Function

    Public Function CekDataKosong(ByVal ListString As List(Of String), _
   ByVal ListNamaData As List(Of String), _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.CekDataKosong(ListString, ListNamaData, _
            TampilkanPesan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objCekData = Nothing
        End Try
    End Function

    Public Function CekDataKosong(ByVal Data As String) As Boolean
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.CekDataKosong(Data)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objCekData = Nothing
        End Try
    End Function

    Public Sub KendalikanControl(ByVal KoleksiControl As Windows.Forms.Control.ControlCollection, _
    ByVal KendalikanText As Boolean, ByVal txt As Boolean, ByVal KendalikanButton As Boolean, ByVal btn As _
    Boolean, ByVal KendalikanDTP As Boolean, ByVal dtp As Boolean, ByVal KendalikanCombo As Boolean, _
    ByVal cbo As Boolean)
        Dim obj As New ClsCekData

        Try
            obj.KendalikanControl(KoleksiControl, KendalikanText, txt, KendalikanButton, btn, _
            KendalikanDTP, dtp, KendalikanCombo, cbo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    Public Function BuatTextFormatInt(ByVal Data As String) As String
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.BuatTextFormatInt(Data)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function BuatTextFormatDecimal(ByVal Data As String) As String
        Dim objCekData As ClsCekData = New ClsCekData
        Try
            Return objCekData.BuatTextFormatDecimal(Data)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Module
