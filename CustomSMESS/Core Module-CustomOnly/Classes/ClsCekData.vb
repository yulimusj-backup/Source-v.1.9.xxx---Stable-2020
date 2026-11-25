Public Class ClsCekData
    Public Function CekNonPositif(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim i As Int16
        Try
            'sebelum cek bilangan non negatif, cek dulu numeriknya
            If Me.CekNumerik(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            Do While i < KoleksiTextBox.Count
                'diasumsikan data yang ada didalam text box sudah pasti numerik, sudah dicek numeriknya
                If Convert.ToDecimal(KoleksiTextBox.Item(i + 1).text) > 0 Then
                    MsgBox("Data " + Convert.ToString(KoleksiNamaData.Item(i + 1)) + _
                    " harus negatif atau nol.", MsgBoxStyle.Critical, "Error")
                    KoleksiTextBox.Item(i + 1).focus()
                    Return False
                End If
                i += 1
            Loop

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function CekNonPositif(ByVal Data As String) As Boolean
        Try
            'sebelum cek bilangan non negatif, cek dulu numeriknya
            If Me.CekNumerik(Data) = False Then
                Return False
            End If

            'diasumsikan data yang ada didalam text box sudah pasti numerik, sudah dicek numeriknya
            If Convert.ToDecimal(Data) > 0 Then
                Return False
            End If

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CekNonPositif(ByVal KoleksiData As List(Of String), _
   ByVal KoleksiNamaData As List(Of String), Optional ByVal tampilkanPesan As _
   Boolean = True) As Boolean
        Dim i As Int16
        Try
            'sebelum cek bilangan non negatif, cek dulu numeriknya
            If Me.CekNumerik(KoleksiData, KoleksiNamaData) = False Then
                Return False
            End If

            Do While i < KoleksiData.Count
                'diasumsikan data yang ada didalam text box sudah pasti numerik, sudah dicek numeriknya
                If Convert.ToDecimal(KoleksiData.Item(i)) > 0 Then
                    If tampilkanPesan Then
                        MsgBox("Data " + KoleksiNamaData.Item(i) + _
                        " harus negatif atau nol.", MsgBoxStyle.Critical, "Error")
                    End If
                    Return False
                End If
                i += 1
            Loop

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Enum EnumTipeRounding As Integer
        Up
        Down
        Middlepoint
    End Enum

    'Berpasangan dengan Get_disk_list
    Private Function Cek_Disc_List_TidakBernilaiNol(ByVal ListString As List(Of String)) As Integer
        Dim ParamStr As String = ""
        Dim VarString As String = ""
        Dim WrongFormat As Integer = 0
        Dim i As Integer = 0
        Try
            'Buang semua list yang bernilai = 0
            i = 0
            While i <= ListString.Count - 1
                ParamStr = ListString.Item(i)
                If ParamStr.StartsWith("++") = True Then
                    If ParamStr.EndsWith("%") = True Then
                        VarString = ParamStr.Substring(ParamStr.IndexOf("+") + 2, ParamStr.IndexOf("%") - 2)
                        If CDec(VarString) <= 0 Then
                            ListString.RemoveAt(i)
                            WrongFormat += 1
                        Else
                            i += 1
                        End If
                    Else
                        'Langsung hapus kalo ada format (10%++500)
                        ListString.RemoveAt(i)
                        WrongFormat += 1
                    End If
                ElseIf ParamStr.StartsWith("+") = True Then
                    If ParamStr.EndsWith("%") = True Then
                        VarString = ParamStr.Substring(ParamStr.IndexOf("+") + 1, ParamStr.IndexOf("%") - 1)
                        If CDec(VarString) <= 0 Then
                            ListString.RemoveAt(i)
                            WrongFormat += 1
                        Else
                            i += 1
                        End If
                    Else
                        VarString = ParamStr.Substring(ParamStr.IndexOf("+") + 1)
                        If CDec(VarString) <= 0 Then
                            ListString.RemoveAt(i)
                            WrongFormat += 1
                        Else
                            i += 1
                        End If
                    End If
                End If
            End While

            'Hasil Pengecekan nilai <> 0
            Return WrongFormat
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Get_Disc_List(ByVal txtNumber As String, ByVal ListString As List(Of String))
        Dim DummyStr As String = ""
        Dim PenandaHapus As Boolean = False
        Try
            ListString.Clear()
            DummyStr = txtNumber
            While DummyStr <> ""
                If DummyStr.IndexOf("+") = -1 Then
                    If PenandaHapus = True Then
                        ListString.Add("++" + DummyStr)
                        DummyStr = ""
                    Else
                        ListString.Add("+" + DummyStr)
                        DummyStr = ""
                    End If
                Else
                    If DummyStr.IndexOf("+") = 0 Then
                        DummyStr = DummyStr.Remove(0, 1)
                        PenandaHapus = True
                    Else
                        If PenandaHapus = False Then
                            ListString.Add("+" + DummyStr.Substring(0, DummyStr.IndexOf("+")))
                            DummyStr = DummyStr.Substring(DummyStr.IndexOf("+") + 1)
                        Else
                            ListString.Add("++" + DummyStr.Substring(0, DummyStr.IndexOf("+")))
                            DummyStr = DummyStr.Substring(DummyStr.IndexOf("+") + 1)
                            PenandaHapus = False
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetDisc_GroupingCekKoma(ByVal TxtNumber As String) As String
        Try
            While TxtNumber.Contains(".") = True
                If TxtNumber.EndsWith("0") Or TxtNumber.EndsWith(".") Then
                    TxtNumber = TxtNumber.Remove(TxtNumber.Length - 1, 1)
                Else
                    Exit While
                End If
            End While
            Return TxtNumber
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function Get_Disc_Grouping(ByVal ListString As List(Of String), ByVal vDef_Prec As Integer, ByVal vObjCD As ClsCekData) As String
        Dim Retval As String = ""
        Dim ParamStr As String = ""
        Dim VarString As String = ""
        Try
            For i As Integer = 0 To ListString.Count - 1
                ParamStr = ListString.Item(i)
                If ParamStr.StartsWith("++") = True Then
                    If ParamStr.EndsWith("%") = True Then
                        VarString = vObjCD.GetDisc_GroupingCekKoma(vObjCD.DigitGrouping(ParamStr.Substring(ParamStr.IndexOf("+") + 2, ParamStr.IndexOf("%") - 2), 2))
                        If String.IsNullOrEmpty(Retval) = False Then
                            Retval += "++" + VarString + "%"
                        End If
                    End If
                ElseIf ParamStr.StartsWith("+") = True Then
                    If ParamStr.EndsWith("%") = True Then
                        VarString = vObjCD.GetDisc_GroupingCekKoma(vObjCD.DigitGrouping(ParamStr.Substring(ParamStr.IndexOf("+") + 1, ParamStr.IndexOf("%") - 1), 2))
                        If String.IsNullOrEmpty(Retval) = False Then
                            Retval += "+" + VarString + "%"
                        Else
                            Retval = VarString + "%"
                        End If
                    Else
                        VarString = vObjCD.GetDisc_GroupingCekKoma(vObjCD.DigitGrouping(ParamStr.Substring(ParamStr.IndexOf("+") + 1), vDef_Prec))
                        If String.IsNullOrEmpty(Retval) = False Then
                            Retval += "+" + VarString
                        Else
                            Retval = VarString
                        End If
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub Disc_Leave(ByVal TxtBox As TextBox, ByVal txtNumber As String, ByVal sender As Object, ByVal e As System.EventArgs, _
    ByVal vDef_Prec As Integer, ByVal vObjCD As Core_Module.ClsCekData, ByVal vBuffer As String)
        Dim ListString As New List(Of String)
        Dim ParamStr As String
        Dim NilaiAcuanAwalPersen As Decimal = 100
        Dim NilaiAcuanAkhirPersen As Decimal = 100
        Try
            'cek text-nya:
            '1. Buang semua spasi
            txtNumber = txtNumber.Trim()
            txtNumber = txtNumber.Replace(" ", "")
            '2. pastikan formatnya benar, jika format salah, maka kembalikan ke nilai buffer
            'format salah jika:
            '   a. format salah jika
            '   Diawali +, ++, %, diakhiri +, ++
            If txtNumber.StartsWith("+") = True Or txtNumber.StartsWith("%") = True Or _
            txtNumber.EndsWith("+") = True Or txtNumber.StartsWith("++") = True Or _
            txtNumber.EndsWith("++") = True Then
                TxtBox.Text = vBuffer
                Exit Sub
            End If

            '   Format yang bernilai = 0 (+0,+0%,++0%)
            Me.Get_Disc_List(txtNumber, ListString)
            If Me.Cek_Disc_List_TidakBernilaiNol(ListString) > 0 Then
                If txtNumber <> "0" Then
                    TxtBox.Text = vBuffer
                    Exit Sub
                Else
                    TxtBox.Text = 0
                    Exit Sub
                End If
            End If

            '   b. nilai persen > 100%
            Dim VarPersen As String = ""
            Dim VarHasil As Decimal = 0
            For Each ParamStr In ListString
                If ParamStr.EndsWith("%") = True Then
                    If ParamStr.StartsWith("++") = True Then
                        If ParamStr.EndsWith("%") = True Then
                            VarPersen = ParamStr.Substring(ParamStr.IndexOf("++") + 2, ParamStr.IndexOf("%") - 2)
                            If CDec(VarPersen) > 0 Then
                                VarHasil = VarPersen / 100 * NilaiAcuanAkhirPersen
                                NilaiAcuanAkhirPersen -= VarHasil
                            End If
                        End If
                    ElseIf ParamStr.StartsWith("+") = True Then
                        If ParamStr.EndsWith("%") = True Then
                            VarPersen = ParamStr.Substring(ParamStr.IndexOf("+") + 1, ParamStr.IndexOf("%") - 1)
                            If CDec(VarPersen) > 0 Then
                                VarHasil = VarPersen / 100 * NilaiAcuanAwalPersen
                                NilaiAcuanAkhirPersen -= VarHasil
                            End If
                        End If
                    End If
                End If
            Next
            If NilaiAcuanAkhirPersen < 0 Then
                txtNumber = vBuffer
                Exit Sub
            End If
            '3. digit grouping dan decimal
            '   a. untuk nominal, pastikan ada digit grouping dan decimalnya
            '   b. untuk persen 2 decimal
            txtNumber = Me.Get_Disc_Grouping(ListString, vDef_Prec, vObjCD)
            TxtBox.Text = txtNumber
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListString = Nothing
        End Try
    End Sub

    Public Function Get_Disc_Amount(ByVal vDiscString As String, ByVal vBaseAmount As Decimal, _
    ByVal vDecimalPoint As Integer, ByVal vRoundingTipe As EnumTipeRounding, _
    ByVal vRoundingBaseNumber As Decimal, ByVal vObjCD As ClsCekData) As String
        Dim Retval As String = "0"
        Dim ParamStr As String = ""
        Dim ListString As New List(Of String)
        Dim NilaiAcuanAkhirPersen As Decimal = 0
        Dim NettAmount As Decimal = 0
        Dim HasilMod As Decimal = 0
        Dim RoundingDecimalUpDown As Decimal = 1
        Try
            If vBaseAmount > 0 Then
                vObjCD.Get_Disc_List(vDiscString, ListString)
                If Me.Cek_Disc_List_TidakBernilaiNol(ListString) <= 0 Then
                    Dim VarPersen As String = ""
                    Dim VarHasil As Decimal = 0
                    For Each ParamStr In ListString
                        If ParamStr.EndsWith("%") = True Then
                            If ParamStr.StartsWith("++") = True Then
                                If ParamStr.EndsWith("%") = True Then
                                    VarPersen = ParamStr.Substring(ParamStr.IndexOf("++") + 2, ParamStr.IndexOf("%") - 2)
                                    If CDec(VarPersen) > 0 Then
                                        VarHasil = VarPersen / 100 * NettAmount
                                        NilaiAcuanAkhirPersen += VarHasil
                                        NettAmount = vBaseAmount - NilaiAcuanAkhirPersen
                                    End If
                                End If
                            ElseIf ParamStr.StartsWith("+") = True Then
                                If ParamStr.EndsWith("%") = True Then
                                    VarPersen = ParamStr.Substring(ParamStr.IndexOf("+") + 1, ParamStr.IndexOf("%") - 1)
                                    If CDec(VarPersen) > 0 Then
                                        VarHasil = VarPersen / 100 * vBaseAmount
                                        NilaiAcuanAkhirPersen += VarHasil
                                        NettAmount = vBaseAmount - NilaiAcuanAkhirPersen
                                    End If
                                End If
                            End If
                        Else
                            VarPersen = ParamStr.Substring(ParamStr.IndexOf("+") + 1)
                            If CDec(VarPersen) > 0 Then
                                'Tidak pakai persen (+1500)
                                VarHasil = VarPersen
                                NilaiAcuanAkhirPersen += VarHasil
                                NettAmount = vBaseAmount - NilaiAcuanAkhirPersen
                            End If
                        End If
                    Next
                    'Sesuaikan hasil perhitungan dengan rounding tipe dan rounding base number 
                    Retval = Me.GetRoundingValue(NilaiAcuanAkhirPersen, vDecimalPoint, _
                    vRoundingTipe, vRoundingBaseNumber, vObjCD)

                    '-----
                    'If vRoundingTipe = EnumTipeRounding.Up Then
                    'Retval = +Math.Ceiling(NilaiAcuanAkhirPersen / vRoundingBaseNumber) * vRoundingBaseNumber
                    'End If
                    'If vRoundingTipe = EnumTipeRounding.Down Then
                    'Retval = Math.Floor(NilaiAcuanAkhirPersen / vRoundingBaseNumber) * vRoundingBaseNumber
                    'End If
                    'If vRoundingTipe = EnumTipeRounding.Middlepoint Then
                    'HasilMod = CDec(NilaiAcuanAkhirPersen) Mod vRoundingBaseNumber
                    'If HasilMod < vRoundingBaseNumber / 2 Then
                    'Retval = Math.Floor(NilaiAcuanAkhirPersen / vRoundingBaseNumber) * vRoundingBaseNumber
                    'ElseIf HasilMod >= vRoundingBaseNumber / 2 Then
                    'Retval = Math.Ceiling(NilaiAcuanAkhirPersen / vRoundingBaseNumber) * vRoundingBaseNumber
                    'End If
                    'End If
                    '-----

                    'digit grouping dan decimal
                    Retval = vObjCD.DigitGrouping(Retval, vDecimalPoint)
                End If
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "0"
        Finally
            ListString = Nothing
        End Try
    End Function

    Public Function GetRoundingValue(ByVal vAmount As Decimal, _
    ByVal vDecimalPoint As Integer, ByVal vRoundingTipe As EnumTipeRounding, _
    ByVal vRoundingBaseNumber As Decimal, ByVal vObjCD As ClsCekData) As Decimal
        Dim Retval As Decimal = 0
        Dim HasilMod As Decimal = 0
        Try
            'Sesuaikan hasil perhitungan dengan rounding tipe dan rounding base number 
            If vRoundingTipe = EnumTipeRounding.Up Then
                Retval = +Math.Ceiling(vAmount / vRoundingBaseNumber) * vRoundingBaseNumber
            End If
            If vRoundingTipe = EnumTipeRounding.Down Then
                Retval = Math.Floor(vAmount / vRoundingBaseNumber) * vRoundingBaseNumber
            End If
            If vRoundingTipe = EnumTipeRounding.Middlepoint Then
                HasilMod = CDec(vAmount) Mod vRoundingBaseNumber
                If HasilMod < vRoundingBaseNumber / 2 Then
                    Retval = Math.Floor(vAmount / vRoundingBaseNumber) * vRoundingBaseNumber
                ElseIf HasilMod >= vRoundingBaseNumber / 2 Then
                    Retval = Math.Ceiling(vAmount / vRoundingBaseNumber) * vRoundingBaseNumber
                End If
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Sub Disc_Enter(ByVal txtNumber As TextBox, ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If txtNumber.Enabled = True And txtNumber.ReadOnly = False Then
                If IsNumeric(txtNumber.Text) = False Then
                    txtNumber.Text = "0"
                End If
                If CDec(txtNumber.Text) = 0 Then
                    txtNumber.Text = "0"
                End If
                While txtNumber.Text.Contains(".") = True
                    If txtNumber.Text.EndsWith("0") Or txtNumber.Text.EndsWith(".") Then
                        txtNumber.Text = txtNumber.Text.Remove(txtNumber.Text.Length - 1, 1)
                    Else
                        Exit While
                    End If
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Disc_KeyDown(ByVal txtNumber As TextBox, ByVal sender As Object, _
    ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim LastDigit As String = ""
        Dim IndexLastDigit As Integer = -1
        Dim i As Integer = 0
        Try
            If txtNumber.Text = "" Then
                txtNumber.Text = "0"
            End If
            If txtNumber.ReadOnly = True Then
                e.SuppressKeyPress = True
                Exit Sub
            End If

            If e.Alt = False Then
                If e.KeyData = Keys.OemPeriod Or e.KeyData = Keys.Decimal Then
                    LastDigit = txtNumber.Text.Substring(txtNumber.Text.Length - 1)
                    If IsNumeric(LastDigit) = True Then
                        'kalo ada tanda plus (boleh pakai koma)
                        i = txtNumber.Text.Length - 1
                        While i > 0
                            LastDigit = txtNumber.Text.Substring(i)
                            If LastDigit.StartsWith("+") = True Then
                                IndexLastDigit = i
                                Exit While
                            Else
                                i -= 1
                            End If
                        End While
                        If IndexLastDigit < 0 Then
                            'berlaku kalo ga ada tanda plus (ga boleh ada koma lagi)
                            If txtNumber.Text.Contains(".") = True Then
                                e.SuppressKeyPress = True
                                Exit Sub
                            End If
                        End If
                    Else
                        'kalo last digit bukan angka, buang
                        e.SuppressKeyPress = True
                        Exit Sub
                    End If
                Else
                    If txtNumber.Text = "0" Then
                        txtNumber.Text = ""
                    End If
                    If e.KeyValue = 9 Then 'tab button
                        If IsNothing(txtNumber.GetNextControl(txtNumber, True)) = False Then
                            If txtNumber.GetNextControl(txtNumber, True).CanFocus = True Then
                                txtNumber.GetNextControl(txtNumber, True).Focus()
                                Exit Sub
                            End If
                        End If
                    ElseIf e.KeyValue = 46 Then 'delete = clear
                        txtNumber.Text = "0"
                        Exit Sub
                    ElseIf e.KeyValue = 8 Then 'backspace
                        Exit Sub
                    End If

                    'Cek Last Digit %, kalo persen maka hanya tanda + / ++
                    If e.KeyCode = Keys.NumPad0 Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.EndsWith("%") = True Then
                            Exit Sub
                        End If
                        txtNumber.Text = txtNumber.Text + "0"
                    ElseIf e.KeyCode = Keys.NumPad1 Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.EndsWith("%") = True Then
                            Exit Sub
                        End If
                        txtNumber.Text = txtNumber.Text + "1"
                    ElseIf e.KeyCode = Keys.NumPad2 Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.EndsWith("%") = True Then
                            Exit Sub
                        End If
                        txtNumber.Text = txtNumber.Text + "2"
                    ElseIf e.KeyCode = Keys.NumPad3 Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.EndsWith("%") = True Then
                            Exit Sub
                        End If
                        txtNumber.Text = txtNumber.Text + "3"
                    ElseIf e.KeyCode = Keys.NumPad4 Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.EndsWith("%") = True Then
                            Exit Sub
                        End If
                        txtNumber.Text = txtNumber.Text + "4"
                    ElseIf e.KeyCode = Keys.NumPad5 Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.EndsWith("%") = True Then
                            Exit Sub
                        End If
                        txtNumber.Text = txtNumber.Text + "5"
                    ElseIf e.KeyCode = Keys.NumPad6 Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.EndsWith("%") = True Then
                            Exit Sub
                        End If
                        txtNumber.Text = txtNumber.Text + "6"
                    ElseIf e.KeyCode = Keys.NumPad7 Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.EndsWith("%") = True Then
                            Exit Sub
                        End If
                        txtNumber.Text = txtNumber.Text + "7"
                    ElseIf e.KeyCode = Keys.NumPad8 Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.EndsWith("%") = True Then
                            Exit Sub
                        End If
                        txtNumber.Text = txtNumber.Text + "8"
                    ElseIf e.KeyCode = Keys.NumPad9 Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.EndsWith("%") = True Then
                            Exit Sub
                        End If
                        txtNumber.Text = txtNumber.Text + "9"
                    ElseIf e.KeyCode = Keys.Add Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.EndsWith("++") = False And txtNumber.Text <> "" And txtNumber.Text.EndsWith(".") = False Then
                            txtNumber.Text = txtNumber.Text + "+"
                        End If
                    ElseIf e.KeyCode = Keys.Oemplus And e.Shift = True Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.EndsWith("++") = False And txtNumber.Text <> "" And txtNumber.Text.EndsWith(".") = False Then
                            txtNumber.Text = txtNumber.Text + "+"
                        End If
                    ElseIf e.KeyCode = Keys.D5 And e.Shift = True Then
                        e.SuppressKeyPress = True
                        If txtNumber.Text.Length <> 0 Then
                            LastDigit = txtNumber.Text.Substring(txtNumber.Text.Length - 1)
                            If IsNumeric(LastDigit) = True Then
                                txtNumber.Text = txtNumber.Text + "%"
                            End If
                        End If
                    ElseIf Char.IsDigit(Chr(e.KeyValue)) = True Then
                        If e.Shift = True And Val(Chr(e.KeyValue)) = 5 Then
                            If txtNumber.Text.EndsWith("%") = True Then
                                e.SuppressKeyPress = True
                                Exit Sub
                            Else
                                'masukin langsung
                                Exit Sub
                            End If
                        ElseIf e.Shift = False Then
                            If txtNumber.Text.EndsWith("%") = True Then
                                e.SuppressKeyPress = True
                                Exit Sub
                            Else
                                'masukin langsung
                                Exit Sub
                            End If
                        Else
                            e.SuppressKeyPress = True
                            Exit Sub
                        End If
                    Else
                        'yang lain tidak dianggap
                        e.SuppressKeyPress = True
                        If txtNumber.Text = "" Then
                            txtNumber.Text = "0"
                        End If
                    End If
                End If
            Else
                e.SuppressKeyPress = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Disc_KeyUp(ByVal txtNumber As TextBox, ByVal sender As Object, _
    ByVal e As System.Windows.Forms.KeyEventArgs)
        Try
            If txtNumber.ReadOnly = True Then
                Exit Sub
            End If

            If e.Alt = False Then
                If txtNumber.Text <> "" Then

                Else
                    txtNumber.Text = "0"
                End If

                'geser kursor ke paling kanan
                txtNumber.DeselectAll()
                txtNumber.SelectionStart = txtNumber.Text.Length
                txtNumber.SelectionLength = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJmlPrecision(ByVal vNumber As Decimal) As Integer
        Dim i As Integer = 0
        Try
            While Math.Floor(vNumber) <> vNumber
                vNumber *= 10
                i += 1
            End While

            Return i
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetDateString_Indo_Panjang(ByVal vTanggal As Date) As String
        Dim Bulan As Integer = vTanggal.Month
        Dim BulanStr As String = ""
        Try
            If Bulan = 1 Then
                BulanStr = "Januari"
            ElseIf Bulan = 2 Then
                BulanStr = "Februari"
            ElseIf Bulan = 3 Then
                BulanStr = "Maret"
            ElseIf Bulan = 4 Then
                BulanStr = "April"
            ElseIf Bulan = 5 Then
                BulanStr = "Mei"
            ElseIf Bulan = 6 Then
                BulanStr = "Juni"
            ElseIf Bulan = 7 Then
                BulanStr = "Juli"
            ElseIf Bulan = 8 Then
                BulanStr = "Agustus"
            ElseIf Bulan = 9 Then
                BulanStr = "September"
            ElseIf Bulan = 10 Then
                BulanStr = "Oktober"
            ElseIf Bulan = 11 Then
                BulanStr = "Nopember"
            Else
                BulanStr = "Desember"
            End If

            Return Format(vTanggal.Date, "dd") + " " + BulanStr + " " + Format(vTanggal.Date, "yyyy")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "1 Januari 1950"
        End Try
    End Function

    Public Function GetDateString_Indo(ByVal vTanggal As Date) As String
        Dim Bulan As Integer = vTanggal.Month
        Dim BulanStr As String = ""
        Try
            If Bulan = 1 Then
                BulanStr = "Jan"
            ElseIf Bulan = 2 Then
                BulanStr = "Feb"
            ElseIf Bulan = 3 Then
                BulanStr = "Mar"
            ElseIf Bulan = 4 Then
                BulanStr = "Apr"
            ElseIf Bulan = 5 Then
                BulanStr = "Mei"
            ElseIf Bulan = 6 Then
                BulanStr = "Jun"
            ElseIf Bulan = 7 Then
                BulanStr = "Jul"
            ElseIf Bulan = 8 Then
                BulanStr = "Agt"
            ElseIf Bulan = 9 Then
                BulanStr = "Sept"
            ElseIf Bulan = 10 Then
                BulanStr = "Okt"
            ElseIf Bulan = 11 Then
                BulanStr = "Nop"
            Else
                BulanStr = "Des"
            End If

            Return Format(vTanggal.Date, "dd") + " " + BulanStr + " " + Format(vTanggal.Date, "yyyy")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "1 Jan 1950"
        End Try
    End Function

    Public Function GetFisrtDateOfTheYear(ByVal vDate As Date) As Date
        Try
            Return CDate("#1/1/" + CStr(vDate.Year) + "#")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetLastDateOfTheYear(ByVal vDate As Date) As Date
        Try
            Return CDate("#12/31/" + CStr(vDate.Year) + "#")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini mengembalikan tanggal pertama dari bulan yang sama dengan 
    ''' tanggal pada parameter
    ''' </summary>
    ''' <param name="Tanggal"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFirstDateofTheMonth(ByVal Tanggal As Date) As Date
        Try
            Return CDate("#" + CStr(Tanggal.Month) + "/1/" + _
            CStr(Tanggal.Year) + "#")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub SetBackColor(ByVal vKoleksiControl As Windows.Forms.Control.ControlCollection, _
    ByVal vColor As Color, ByVal vConvertReadOnly As Boolean)
        Dim ctrl As Control
        Dim txt As TextBox

        Try
            For Each ctrl In vKoleksiControl
                txt = TryCast(ctrl, TextBox)
                If IsNothing(txt) = False Then
                    If txt.Visible = True Then
                        If vConvertReadOnly = True Then
                            txt.BackColor = vColor
                        Else
                            If txt.ReadOnly = False Then
                                txt.BackColor = vColor
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetMonthOfTheYear(ByVal vMonth As Integer, Optional ByVal _
    UseEnglish As Boolean = True) As String
        Try
            If UseEnglish = True Then
                Select Case vMonth
                    Case 1
                        Return "January"
                    Case 2
                        Return "February"
                    Case 3
                        Return "March"
                    Case 4
                        Return "April"
                    Case 5
                        Return "May"
                    Case 6
                        Return "June"
                    Case 7
                        Return "Juli"
                    Case 8
                        Return "August"
                    Case 9
                        Return "September"
                    Case 10
                        Return "October"
                    Case 11
                        Return "November"
                    Case 12
                        Return "December"
                End Select
            Else
                Select Case vMonth
                    Case 1
                        Return "Januari"
                    Case 2
                        Return "Februari"
                    Case 3
                        Return "Maret"
                    Case 4
                        Return "April"
                    Case 5
                        Return "Mei"
                    Case 6
                        Return "Juni"
                    Case 7
                        Return "Juli"
                    Case 8
                        Return "Agustus"
                    Case 9
                        Return "September"
                    Case 10
                        Return "Oktober"
                    Case 11
                        Return "Nopember"
                    Case 12
                        Return "Desember"
                End Select
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
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
        Dim enddate As Date
        Try
            If Tanggal.Month < 12 Then
                enddate = CDate("#" + CStr(Tanggal.Month + 1) + "/1/" + _
                CStr(Tanggal.Year) + "#")
            Else
                enddate = CDate("#1/1/" + CStr(Tanggal.Year + 1) + "#")
            End If
            enddate = enddate - TimeSpan.FromDays(1)
            Return enddate
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'kelas yang berfungsi untuk melakukan fungsi pengecekan data
    Public Function CekNumerik(ByVal Data As String) As Boolean
        Try
            'sebelum cek numerik, cek apakah kosong
            If Me.CekDataKosong(Data) = False Then
                Return False
            End If

            'cek numerik
            If IsNumeric(Data) = False Then
                Return False
            End If

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CekDataKosong(ByVal Data As String, Optional ByVal vShowMessage As Boolean = True) As Boolean
        Try
            'hapus spasi didepan data
            Data = HapusSpasi(Data)

            'cek data
            If Data.Length = 0 Then
                Return False
            End If

            'cek berkaitan dengan karakter yang tidak diingikan --> ', *, %
            If Data.Contains("'") = True Or Data.Contains("*") = True Or Data.Contains("%") Or Data.Contains("#") = True Then
                Return False
            End If

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CekNonNegatif(ByVal Data As String) As Boolean
        Try
            'sebelum cek bilangan non negatif, cek dulu numeriknya
            If Me.CekNumerik(Data) = False Then
                Return False
            End If

            'diasumsikan data yang ada didalam text box sudah pasti numerik, sudah dicek numeriknya
            If Convert.ToDecimal(Data) < 0 Then
                Return False
            End If


            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CekBilanganPositif(ByVal Data As String) As Boolean
        Try
            'sebelum cek bilangan positif, cek dulu numeriknya
            If Me.CekNumerik(Data) = False Then
                Return False
            End If

            'diasumsikan data yang ada didalam text box sudah pasti numerik, sudah dicek numeriknya
            If Convert.ToDecimal(Data) <= 0 Then
                Return False
            End If

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CekNumerik(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim i As Int16
        Try
            'sebelum cek numerik, cek apakah kosong
            If Me.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            Do While i < KoleksiTextBox.Count
                'cek numerik
                If IsNumeric(KoleksiTextBox.Item(i + 1).text) = False Then
                    MsgBox("Data " + Convert.ToString(KoleksiNamaData.Item(i + 1)) + " harus berupa angka.", MsgBoxStyle.Critical, "Error")
                    KoleksiTextBox.Item(i + 1).focus()
                    Return False
                End If
                i += 1
            Loop

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CekNumerik(ByVal KoleksiData As List(Of String), _
    ByVal KoleksiNamaData As List(Of String), Optional ByVal TampilkanPesan _
    As Boolean = True) As Boolean
        Dim i As Int16
        Try
            'sebelum cek numerik, cek apakah kosong
            If Me.CekDataKosong(KoleksiData, KoleksiNamaData) = False Then
                Return False
            End If

            Do While i < KoleksiData.Count
                'hapus minus dibelakang
                KoleksiData.Item(i) = _
                Me.HapusNonNumerikBelakang(KoleksiData.Item(i))

                'cek numerik
                If IsNumeric(KoleksiData.Item(i)) = False Then
                    If TampilkanPesan Then
                        MsgBox("Data " + KoleksiNamaData.Item(i) + _
                        " harus berupa angka.", MsgBoxStyle.Critical, "Error")
                    End If
                    Return False
                End If
                i += 1
            Loop

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' Cek bilangan positif adalah fungsi untuk mengecek apakah isi dari sekumpulan cek box adalah bilang 
    ''' positif. Sudah termasuk pengecekan numerik.
    ''' </summary>
    Public Function CekBilanganPositif(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim i As Int16
        Try
            'sebelum cek bilangan positif, cek dulu numeriknya
            If Me.CekNumerik(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            Do While i < KoleksiTextBox.Count
                'diasumsikan data yang ada didalam text box sudah pasti numerik, sudah dicek numeriknya
                If Convert.ToDecimal(KoleksiTextBox.Item(i + 1).text) <= 0 Then
                    MsgBox("Data " + Convert.ToString(KoleksiNamaData.Item(i + 1)) + _
                    " harus positif dan tidak boleh nol.", MsgBoxStyle.Critical, "Error")
                    KoleksiTextBox.Item(i + 1).focus()
                    Return False
                End If
                i += 1
            Loop

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Cek bilangan positif adalah fungsi untuk mengecek apakah isi dari sekumpulan cek box adalah bilang 
    ''' positif. Sudah termasuk pengecekan numerik.
    ''' </summary>
    Public Function CekBilanganPositif(ByVal KoleksiData As List(Of String), _
    ByVal KoleksiNamaData As List(Of String), Optional ByVal TampilkanPesan _
    As Boolean = True) As Boolean
        Dim i As Int16
        Try
            'sebelum cek bilangan positif, cek dulu numeriknya
            If Me.CekNumerik(KoleksiData, KoleksiNamaData) = False Then
                Return False
            End If

            Do While i < KoleksiData.Count
                'diasumsikan data yang ada didalam text box sudah pasti numerik, sudah dicek numeriknya
                If Convert.ToDecimal(KoleksiData.Item(i)) <= 0 Then
                    If TampilkanPesan Then
                        MsgBox("Data " + KoleksiNamaData.Item(i) + _
                        " harus positif dan tidak boleh nol.", MsgBoxStyle.Critical, "Error")
                    End If
                    Return False
                End If
                i += 1
            Loop

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function CekNonNegatif(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim i As Int16
        Try
            'sebelum cek bilangan non negatif, cek dulu numeriknya
            If Me.CekNumerik(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            Do While i < KoleksiTextBox.Count
                'diasumsikan data yang ada didalam text box sudah pasti numerik, sudah dicek numeriknya
                If Convert.ToDecimal(KoleksiTextBox.Item(i + 1).text) < 0 Then
                    MsgBox("Data " + Convert.ToString(KoleksiNamaData.Item(i + 1)) + _
                    " harus positif atau nol.", MsgBoxStyle.Critical, "Error")
                    KoleksiTextBox.Item(i + 1).focus()
                    Return False
                End If
                i += 1
            Loop

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function CekNonNegatif(ByVal KoleksiData As List(Of String), _
    ByVal KoleksiNamaData As List(Of String), Optional ByVal tampilkanPesan As _
    Boolean = True) As Boolean
        Dim i As Int16
        Try
            'sebelum cek bilangan non negatif, cek dulu numeriknya
            If Me.CekNumerik(KoleksiData, KoleksiNamaData) = False Then
                Return False
            End If

            Do While i < KoleksiData.Count
                'diasumsikan data yang ada didalam text box sudah pasti numerik, sudah dicek numeriknya
                If Convert.ToDecimal(KoleksiData.Item(i)) < 0 Then
                    If tampilkanPesan Then
                        MsgBox("Data " + KoleksiNamaData.Item(i) + _
                        " harus positif atau nol.", MsgBoxStyle.Critical, "Error")
                    End If
                    Return False
                End If
                i += 1
            Loop

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function CekDataKosong(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim i As Int16
        Dim data As String
        Try
            Do While i < KoleksiTextBox.Count
                'salin data
                data = KoleksiTextBox.Item(i + 1).text
                data = HapusSpasi(data)

                'cek data
                If data.Length = 0 Then
                    MsgBox(Convert.ToString(KoleksiNamaData.Item(i + 1)) + " cannot be an empty string.", MsgBoxStyle.Critical, "Error")
                    KoleksiTextBox.Item(i + 1).focus()
                    Return False
                End If

                'cek berkaitan dengan karakter yang tidak diingikan --> ', *, %
                If data.Contains("'") = True Or data.Contains("*") = True Or data.Contains("%") = True Or data.Contains("#") = True Then
                    MsgBox("" + KoleksiNamaData.Item(i + 1) _
                        + " data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Critical, "Error")
                    Return False
                End If
                i += 1
            Loop

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function CekDataKosong(ByVal ListString As List(Of String), _
       ByVal ListNamaData As List(Of String), Optional ByVal TampilkanPesan _
       As Boolean = True) As Boolean
        Dim i As Int16
        Dim data As String
        Try
            Do While i < ListString.Count
                'salin data
                data = ListString.Item(i)
                data = HapusSpasi(data)

                'cek data
                If data.Length = 0 Then
                    If TampilkanPesan Then
                        MsgBox("Data " + ListNamaData.Item(i) _
                        + " harus diisi.", MsgBoxStyle.Critical, "Error")
                    End If
                    Return False
                End If

                'cek berkaitan dengan karakter yang tidak diingikan --> ', *, %
                If data.Contains("'") = True Or data.Contains("*") = True Or data.Contains("%") Or data.Contains("#") = True Then
                    If TampilkanPesan Then
                        MsgBox("Data " + ListNamaData.Item(i) _
                        + " tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Critical, "Error")
                    End If
                    Return False
                End If
                i += 1
            Loop

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function HapusSpasi(ByVal Data As String) As String
        Try
            Do While Data.EndsWith(" ") And Data.Length > 0
                Data = Data.Remove(Data.Length - 1, 1)
            Loop
            Return Data
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function HapusNonNumerikBelakang(ByVal Data As String) As String
        Try
            Dim i As String
            i = Data.Chars(Data.Length - 1)

            If IsNumeric(i) = False Then
                Data = Data.Remove(Data.Length - 1, 1)
            End If
            Return Data
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub KendalikanControl(ByVal KoleksiControl As Windows.Forms.Control.ControlCollection, _
    ByVal KendalikanText As Boolean, ByVal txt As Boolean, ByVal KendalikanButton As Boolean, ByVal btn As _
    Boolean, ByVal KendalikanDTP As Boolean, ByVal dtp As Boolean, ByVal KendalikanCombo As _
    Boolean, ByVal cbo As Boolean, Optional ByVal KendalikanCheckbox As Boolean = False, Optional ByVal chk As Boolean = False)
        Dim ctrl As Control
        Dim cek As Boolean = True

        Try
            For Each ctrl In KoleksiControl
                If ctrl.Name.Length > 0 Then
                    cek = True

                    If KendalikanText = True And cek = True Then
                        If ctrl.Name.Substring(0, 3).ToLower = "txt" Then
                            If txt = True Then
                                ctrl.Enabled = True
                            Else
                                ctrl.Enabled = False
                            End If
                            cek = False
                        End If
                    End If

                    If KendalikanButton = True And cek = True Then
                        If ctrl.Name.Substring(0, 3).ToLower = "btn" Or ctrl.Name.Substring(0, 3).ToLower = _
                        "cmd" Then
                            If btn = True Then
                                ctrl.Enabled = True
                            Else
                                ctrl.Enabled = False
                            End If
                            cek = False
                        End If
                    End If

                    If KendalikanDTP = True And cek = True Then
                        If ctrl.Name.Substring(0, 3).ToLower = "dtp" Then
                            If dtp = True Then
                                ctrl.Enabled = True
                            Else
                                ctrl.Enabled = False
                            End If
                            cek = False
                        End If
                    End If

                    If KendalikanCombo = True And cek = True Then
                        If ctrl.Name.Substring(0, 3).ToLower = "cbo" Then
                            If cbo = True Then
                                ctrl.Enabled = True
                            Else
                                ctrl.Enabled = False
                            End If
                            cek = False
                        End If
                    End If

                    If KendalikanCheckbox = True And cek = True Then
                        If ctrl.Name.Substring(0, 3).ToLower = "chk" Then
                            If cbo = True Then
                                ctrl.Enabled = True
                            Else
                                ctrl.Enabled = False
                            End If
                            cek = False
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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
        Dim ctrl As Control
        Dim txt As TextBox
        Dim dtp As DateTimePicker
        Dim cbo As ComboBox
        Dim cek As Boolean = True

        Try
            For Each ctrl In KoleksiControl
                If ctrl.Name.Length > 0 Then
                    cek = True

                    If KosongkanText = True And cek = True Then
                        If ctrl.Name.Substring(0, 3).ToLower = "txt" Then
                            txt = ctrl
                            txt.Clear()
                        End If
                    End If

                    If KosongkanDTP = True And cek = True Then
                        If ctrl.Name.Substring(0, 3).ToLower = "dtp" Then
                            dtp = ctrl
                            dtp.Value = Date.Today
                            cek = False
                        End If
                    End If

                    If KosongkanCombo = True And cek = True Then
                        If ctrl.Name.Substring(0, 3).ToLower = "cbo" Then
                            cbo = ctrl
                            If cbo.Items.Count > 0 Then
                                cbo.SelectedIndex = 0
                            End If
                            cek = False
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' Fungsi ini mengubah format angka menjadi angka decimal dengan jumlah angka decimal sesuai
    ''' dengan yang diinginkan
    ''' </summary>
    ''' <param name="KoleksiTextBox">Text box yang textnya akan diformat datanya</param>
    ''' <param name="AngkaDecimal">Jumlah angka dibelakang koma yang diinginkan</param>
    ''' <remarks></remarks>
    Public Sub FormatAngkaINA(ByVal KoleksiTextBox As Collection, ByVal AngkaDecimal As Integer)
        Dim txtbox As TextBox

        Try
            For Each txtbox In KoleksiTextBox
                txtbox.Text = Me.FormatAngka(txtbox.Text, AngkaDecimal)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' Fungsi ini mengubah format angka menjadi angka decimal dengan jumlah angka decimal sesuai
    ''' dengan yang diinginkan
    ''' </summary>
    ''' <param name="Data">Data numerik yang akan diformat, bila bukan numerik maka akan memberi
    ''' nilai balikan 0</param>
    ''' <param name="AngkaDecimal">Jumlah angka dibelakang koma yang diinginkan</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FormatAngkaINA(ByVal Data As String, ByVal AngkaDecimal As Integer) As Decimal
        Dim index As Integer
        Dim dummy As Decimal

        Try
            If Me.CekNumerik(Data) Then
                'cek jumlah karakter dibelakang decimal
                index = Data.IndexOfAny(",")
                dummy = 0
                If index = -1 Then
                    If Data = "" Then
                        Data = "0"
                    End If
                    Data = Data + ","
                    For dummy = 1 To AngkaDecimal
                        Data = Data + "0"
                    Next
                Else
                    Do While Data.Length - index - 1 < AngkaDecimal + 1
                        Data = Data + "0"
                        index = Data.IndexOfAny(",")
                    Loop

                    'potong karakter decimal
                    Data = Data.Substring(0, index + AngkaDecimal + 2)

                    'pembulatan
                    If Convert.ToInt16(Data.Substring(Data.Length - 1, 1)) >= 5 Then
                        dummy = 10 ^ AngkaDecimal
                        dummy = 1 / dummy
                        Data = Convert.ToDecimal(Data) + dummy
                    End If

                    'potong karakter decimal
                    Data = Data.Substring(0, index + AngkaDecimal + 1)
                End If
            Else
                Data = "0"
                Data = Data + ","
                For dummy = 1 To AngkaDecimal
                    Data = Data + "0"
                Next
            End If

            Return Convert.ToDecimal(Data)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini melakukan format text yang merupakan angka dalam harga rupiah
    ''' </summary>
    ''' <param name="Angka">Text yang berupa angka, jika angka maka akan error. Angka yang diterima harus
    ''' dalam format #,00</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function FormatHargaINA(ByVal Angka As String) As String
        Dim idKoma As Integer = Angka.IndexOf(",")
        Dim idTitik As Integer

        Try
            If idKoma = -1 Then
                'tambahkan koma dan digit dibelakang koma sebanyak yang diminta
                Angka = Angka + ",00"
            End If

            idKoma = Angka.IndexOf(",")
            'tambahkan titik didepan koma
            If idKoma > 3 Then
                Angka = Angka.Insert(idKoma - 3, ".")
                idTitik = Angka.IndexOf(".")

                Do While idTitik > 3
                    Angka = Angka.Insert(idTitik - 3, ".")
                    idTitik = Angka.IndexOf(".")
                Loop
            End If

            Return "Rp." + Angka
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Rp. 0,00"
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini melakukan format text yang merupakan angka dalam harga rupiah
    ''' </summary>
    ''' <param name="KoleksiTextBox">Koleksi yang berisikan text box yang textnya sudah dalam format
    ''' angka semua</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FormatHargaINA(ByVal KoleksiTextBox As Collection) As Boolean
        Try
            Dim txt As TextBox

            For Each txt In KoleksiTextBox
                txt.Text = Me.FormatHargaINA(txt.Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini mengubah format angka menjadi angka decimal dengan jumlah angka decimal sesuai
    ''' dengan yang diinginkan. Format dalam US
    ''' </summary>
    ''' <param name="KoleksiTextBox">Text box yang textnya akan diformat datanya</param>
    ''' <param name="AngkaDecimal">Jumlah angka dibelakang koma yang diinginkan</param>
    ''' <remarks></remarks>
    Public Sub FormatAngka(ByVal KoleksiTextBox As Collection, ByVal AngkaDecimal As Integer)
        Dim txtbox As TextBox

        Try
            For Each txtbox In KoleksiTextBox
                txtbox.Text = Me.FormatAngka(txtbox.Text, AngkaDecimal)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' Fungsi ini mengubah format angka menjadi angka decimal dengan jumlah angka decimal sesuai
    ''' dengan yang diinginkan. Format dalam US
    ''' </summary>
    ''' <param name="Data">Data numerik yang akan diformat, bila bukan numerik maka akan memberi
    ''' nilai balikan 0</param>
    ''' <param name="AngkaDecimal">Jumlah angka dibelakang koma yang diinginkan</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FormatAngka(ByVal Data As String, ByVal AngkaDecimal As Integer) As String
        Dim index As Integer
        Dim dummy As Decimal

        Try
            If Me.CekNumerik(Data) Then
                'cek jumlah karakter dibelakang decimal
                index = Data.IndexOfAny(".")
                dummy = 0
                If index = -1 Then
                    If Data = "" Then
                        Data = "0"
                    End If
                    Data = Data + "."
                    For dummy = 1 To AngkaDecimal
                        Data = Data + "0"
                    Next
                Else
                    Do While Data.Length - index - 1 < AngkaDecimal + 1
                        Data = Data + "0"
                        index = Data.IndexOfAny(".")
                    Loop

                    'potong karakter decimal
                    Data = Data.Substring(0, index + AngkaDecimal + 2)

                    'pembulatan
                    If Convert.ToInt16(Data.Substring(Data.Length - 1, 1)) >= 5 Then
                        dummy = 10 ^ AngkaDecimal
                        dummy = 1 / dummy
                        Data = Convert.ToDecimal(Data) + dummy
                    End If

                    'potong karakter decimal
                    Data = Data.Substring(0, index + AngkaDecimal + 1)
                End If
            Else
                Data = "0"
                Data = Data + "."
                For dummy = 1 To AngkaDecimal
                    Data = Data + "0"
                Next
            End If

            Return Data
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini melakukan format text yang merupakan angka dalam harga US
    ''' </summary>
    ''' <param name="Angka">Text yang berupa angka, jika angka maka akan error. Angka yang diterima harus
    ''' dalam format #.00</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FormatHargaUS(ByVal Angka As String, ByVal AngkaDecimal As Integer, _
    Optional ByVal TambahkanSatuan As Boolean = False, Optional ByVal Satuan As String = "Rp. ") As String
        Dim idKoma As Integer = Angka.IndexOf(".")
        Dim idTitik As Integer

        Try
            If idKoma = -1 Then
                'tambahkan koma dan digit dibelakang koma sebanyak yang diminta
                Angka = Angka + ".0"
                idKoma = Angka.IndexOf(".")
                Do While Angka.Length - idKoma <= AngkaDecimal
                    Angka = Angka + "0"
                    idKoma = Angka.IndexOf(".")
                Loop
            ElseIf Angka.Length - idKoma - 1 > AngkaDecimal Then
                'jika angka dibelakang koma lebih dari angka decimal yang diminta
                Dim dummy As Integer = CInt(Angka.Substring(Angka.Length - 1))
                If dummy >= 5 Then
                    Angka = CDbl(Angka) + 1 / 10 ^ AngkaDecimal
                End If
                Angka = Angka.Substring(0, idKoma + 1 + AngkaDecimal)
            ElseIf Angka.Length - idKoma <= AngkaDecimal Then
                'jika angka dibelakang koma kurang dari angka decimal yag diminta
                Do While Angka.Length - idKoma <= AngkaDecimal
                    Angka = Angka + "0"
                    idKoma = Angka.IndexOf(".")
                Loop
            End If

            idKoma = Angka.IndexOf(".")
            'tambahkan titik didepan koma
            If idKoma > 3 Then
                Angka = Angka.Insert(idKoma - 3, ",")
                idTitik = Angka.IndexOf(",")

                Do While idTitik > 3
                    Angka = Angka.Insert(idTitik - 3, ",")
                    idTitik = Angka.IndexOf(",")
                Loop
            End If

            If TambahkanSatuan Then
                Return Satuan + Angka
            Else
                Return Angka
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-"
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini melakukan format text yang merupakan angka dalam harga US. Semua angka dalam 
    ''' textbox yang dimasukkan harus dalam format #,###.00
    ''' </summary>
    ''' <param name="KoleksiTextBox">Koleksi yang berisikan text box yang textnya sudah dalam format
    ''' angka semua</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FormatHargaUS(ByVal KoleksiTextBox As Collection, ByVal AngkaDecimal As Integer, _
    Optional ByVal TampilkanSatuan As Boolean = False, Optional ByVal Satuan As String = "Rp. ") As Boolean
        Try
            Dim txt As TextBox

            For Each txt In KoleksiTextBox
                txt.Text = Me.FormatHargaUS(txt.Text, AngkaDecimal, TampilkanSatuan, Satuan)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini hanya khusus untuk Program Sales IWAKI
    ''' </summary>
    ''' <param name="Data"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function BuatTextFormatInt(ByVal Data As String) As String
        Dim Angka As Int64
        Dim index As Integer
        Try
            If Not Data = "" Then
                If IsNumeric(Data) = False Then
                    Data = Data.Remove(Data.Length - 1, 1)
                    If IsNumeric(Data) = False Then
                        Data = "0"
                    End If
                Else
                    If Data.Contains(",") = True Then
                        index = Data.IndexOf(",")
                        Data = Data.Remove(index, Data.Length - index)
                    End If
                    If Data.Contains(".") = True Then
                        index = Data.IndexOf(".")
                        Data = Data.Remove(index, Data.Length - index)
                    End If
                    Angka = Convert.ToInt64(Data)
                    Data = Convert.ToString(Angka)
                End If
            Else
                Data = "0"
            End If
            Return Data
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "0"
        End Try
    End Function

    Public Function BuatTextFormatDecimal(ByVal Data As String) As String
        Dim Angka As Decimal
        Try
            If Not Data = "" Then
                If IsNumeric(Data) = False Then
                    Data = Data.Remove(Data.Length - 1, 1)
                    If IsNumeric(Data) = False Then
                        Data = "0"
                    End If
                ElseIf Data.Chars(Data.Length - 1) <> "." And Data.Chars(Data.Length - 1) <> "," Then
                    Angka = Convert.ToDecimal(Data)
                    Data = Convert.ToString(Angka)
                End If
            Else
                Data = "0"
            End If
            Return Data
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "0"
        End Try
    End Function

    Public Sub SetNumber(ByVal dgv As DataGridView, Optional ByVal vSetAutosize As Boolean = True)
        Try
            If vSetAutosize = True Then
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            End If

            For i As Integer = 1 To dgv.Rows.Count
                dgv.Item(0, i - 1).Value = i
            Next

            If vSetAutosize = True Then
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    ''' <summary>
    ''' Fungsi ini melakukan digit grouping seperti pada kalkulator, dengan angka decimal
    ''' </summary>
    ''' <param name="Par"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DigitGrouping(ByVal Par As String, ByVal Dec As Integer, _
    Optional ByVal vEnumTipeRounding As EnumTipeRounding = EnumTipeRounding.Middlepoint) As String
        Dim RoundingBaseNumber As Decimal = 1
        Try
            Dim Isnegatif As Boolean = False

            Dim Par1 As String = Par
            If Par1 <> "" Then
                If CDbl(Par) < 0 Then
                    Isnegatif = True
                    Par = Math.Abs(CDbl(Par))
                    Par1 = Par
                End If

                'lakukan pembulatan
                If vEnumTipeRounding = EnumTipeRounding.Middlepoint Then
                    Par1 = Math.Round(CDec(Par1), Dec, MidpointRounding.AwayFromZero).ToString
                ElseIf vEnumTipeRounding = EnumTipeRounding.Down Then
                    RoundingBaseNumber = 1 / (10 ^ Dec)
                    Par1 = Math.Floor(Par1 / RoundingBaseNumber) * RoundingBaseNumber
                ElseIf vEnumTipeRounding = EnumTipeRounding.Up Then
                    RoundingBaseNumber = 1 / (10 ^ Dec)
                    Par1 = Math.Ceiling(Par1 / RoundingBaseNumber) * RoundingBaseNumber
                End If
                'tambahkan nol jika perlu
                If Dec <> 0 Then
                    If Par1.IndexOf(".") = -1 Then
                        'nol yang pertama
                        Par1 = Par1 + ".0"
                    End If
                    While Par1.Length - Par1.LastIndexOf(".") - 1 < Dec
                        Par1 = Par1 + "0"
                    End While
                End If

                'hapuskan koma yang ada
                While Par1.Contains(",") = True
                    Par1 = Par1.Remove(Par1.IndexOf(","), 1)
                End While

                'digit grouping
                Dim Start As Integer = 9
                Dim Finish As Boolean = False

                'tambahkan koma
                While Start - 3 > 0
                    'identifikasi titik
                    If Par1.IndexOf(".") <> -1 Then
                        Start = Par1.IndexOf(".")
                    Else
                        Start = Par1.Length
                    End If

                    If Par1.Contains(",") = True Then
                        Start = Par1.IndexOf(",")
                    End If

                    If Start - 3 > 0 Then
                        If Par1.Substring(Start - 3) <> "," Then
                            Par1 = Par1.Insert(Start - 3, ",")
                            Start -= 3
                        Else
                            Start -= 3
                        End If
                    Else
                        Exit While
                    End If
                End While
            Else
                Par1 = "0"
            End If

            If Isnegatif = True Then
                Par1 = "-" + Par1
            End If


            Return Par1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "0"
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini melakukan digit grouping seperti pada kalkulator
    ''' </summary>
    ''' <param name="Par"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DigitGrouping(ByVal Par As String) As String
        Try
            Dim Isnegatif As Boolean = False

            Dim Par1 As String = Par
            If Par1 <> "" Then
                If CDbl(Par) < 0 Then
                    Isnegatif = True
                    Par = Math.Abs(CDbl(Par))
                    Par1 = Par
                End If

                'hapuskan koma yang ada
                While Par1.Contains(",") = True
                    Par1 = Par1.Remove(Par1.IndexOf(","), 1)
                End While

                'digit grouping
                Dim Start As Integer = 9
                Dim Finish As Boolean = False

                'tambahkan koma
                While Start - 3 > 0
                    'identifikasi titik
                    If Par1.IndexOf(".") <> -1 Then
                        Start = Par1.IndexOf(".")
                    Else
                        Start = Par1.Length
                    End If

                    If Par1.Contains(",") = True Then
                        Start = Par1.IndexOf(",")
                    End If

                    If Start - 3 > 0 Then
                        If Par1.Substring(Start - 3) <> "," Then
                            Par1 = Par1.Insert(Start - 3, ",")
                            Start -= 3
                        Else
                            Start -= 3
                        End If
                    Else
                        Exit While
                    End If
                End While
            Else
                Par1 = "0"
            End If

            If Isnegatif = True Then
                Par1 = "-" + Par1
            End If
            Return Par1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "0"
        End Try
    End Function

    Public Function SetDecimal(ByVal Number As String, ByVal Dec As Integer) As String
        Try
            'pastikan paramter yang diberikan benar
            '1. angka no negatif 
            If Me.CekNonNegatif(Dec) = False Then
                Return ""
            End If
            If Me.CekNonNegatif(Number) = False Then
                Return ""
            End If
            'lakukan pembulatan
            Number = Math.Round(CDbl(Number), Dec, MidpointRounding.AwayFromZero)
            'tambahkan nol jika perlu
            If Dec <> 0 Then
                If Number.IndexOf(".") = -1 Then
                    'nol yang pertama
                    Number = Number + ".0"
                End If
                While Number.Length - Number.LastIndexOf(".") - 1 < Dec
                    Number = Number + "0"
                End While
            End If

            Return Number
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    'Fungsi Number2x
    Public Sub Number_Click(ByVal TxtNumber As System.Windows.Forms.TextBox)
        Try
            TxtNumber.DeselectAll()
            TxtNumber.SelectionStart = TxtNumber.Text.Length
            TxtNumber.SelectionLength = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Number_Enter(ByVal txtNumber As TextBox, ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If txtNumber.Enabled = True And txtNumber.ReadOnly = False Then
                If IsNumeric(txtNumber.Text) = False Then
                    txtNumber.Text = "0"
                End If
                If CDec(txtNumber.Text) = 0 Then
                    txtNumber.Text = "0"
                End If
                While txtNumber.Text.Contains(".") = True
                    If txtNumber.Text.EndsWith("0") Or txtNumber.Text.EndsWith(".") Then
                        txtNumber.Text = txtNumber.Text.Remove(txtNumber.Text.Length - 1, 1)
                    Else
                        Exit While
                    End If
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Number_Leave(ByVal txtNumber As TextBox, ByVal sender As Object, ByVal e As System.EventArgs, _
    ByVal vDef_Prec As Integer, ByVal vObjCD As Core_Module.ClsCekData)
        Try
            If txtNumber.Enabled = True And txtNumber.ReadOnly = False Then
                'digit grouping
                txtNumber.Text = vObjCD.DigitGrouping(txtNumber.Text, vDef_Prec)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Number_KeyDown(ByVal txtNumber As TextBox, ByVal sender As Object, _
    ByVal e As System.Windows.Forms.KeyEventArgs, Optional ByVal Max As Double = 0)
        Try
            If txtNumber.Text = "" Then
                txtNumber.Text = "0"
            End If
            If txtNumber.ReadOnly = True Then
                e.SuppressKeyPress = True
                Exit Sub
            End If

            If e.Alt = False Then
                If e.KeyData = Keys.OemPeriod Or e.KeyData = Keys.Decimal Then
                    If txtNumber.Text.Contains(".") = True Then
                        e.SuppressKeyPress = True
                    End If
                Else
                    If txtNumber.Text = "0" Then
                        txtNumber.Text = ""
                    End If

                    If e.KeyValue = 9 Then 'tab button
                        If IsNothing(txtNumber.GetNextControl(txtNumber, True)) = False Then
                            If txtNumber.GetNextControl(txtNumber, True).CanFocus = True Then
                                txtNumber.GetNextControl(txtNumber, True).Focus()
                                Exit Sub
                            End If
                        End If
                    ElseIf e.KeyValue = 46 Then 'delete = clear
                        txtNumber.Text = "0"
                        Exit Sub
                    ElseIf e.KeyValue = 8 Then 'backspace
                        Exit Sub
                    ElseIf e.KeyValue = 189 Or e.KeyValue = 109 Then
                        If IsNumeric(txtNumber.Text) = True Then
                            If txtNumber.Text <> "0" Then
                                If CDbl(txtNumber.Text) < 0 Then
                                    'kalo negatif, buang minus
                                    txtNumber.Text = txtNumber.Text.Substring(1)
                                Else
                                    'kalo poisitif, tambah minus
                                    txtNumber.Text = "-" + txtNumber.Text
                                End If
                            End If
                        End If
                    End If

                    'pembatasan jumlah maksimum
                    If txtNumber.Text.Length >= txtNumber.MaxLength Then
                        e.SuppressKeyPress = True
                        Exit Sub
                    End If
                    'pembatasan bilangan decimal maksimum
                    If txtNumber.Text.Length - txtNumber.Text.LastIndexOf(".") - 1 = 6 _
                    And txtNumber.Text.LastIndexOf(".") <> -1 Then
                        e.SuppressKeyPress = True
                        Exit Sub
                    End If

                    If e.KeyCode = Keys.NumPad0 Then
                        e.SuppressKeyPress = True
                        If Max <> 0 And CDbl(txtNumber.Text + "0") > Max Then
                            Exit Sub
                        Else
                            txtNumber.Text = txtNumber.Text + "0"
                        End If
                    ElseIf e.KeyCode = Keys.NumPad1 Then
                        e.SuppressKeyPress = True
                        If Max <> 0 And CDbl(txtNumber.Text + "1") > Max Then
                            Exit Sub
                        Else
                            txtNumber.Text = txtNumber.Text + "1"
                        End If
                    ElseIf e.KeyCode = Keys.NumPad2 Then
                        e.SuppressKeyPress = True
                        If Max <> 0 And CDbl(txtNumber.Text + "2") > Max Then
                            Exit Sub
                        Else
                            txtNumber.Text = txtNumber.Text + "2"
                        End If
                    ElseIf e.KeyCode = Keys.NumPad3 Then
                        e.SuppressKeyPress = True
                        If Max <> 0 And CDbl(txtNumber.Text + "3") > Max Then
                            Exit Sub
                        Else
                            txtNumber.Text = txtNumber.Text + "3"
                        End If
                    ElseIf e.KeyCode = Keys.NumPad4 Then
                        e.SuppressKeyPress = True
                        If Max <> 0 And CDbl(txtNumber.Text + "4") > Max Then
                            Exit Sub
                        Else
                            txtNumber.Text = txtNumber.Text + "4"
                        End If
                    ElseIf e.KeyCode = Keys.NumPad5 Then
                        e.SuppressKeyPress = True
                        If Max <> 0 And CDbl(txtNumber.Text + "5") > Max Then
                            Exit Sub
                        Else
                            txtNumber.Text = txtNumber.Text + "5"
                        End If
                    ElseIf e.KeyCode = Keys.NumPad6 Then
                        e.SuppressKeyPress = True
                        If Max <> 0 And CDbl(txtNumber.Text + "6") > Max Then
                            Exit Sub
                        Else
                            txtNumber.Text = txtNumber.Text + "6"
                        End If
                    ElseIf e.KeyCode = Keys.NumPad7 Then
                        e.SuppressKeyPress = True
                        If Max <> 0 And CDbl(txtNumber.Text + "7") > Max Then
                            Exit Sub
                        Else
                            txtNumber.Text = txtNumber.Text + "7"
                        End If
                    ElseIf e.KeyCode = Keys.NumPad8 Then
                        e.SuppressKeyPress = True
                        If Max <> 0 And CDbl(txtNumber.Text + "8") > Max Then
                            Exit Sub
                        Else
                            txtNumber.Text = txtNumber.Text + "8"
                        End If
                    ElseIf e.KeyCode = Keys.NumPad9 Then
                        e.SuppressKeyPress = True
                        If Max <> 0 And CDbl(txtNumber.Text + "9") > Max Then
                            Exit Sub
                        Else
                            txtNumber.Text = txtNumber.Text + "9"
                        End If
                    ElseIf Char.IsDigit(Chr(e.KeyValue)) = True Then
                        'masukin langsung
                        If Max <> 0 And CDbl(txtNumber.Text + Chr(e.KeyValue)) > Max Then
                            e.SuppressKeyPress = True
                            Exit Sub
                        End If
                    Else
                        'yang lain tidak dianggap
                        e.SuppressKeyPress = True
                        If txtNumber.Text = "" Then
                            txtNumber.Text = "0"
                        End If
                    End If
                End If
            Else
                e.SuppressKeyPress = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Number_KeyUp(ByVal txtNumber As TextBox, ByVal sender As Object, _
    ByVal e As System.Windows.Forms.KeyEventArgs)
        Try
            Dim Isnegatif As Boolean = False

            If txtNumber.ReadOnly = True Then
                Exit Sub
            End If

            If e.Alt = False Then
                If txtNumber.Text <> "" And txtNumber.Text <> "-" And IsNumeric(txtNumber.Text) = True Then
                    If CDbl(txtNumber.Text) < 0 Then
                        Isnegatif = True
                        txtNumber.Text = txtNumber.Text.Substring(1)
                    End If

                    'hapuskan koma yang ada
                    While txtNumber.Text.Contains(",") = True
                        txtNumber.Text = txtNumber.Text.Remove(txtNumber.Text.IndexOf(","), 1)
                    End While

                    'digit grouping
                    Dim Start As Integer = 9
                    Dim Finish As Boolean = False

                    'tambahkan koma
                    While Start - 3 > 0
                        'identifikasi titik
                        If txtNumber.Text.IndexOf(".") <> -1 Then
                            Start = txtNumber.Text.IndexOf(".")
                        Else
                            Start = txtNumber.Text.Length
                        End If

                        If txtNumber.Text.Contains(",") = True Then
                            Start = txtNumber.Text.IndexOf(",")
                        End If

                        If Start - 3 > 0 Then
                            If txtNumber.Text.Substring(Start - 3) <> "," Then
                                txtNumber.Text = txtNumber.Text.Insert(Start - 3, ",")
                                Start -= 3
                            Else
                                Start -= 3
                            End If
                        Else
                            Exit While
                        End If
                    End While
                Else
                    txtNumber.Text = "0"
                End If

                If Isnegatif = True Then
                    txtNumber.Text = "-" + txtNumber.Text
                End If

                'geser kursor ke paling kanan
                txtNumber.DeselectAll()
                txtNumber.SelectionStart = txtNumber.Text.Length
                txtNumber.SelectionLength = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
