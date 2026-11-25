Public Module ModUser
    Public Function IsiComboBox(ByVal CBO As ComboBox) As Boolean
        Dim obj As New ClsUser

        Try
            Return obj.IsiComboBox(CBO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Function
    Public Function Login(ByVal TxtKodeuser As TextBox, ByVal TxtPassword As TextBox) As Boolean
        Dim objuser As ClsUser = New ClsUser
        Try
            Return objuser.Login(TxtKodeuser.Text, TxtPassword.Text, TxtKodeuser, TxtPassword)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objuser = Nothing
        End Try
    End Function

    Public Function GetNama(ByVal Kodeuser As String) As String
        Dim objuser As ClsUser = New ClsUser
        Try
            Return objuser.GetNama(Kodeuser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            objuser = Nothing
        End Try
    End Function

    Public Function GantiPassword(ByVal Txtkodeuser As TextBox, ByVal TxtPasswordLama As TextBox, ByVal TxtPasswordBaru As TextBox, _
    ByVal TxtKonfirmasiPassword As TextBox, ByVal KodeUser As String) As Boolean
        Dim objuser As ClsUser = New ClsUser
        Try
            Return objuser.SetPassword(Txtkodeuser.Text, TxtPasswordLama, TxtPasswordBaru, _
            TxtKonfirmasiPassword)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objuser = Nothing
        End Try
    End Function

    Public Function Simpan(ByVal objuser As ClsUser, ByVal txtKodeuser As TextBox, ByVal txtnamauser As TextBox, _
    ByVal Password As TextBox, ByVal Keterangan As TextBox, ByVal txtPrimaryGroup As Double) As Boolean
        Try
            Return objuser.Simpan(txtKodeuser, txtnamauser, Password, _
            objuser.Keterangan, CDbl(txtPrimaryGroup))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal objuser As ClsUser) As Boolean
        Try
            If objuser.UseTempPass = True Then
                Return objuser.SimpanPerubahan(objuser.KodeUser, objuser.NamaUser, objuser.Password, _
                objuser.Keterangan, objuser.PrimaryGroup, 1)
            Else
                Return objuser.SimpanPerubahan(objuser.KodeUser, objuser.NamaUser, objuser.Password, _
                objuser.Keterangan, objuser.PrimaryGroup, 0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function Hapus(ByVal Kodeuser As String) As Boolean
        Dim objuser As ClsUser = New ClsUser
        Try
            Return objuser.Hapus(Kodeuser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objuser = Nothing
        End Try
    End Function

    Public Function CariDetail(ByVal Kodeuser As String, ByVal Objuser As ClsUser) As Boolean
        Try
            Return Objuser.CariDetail(Kodeuser, Objuser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function UpdateHakAkses(ByVal Objuser As ClsUser, ByVal KoleksiHakAkses As Collection) As Boolean
        Try
            Return Objuser.SimpanHakAkses(Objuser.KodeUser, KoleksiHakAkses)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
    Public Function UpdateHakKhusus(ByVal Objuser As ClsUser, ByVal KoleksiHakKhusus As Collection) As Boolean
        Try
            Return Objuser.SimpanHakKhusus(Objuser.KodeUser, KoleksiHakKhusus)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function IsiDataGridView(ByVal Dgv As DataGridView, ByVal View As DataView, _
    ByVal TextCari As String, ByVal BerdasarkanID As Boolean) As Boolean
        Dim objuser As ClsUser = New ClsUser
        Try
            Return objuser.IsiDGV(Dgv, View, TextCari, BerdasarkanID)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objuser = Nothing
        End Try
    End Function

    Public Function ApakahPunyaHakKhusus(ByVal KodeUser As String, ByVal NamaHakKhusus As String) As Boolean
        Dim obj As New ClsUser
        Try
            Return obj.GetHakKhusus(KodeUser, NamaHakKhusus)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Function
End Module
