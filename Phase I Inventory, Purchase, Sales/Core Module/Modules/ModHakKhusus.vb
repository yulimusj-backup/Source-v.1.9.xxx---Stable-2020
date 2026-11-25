Public Module ModHakKhusus
    ''Module ini di gunakan pada Frm Management hak akses
    Public Function CariDetail(ByVal Obj As ClsHakKhusus) As Boolean
        Try
            Return Obj.CariDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function IsiListBox(ByVal Kode As String, ByVal LB As ListBox) As Boolean
        Dim objhakkhusus As ClsHakKhusus = New ClsHakKhusus
        Try
            Return objhakkhusus.IsiListBox(Kode, LB)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objhakkhusus = Nothing
        End Try
    End Function

    Public Function IsiListBox(ByVal LB As ListBox) As Boolean
        Dim objhakkhusus As ClsHakKhusus = New ClsHakKhusus
        Try
            Return objhakkhusus.IsiListBox(LB)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objhakkhusus = Nothing
        End Try
    End Function

    Public Function Simpan(ByVal objhakkhusus As ClsHakKhusus, ByVal txtKodehakkhusus As TextBox, ByVal txtnamahakkhusus As TextBox, _
  ByVal Deskripsi As TextBox) As Boolean
        Try
            Return objhakkhusus.Simpan(txtKodehakkhusus)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal objhakkhusus As ClsHakKhusus, ByVal txtKodehakkhusus As TextBox, ByVal txtnamahakkhusus As TextBox, _
  ByVal Deskripsi As TextBox) As Boolean
        Try
            Return objhakkhusus.SimpanPerubahan()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function Hapus(ByVal Kodehakkhusus As String) As Boolean
        Dim obj As New ClsHakKhusus

        Try
            Return obj.Hapus(Kodehakkhusus)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal DGV As DataGridView, ByVal View As DataView, _
    ByVal Text As String, ByVal DenganID As Boolean) As Boolean
        Dim obj As New ClsHakKhusus

        Try
            Return obj.IsiDGV(DGV, View, Text, DenganID)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Module
