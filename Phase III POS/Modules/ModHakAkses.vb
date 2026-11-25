Option Explicit On
Option Strict Off
Imports Core_Module

Public Module ModHakAkses
    ''Module ini di gunakan pada Frm Management hak akses
    Public Function CariDetail(ByVal Obj As ClsHakAkses) As Boolean
        Try
            Return Obj.CariDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function IsiListBox(ByVal Kode As String, ByVal LB As ListBox) As Boolean
        Dim objHakAkses As ClsHakAkses = New ClsHakAkses
        Try
            Return objHakAkses.IsiListBox(Kode, LB)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objHakAkses = Nothing
        End Try
    End Function

    Public Function IsiListBox(ByVal LB As ListBox) As Boolean
        Dim objHakAkses As ClsHakAkses = New ClsHakAkses
        Try
            Return objHakAkses.IsiListBox(LB)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objHakAkses = Nothing
        End Try
    End Function

    Public Function Simpan(ByVal objhakakses As ClsHakAkses, ByVal txtKodehakakses As TextBox, ByVal txtnamahakakses As TextBox, _
  ByVal Deskripsi As TextBox) As Boolean
        Try
            Return objhakakses.Simpan(txtKodehakakses)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal objhakakses As ClsHakAkses, ByVal txtKodehakakses As TextBox, ByVal txtnamahakakses As TextBox, _
  ByVal Deskripsi As TextBox) As Boolean
        Try
            Return objhakakses.SimpanPerubahan()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function Hapus(ByVal KodeHakAkses As String) As Boolean
        Dim obj As New ClsHakAkses

        Try
            Return obj.Hapus(KodeHakAkses)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal DGV As DataGridView, ByVal View As DataView, _
    ByVal Text As String, ByVal DenganID As Boolean) As Boolean
        Dim obj As New ClsHakAkses

        Try
            Return obj.IsiDGV(DGV, View, Text, DenganID)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Module
