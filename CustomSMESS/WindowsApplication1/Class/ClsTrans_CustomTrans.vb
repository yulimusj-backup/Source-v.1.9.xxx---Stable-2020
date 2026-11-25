Imports System.Windows.Forms

''' <summary>
''' kelas ini mewakili custom transaction dan custom master data
''' </summary>
''' <remarks></remarks>
Public Class ClsTrans_CustomTrans
    Inherits ToolStripMenuItem
    Private mParentForm As Windows.Forms.Form
    Private mLoginForm As Windows.Forms.Form
    Private mUserCode As String = ""

    Public Enum EnumTipeMenu As Integer
        Master
        Transaction
    End Enum

    Public Sub IsikanMenuCustom(ByVal vParentMenu As ToolStripMenuItem, ByVal vParentForm As Windows.Forms.Form, _
    ByVal vLoginForm As Windows.Forms.Form, ByVal vTipeMenu As EnumTipeMenu, ByVal vUserCode As String)
        Dim MI As ClsTrans_CustomTrans
        Dim row As DseCustomMenu.tbltrans_customtransRow
        Dim TBL As New DseCustomMenu.tbltrans_customtransDataTable
        Dim DTA As New DseCustomMenuTableAdapters.tbltrans_customtransTableAdapter

        Try
            'kosongkan menu anak-nya
            vParentMenu.DropDownItems.Clear()

            'tambahkan menu anak
            DTA.FillBy_Tipe(TBL, vTipeMenu)
            For Each row In TBL.Rows
                MI = New ClsTrans_CustomTrans(vUserCode)
                MI.Text = row.menu_text
                MI.Name = row.menu_name + "1"
                MI.mParentForm = vParentForm
                MI.mLoginForm = vLoginForm

                vParentMenu.DropDownItems.Add(MI)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Private Sub ClsTrans_CustomTrans_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        Try
            'isikan tergantung untuk masing2 customer, diisikan dengan coding yang membuka form ybs, lihat custom gats sebagai referensi

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub New(ByVal vUserCode As String)
        Try
            Me.mUserCode = vUserCode
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
