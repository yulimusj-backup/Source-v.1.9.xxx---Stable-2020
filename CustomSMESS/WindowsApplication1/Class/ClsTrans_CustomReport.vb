Imports System.Windows.Forms

''' <summary>
''' kelas ini mewakili custom report yang dibuatkan untuk user tertentu saja
''' </summary>
''' <remarks></remarks>
Public Class ClsTrans_CustomReport
    Inherits ToolStripMenuItem
    Private mParentForm As Windows.Forms.Form
    Private mLoginForm As Windows.Forms.Form
    Private mUserCode As String = ""

    Public Sub IsikanMenuCustom(ByVal vParentMenu As ToolStripMenuItem, ByVal vParentForm As Windows.Forms.Form, ByVal vLoginForm As Windows.Forms.Form)
        Dim MI As ClsTrans_CustomReport
        Dim row As DseCustomMenu.tbltrans_customreportRow
        Dim TBL As New DseCustomMenu.tbltrans_customreportDataTable
        Dim DTA As New DseCustomMenuTableAdapters.tbltrans_customreportTableAdapter

        Try
            'kosongkan menu anak-nya
            vParentMenu.DropDownItems.Clear()

            'tambahkan menu anak
            DTA.Fill(TBL)
            For Each row In TBL.Rows
                MI = New ClsTrans_CustomReport(Me.mUserCode)
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

    Private Sub ClsTrans_CustomReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        Try
            'isikan tergantung untuk masing2 customer
           
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub New(ByVal vUserCode As String)
        Me.mUserCode = vUserCode
    End Sub
End Class
