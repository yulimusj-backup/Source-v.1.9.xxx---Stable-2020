Option Explicit On
Option Strict Off

Public Class ClsDGVCustom
    Inherits DataGridView

    Private Sub ClsDGVCustom_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles Me.CellBeginEdit
        Try
            If e.ColumnIndex = 0 Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ClsDGVCustom_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles Me.RowsAdded
        Dim obj As New Core_Module.ClsCekData
        Try
            obj.SetNumber(Me, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    Private Sub ClsDGVCustom_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles Me.RowsRemoved
        Dim obj As New Core_Module.ClsCekData
        Try
            obj.SetNumber(Me, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    Private Sub ClsDGVCustom_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Sorted
        Dim obj As New Core_Module.ClsCekData
        Try
            obj.SetNumber(Me, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    Public Sub New()
        Try
            Me.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
