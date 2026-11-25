Option Explicit On
Option Strict Off

Partial Public Class FrmTools_MonthlyClosing
    Private Sub T4_BtnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles T4_BtnFilter.Click
        Try
            If Me.T4_TxtPrimaryClosing.Text <> "-1" Then

            Else
                MsgBox("Please specify a valid monthly closing primary")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
