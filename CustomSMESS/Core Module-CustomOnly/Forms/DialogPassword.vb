Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class DialogPassword
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If ModUser.Login(Me.txtUserName, Me.txtPassword) = True Then
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Hide()
        Else
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Hide()
    End Sub
End Class
