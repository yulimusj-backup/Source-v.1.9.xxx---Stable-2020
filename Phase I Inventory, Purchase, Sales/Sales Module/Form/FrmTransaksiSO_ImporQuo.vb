Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class FrmTransaksiSO_ImporQuo

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub rboHeaderOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboHeaderOnly.CheckedChanged
        Try
            Me.SetGboitem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboItemOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboItemOnly.CheckedChanged
        Try
            SetGboitem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboAll.CheckedChanged
        Try
            SetGboitem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGboitem()
        Try
            If Me.rboAll.Checked = True Or Me.rboItemOnly.Checked = True Then
                Me.gboItemMode.Enabled = True
            Else
                Me.gboItemMode.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
