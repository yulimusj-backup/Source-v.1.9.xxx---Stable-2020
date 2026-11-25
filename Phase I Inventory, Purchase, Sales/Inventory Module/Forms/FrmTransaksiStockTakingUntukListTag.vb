Option Explicit On
Option Strict Off

Public Class FrmTransaksiStockTakingUntukListTag

    Private Sub btn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1.Click
        Try
            If Me.SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                If Me.SaveFileDialog1.FileName <> "" Then
                    Dim I As String
                    Dim IsiNotePad As String = ""
                    For Each I In Me.lboItems.Items
                        IsiNotePad = IsiNotePad + I + vbCrLf
                    Next
                    My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, IsiNotePad, False)
                    MsgBox("Export to Notepad completed.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub
End Class