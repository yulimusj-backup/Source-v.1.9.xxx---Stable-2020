Public Class FrmUpdate_ClearTabel_new
    Dim DTA As New DseUpdate_ClearTableTableAdapters.tbltrans_clear_dataTableAdapter
    Public vBaru As Boolean = True
    Public vPrimaryAcuan As Double = -1

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim Jml As Nullable(Of Long)
        Dim Primary1 As Nullable(Of Double)

        Try
            Me.TextBox1.Text = Me.TextBox1.Text.Trim
            If Me.TextBox1.Text <> "" Then
                'cek apakah sudah ada
                Jml = DTA.ScalarQuery_GetJml(Me.TextBox1.Text)
                If Me.vBaru = True Then
                    'tidak boleh ada di dalam database sama sekali
                    If Jml.HasValue = True Then
                        If Jml.Value > 0 Then
                            MsgBox("Scheme description duplication", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If
                Else
                    'jika ada di dalam database, maka primary acuannya tidak boleh berbeda
                    Primary1 = DTA.ScalarQuery_GetPrimary(Me.TextBox1.Text)
                    If Primary1.HasValue = True Then
                        If Primary1.Value <> Me.vPrimaryAcuan Then
                            MsgBox("Scheme description duplication", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If
                End If
            Else
                MsgBox("Please specify scheme name", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'kalo sampe sini maka close dan dialog result = ok
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class