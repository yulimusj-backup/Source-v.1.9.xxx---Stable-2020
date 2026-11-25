Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing

Public Class FrmTrans_Gaji_MsgBox
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_MsgBox_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
            If Me.dgvItem.Rows.Count > 0 Then
                rowDGV = Me.dgvItem.Rows(0)
                If IsNothing(rowDGV.Cells(1).Value) = True Then
                    Me.dgvItem.Columns(1).Visible = False
                Else
                    If IsDBNull(rowDGV.Cells(1).Value) = True Then
                        Me.dgvItem.Columns(1).Visible = False
                    ElseIf rowDGV.Cells(1).Value = "" Then
                        Me.dgvItem.Columns(1).Visible = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class