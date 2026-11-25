Option Explicit On
Option Strict Off

Public Class FrmCariMasterUser
    Public Kodeuser As String = ""
    Public backColorAll As Color = Core_Module.GetSMESSBackColorAll
    Private view As DataView

    Private Sub RBDenganID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBDenganID.CheckedChanged
        CekPencarian()
    End Sub

    Private Sub CekPencarian()
        Try
            ResetField()
            If Me.RBDenganID.Checked = True Then
                Me.TxtCari.CharacterCasing = CharacterCasing.Upper
                Me.LblTxtCari.Text = "User Code"
            Else
                Me.TxtCari.CharacterCasing = CharacterCasing.Normal
                Me.LblTxtCari.Text = "User Name"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ResetField()
        Dim row As DataGridViewRow
        For Each row In Me.DGHasil.Rows
            Me.DGHasil.Rows.Remove(row)
        Next
        For Each row In Me.DGHasil.SelectedRows
            Me.DGHasil.Rows.Remove(row)
        Next
        Me.TxtCari.Clear()
    End Sub

    Private Sub RBDenganNama_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBDenganNama.CheckedChanged
        CekPencarian()
    End Sub

    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click
        ModUser.IsiDataGridView(Me.DGHasil, view, Me.TxtCari.Text, Me.RBDenganID.Checked)
    End Sub

    Private Sub DGHasil_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGHasil.CellContentDoubleClick
        Me.Pilih()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Try
            Pilih()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cmdKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKeluar.Click
        Me.Kodeuser = ""
        Me.Close()
    End Sub

    Private Sub Pilih()
        If Me.DGHasil.SelectedRows.Count > 0 Then
            Me.Kodeuser = Me.DGHasil.SelectedRows.Item(0).Cells(0).Value
        End If
        Me.Close()
    End Sub

    Private Sub TxtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtCari.KeyDown
        If e.KeyValue = 13 Then
            Me.cmdSearch.PerformClick()
        ElseIf e.KeyValue = 40 Then
            If Me.DGHasil.RowCount > 0 Then
                Me.DGHasil.Focus()
            End If
        End If
    End Sub

    Private Sub DGHasil_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DGHasil.KeyDown
        If e.KeyValue = 13 Then
            Pilih()
        End If
    End Sub

    Private Sub TxtCari_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCari.GotFocus
        TxtCari.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub TxtCari_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCari.LostFocus
        TxtCari.BackColor = Me.txtPrimaryGroup.BackColor
    End Sub

    Private Sub FrmCariMasterUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ShowIcon = False
        Me.DGHasil.AlternatingRowsDefaultCellStyle.BackColor = Me.backColorAll
        If Me.TxtCari.Text <> "" Then
            Me.cmdSearch.PerformClick()
        End If
        Me.TxtCari.Focus()
    End Sub
End Class