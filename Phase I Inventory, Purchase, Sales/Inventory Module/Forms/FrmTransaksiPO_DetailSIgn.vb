Option Explicit On
Option Strict Off

Public Class FrmTransaksiPO_DetailSIgn
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub TxtPreparedBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPreparedBy.GotFocus
        Try
            If Me.TxtPreparedBy.ReadOnly = False Then
                Me.TxtPreparedBy.BackColor = Me.BackColorAll
            Else
                Me.TxtPreparedBy.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPreparedBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPreparedBy.LostFocus
        Try
            If Me.TxtPreparedBy.ReadOnly = False Then
                Me.TxtPreparedBy.BackColor = Me.TxtPrimary.BackColor
            Else
                Me.TxtPreparedBy.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPreparedBy_Jabatan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPreparedBy_Jabatan.GotFocus
        Try
            If Me.TxtPreparedBy_Jabatan.ReadOnly = False Then
                Me.TxtPreparedBy_Jabatan.BackColor = Me.BackColorAll
            Else
                Me.TxtPreparedBy_Jabatan.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPreparedBy_Jabatan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPreparedBy_Jabatan.LostFocus
        Try
            If Me.TxtPreparedBy_Jabatan.ReadOnly = False Then
                Me.TxtPreparedBy_Jabatan.BackColor = Me.TxtPrimary.BackColor
            Else
                Me.TxtPreparedBy_Jabatan.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtApprovedBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtApprovedBy.GotFocus
        Try
            If Me.TxtApprovedBy.ReadOnly = False Then
                Me.TxtApprovedBy.BackColor = Me.BackColorAll
            Else
                Me.TxtApprovedBy.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtApprovedBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtApprovedBy.LostFocus
        Try
            If Me.TxtApprovedBy.ReadOnly = False Then
                Me.TxtApprovedBy.BackColor = Me.TxtPrimary.BackColor
            Else
                Me.TxtApprovedBy.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtApprovedBy_Jabatan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtApprovedBy_Jabatan.GotFocus
        Try
            If Me.TxtApprovedBy_Jabatan.ReadOnly = False Then
                Me.TxtApprovedBy_Jabatan.BackColor = Me.BackColorAll
            Else
                Me.TxtApprovedBy_Jabatan.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtApprovedBy_Jabatan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtApprovedBy_Jabatan.LostFocus
        Try
            If Me.TxtApprovedBy_Jabatan.ReadOnly = False Then
                Me.TxtApprovedBy_Jabatan.BackColor = Me.TxtPrimary.BackColor
            Else
                Me.TxtApprovedBy_Jabatan.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAcknowledgeBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAcknowledgeBy.GotFocus
        Try
            If Me.TxtAcknowledgeBy.ReadOnly = False Then
                Me.TxtAcknowledgeBy.BackColor = Me.BackColorAll
            Else
                Me.TxtAcknowledgeBy.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAcknowledgeBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAcknowledgeBy.LostFocus
        Try
            If Me.TxtAcknowledgeBy.ReadOnly = False Then
                Me.TxtAcknowledgeBy.BackColor = Me.TxtPrimary.BackColor
            Else
                Me.TxtAcknowledgeBy.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAcknowledgeBy_Jabatan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAcknowledgeBy_Jabatan.GotFocus
        Try
            If Me.TxtAcknowledgeBy_Jabatan.ReadOnly = False Then
                Me.TxtAcknowledgeBy_Jabatan.BackColor = Me.BackColorAll
            Else
                Me.TxtAcknowledgeBy_Jabatan.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAcknowledgeBy_Jabatan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAcknowledgeBy_Jabatan.LostFocus
        Try
            If Me.TxtAcknowledgeBy_Jabatan.ReadOnly = False Then
                Me.TxtAcknowledgeBy_Jabatan.BackColor = Me.TxtPrimary.BackColor
            Else
                Me.TxtAcknowledgeBy_Jabatan.BackColor = Me.TxtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class