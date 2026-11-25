Option Explicit On
Option Strict Off

Public Class FrmTransaksiSettingPO
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim Ask_Before_Closing As Boolean = True

    Private Sub FrmTransaksiSettingPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim hasil As MsgBoxResult

        Try
            If Me.Ask_Before_Closing = True Then
                hasil = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    My.Settings.PO_Prepared_By = Me.TxtPreparedBy.Text
                    My.Settings.PO_Prepared_Position = Me.TxtPreparedBy_Jabatan.Text
                    My.Settings.PO_Acknowled_By = Me.TxtAcknowledgeBy.Text
                    My.Settings.PO_Acknowledge_Position = Me.TxtAcknowledgeBy_Jabatan.Text
                    My.Settings.PO_Approved_By = Me.TxtApprovedBy.Text
                    My.Settings.PO_Approved_Position = Me.TxtApprovedBy_Jabatan.Text

                    My.Settings.Save()

                    'update data settings po pada inventory module juga
                    Inventory_Module.UpdateSettingPO(Me.TxtPreparedBy.Text, Me.TxtPreparedBy_Jabatan.Text, Me.TxtAcknowledgeBy.Text, _
                    Me.TxtAcknowledgeBy_Jabatan.Text, Me.TxtApprovedBy.Text, Me.TxtApprovedBy_Jabatan.Text)
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSettingPO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.TxtPreparedBy.Text = My.Settings.PO_Prepared_By
            Me.TxtPreparedBy_Jabatan.Text = My.Settings.PO_Prepared_Position
            Me.TxtAcknowledgeBy.Text = My.Settings.PO_Acknowled_By
            Me.TxtAcknowledgeBy_Jabatan.Text = My.Settings.PO_Acknowledge_Position
            Me.TxtApprovedBy.Text = My.Settings.PO_Approved_By
            Me.TxtApprovedBy_Jabatan.Text = My.Settings.PO_Approved_Position
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Private Sub TxtPreparedBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPreparedBy.GotFocus
        Try
            Me.TxtPreparedBy.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPreparedBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPreparedBy.LostFocus
        Try
            Me.TxtPreparedBy.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPreparedBy_Jabatan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPreparedBy_Jabatan.GotFocus
        Try
            Me.TxtPreparedBy_Jabatan.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPreparedBy_Jabatan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPreparedBy_Jabatan.LostFocus
        Try
            Me.TxtPreparedBy_Jabatan.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtApprovedBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtApprovedBy.GotFocus
        Try
            Me.TxtApprovedBy.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtApprovedBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtApprovedBy.LostFocus
        Try
            Me.TxtApprovedBy.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtApprovedBy_Jabatan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtApprovedBy_Jabatan.GotFocus
        Try
            Me.TxtApprovedBy_Jabatan.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtApprovedBy_Jabatan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtApprovedBy_Jabatan.LostFocus
        Try
            Me.TxtApprovedBy_Jabatan.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAcknowledgeBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAcknowledgeBy.GotFocus
        Try
            Me.TxtAcknowledgeBy.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAcknowledgeBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAcknowledgeBy.LostFocus
        Try
            Me.TxtAcknowledgeBy.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAcknowledgeBy_Jabatan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAcknowledgeBy_Jabatan.GotFocus
        Try
            Me.TxtAcknowledgeBy_Jabatan.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAcknowledgeBy_Jabatan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAcknowledgeBy_Jabatan.LostFocus
        Try
            Me.TxtAcknowledgeBy_Jabatan.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                'simpan settings
                My.Settings.PO_Prepared_By = Me.TxtPreparedBy.Text
                My.Settings.PO_Prepared_Position = Me.TxtPreparedBy_Jabatan.Text
                My.Settings.PO_Acknowled_By = Me.TxtAcknowledgeBy.Text
                My.Settings.PO_Acknowledge_Position = Me.TxtAcknowledgeBy_Jabatan.Text
                My.Settings.PO_Approved_By = Me.TxtApprovedBy.Text
                My.Settings.PO_Approved_Position = Me.TxtApprovedBy_Jabatan.Text
                My.Settings.Save()

                'update juga langsung ke inventory module
                UpdateSettingsPO()

                'update data settings po pada inventory module juga
                Inventory_Module.UpdateSettingPO(Me.TxtPreparedBy.Text, Me.TxtPreparedBy_Jabatan.Text, Me.TxtAcknowledgeBy.Text, _
                Me.TxtAcknowledgeBy_Jabatan.Text, Me.TxtApprovedBy.Text, Me.TxtApprovedBy_Jabatan.Text)

                'tutup form
                Me.Ask_Before_Closing = False
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class