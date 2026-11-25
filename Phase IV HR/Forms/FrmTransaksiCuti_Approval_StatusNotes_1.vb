Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class FrmTransaksiCuti_Approval_StatusNotes_1
    Public mPrimaryTrans As Double = -1
    Public FrmUtama As New FrmTransaksiCuti_Approval
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll

    Dim RunEventClosing As Boolean = True
    Dim ObjApp As New ClsTrans_Cuti_Approval
    Dim ObjPengajuan As New ClsTrans_Cuti_Pengajuan

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.RunEventClosing = False
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'konfirmasi disimpan atau tidak
            hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
            If hasil = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjApp.UpdateStatusNotes(Me.TxtStatusNotes.Text, Me.mPrimaryTrans) = True Then
                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                    Me.RunEventClosing = False
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Approval_StatusNotes_1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.RunEventClosing = True Then
                If Me.TxtStatusNotes.Text <> Me.ObjPengajuan.GetStatusNotes(Me.mPrimaryTrans) Then
                    Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjApp.UpdateStatusNotes(Me.TxtStatusNotes.Text, Me.mPrimaryTrans) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Else
                            e.Cancel = True
                        End If
                    ElseIf hasil = MsgBoxResult.Cancel Then
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Approval_StatusNotes_1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.RunEventClosing = True
            Me.TxtStatusNotes.Text = Me.ObjApp.GetStatusNotes(Me.mPrimaryTrans)
            Me.TxtStatusNotes.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtStatusNotes_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStatusNotes.GotFocus
        Try
            Me.TxtStatusNotes.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtStatusNotes_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStatusNotes.LostFocus
        Try
            Me.TxtStatusNotes.BackColor = Drawing.Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class