Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Public Class FrmTransaksiCuti_Pengajuan_ExplainDuration
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Public Half1 As Boolean = False
    Public Half2 As Boolean = False
    Dim ObjCD As New Core_Module.ClsCekData
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub dgvExplain_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvExplain.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColStatus.Index Then
                e.Cancel = True
            Else
                If Me.dgvExplain.Item(Me.ColDuration.Index, e.RowIndex).Value = 0 And Me.dgvExplain.Item(Me.ColStatus.Index, e.RowIndex).Value = "Leave" Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub dgvExplain_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvExplain.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColStatus.Index Then
                If Me.dgvExplain.Item(Me.ColStatus.Index, e.RowIndex).Value <> "Leave" Then
                    Me.dgvExplain.Item(Me.ColDuration.Index, e.RowIndex).Value = "0.0"
                Else
                    If e.RowIndex = 0 Then
                        If Me.Half1 = True Then
                            Me.dgvExplain.Item(Me.ColDuration.Index, e.RowIndex).Value = "0.5"
                        Else
                            Me.dgvExplain.Item(Me.ColDuration.Index, e.RowIndex).Value = "1.0"
                        End If
                    ElseIf e.RowIndex = Me.dgvExplain.RowCount - 1 Then
                        If Me.Half2 = True Then
                            Me.dgvExplain.Item(Me.ColDuration.Index, e.RowIndex).Value = "0.5"
                        Else
                            Me.dgvExplain.Item(Me.ColDuration.Index, e.RowIndex).Value = "1.0"
                        End If
                    Else
                        Me.dgvExplain.Item(Me.ColDuration.Index, e.RowIndex).Value = "1.0"
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub FrmTransaksiCuti_Pengajuan_ExplainDuration_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.btnClose.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub FrmTransaksiCuti_Pengajuan_ExplainDuration_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ShowIcon = False
            Me.dgvExplain.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub dgvExplain_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvExplain.DataError

    End Sub
End Class