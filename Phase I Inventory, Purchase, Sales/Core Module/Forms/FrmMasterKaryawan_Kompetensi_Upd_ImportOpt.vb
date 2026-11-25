Option Explicit On
Option Strict Off
Public Class FrmMasterKaryawan_Kompetensi_Upd_ImportOpt

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub FrmMasterKaryawan_Kompetensi_Upd_ImportOpt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Text = "Import View"
            Me.DgvImport.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If GetJmlDgvImport() = True Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            Else
                MsgBox("Not found selected import. please change/checked then true.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlDgvImport() As Boolean
        Try
            For Each row As DataGridViewRow In Me.DgvImport.Rows
                With row
                    If .Cells(Me.ColImp_Import.Index).Value = True Then
                        Return True
                    End If
                End With
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.No
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbUnchkAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUnchkAll.LinkClicked
        Try
            For i As Integer = 0 To Me.DgvImport.Rows.Count - 1
                Me.DgvImport.Item(Me.ColImp_Import.Index, i).Value = False
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbChkAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbChkAll.LinkClicked
        Try
            For i As Integer = 0 To Me.DgvImport.Rows.Count - 1
                Me.DgvImport.Item(Me.ColImp_Import.Index, i).Value = True
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class