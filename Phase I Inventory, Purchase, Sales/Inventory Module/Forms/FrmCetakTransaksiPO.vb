Option Explicit On
Option Strict Off

Public Class FrmCetakTransaksiPO
    Public User As String = ""
    Public PrimaryTransaksi As Double = -1
    Public ListPrimaryPo As New List(Of Double)
    Public Printed As Boolean = False
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub FrmCetakTransaksiPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim DTAPO As New DseTransaksiPOTableAdapters.tbltrans_po_NewTableAdapter
        Dim Result As New MsgBoxResult
        Try
            Result = MsgBox("Is PO printed correctly?", MsgBoxStyle.YesNo, "Confirm")
            If Result = MsgBoxResult.Yes Then
                If Me.ListPrimaryPo.Count > 0 Then
                    For intA As Integer = 0 To Me.ListPrimaryPo.Count - 1
                        DTAPO.UpdateQuery_PrintedBy(Me.User, Now, Me.ListPrimaryPo.Item(intA))
                    Next
                Else
                    DTAPO.UpdateQuery_PrintedBy(Me.User, Now, PrimaryTransaksi)
                End If
                Printed = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPO.Dispose()
        End Try
    End Sub

    Private Sub FrmCetakTransaksiPO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class