Option Explicit On
Option Strict Off

Public Class FrmTransaksiPO_SetOB
    Dim ObjPOItem As New ClsTransaksiPOItem
    Dim TBLPO As New DseTransaksiPO.tbltrans_poDataTable
    Dim DTAPO As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
    Dim rowPO As DseTransaksiPO.tbltrans_poRow
    Dim SW As New Stopwatch

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Try
            Me.Prb01.Value = 0
            Me.Prb01.Maximum = DTAPO.Fill(TBLPO) + 1
            Me.Enabled = False
            Control.CheckForIllegalCrossThreadCalls = False
            SW.Start()
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim counter As Double = 1

        Try
            For Each rowPO In Me.TBLPO.Rows
                Me.ObjPOItem.SetOrderBalance(rowPO.Primary)
                Me.Prb01.Value += 1
                Me.Text = "Progress: " + counter.ToString + " of " + Me.TBLPO.Rows.Count.ToString
                counter += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.Enabled = True
            MsgBox("Finish in " + (SW.ElapsedTicks / TimeSpan.TicksPerSecond).ToString + " seconds, average: " _
            + ((SW.ElapsedTicks / TimeSpan.TicksPerSecond) / Me.TBLPO.Rows.Count).ToString + " seconds", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ObjPOItem.SetOrderBalance(451)
    End Sub
End Class