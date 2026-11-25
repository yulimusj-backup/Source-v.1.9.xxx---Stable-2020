Option Explicit On
Option Strict Off

Public Class FrmTransaksiSO_SetSOB
    Dim DTASO As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
    Dim TBLSO As New DseTransaksiSO.tbltrans_soDataTable
    Dim rowSO As DseTransaksiSO.tbltrans_soRow
    Dim ObjSOItem As New ClsTransaksiSalesOrderItem
    Dim SW As New Stopwatch

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Me.Prb01.Value = 0
            Me.Prb01.Maximum = DTASO.Fill(TBLSO) + 1
            Me.Button1.Enabled = False
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
            For Each rowSO In Me.TBLSO.Rows
                Me.ObjSOItem.Set_Order_Balance(rowSO.Primary)
                Me.Prb01.Value += 1
                Me.Text = "Progress: " + counter.ToString + " of " + Me.TBLSO.Rows.Count.ToString
                counter += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Control.CheckForIllegalCrossThreadCalls = True
            MsgBox("Selesai dalam " + (SW.ElapsedTicks / TimeSpan.TicksPerMinute).ToString + " minutes. Average: " + _
            (Math.Round((SW.ElapsedTicks / TimeSpan.TicksPerSecond) / Me.TBLSO.Rows.Count)).ToString, MsgBoxStyle.Information, "Information")
            Me.Button1.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class