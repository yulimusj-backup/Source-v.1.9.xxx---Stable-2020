'form ini mencari AOP pelunasan, yang sub total b4 dp = sub total, padahal nilai downpaymentnya > 0
Public Class FrmRepairAOP
    Public BackColorAll As Color = Color.LightGoldenrodYellow

    Dim DTA As New DseUpdate_v2018TableAdapters.tbltrans_aopTableAdapter
    Dim TBL As New DseUpdate_v2018.tbltrans_aopDataTable
    Dim row As DseUpdate_v2018.tbltrans_aopRow
    Dim BgwProgress As Double = 0

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Control.CheckForIllegalCrossThreadCalls = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Me.TBL.Rows.Clear()
            Me.DTA.Fill(Me.TBL, #1/1/1950#)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheck.Click
        Try
            If MsgBox("Start data check?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnRepair_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRepair.Click
        Try
            If MsgBox("Start data repair?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DTA.UpdateQuery()
                MsgBox("Data repair finished", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            BgwProgress = 1
            Me.BackgroundWorker1.ReportProgress(1)

            Me.LoadData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            If Me.BgwProgress = 1 Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim rowDGV As DataGridViewRow

        Try
            For Each row In Me.TBL.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)

                rowDGV.Cells(Me.ColDate.Index).Value = Me.row.Tanggal_Transaksi
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                rowDGV.Cells(Me.ColNoAOP.Index).Value = Me.row.No_Transaksi
                rowDGV.Cells(Me.ColPrimary.Index).Value = Me.row.Primary
            Next

            Me.BgwProgress = 0

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            MsgBox("Data check finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
