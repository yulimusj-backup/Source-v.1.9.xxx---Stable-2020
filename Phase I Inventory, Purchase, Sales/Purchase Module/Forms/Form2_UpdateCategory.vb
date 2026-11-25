Public Class Form2_UpdateCategory
    Public vPeriod As Date = #1/1/1950#

    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub Form2_UpdateCategory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.bgw1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Dim DTA As New Dse4TableAdapters.DataTable1TableAdapter
        Dim TBL As New Dse4.DataTable1DataTable
        Dim row As Dse4.DataTable1Row

        Try
            Me.ObjCD.UpdateThreading()

            'load data tabel
            Me.lbl1.Text = "Preparing data..."

            DTA.Fill(TBL, Me.vPeriod)

            'update ke db
            Me.prg1.Value = 0
            Me.prg1.Maximum = TBL.Rows.Count
            For Each row In TBL.Rows
                Me.prg1.Value += 1
                Me.lbl1.Text = "Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString

                DTA.UpdateQuery_SetCategory(row.Category_Code, row.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class