Option Explicit On
Option Strict Off
Partial Public Class FrmUpdate

    Private Sub btn_v170_quotation_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_v170_quotation.Click
        Try
            If MsgBox("Start quotation update?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker16.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_v170_Quotation()
        Dim DTA As New DseUpdate_v170TableAdapters.tbltrans_quotation_itemTableAdapter
        Dim TBL As New DseUpdate_v170.tbltrans_quotation_itemDataTable
        Dim row As DseUpdate_v170.tbltrans_quotation_itemRow
        Dim Dummy As Integer = 0
        Dim PrimaryQuotationBefore As Double = -1
        Dim DTA1 As New DseUpdate_v170TableAdapters.tbltrans_quotation_itemdeliveryTableAdapter
        Dim ObjPriority As New Core_Module.ClsMasterPriorityLevel
        Dim DefPriority As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            Me.prg_v170_quotation.Value = 0

            DTA.Fill(TBL)

            Me.prg_v170_quotation.Maximum = TBL.Rows.Count
            DefPriority = ObjPriority.GetPrimaryDefault

            For Each row In TBL.Rows
                If PrimaryQuotationBefore <> row.Primary_Transaksi Then
                    Dummy = 0
                    PrimaryQuotationBefore = row.Primary_Transaksi
                    DTA1.Delete(row.Primary_Transaksi)
                End If
                Dummy += 1

                Me.prg_v170_quotation.Value += 1
                Me.lbl_v170_quotation.Text = "Updating Data " + Me.prg_v170_quotation.Value.ToString + " of " + Me.prg_v170_quotation.Maximum.ToString
                DTA.UpdateQuery_SetIdentifier(row.Primary_Transaksi.ToString + "-" + Dummy.ToString, row.Primary)

                If row.Delivery_Date_Cont = True Then
                    'tambahkan item delivery date control
                    DTA1.Insert(row.Primary_Transaksi, row.Primary_Transaksi.ToString + "-" + _
                    Dummy.ToString, row.Qty, row.Delivery_Date, "", DefPriority)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA1.Dispose()
            ObjPriority = Nothing
        End Try
    End Sub

    Private Sub BackgroundWorker16_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker16.DoWork
        Try
            Me.Update_v170_Quotation()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker16_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker16.RunWorkerCompleted
        Try
            Me.lbl_v170_quotation.Text = "Data Updated"
            MsgBox("Update completed.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
