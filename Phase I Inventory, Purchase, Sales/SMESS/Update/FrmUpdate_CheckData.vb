Option Explicit On
Option Strict Off

Partial Public Class FrmUpdate
    Dim TblRef As New DseUpdate_CheckData.tbldo_itemDataTable
    Dim rowRef As DseUpdate_CheckData.tbldo_itemRow

    Private Sub btnCheckItemDO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckItemDO.Click
        Try
            If MsgBox("Start data check?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.Bgw_Check_ItemDO.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CekDataDOItem()
        Dim DTA As New DseUpdate_CheckDataTableAdapters.tbldo_itemTableAdapter
        Dim TBL As New DseUpdate_CheckData.tbldo_itemDataTable
        Dim row As DseUpdate_CheckData.tbldo_itemRow

        Dim DTA1 As New DseUpdate_CheckDataTableAdapters.tbltrans_stockinoutTableAdapter
        Dim TBL1 As New DseUpdate_CheckData.tbltrans_stockinoutDataTable
        Dim row1 As DseUpdate_CheckData.tbltrans_stockinoutRow

        Try
            Me.TblRef.Rows.Clear()

            DTA.FillBy_TGLDari(TBL, Me.dtpCheck_DOItem.Value.Date)
            DTA1.Fill(TBL1, 3, Me.dtpCheck_DOItem.Value.Date)

            Me.prgCheck_ItemDO.Value = 0
            Me.prgCheck_ItemDO.Maximum = TBL.Rows.Count + TBL1.Rows.Count

            For Each row In TBL.Rows
                Me.prgCheck_ItemDO.Value += 1

                If Me.prgCheck_ItemDO.Value > 500 Then
                    If Math.IEEERemainder(Me.prgCheck_ItemDO.Value, 50) = 0 Then
                        Me.lblCheck_ItemDO.Text = "Checking data " + Me.prgCheck_ItemDO.Value.ToString + " of " + Me.prgCheck_ItemDO.Maximum.ToString
                    End If
                Else
                    Me.lblCheck_ItemDO.Text = "Checking data " + Me.prgCheck_ItemDO.Value.ToString + " of " + Me.prgCheck_ItemDO.Maximum.ToString
                End If


                If row.Total_Qty <> CDec(ObjCD.DigitGrouping(row.Total_Qty_Multi * _
                    row.Ratio_Default / row.Ratio_Reference, row._Decimal)) Then

                    Me.rowRef = Me.TblRef.Newtbldo_itemRow
                    Me.rowRef.No_Transaksi = row.No_Transaksi
                    Me.rowRef.Product_Code = row.Product_Code
                    Me.rowRef.Primary_DO = -1
                    Me.rowRef.Primary_Product = -1
                    Me.rowRef.Ratio_Default = 1
                    Me.rowRef.Ratio_Reference = 1
                    Me.rowRef.Total_Qty = 1
                    Me.rowRef.Total_Qty_Multi = 0
                    Me.rowRef._Decimal = 0

                    Me.TblRef.Rows.Add(Me.rowRef) 
                End If
            Next

            For Each row1 In TBL1.Rows
                Me.prgCheck_ItemDO.Value += 1
                If Me.prgCheck_ItemDO.Value > 500 Then
                    If Me.prgCheck_ItemDO.Value \ 50 = 0 Then
                        Me.lblCheck_ItemDO.Text = "Checking data " + Me.prgCheck_ItemDO.Value.ToString + " of " + Me.prgCheck_ItemDO.Maximum.ToString
                    End If
                Else
                    Me.lblCheck_ItemDO.Text = "Checking data " + Me.prgCheck_ItemDO.Value.ToString + " of " + Me.prgCheck_ItemDO.Maximum.ToString
                End If


                If row1.Qty <> CDec(ObjCD.DigitGrouping(row1.Qty_Multi_Unit * _
                    row1.Ratio_Default / row1.Rasio, row1._Decimal)) Then

                    Me.rowRef = Me.TblRef.Newtbldo_itemRow
                    Me.rowRef.No_Transaksi = row1.No_Transaksi
                    Me.rowRef.Product_Code = row1.Product_Code
                    Me.rowRef.Primary_DO = -1
                    Me.rowRef.Primary_Product = -1
                    Me.rowRef.Ratio_Default = 1
                    Me.rowRef.Ratio_Reference = 1
                    Me.rowRef.Total_Qty = 1
                    Me.rowRef.Total_Qty_Multi = 0
                    Me.rowRef._Decimal = 0

                    Me.TblRef.Rows.Add(Me.rowRef)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub Bgw_Check_ItemDO_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles Bgw_Check_ItemDO.DoWork
        Try
            Me.CekDataDOItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Bgw_Check_ItemDO_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles Bgw_Check_ItemDO.RunWorkerCompleted
        Try
            If Me.TblRef.Rows.Count = 0 Then
                MsgBox("Data checked", MsgBoxStyle.Information, "Information")
            Else
                MsgBox("Error data found.", MsgBoxStyle.Information, "Information")
                Dim ij As New FrmUpdate_CheckData_ItemDO
                ij.TblRef = Me.TblRef

                ij.ShowDialog()
            End If
            Me.lblCheck_ItemDO.Text = "Check."
            Me.prgCheck_ItemDO.Value = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
