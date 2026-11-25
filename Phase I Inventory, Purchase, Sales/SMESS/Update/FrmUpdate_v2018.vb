Option Explicit On
Option Strict Off

Partial Public Class FrmUpdate
    Private Sub btnUpdate2018_AOP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate2018_AOP.Click
        Dim ij As New FrmRepairAOP

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn2018_UpdateSODO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2018_UpdateSODO.Click
        Dim DTA As New DseUpdate_v2018TableAdapters.tbltrans_moTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery1()
            DTA.UpdateQuery2()

            Me.LblProgress2018_UpdateSODO.Text = "Done"

            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub Update2()
        Dim DTA1 As New DseUpdate_v180TableAdapters.tbltrans_purchase_closing_item_akunTableAdapter
        Dim DTA2 As New DseUpdate_v180TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter

        Try
            'tipe item 8 -> -8
            DTA1.UpdateQuery(-8, 8)
            DTA2.UpdateQuery(-8, 8)
            'tipe item '7->8
            DTA1.UpdateQuery(8, 7)
            DTA2.UpdateQuery(8, 7)
            'tipe item 6->7
            DTA1.UpdateQuery(7, 6)
            DTA2.UpdateQuery(7, 6)
            'tipe item 5->6
            DTA1.UpdateQuery(6, 5)
            DTA2.UpdateQuery(6, 5)
            'tipe item 4->5
            DTA1.UpdateQuery(5, 4)
            DTA2.UpdateQuery(5, 4)
            'tipe item  -8 ->4
            DTA1.UpdateQuery(4, -8)
            DTA2.UpdateQuery(4, -8)

            MsgBox("Selesai")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button15_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Update2()
    End Sub

    Private Sub btn2018_UpdateMO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2018_UpdateMO.Click
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.btn2018_UpdateMO.Enabled = False
            Me.Bgw2018_UpdateMO.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Bgw2018_UpdateMO_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles Bgw2018_UpdateMO.DoWork
        Dim DTA As New DseUpdate_v2018TableAdapters.tbltrans_moTableAdapter
        Dim TBL As New DseUpdate_v2018.tbltrans_moDataTable
        Dim row As DseUpdate_v2018.tbltrans_moRow
        Dim ObjUser As New Core_Module.ClsUser

        Try
            Me.ObjCD.UpdateThreading()
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'load data mo
            DTA.Fill(TBL)
            Me.prg2018_UpdateMO.Value = 0
            Me.prg2018_UpdateMO.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                Me.prg2018_UpdateMO.Value += 1
                Me.LblProgress2018_UpdateMO.Text = "Data " + Me.prg2018_UpdateMO.Value.ToString + " of " + Me.prg2018_UpdateMO.Maximum.ToString + " : " + row.MO_Number
                If row.Primary_User_Auth = -1 Then
                    DTA.UpdateQuery("-", row.Primary)
                Else
                    DTA.UpdateQuery(ObjUser.GetUserCode(row.Primary_User_Auth) + " - " + Format(row.Auth_Time, "dd MMM yyyy HH.mm"), row.Primary)
                End If
            Next
            'update data auth str mo

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Bgw2018_UpdateMO_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles Bgw2018_UpdateMO.RunWorkerCompleted
        Try
            Me.btn2018_UpdateMO.Enabled = True
            Me.LblProgress2018_UpdateMO.Text = "done"
            Me.prg2018_UpdateMO.Value = 0
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdate2018_ItemKartuAPAR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate2018_ItemKartuAPAR.Click
        Try
            Dim dta As New DseUpdate_v2018TableAdapters.tbltransaksi_kartu_ar_ap1TableAdapter
            Dim tbl As New DseUpdate_v2018.tbltransaksi_kartu_ar_ap1DataTable
            Dim row As DseUpdate_v2018.tbltransaksi_kartu_ar_ap1Row

            Try
                Windows.Forms.Cursor.Current = Cursors.WaitCursor

                dta.Connection.ConnectionString = Core_Module.GetConstring
                dta.Fill(tbl)
                For Each row In tbl
                    With row
                        'switch primary akun vb exchange rate
                        dta.UpdateQuery(.primary_akun, .exchange_rate, .primary_currency, .exchange_rate, .primary_akun)
                    End With
                Next

                MsgBox("Done", MsgBoxStyle.Information, "Information")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Finally
                Windows.Forms.Cursor.Current = Cursors.Default
            End Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
