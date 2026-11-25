Option Explicit On
Option Strict Off

Partial Public Class FrmUpdate
    Private Sub btnUpdate_UpdatePOB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_UpdatePOB.Click
        Try
            If MsgBox("Update POB?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgwUpdatePOB.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwUpdatePOB_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwUpdatePOB.DoWork
        Dim DTA_PO As New DseUpdate_OtherTableAdapters.tbltrans_poTableAdapter
        Dim TBL_PO As New DseUpdate_Other.tbltrans_poDataTable
        Dim row_PO As DseUpdate_Other.tbltrans_poRow
        Dim ObjPO As New Purchase_Module.ClsTransaksiPOItem

        Try
            DTA_PO.Fill(TBL_PO)
            Me.PrgUpdate_UpdatePOB.Value = 0
            Me.PrgUpdate_UpdatePOB.Maximum = TBL_PO.Rows.Count

            For Each row_PO In TBL_PO.Rows
                Me.PrgUpdate_UpdatePOB.Value += 1
                Me.lblProgress_UpdatePOB.Text = "Updating POB - data:" + Me.PrgUpdate_UpdatePOB.Value.ToString + _
                " of " + Me.PrgUpdate_UpdatePOB.Maximum.ToString + " - " + row_PO.No_Transaksi
                ObjPO.SetOrderBalance(row_PO.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_PO.Dispose()
            TBL_PO.Dispose()
            ObjPO = Nothing
        End Try
    End Sub

    Private Sub bgwUpdatePOB_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwUpdatePOB.RunWorkerCompleted
        Try
            MsgBox("Update completed.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDeletePackingList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeletePackingList.Click
        Control.CheckForIllegalCrossThreadCalls = False
        Me.BackgroundWorker12.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker12_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker12.DoWork
        Dim DTADO As New DseUpdate_OtherTableAdapters.tbltrans_doTableAdapter
        Dim TBLDO As New DseUpdate_Other.tbltrans_doDataTable
        Dim rowDO As DseUpdate_Other.tbltrans_doRow
        Dim ObjDO As New Sales_Module.ClsTransaksiDO

        Try
            DTADO.Fill(TBLDO)
            Me.prgPackingList.Value = 0
            Me.prgPackingList.Maximum = TBLDO.Rows.Count
            For Each rowDO In TBLDO.Rows
                Me.prgPackingList.Value += 1
                Me.lblPackingList.Text = "Delete packing list - Data " + Me.prgPackingList.Value.ToString + " of " + Me.prgPackingList.Maximum.ToString + " - " + rowDO.No_Transaksi
                ObjDO.Hapus(rowDO.Primary, "USER", False)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTADO.Dispose()
            TBLDO.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker12_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker12.RunWorkerCompleted
        Try
            Me.prgPackingList.Value = 0
            Me.lblPackingList.Text = "Finish"
            MsgBox("Finish")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateSalesPerson.Click
        Control.CheckForIllegalCrossThreadCalls = False
        Me.BackgroundWorker11.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker11_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker11.DoWork
        Dim DTASalesP As New DseUpdate_OtherTableAdapters.tblmaster_salespersonTableAdapter
        Dim TBLSalesP As New DseUpdate_Other.tblmaster_salespersonDataTable
        Dim rowSalesP As DseUpdate_Other.tblmaster_salespersonRow
        Dim DTACust As New DseUpdate_OtherTableAdapters.tblmaster_customerTableAdapter
        Dim TBLCust As New DseUpdate_Other.tblmaster_customerDataTable
        Dim rowCust As DseUpdate_Other.tblmaster_customerRow
        Dim DTASO As New DseUpdate_OtherTableAdapters.tbltrans_soTableAdapter
        Dim TBLSO As New DseUpdate_Other.tbltrans_soDataTable
        Dim rowSO As DseUpdate_Other.tbltrans_soRow
        Dim Rdn As New Random
        Dim Index As Integer = 0

        Try
            DTASalesP.Fill(TBLSalesP)
            DTACust.Fill(TBLCust)
            Me.prbUpdateSalesPerson.Value = 0
            Me.prbUpdateSalesPerson.Maximum = TBLCust.Rows.Count
            For Each rowCust In TBLCust.Rows
                Me.prbUpdateSalesPerson.Value += 1
                Me.LblUpdateSalesPerson.Text = "Phase 1 of 2: Update customer - Data " + Me.prbUpdateSalesPerson.Value.ToString + " of " + Me.prbUpdateSalesPerson.Maximum.ToString + " - " + rowCust.Customer_Name
                Index = Rdn.Next(0, TBLSalesP.Rows.Count - 1)
                rowSalesP = TBLSalesP.Rows(Index)
                DTACust.UpdateQuery_UpdateSalesPerson(rowSalesP.Primary, rowCust.Primary)
            Next

            DTASO.Fill(TBLSO)
            Me.prbUpdateSalesPerson.Value = 0
            Me.prbUpdateSalesPerson.Maximum = TBLSO.Rows.Count
            For Each rowSO In TBLSO.Rows
                Me.prbUpdateSalesPerson.Value += 1
                Me.LblUpdateSalesPerson.Text = "Phase 1 of 2: Update customer - Data " + Me.prbUpdateSalesPerson.Value.ToString + " of " + _
                Me.prbUpdateSalesPerson.Maximum.ToString + " - " + rowSO.No_Transaksi
                Index = Rdn.Next(0, TBLSalesP.Rows.Count - 1)
                rowSalesP = TBLSalesP.Rows(Index)
                DTASO.UpdateQuery_UpdateSalesPerson(rowSalesP.Primary, rowSO.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker11_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker11.RunWorkerCompleted
        Try
            Me.prbUpdateSalesPerson.Value = 0
            Me.LblUpdateSalesPerson.Text = "Finish"
            MsgBox("Finish")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdateEmail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdateEmail.Click
        Try
            If MsgBox("Update email contact person?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker13.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker13_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker13.DoWork
        Dim ListEmail As New List(Of String)
        Dim Rdn As New Random
        Dim DummyIdx As Integer = 0
        Dim DTA As New DseUpdate_OtherTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim TBL As New DseUpdate_Other.tblmaster_customer_itemcpDataTable
        Dim row As DseUpdate_Other.tblmaster_customer_itemcpRow
        Dim EmailAdd As String = ""

        Try
            ListEmail.Add("yulimus.j@gmail.com")
            ListEmail.Add("yulimus@arthajina.com")
            ListEmail.Add("yulimus@yahoo.com")
            ListEmail.Add("yulimus.j@gmail.com")
            ListEmail.Add("email.testing597@gmail.com")

            DTA.Fill(TBL)
            Me.prgUpdateEmail.Value = 0
            Me.prgUpdateEmail.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                Me.prgUpdateEmail.Value += 1
                Me.LblUpdateEmail.Text = "Updating data " + Me.prgUpdateEmail.Value.ToString + " of " + Me.prgUpdateEmail.Maximum.ToString + ": " + row.Customer_Name + " - " + row.Name
                DummyIdx = Rdn.Next(0, ListEmail.Count - 1)
                EmailAdd = ListEmail.Item(DummyIdx)

                DTA.UpdateQuery_EmailCP(EmailAdd, 1, row.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker13_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker13.RunWorkerCompleted
        Try
            Me.prgUpdateEmail.Value = 0
            Me.LblUpdateEmail.Text = "Finish"
            MsgBox("Email update finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdate_NPWP_AOP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_NPWP_AOP.Click
        Try
            If MsgBox("Update NPWP AOP?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker14.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker14_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker14.DoWork
        Dim DTA As New DseUpdate_OtherTableAdapters.tbltrans_aopTableAdapter
        Dim TBL As New DseUpdate_Other.tbltrans_aopDataTable
        Dim row As DseUpdate_Other.tbltrans_aopRow
        Dim Urutan As Double = 17177307
        Dim Jml As Nullable(Of Long)

        Try
            DTA.Fill(TBL)
            Me.PrgUpdate_NPWP_AOP.Value = 0
            Me.PrgUpdate_NPWP_AOP.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                Me.PrgUpdate_NPWP_AOP.Value += 1
                If row.No_NPWP.Length <> ("01.555.327.4-446.000").Length Or row.No_NPWP = "00.000.000.0-000.000" Then
                    row.No_NPWP = "01.555.327.4-446.000"
                    DTA.UpdateQuery(row.No_NPWP, row.Primary)
                End If
                If row.Tax_Invoice_Number.Length <> ("010.000-14.17177307").Length Or row.Tax_Invoice_Number.Contains("YY") = True Then
                    row.Tax_Invoice_Number = "010.000-" + Format(row.Tanggal_Transaksi, "yy") + "." + Urutan.ToString
                    Jml = DTA.ScalarQuery_GetJml(row.Tax_Invoice_Number)

                    While Jml.Value > 0
                        Urutan += 1
                        row.Tax_Invoice_Number = "010.000-" + Format(row.Tanggal_Transaksi, "yy") + "." + Urutan.ToString
                        Jml = DTA.ScalarQuery_GetJml(row.Tax_Invoice_Number)
                    End While

                    DTA.UpdateQuery1(row.Tax_Invoice_Number, row.Primary)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker14_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker14.RunWorkerCompleted
        Try
            MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdate_UpdatePrOB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_UpdatePrOB.Click
        Try
            If MsgBox("Update PrOB?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgwUpdatePrOB.RunWorkerAsync()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdatePrOB()
        Dim DTA As New DseUpdate_OtherTableAdapters.tbltrans_purchase_reqTableAdapter
        Dim TBL As New DseUpdate_Other.tbltrans_purchase_reqDataTable
        Dim row As DseUpdate_Other.tbltrans_purchase_reqRow
        Dim ObjPR As New Purchase_Module.ClsTransaksiPR

        Try
            'load data pr
            DTA.Fill(TBL)

            'panggil fungsi set order balance
            Me.PrgUpdate_PrOB.Value = 0
            Me.PrgUpdate_PrOB.Maximum = TBL.Rows.Count
            For Each row In TBL.Rows
                Me.PrgUpdate_PrOB.Value += 1
                Me.lblProgress_UpdatePrOB.Text = "Data " + Me.PrgUpdate_PrOB.Value.ToString + " of " + Me.PrgUpdate_PrOB.Maximum.ToString + " - " + row.no_transaksi
                ObjPR.SetOrderBalance1(row.primary, row.ready_for_purchase)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjPR = Nothing
        End Try
    End Sub

    Private Sub bgwUpdatePrOB_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwUpdatePrOB.DoWork
        Try
            Me.UpdatePrOB()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwUpdatePrOB_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwUpdatePrOB.RunWorkerCompleted
        MsgBox("Update completed.", MsgBoxStyle.Information, "Information")
    End Sub
End Class
