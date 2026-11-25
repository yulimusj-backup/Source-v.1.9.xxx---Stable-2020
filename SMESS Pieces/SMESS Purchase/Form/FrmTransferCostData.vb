Option Explicit On
Option Strict Off

Public Class FrmTransferCostData
    Dim DTA As New DseTransferClosingTableAdapters.tbltrans_purchase_closingTableAdapter
    Dim TBL As New DseTransferClosing.tbltrans_purchase_closingDataTable
    Dim row As DseTransferClosing.tbltrans_purchase_closingRow
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCLosing As New Purchase_Module.ClsTransaksiMonthlyClosing

    Private Sub FrmTransferCostData_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            DTA.Fill(TBL)

            Me.cboClosing.DataSource = TBL.DefaultView
            Me.cboClosing.DisplayMember = "No Transaksi"

            If Me.TBL.Rows.Count > 0 Then
                Me.cboClosing.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If Me.cboClosing.SelectedIndex <> -1 Then
                Me.row = TBL.DefaultView(Me.cboClosing.SelectedIndex).Row
                Me.ObjCLosing.SimpanDataCost(row.Primary, row.Closing_Period)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckKonsistensi()
        Dim DTA1 As New DseTransferClosingTableAdapters.tbltrans_salesinvoice_item_baruTableAdapter
        Dim TBL1 As New DseTransferClosing._tbltrans_salesinvoice_item_baruDataTable
        Dim row1 As DseTransferClosing._tbltrans_salesinvoice_item_baruRow
        Dim DTA2 As New DseTransferClosingTableAdapters.tbltrans_salesinvoice_item1_lamaTableAdapter
        Dim TBL2 As New DseTransferClosing._tbltrans_salesinvoice_item1_lamaDataTable
        Dim row2 As DseTransferClosing._tbltrans_salesinvoice_item1_lamaRow
        Dim row As DseTransferClosing.tbltrans_purchase_closingRow
        Dim Filter As String = ""

        Try
            While Me.cboClosing.SelectedIndex <= Me.cboClosing.Items.Count - 1
                row = Me.TBL.DefaultView(Me.cboClosing.SelectedIndex).Row
                DTA2.Fill(TBL2, Me.ObjCD.GetLasDateOfTheMonth(row.Closing_Period), Me.ObjCD.GetFirstDateofTheMonth(row.Closing_Period))
                DTA1.Fill(TBL1, Me.ObjCD.GetLasDateOfTheMonth(row.Closing_Period), Me.ObjCD.GetFirstDateofTheMonth(row.Closing_Period))

                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = TBL1.Rows.Count
                For Each row1 In TBL1.Rows
                    Filter = "Primary = " + row1.Primary.ToString

                    TBL2.DefaultView.RowFilter = Filter
                    If TBL2.DefaultView.Count = 0 Then
                        MsgBox("Salah pada " + row1.Primary.ToString)
                    Else
                        row2 = TBL2.DefaultView(0).Row
                        If row1.Cost <> row2.Cost Then
                            MsgBox("Salah pada " + row1.Primary.ToString)
                        End If
                    End If

                    Me.ProgressBar1.Value += 1
                    Me.Text = "Checking Phase 1 (" + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                Next

                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = TBL1.Rows.Count
                For Each row2 In TBL2.Rows
                    Filter = "Primary = " + row2.Primary.ToString

                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count = 0 Then
                        MsgBox("Salah pada " + row2.Primary.ToString)
                    Else
                        row1 = TBL1.DefaultView(0).Row
                        If row2.Cost <> row1.Cost Then
                            MsgBox("Salah pada " + row2.Primary.ToString)
                        End If
                    End If

                    Me.ProgressBar1.Value += 1
                    Me.Text = "Checking Phase 2 (" + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                Next

                If Me.cboClosing.SelectedIndex + 1 <= Me.cboClosing.Items.Count - 1 Then
                    Me.cboClosing.SelectedIndex += 1
                Else
                    Exit Sub
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
            DTA2.Dispose()
            TBL2.Dispose()
        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Dim i As New Threading.Thread(AddressOf CheckKonsistensi)
            i.Start()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class