Public Class Form2_Cek
    Public ClosingPeriod As Date
    Public FrmParent As Form2
    Public DGVRekap As DataGridView
    Public DGVItemDetail As DataGridView
    Public DGVDetailUsing As DataGridView
    Public DGVDetailWarehouse As DataGridView

    Private ListCodeNo As New List(Of String)
    Private vChkRst_CodeNo As String
    Private vChkRst_Result As String
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub AddRowResult()
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvCheckResult.Rows.Add()
            rowDGV = Me.dgvCheckResult.Rows(Me.dgvCheckResult.Rows.Count - 1)
            rowDGV.Cells(Me.ColCodeNo.Index).Value = vChkRst_CodeNo
            rowDGV.Cells(Me.ColCheckResult.Index).Value = vChkRst_Result
            rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Try
            If MsgBox("Perform data check?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.dgvCheckResult.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvCheckResult.Rows.Clear()

                Me.KondisikanDataAwal()

                Control.CheckForIllegalCrossThreadCalls = False
                Me.prg1.Visible = True
                Me.lblProgress.Visible = True
                Me.bgw1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KondisikanDataAwal()
        Try
            Me.lblProgress.Text = "Collecting code no..."
            Me.prg1.Value = 0

            DGVRekap.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGVItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGVDetailUsing.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGVDetailWarehouse.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'urutkan dgv menurut kode barang
            DGVRekap.Sort(FrmParent.ColCodeNo, System.ComponentModel.ListSortDirection.Ascending)
            DGVItemDetail.Sort(FrmParent.ColItem_CodeNo, System.ComponentModel.ListSortDirection.Ascending)
            DGVDetailUsing.Sort(FrmParent.ColU_CodeNo, System.ComponentModel.ListSortDirection.Ascending)
            DGVDetailWarehouse.Sort(FrmParent.ColW_CodeNo, System.ComponentModel.ListSortDirection.Ascending)

            'buat list code no
            Me.ListCodeNo.Clear()
            Me.prg1.Maximum = DGVRekap.Rows.Count
            For Each rowdgv As DataGridViewRow In DGVRekap.Rows
                Me.ListCodeNo.Add(rowdgv.Cells(Me.ColCodeNo.Index).Value.ToString)
                Me.prg1.Value += 1
            Next

            DGVRekap.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGVItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGVDetailUsing.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGVDetailWarehouse.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            Me.ObjCD.UpdateThreading()

            If Me.chk2.Checked = True Then
                Me.Check_Qty()
            End If
            If Me.chk1.Checked = True Then
                Me.Check_Rekap_VS_Detail()
            End If
            If Me.chk3.Checked = True Then
                Me.Check_RekapEnd_VS_Detail_Warehouse()
            End If
            If Me.chk4.Checked = True Then
                Me.Check_Detail_Vs_Detail_Using()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgw1.ProgressChanged
        Try
            Me.AddRowResult()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            Me.prg1.Visible = False
            Me.lblProgress.Visible = False
            Me.dgvCheckResult.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            If Me.dgvCheckResult.Rows.Count = 0 Then
                MsgBox("Data check", MsgBoxStyle.Information, "Information")
            Else
                MsgBox("A few discrepancy found, please resolve", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Check_Detail_Vs_Detail_Using()
        Dim rowDGV As DataGridViewRow
        Dim QtyEnd, AmountEnd As Decimal

        Try
            'untuk setiap code no, pastikan qty end dan amount end, sama dengan detail warehouse-nya
            Me.lblProgress.Text = "Checking end recap out data vs detail using consistency"
            Me.prg1.Value = 0
            Me.prg1.Maximum = Me.DGVRekap.Rows.Count

            For Each rowDGV In Me.DGVRekap.Rows
                Me.prg1.Value += 1
                If Me.prg1.Value > 25 Then
                    If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                        Me.lblProgress.Text = "Checking end recap out data vs detail using consistency - " + rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value.ToString _
                + " - data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                    ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                        Me.lblProgress.Text = "Checking end recap out data vs detail using consistency - " + rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value.ToString _
                 + " - data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                    End If
                Else
                    Me.lblProgress.Text = "Checking end recap out data vs detail using consistency - " + rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value.ToString _
                + " - data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                End If

                Me.Get_Qty_Amount_Using(rowDGV.Cells(FrmParent.ColCodeNo.Index).Value, QtyEnd, AmountEnd)
                If QtyEnd <> CDec(rowDGV.Cells(FrmParent.ColOutQty.Index).Value) Or _
                AmountEnd <> CDec(rowDGV.Cells(FrmParent.ColAmountOut.Index).Value) Then
                    Me.vChkRst_CodeNo = rowDGV.Cells(Me.FrmParent.ColCodeNo.Index).Value
                    Me.vChkRst_Result = "Inconsistent out data vs detail using"
                    Me.bgw1.ReportProgress(0)
                    Exit Sub
                    'Me.dgvCheckResult.Rows.Add()
                    'rowDGV = Me.dgvCheckResult.Rows(Me.dgvCheckResult.Rows.Count - 1)
                    'rowDGV.Cells(Me.ColCodeNo.Index).Value = rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value
                    'rowDGV.Cells(Me.ColCheckResult.Index).Value = "Inconsistent end data vs detail warehouse"
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Check_RekapEnd_VS_Detail_Warehouse()
        Dim rowDGV As DataGridViewRow
        Dim QtyEnd, AmountEnd As Decimal

        Try
            'untuk setiap code no, pastikan qty end dan amount end, sama dengan detail warehouse-nya
            Me.lblProgress.Text = "Checking end recap data vs detail consistency"
            Me.prg1.Value = 0
            Me.prg1.Maximum = Me.DGVRekap.Rows.Count
            For Each rowDGV In Me.DGVRekap.Rows
                Me.prg1.Value += 1
                If Me.prg1.Value > 25 Then
                    If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                        Me.lblProgress.Text = "Checking end recap data vs detail consistency - " + rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value.ToString _
                        + " - data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                    ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                        Me.lblProgress.Text = "Checking end recap data vs detail consistency - " + rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value.ToString _
                        + " - data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                    End If
                Else
                    Me.lblProgress.Text = "Checking end recap data vs detail consistency - " + rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value.ToString _
                    + " - data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                End If

               

                Me.Get_Qty_Amount_Warehouse(rowDGV.Cells(FrmParent.ColCodeNo.Index).Value, QtyEnd, AmountEnd)
                If QtyEnd <> CDec(rowDGV.Cells(FrmParent.ColQtyEnd.Index).Value) Or _
                AmountEnd <> CDec(rowDGV.Cells(FrmParent.ColAmountEnd.Index).Value) Then
                    Me.vChkRst_CodeNo = rowDGV.Cells(Me.FrmParent.ColCodeNo.Index).Value
                    Me.vChkRst_Result = "Inconsistent end data vs detail warehouse"
                    Me.bgw1.ReportProgress(0)
                    'Me.dgvCheckResult.Rows.Add()
                    'rowDGV = Me.dgvCheckResult.Rows(Me.dgvCheckResult.Rows.Count - 1)
                    'rowDGV.Cells(Me.ColCodeNo.Index).Value = rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value
                    'rowDGV.Cells(Me.ColCheckResult.Index).Value = "Inconsistent end data vs detail warehouse"
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Get_Qty_Amount_Using(ByVal vCodeNo As String, ByRef vQty As Decimal, ByRef vAmount As Decimal)
        Dim rowDGV As DataGridViewRow
        Dim Keluar As Boolean = False

        Try
            vQty = 0
            vAmount = 0

            For Each rowDGV In Me.DGVDetailUsing.Rows
                If rowDGV.Cells(FrmParent.ColU_CodeNo.Index).Value.ToString = vCodeNo Then
                    Keluar = True
                    vQty += CDec(rowDGV.Cells(FrmParent.ColU_QtyUsing.Index).Value)
                    vAmount += CDec(rowDGV.Cells(FrmParent.ColU_AmountEnd.Index).Value)
                Else
                    If Keluar = True Then
                        Exit Sub
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Get_Qty_Amount_Warehouse(ByVal vCodeNo As String, ByRef vQty As Decimal, ByRef vAmount As Decimal)
        Dim rowDGV As DataGridViewRow
        Dim Keluar As Boolean = False

        Try
            vQty = 0
            vAmount = 0

            For Each rowDGV In Me.DGVDetailWarehouse.Rows
                If rowDGV.Cells(FrmParent.ColW_CodeNo.Index).Value.ToString = vCodeNo Then
                    Keluar = True
                    vQty += CDec(rowDGV.Cells(FrmParent.ColW_QtyEnd.Index).Value)
                    vAmount += CDec(rowDGV.Cells(FrmParent.ColW_AmountEnd.Index).Value)
                Else
                    If Keluar = True Then
                        Exit Sub
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Check_Rekap_VS_Detail()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Check_Qty()
        Dim DTARekap As New DseCekTableAdapters.DataTable1TableAdapter
        Dim TBLRekap As New DseCek.DataTable1DataTable
        Dim rowRekap As DseCek.DataTable1Row
        Dim rowDGV As DataGridViewRow
        Dim DTADetail As New DseCekTableAdapters.DataTable2TableAdapter
        Dim TBLDetail As New DseCek.DataTable2DataTable
        Dim rowDetail As DseCek.DataTable2Row

        Try
            Me.lblProgress.Text = "Correcting code no"
            Me.prg1.Value = 0
            Me.prg1.Maximum = Me.DGVRekap.Rows.Count + Me.DGVItemDetail.Rows.Count + Me.DGVDetailWarehouse.Rows.Count + Me.DGVDetailUsing.Rows.Count
            For Each rowDGV In Me.DGVRekap.Rows
                Me.prg1.Value += 1
                If rowDGV.Cells(Me.FrmParent.ColCodeNo.Index).Value.ToString.Contains("'") = True Then
                    rowDGV.Cells(Me.FrmParent.ColCodeNo.Index).Value = rowDGV.Cells(Me.FrmParent.ColCodeNo.Index).Value.ToString.Replace("'", "")
                End If
                rowDGV.Cells(Me.FrmParent.ColCodeNo.Index).Value = rowDGV.Cells(Me.FrmParent.ColCodeNo.Index).Value.ToString.Trim
            Next
            For Each rowDGV In Me.DGVItemDetail.Rows
                Me.prg1.Value += 1
                If rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value.ToString.Contains("'") = True Then
                    rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value = rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value.ToString.Replace("'", "")
                End If
                rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value = rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value.ToString.Trim
            Next
            For Each rowDGV In Me.DGVDetailWarehouse.Rows
                Me.prg1.Value += 1
                If rowDGV.Cells(Me.FrmParent.ColW_CodeNo.Index).Value.ToString.Contains("'") = True Then
                    rowDGV.Cells(Me.FrmParent.ColW_CodeNo.Index).Value = rowDGV.Cells(Me.FrmParent.ColW_CodeNo.Index).Value.ToString.Replace("'", "")
                End If
                rowDGV.Cells(Me.FrmParent.ColW_CodeNo.Index).Value = rowDGV.Cells(Me.FrmParent.ColW_CodeNo.Index).Value.ToString.Trim
            Next
            For Each rowDGV In Me.DGVDetailUsing.Rows
                Me.prg1.Value += 1
                If rowDGV.Cells(Me.FrmParent.ColU_CodeNo.Index).Value.ToString.Contains("'") = True Then
                    rowDGV.Cells(Me.FrmParent.ColU_CodeNo.Index).Value = rowDGV.Cells(Me.FrmParent.ColU_CodeNo.Index).Value.ToString.Replace("'", "")
                End If
                rowDGV.Cells(Me.FrmParent.ColU_CodeNo.Index).Value = rowDGV.Cells(Me.FrmParent.ColU_CodeNo.Index).Value.ToString.Trim
            Next

            'bandingkan jumlah data rekap vs tabel item closing
            Me.lblProgress.Text = "Checking recap data consistency"
            DTARekap.Fill(TBLRekap, Me.ClosingPeriod)
            Me.prg1.Value = 0
            Me.prg1.Maximum = TBLRekap.Rows.Count
            If TBLRekap.Rows.Count <> Me.ListCodeNo.Count Then
                'jika ada item yang tidak cocok - ada di db tapi tidak ada di dgv - maka tampilkan ke dalam dgv -> code no
                For Each rowRekap In TBLRekap.Rows
                    If rowRekap.Product_Code.Contains("'") = True Then
                        rowRekap.Product_Code = rowRekap.Product_Code.Replace("'", "")
                    End If
                    rowRekap.Product_Code = rowRekap.Product_Code.Trim

                    Me.prg1.Value += 1
                    If Me.prg1.Value > 25 Then
                        If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                            Me.lblProgress.Text = "Checking recap data consistency - " + rowRekap.Product_Code + " - data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                        ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                            Me.lblProgress.Text = "Checking recap data consistency - " + rowRekap.Product_Code + " - data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                        End If
                    Else
                        Me.lblProgress.Text = "Checking recap data consistency - " + rowRekap.Product_Code + " - data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                    End If

                    If ListCodeNo.Contains(rowRekap.Product_Code) = False Then
                        Me.vChkRst_CodeNo = rowRekap.Product_Code
                        Me.vChkRst_Result = "Code no not found in dgv rekap"
                        Me.bgw1.ReportProgress(0)

                        'Me.dgvCheckResult.Rows.Add()
                        'rowDGV = Me.dgvCheckResult.Rows(Me.dgvCheckResult.Rows.Count - 1)
                        'rowDGV.Cells(Me.ColCodeNo.Index).Value = rowRekap.Product_Code
                        'rowDGV.Cells(Me.ColCheckResult.Index).Value = "Code no not found in dgv rekap"
                    End If
                Next
            End If


            'banding jumlah data detail vs tabel item detail closing
            Me.lblProgress.Text = "Checking detail data consistency"
            Me.prg1.Value = 0
            DTADetail.Fill(TBLDetail, Me.ClosingPeriod)
            Me.prg1.Maximum = TBLDetail.Rows.Count
            For Each rowDetail In TBLDetail.Rows
                If rowDetail.Product_Code.Contains("'") = True Then
                    rowDetail.Product_Code = rowDetail.Product_Code.Replace("'", "")
                End If
                rowDetail.Product_Code = rowDetail.Product_Code.Trim

                Me.prg1.Value += 1
                Me.lblProgress.Text = "Checking detail data consistency - " + rowDetail.Product_Code + " - data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString

                If Me.GetJmlBarisDetail(rowDetail.Product_Code) <> rowDetail.JmlItem Then
                    Me.vChkRst_CodeNo = rowDetail.Product_Code
                    Me.vChkRst_Result = "Code no not found in dgv item/inconsistent item vs rekap data"
                    Me.bgw1.ReportProgress(0)

                    'Me.dgvCheckResult.Rows.Add()
                    'rowDGV = Me.dgvCheckResult.Rows(Me.dgvCheckResult.Rows.Count - 1)
                    'rowDGV.Cells(Me.ColCodeNo.Index).Value = rowDetail.Product_Code
                    'rowDGV.Cells(Me.ColCheckResult.Index).Value = "Code no not found in dgv item/inconsistent item vs rekap data"
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlBarisDetail(ByVal vCodeNo As String) As Double
        Dim rowDGV As DataGridViewRow
        Dim Keluar As Boolean = False
        Dim RetVal As Integer = 0

        Try
            For Each rowDGV In Me.DGVItemDetail.Rows
                If rowDGV.Cells(Me.FrmParent.ColItem_CodeNo.Index).Value = vCodeNo Then
                    RetVal += 1
                Else
                    If Keluar = True Then
                        Return RetVal
                    End If
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Form2_Cek_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DGVRekap = Me.FrmParent.DataGridView1
            Me.DGVItemDetail = Me.FrmParent.DataGridView2
            Me.DGVDetailWarehouse = Me.FrmParent.DataGridView3
            Me.DGVDetailUsing = Me.FrmParent.DataGridView4
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class