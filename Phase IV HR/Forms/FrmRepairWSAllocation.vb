'form ini digunakan untuk memasukkan data ws allocation berdasarkan
'data attendance karyawan. dibuat untuk handle masalah di taci lembur, ketika transfer master ws
Public Class FrmRepairWSAllocation
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Try
            If MsgBox("Start data update?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker1.RunWorkerAsync()
                'UpdateWSAllocation()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateWSAllocation()
        Dim DTA As New DseRepairWSAllocationTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseRepairWSAllocation.DataTable1DataTable
        Dim row, rowB4 As DseRepairWSAllocation.DataTable1Row
        Dim DTA1 As New DseRepairWSAllocationTableAdapters.tbltrans_attendanceTableAdapter
        Dim TBL1 As New DseRepairWSAllocation.tbltrans_attendanceDataTable
        Dim row1 As DseRepairWSAllocation.tbltrans_attendanceRow
        Dim DTA2 As New DseRepairWSAllocationTableAdapters.tblmaster_workshift_itemallocateTableAdapter
        Dim TBL2 As New DseRepairWSAllocation.tblmaster_workshift_itemallocateDataTable
        Dim row2 As DseRepairWSAllocation.tblmaster_workshift_itemallocateRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.Fill(TBL, Me.dtpEnd.Value.Date, Me.dtpStart.Value.Date)
            DTA1.Fill(TBL1, Me.dtpEnd.Value.Date, Me.dtpStart.Value.Date)

            TBL.DefaultView.Sort = "[primary karyawan], date"
            Me.prgBar.Value = 0
            Me.prgBar.Maximum = TBL1.Rows.Count
            For Each row1 In TBL1.Rows
                Me.prgBar.Value += 1
                Me.LblProgress.Text = "Updating data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString + " - " + row1.Full_Name
                TBL2.Rows.Clear()

                'hapus data lama
                DTA2.DeleteQuery1(row1.primary_karyawan, 0, Me.dtpEnd.Value.Date, Me.dtpStart.Value.Date)
                DTA2.DeleteQuery2(row1.primary_karyawan, 0, Me.dtpEnd.Value.Date, Me.dtpStart.Value.Date)

                'olah data
                TBL.DefaultView.RowFilter = "[primary karyawan] = " + row1.primary_karyawan.ToString
                For ij As Double = 0 To TBL.DefaultView.Count - 1
                    row = TBL.DefaultView(ij).Row
                    If ij = 0 Then
                        'tambahkan baris baru
                        row2 = TBL2.Newtblmaster_workshift_itemallocateRow()
                        With row2
                            .end_date = row._date
                            .primary = -1
                            .primary_entity = row.primary_karyawan
                            .primary_workshift = row.primary_workshift
                            .remarks = ""
                            .start_date = row._date
                            .tipe_entity = 0
                        End With
                        TBL2.Rows.Add(row2)
                    Else
                        'cek dengan baris sebelumnya, apakah primary workshift sama
                        rowB4 = TBL.DefaultView(ij - 1).Row
                        If row.primary_workshift = rowB4.primary_workshift Then
                            'primary ws sama, ganti end date
                            row2 = TBL2.Rows(TBL2.Rows.Count - 1)
                            row2.end_date = row._date
                        Else
                            'primary ws tidak sama, tambahkan baris baru
                            row2 = TBL2.Newtblmaster_workshift_itemallocateRow()
                            With row2
                                .end_date = row._date
                                .primary = -1
                                .primary_entity = row.primary_karyawan
                                .primary_workshift = row.primary_workshift
                                .remarks = ""
                                .start_date = row._date
                                .tipe_entity = 0
                            End With
                            TBL2.Rows.Add(row2)
                        End If
                    End If
                Next

                'simpan data
                For Each row2 In TBL2.Rows
                    With row2
                        DTA2.InsertQuery(.primary_entity, .tipe_entity, .start_date, .end_date, .primary_workshift, .remarks)
                    End With
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.ObjCD.UpdateThreading()
            Me.UpdateWSAllocation()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.LblProgress.Text = "Data updated"
            Me.prgBar.Value = 0
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmRepairWSAllocation_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DTA As New DseRepairWSAllocationTableAdapters.DataTable1TableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetMinDate()
            If RetVal.HasValue = True Then
                Me.dtpStart.Value = RetVal.Value.Date
                Me.dtpEnd.Value = DTA.ScalarQuery_GetMaxDate().Value.Date
            End If
            Me.LblProgress.Text = "Ready..."
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class