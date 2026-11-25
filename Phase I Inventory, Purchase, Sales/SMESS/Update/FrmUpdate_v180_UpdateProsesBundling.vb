Option Explicit On
Option Strict Off

Public Class FrmUpdate_v180_UpdateProsesBundling
    Dim Buffer As String = ""
    Dim ObjWorkUnit As New Core_Module.ClsMasterWorkUnit
    Dim ObjProses As New Core_Module.ClsMasterProses
    Dim BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim DTABundling1 As New DseUpdate_v180TableAdapters.DataTable1TableAdapter
        Dim TBLBundling1 As New DseUpdate_v180.DataTable1DataTable
        Dim rowBundling1 As DseUpdate_v180.DataTable1Row
        Dim DTAAcuan As New DseUpdate_v180TableAdapters.tbltrans_custom_update_bundlingTableAdapter
        Dim TBLAcuan As New DseUpdate_v180.tbltrans_custom_update_bundlingDataTable
        Dim DTAAcuan1 As New DseUpdate_v180TableAdapters.DataTable2TableAdapter
        Dim TBLAcuan1 As New DseUpdate_v180.DataTable2DataTable
        Dim rowAcuan1 As DseUpdate_v180.DataTable2Row
        Dim rowDGV As DataGridViewRow
        Dim Lengkap As Boolean = True

        Try
            If MsgBox("Start proces update?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                Exit Sub
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'load ualng dgv dari db
            'cek kode category yang ada di bundling dengan primary work unit =-1 atau primary process =-1
            'jika masih ada yang belum lengkap dalam dgv, maka tampilkan dan minta user melengkapi
            DTABundling1.Fill(TBLBundling1)
            DTAAcuan.Fill(TBLAcuan)
            For Each rowBundling1 In TBLBundling1
                TBLAcuan.DefaultView.RowFilter = "[primary category] = " + rowBundling1.Primary_Category.ToString
                If TBLAcuan.DefaultView.Count = 0 Then
                    'tambahkan baris baru di tabel acuan
                    DTAAcuan.InsertQuery(rowBundling1.Primary_Category, rowBundling1.primary_process, rowBundling1.primary_work_unit)
                End If
            Next

            'cek baris dalam dgv, pastikan tidak ada primary work unit dan primary process =-1
            DTAAcuan1.Fill(TBLAcuan1)
            Me.dgv1.Rows.Clear()
            Me.dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgv1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each rowAcuan1 In TBLAcuan1.Rows
                Me.dgv1.Rows.Add()
                rowDGV = Me.dgv1.Rows(Me.dgv1.Rows.Count - 1)
                rowDGV.Cells(Me.ColCategoryCode.Index).Value = rowAcuan1.Category_Code
                rowDGV.Cells(Me.ColPrimaryCategory.Index).Value = rowAcuan1.primary_category
                rowDGV.Cells(Me.ColPrimaryProcess.Index).Value = rowAcuan1.primary_process
                rowDGV.Cells(Me.ColPrimaryWorkUnit.Index).Value = rowAcuan1.primary_work_unit
                rowDGV.Cells(Me.ColProcessCode.Index).Value = rowAcuan1.Process_Code
                rowDGV.Cells(Me.ColWorkUnitCode.Index).Value = rowAcuan1.work_unit_code
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgv1.Rows.Count
            Next
            Me.dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'jika semua sudah lengkap, maka update ke db sesuai data yang ada
            For Each rowDGV In Me.dgv1.Rows
                If rowDGV.Cells(Me.ColWorkUnitCode.Index).Value = "" Or rowDGV.Cells(Me.ColProcessCode.Index).Value = "" Then
                    Lengkap = False
                End If
            Next
            If Lengkap = False Then
                MsgBox("Mohon lengkapi data dahulu.")
                Exit Sub
            End If

            'sampai sini, mulai proses updatenya.
            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAAcuan1.Dispose()
            TBLAcuan1.Dispose()
            DTABundling1.Dispose()
            TBLBundling1.Dispose()
            DTAAcuan.Dispose()
            TBLAcuan.Dispose()
        End Try
    End Sub

    Private Sub dgv1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgv1.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColProcessCode.Index Or e.ColumnIndex = Me.ColWorkUnitCode.Index Then
                Me.Buffer = Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgv1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgv1.Rows(e.RowIndex)
        Dim Index As Integer = -1
        Dim s As String = ""
        Dim DTAAcuan As New DseUpdate_v180TableAdapters.tbltrans_custom_update_bundlingTableAdapter

        Try
            With Row
                If IsNothing(Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If

                    ''-----Cell kode work unit-----
                    If e.ColumnIndex = Me.ColWorkUnitCode.Index Then
                        Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjWorkUnit.ApakahSudahAda(Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode unit yang diketik ada --> load data kode work unit + primary
                            Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper.Trim
                            Me.dgv1.Item(Me.ColPrimaryWorkUnit.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetPrimary(Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value)
                        Else
                            '2. kode unit yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New Core_Module.FrmMasterWorkUnit_Cari
                            ij.BackColorAll = Me.BackColorAll
                            ij.txtCari.Text = Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()

                            If ij.PrimaryWorkUnit <> -1 Then
                                '2.a. kode unit yang dicari ada --> load data deskripsi + primary unit + kode unit
                                Me.dgv1.Item(Me.ColWorkUnitCode.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetCode(ij.PrimaryWorkUnit)
                                Me.dgv1.Item(Me.ColPrimaryWorkUnit.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
                            ElseIf ij.PrimaryWorkUnit = -1 Then
                                Me.dgv1.Item(Me.ColWorkUnitCode.Index, e.RowIndex).Value = Me.Buffer
                            End If
                            ij = Nothing
                        End If
                    ElseIf e.ColumnIndex = Me.ColProcessCode.Index Then
                        'kode process
                        Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjProses.ApakahSudahAda(Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode proses yang diketik ada --> load data kode proses + primary
                            Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper.Trim
                            Me.dgv1.Item(Me.ColPrimaryProcess.Index, e.RowIndex).Value = Me.ObjProses.GetPrimary(Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value)
                        Else
                            '2. kode proses yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New Core_Module.FrmMasterProsesCari
                            ij.BackColorAll = Me.BackColorAll
                            ij.txtCari.Text = Me.dgv1.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()

                            If ij.mPrimary <> -1 Then
                                '2.a. kode proses yang dicari ada --> load data deskripsi + primary unit + kode proses
                                Me.dgv1.Item(Me.ColProcessCode.Index, e.RowIndex).Value = Me.ObjProses.GetCode(ij.mPrimary)
                                Me.dgv1.Item(Me.ColPrimaryProcess.Index, e.RowIndex).Value = ij.mPrimary
                            ElseIf ij.mPrimary = -1 Then
                                Me.dgv1.Item(Me.ColProcessCode.Index, e.RowIndex).Value = Me.Buffer
                            End If
                            ij = Nothing
                        End If
                    End If
                End If

                'jika data wotk unit dan process code tepat, maka update ke db
                If Me.dgv1.Item(Me.ColPrimaryWorkUnit.Index, e.RowIndex).Value <> -1 And _
                Me.dgv1.Item(Me.ColPrimaryProcess.Index, e.RowIndex).Value <> -1 Then
                    DTAAcuan.UpdateQuery(Me.dgv1.Item(Me.ColPrimaryProcess.Index, e.RowIndex).Value, Me.dgv1.Item(Me.ColPrimaryWorkUnit.Index, e.RowIndex).Value, _
                    Me.dgv1.Item(Me.ColPrimaryCategory.Index, e.RowIndex).Value)
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.LblProgress.Text = ""
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim DTAAcuan As New DseUpdate_v180TableAdapters.tbltrans_custom_update_bundlingTableAdapter
        Dim TBLAcuan As New DseUpdate_v180.tbltrans_custom_update_bundlingDataTable
        Dim rowAcuan As DseUpdate_v180.tbltrans_custom_update_bundlingRow
        Dim ListPrimaryTransaksi As New List(Of Double)
        Dim DTA As New DseUpdate_v180TableAdapters.DataTable3TableAdapter
        Dim TBL As New DseUpdate_v180.DataTable3DataTable
        Dim row As DseUpdate_v180.DataTable3Row

        Try
            'load semua data transaksi bundling yang primary work unit/primary process =-1
            DTAAcuan.Fill(TBLAcuan)
            DTA.Fill(TBL)
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                Me.ProgressBar1.Value += 1

                If ListPrimaryTransaksi.Contains(row.Primary_Transaksi) = False Then
                    ListPrimaryTransaksi.Add(row.Primary_Transaksi)
                    TBLAcuan.DefaultView.RowFilter = "[Primary Category] = " + row.Primary_Category.ToString
                    If TBLAcuan.DefaultView.Count = 0 Then
                        Beep()
                        MsgBox("Logic error")
                    Else
                        rowAcuan = TBLAcuan.DefaultView(0).Row
                        DTA.UpdateQuery(rowAcuan.primary_process, rowAcuan.primary_work_unit, row.Primary_Transaksi)
                    End If
                End If
            Next

            'cocokan datanya berdasarkan tabel acuan

            'update ke dalam db

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryTransaksi = Nothing
            DTAAcuan.Dispose()
            TBLAcuan.Dispose()
        End Try
    End Sub
End Class