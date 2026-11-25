Option Explicit On
Option Strict Off

Public Class Form2_CekBundling
    Dim ObjCD As New Core_Module.ClsCekData

    Dim DTA As New DataSet2_BundlingTableAdapters.tbltransaksi_stockbundling1TableAdapter
    Dim TBL As New DataSet2_Bundling.tbltransaksi_stockbundling1DataTable
    Dim row As DataSet2_Bundling.tbltransaksi_stockbundling1Row
    Dim DTAItem As New DataSet2_BundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
    Dim TBLItem As New DataSet2_Bundling.tbltransaksi_stockbundling_item_beforeDataTable
    Dim rowItem As DataSet2_Bundling.tbltransaksi_stockbundling_item_beforeRow

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If MsgBox("Check suspisious bundling on " + Format(Me.dtpPeriod.Value, "MMM yyyy"), MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.Rows.Clear()
                Me.bgw1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Dim StartDate, EndDate As Date
        Dim CurIdx As Integer = 0

        Try
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
            EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date)

            Me.TBL.Rows.Clear()
            Me.TBLItem.Rows.Clear()
            'fill tabel bundling
            Me.LblProgress.Text = "Preparing data"
            Me.DTA.Fill(Me.TBL, EndDate, StartDate)
            'fill item before
            Me.DTAItem.FillBy_TglSampaiDari(Me.TBLItem, EndDate, StartDate)

            'untuk setiap baris item bundling, cari yang: cuma terdiri 1 baris dan qtynya = 1
            Me.LblProgress.Text = "Checking Data"
            Me.prg1.Value = 0
            Me.prg1.Maximum = Me.TBL.Rows.Count
            While CurIdx <= Me.TBL.Rows.Count - 1
                Me.prg1.Value += 1
                If Me.prg1.Value > 25 Then
                    If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Checking Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                    ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                        Me.LblProgress.Text = "Checking Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Checking Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                End If

                Me.row = Me.TBL.Rows(CurIdx)
                Me.TBLItem.DefaultView.RowFilter = "[Primary Transaksi] = " + Me.row.Primary.ToString
                If Me.TBLItem.DefaultView.Count <= 1 Then
                    Me.rowItem = Me.TBLItem.DefaultView(0).Row
                    If Me.rowItem.Qty < 10 And rowItem.Primary_Units = 1 Then
                        CurIdx += 1
                    Else
                        Me.TBL.Rows.RemoveAt(CurIdx)
                    End If
                Else
                    Me.TBL.Rows.RemoveAt(CurIdx)
                End If
                While Me.TBLItem.DefaultView.Count > 0
                    Me.TBLItem.Rows.Remove(Me.TBLItem.DefaultView(0).Row)
                End While
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Form2_CekBundling_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dtpPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.LblProgress.Text = "Ready"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.Rows.Clear()
            While Me.TBL.Rows.Count > 0
                Me.row = Me.TBL.Rows(0)
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(Me.ColNoBundling.Index).Value = Me.row.No_Transaksi
                Me.TBL.Rows.Remove(Me.row)
            End While

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.LblProgress.Text = "Ready"
            Me.prg1.Value = 0
            MsgBox("Data checked", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub prg1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles prg1.Click

    End Sub
End Class