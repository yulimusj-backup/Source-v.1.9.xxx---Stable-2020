Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonthlyClosing_correction
    Public UserCode As String = ""
    Public BAckColorAll As Color = BackColor
    Public DefaultPrecision As Integer = 0
    Public vfrm As New FrmTransaksiMonthlyClosing2
    Public ObjCD As New Core_Module.ClsCekData
    Public TblA As New DseMonthlyClosingNew.Adjustment_CorrectionDataTable
    Public rowA As DseMonthlyClosingNew.Adjustment_CorrectionRow
    Public IsReadOnly As Boolean = False
    Public IsikanDataUnitPrice As Boolean = False
    Public RefPeriod As Date = #1/1/1950#
    Public ClosingPertama As Boolean = False

    Dim objCurrency As New Core_Module.clsMasterCurrency
    Dim CekBeforeClose As Boolean = True
    Dim Buffer As String = ""
    Dim baru As Boolean = True
    Dim ListPrimaryCodeNo_Bundling As New List(Of Double)
    Dim ObjClosing As New ClsTransaksiMonthlyClosing
    Dim TglMax As Date = Date.MinValue
    Dim TglMin As Date = Date.MaxValue

    Dim TipePrgBarStyle As Windows.Forms.ProgressBarStyle = ProgressBarStyle.Marquee
    Dim TipeProgress As EnumTipeProgress = EnumTipeProgress.ChangePrgBarStyle

    Enum EnumTipeProgress As Integer
        ChangePrgBarStyle
    End Enum

    'terkait unit price
    Dim DTAPeriod As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
    Dim TBLPeriod As New DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item_detailDataTable
    Dim RowPeriod As DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item_detailRow
    Dim DTAItemDetail As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing_item_detail1TableAdapter
    Dim TBLItemDetail As New DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item_detail1DataTable
    Dim RowItemDetail As DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item_detail1Row
    Dim SudahLoadAOP, SudahLoadSR, SudahLoadAdjust, SudahLoadBundling, SudahLoadStockTaking As Boolean
    Dim DTAAOP As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_aop1TableAdapter
    Dim TBLAOP As New DseTransaksiMonthlyClosing.tbltrans_aop1DataTable
    Dim RowAOP As DseTransaksiMonthlyClosing.tbltrans_aop1Row
    Dim DTASR As New DseTransaksiMonthlyClosingTableAdapters.tbltransaksi_stockinout3TableAdapter
    Dim TBLSR As New DseTransaksiMonthlyClosing.tbltransaksi_stockinout3DataTable
    Dim RowSR As DseTransaksiMonthlyClosing.tbltransaksi_stockinout3Row
    Dim TBLAdjust As New DseTransaksiMonthlyClosing.tbltransaksi_stockinout3DataTable
    Dim RowAdjust As DseTransaksiMonthlyClosing.tbltransaksi_stockinout3Row
    Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.tbltransaksi_stockbundlingTableAdapter
    Dim TBLBundling As New DseTransaksiMonthlyClosing.tbltransaksi_stockbundlingDataTable
    Dim RowBundling As DseTransaksiMonthlyClosing.tbltransaksi_stockbundlingRow
    Dim DTAStockTaking As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_stock_takingTableAdapter
    Dim TBLStockTaking As New DseTransaksiMonthlyClosing2.tbltrans_stock_takingDataTable
    Dim RowStockTaking As DseTransaksiMonthlyClosing2.tbltrans_stock_takingRow

    Private Sub FrmTransaksiMonthlyClosing_correction_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.CekBeforeClose = True Then
                If MsgBox("Cancel data entry?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    Me.CekBeforeClose = False
                    Me.Close()
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanForm()
        Try
            Me.ControlBox = True
            Me.dgvItemAdjust.Enabled = True
            Me.btnCancel.Enabled = True
            Me.btnOK.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanForm()
        Try
            Me.ControlBox = False
            Me.dgvItemAdjust.Enabled = False
            Me.btnCancel.Enabled = False
            Me.btnOK.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemAdjust_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemAdjust.CellBeginEdit
        Try
            If Me.IsReadOnly = True Then
                e.Cancel = True
            End If
            'yang bisa diedit hanya unit price dan remarks saja, untuk data yg bukan begining balance
            If Me.dgvItemAdjust.Item(Me.ColRemarksAdjust.Index, e.RowIndex).Value = "Begining Balance" And _
            ((Me.ObjClosing.GetJml = 0 And Me.baru = True) Or (Me.ObjClosing.GetJml = 1 And Me.baru = False)) And _
            e.ColumnIndex = Me.ColUnitPriceAdjust.Index Then
                e.Cancel = True
            Else
                If e.ColumnIndex = Me.ColUnitPriceAdjust.Index Or e.ColumnIndex = Me.ColRemarksAdjust.Index Then
                    Me.Buffer = Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value
                    If e.ColumnIndex = Me.ColUnitPriceAdjust.Index Then
                        Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erro")
        End Try
    End Sub

    Private Sub dgvItemAdjust_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemAdjust.CellEndEdit
        Dim JenisDocument As ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockBundlingAdjustment
        Dim vDefPrec As Integer = 0

        Try
            vDefPrec = Me.objCurrency.GetDecimal_DefaultCurrency

            'jika yang diedit unit price
            If e.ColumnIndex = Me.ColUnitPriceAdjust.Index Then
                'pastikan tidak nothing
                If IsNothing(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'pastikan nilainya tidak negatif
                If Me.ObjCD.CekNonNegatif(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'digit grouping
                Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefaultPrecision)

                'hitung ulang unit price untuk item yang bersangkutan
                Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value)
            Else
                If IsNothing(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvItemAdjust.Rows.Count > 0 Then
                Dim strCode As String = ""
                For Each rowDGV In Me.dgvItemAdjust.Rows
                    With rowDGV
                        If .Cells(Me.ColUnitPriceAdjust.Index).Value = 0 Then
                            If strCode = "" Then
                                strCode &= .Cells(Me.ColCodeNoAdjust.Index).Value
                            Else
                                strCode &= ", " & .Cells(Me.ColCodeNoAdjust.Index).Value
                            End If
                        End If
                    End With
                Next

                If strCode <> "" Then
                    MsgBox("Amount of Code No : " & strCode & " cannot fill by zero, please input correctly. ", MsgBoxStyle.Information, "Information")
                    Exit Sub
                Else
                    Me.TblA.Rows.Clear()
                    For Each rowDGV In Me.dgvItemAdjust.Rows
                        Me.rowA = Me.TblA.NewAdjustment_CorrectionRow
                        With Me.rowA
                            .Primary = -1
                            .Primary_Transaksi = -1
                            .Document = rowDGV.Cells(Me.ColJenisDocumentAdjustStr.Index).Value
                            .Document_No = rowDGV.Cells(Me.ColDocumentNo.Index).Value
                            .Code_No = rowDGV.Cells(Me.ColCodeNoAdjust.Index).Value
                            .Description = rowDGV.Cells(Me.ColDescAdjust.Index).Value
                            .Qty_In_Adjust = rowDGV.Cells(Me.ColQty.Index).Value
                            .Units = rowDGV.Cells(Me.ColUnitAdjust.Index).Value
                            .Unit_Price = rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value
                            .Primary_Code_No = rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value
                            .Primary_Unit = rowDGV.Cells(Me.ColPrimaryUnitAdjust.Index).Value
                            .Primary_Documen_Acuan = rowDGV.Cells(Me.ColPrimaryStockInAdjust.Index).Value
                            .Jenis_Dokumen = rowDGV.Cells(Me.ColJenisDokumen.Index).Value
                            .Remarks = rowDGV.Cells(Me.ColRemarksAdjust.Index).Value
                            .Jenis_Dokumen = rowDGV.Cells(Me.ColJenisDokumen.Index).Value
                            .Tanggal_Documen_Acuan = rowDGV.Cells(Me.ColTglDocAcuan.Index).Value
                            .Tipe_item_detail = rowDGV.Cells(Me.ColTipeItemDetail.Index).Value
                            .Primary_Gain_Account = rowDGV.Cells(Me.ColPrimaryGainAccount.Index).Value
                            .Primary_Loss_Account = rowDGV.Cells(Me.ColPrimaryLossAccount.Index).Value
                            .Primary_Adjustment_type = rowDGV.Cells(Me.ColPrimaryAdjustmentType.Index).Value
                            .Adjust_Type_Desc = rowDGV.Cells(Me.ColAdjustTypeDesc.Index).Value
                            .remarks_adjust = rowDGV.Cells(Me.ColRemarks_Adjust_detail.Index).Value
                            .DeffPrec = rowDGV.Cells(Me.ColDeffPrec.Index).Value
                            .Tipe_Transaksi_Asal = rowDGV.Cells(Me.ColTipeTransaksiAsal.Index).Value
                        End With
                        Me.TblA.Rows.Add(Me.rowA)
                    Next
                    Me.ObjCD.SetNumber(Me.dgvItemAdjust)
                End If
            End If

            If MsgBox("Accept in adjustment cost data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.CekBeforeClose = False
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonthlyClosing_correction_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.LblProgress.Text = ""
            Me.prgBar.Value = 0
            Me.DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTABundling.Connection.ConnectionString = Me.DTAAOP.Connection.ConnectionString
            Me.DTAItemDetail.Connection.ConnectionString = Me.DTAAOP.Connection.ConnectionString
            Me.DTAPeriod.Connection.ConnectionString = Me.DTAAOP.Connection.ConnectionString
            Me.DTASR.Connection.ConnectionString = Me.DTAAOP.Connection.ConnectionString
            Me.DTAStockTaking.Connection.ConnectionString = Me.DTAAOP.Connection.ConnectionString

            Me.dgvItemAdjust.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            'Me.LinkLabel1.Visible = False
            If Me.UserCode <> "USER" Then
                Me.LinkLabel1.Visible = False
            End If

            If Me.dgvItemAdjust.Rows.Count > 0 And Me.ClosingPertama = False And IsReadOnly = False Then
                Me.IsikanDataReferenceUnitPrice()
            Else
                Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim X As Decimal = 200
        Dim rowDGV As DataGridViewRow
        Dim Defprec As Integer = 0

        Try
            If MsgBox("Update data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                defprec = Me.objCurrency.GetDecimal_DefaultCurrency
                Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItemAdjust.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                For Each rowDGV In Me.dgvItemAdjust.Rows
                    If CDec(rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value) = 0 Then
                        rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(X, defprec)
                    End If
                Next
                Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        'Try
        '    If MsgBox("Cancel data entry?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
        '        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        '        Me.CekBeforeClose = False
        '        Me.Close()
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub dgvItemAdjust_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemAdjust.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemAdjust.Rows(Me.dgvItemAdjust.Rows.Count - 1)
            rowDGV.Cells(Me.ColReference.Index).Value = "Set"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDataReferenceUnitPrice()
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.NonAktifkanForm()
            Me.Panel1.Visible = True
            Me.bgw1.RunWorkerAsync()
            'Me.IsikanRefData()
            'Me.AktifkanForm()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanUnitPrice()
        Try
            If Me.TBLItemDetail.Rows.Count > 0 And Me.dgvItemAdjust.Rows.Count > 0 Then
                Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItemAdjust.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgw2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemAdjust_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemAdjust.CellContentClick
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            Me.ObjCD.UpdateThreading()
            Me.IsikanRefData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanRefData()
        Dim ListPrimary As New List(Of Double)
        Dim rowDGV As DataGridViewRow

        Try
            Me.TglMax = #1/1/1950#
            Me.TglMin = Date.MaxValue

            Me.LblProgress.Text = ""
            Me.prgBar.Value = 0

            'tentukan period max dan min
            Me.TipePrgBarStyle = ProgressBarStyle.Marquee
            Me.bgw1.ReportProgress(0)
            Me.LblProgress.Text = "Finding relevant code no"
            Me.DTAPeriod.Fill(Me.TBLPeriod, Me.RefPeriod - TimeSpan.FromDays(1))
            Me.TipePrgBarStyle = ProgressBarStyle.Continuous
            Me.bgw1.ReportProgress(0)
            Me.prgBar.Value = 0
            Me.prgBar.Maximum = Me.dgvItemAdjust.Rows.Count
            For Each rowDGV In Me.dgvItemAdjust.Rows
                Me.prgBar.Value += 1
                If Me.prgBar.Value > 25 Then
                    If Math.IEEERemainder(Me.prgBar.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Finding relevant code no - data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString
                    ElseIf Me.prgBar.Value = Me.prgBar.Maximum Then
                        Me.LblProgress.Text = "Finding relevant code no - data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Finding relevant code no - data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString
                End If

                If ListPrimary.Contains(rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value) = Boolean.FalseString Then
                    ListPrimary.Add(rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value)
                End If
            Next

            Me.LblProgress.Text = "Finding reference period"
            ListPrimary.Sort()
            Me.prgBar.Value = 0
            Me.prgBar.Maximum = ListPrimary.Count
            For Each PrimaryProduct As Double In ListPrimary
                Me.prgBar.Value += 1
                If Me.prgBar.Value > 25 Then
                    If Math.IEEERemainder(Me.prgBar.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Finding reference period - data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString
                    ElseIf Me.prgBar.Value = Me.prgBar.Maximum Then
                        Me.LblProgress.Text = "Finding reference period - data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Finding reference period - data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString
                End If

                If Me.TBLPeriod.Rows.Count > 0 Then
                    RowPeriod = Me.TBLPeriod.Rows(0)
                    If PrimaryProduct < RowPeriod.primary_code_no Then
                        'cek primary product selanjutnya, karena belum pernah di closing
                    Else
                        While PrimaryProduct > RowPeriod.primary_code_no
                            'buang baris pada tabel period, untuk mempercepat kinerja 
                            TBLPeriod.Rows.Remove(Me.RowPeriod)
                            If Me.TBLPeriod.Rows.Count > 0 Then
                                RowPeriod = Me.TBLPeriod.Rows(0)
                            Else
                                Exit For
                            End If
                        End While
                        If PrimaryProduct = RowPeriod.primary_code_no Then
                            If TglMin = #1/1/1950# Then
                                TglMin = RowPeriod.tgl_transaksi_asal.Date
                            Else
                                If TglMin.Date > RowPeriod.tgl_transaksi_asal.Date Then
                                    TglMin = RowPeriod.tgl_transaksi_asal.Date
                                End If
                            End If
                            If TglMax.Date < RowPeriod.tgl_transaksi_asal Then
                                TglMax = RowPeriod.tgl_transaksi_asal.Date
                            End If
                        End If
                        'buang baris pada tabel period, untuk mempercepat kinerja 
                        TBLPeriod.Rows.Remove(Me.RowPeriod)
                    End If
                End If
            Next

            'isikan item detail purchase closing yang tipe-nya in, mulai dari tgl min - tgl max
            If TglMax = #1/1/1950# Or TglMin = Date.MaxValue Then
                'do nothing, ga ada reference yang dicari
            Else
                Me.TipePrgBarStyle = ProgressBarStyle.Marquee
                Me.bgw1.ReportProgress(0)
                Me.LblProgress.Text = "Collecting reference data"
                Me.prgBar.Value = 0
                Me.TBLItemDetail.Rows.Clear()
                Me.DTAItemDetail.Fill(Me.TBLItemDetail, Me.TglMax, Me.TglMin)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub bgw1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgw1.ProgressChanged
        Try
            If Me.TipeProgress = EnumTipeProgress.ChangePrgBarStyle Then
                Me.prgBar.Style = Me.TipePrgBarStyle
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            If Me.TBLItemDetail.Rows.Count > 0 And Me.dgvItemAdjust.Rows.Count > 0 Then
                If Me.IsikanDataUnitPrice = Boolean.TrueString And IsReadOnly = False Then
                    Me.IsikanUnitPrice()
                Else
                    Me.AktifkanForm()
                    Me.Panel1.Visible = False
                End If
            Else
                Me.AktifkanForm()
                Me.Panel1.Visible = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw2.DoWork
        Dim ListPrimary As New List(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim Remarks As String = ""

        Try
            Me.ObjCD.UpdateThreading()

            'kumpulkan primary product
            For Each rowDGV In Me.dgvItemAdjust.Rows
                If ListPrimary.Contains(rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value) = False Then
                    ListPrimary.Add(rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value)
                End If
            Next

            'untuk setiap primary product:
            Me.LblProgress.Text = "Updating unit price" 
            Me.prgBar.Value = 0
            Me.prgBar.Maximum = ListPrimary.Count
            Me.TBLItemDetail.DefaultView.Sort = "[primary code no], [tgl transaksi asal] DESC, [tipe transaksi asal]"
            For Each PrimaryProduct As Double In ListPrimary
                Me.prgBar.Value += 1
                If Me.prgBar.Value > 25 Then
                    If Math.IEEERemainder(Me.prgBar.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Updating unit price - data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString
                    ElseIf Me.prgBar.Value = Me.prgBar.Maximum Then
                        Me.LblProgress.Text = "Updating unit price - data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Updating unit price - data " + Me.prgBar.Value.ToString + " of " + Me.prgBar.Maximum.ToString
                End If

                '1. tentukan baris acuan pada item detail
                Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + PrimaryProduct.ToString
                If Me.TBLItemDetail.DefaultView.Count > 0 Then
                    Me.RowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
                    Remarks = Me.GetRemarks(Me.RowItemDetail)
                    '2. update nilainya ke dgv
                    For Each rowDGV In Me.dgvItemAdjust.Rows
                        If rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = PrimaryProduct Then
                            rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(Me.RowItemDetail.unit_price, Me.DefaultPrecision)
                            rowDGV.Cells(Me.ColRemarksAdjust.Index).Value = Remarks
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Function GetRemarks(RowItemDetail As DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item_detail1Row) As String
        Dim TipeTrans As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal
        Dim RetVal As String = ""

        Try
            'pastikan data acuan sudah terisikan
            TipeTrans = RowItemDetail.tipe_transaksi_asal
            If TipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                If Me.SudahLoadAdjust = False Then
                    Me.SudahLoadAdjust = True
                    Me.DTASR.Fill(Me.TBLAdjust, Me.TglMax, Me.TglMin, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust)
                End If
                Me.TBLAdjust.DefaultView.RowFilter = "[Primary] = " + RowItemDetail.primary_transaksi_asal.ToString
                If TBLAdjust.DefaultView.Count > 0 Then
                    RowAdjust = TBLAdjust.DefaultView(0).Row
                    RetVal = "Refer to stock in adjust: " + RowAdjust.No_Transaksi + " - " + Format(RowAdjust.Tanggal_Transaksi, "dd MMM yyyy")
                End If
            ElseIf TipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                If Me.SudahLoadBundling = False Then
                    Me.SudahLoadBundling = True
                    Me.DTABundling.Fill(Me.TBLBundling, Me.TglMax, Me.TglMin)
                End If
                Me.TBLBundling.DefaultView.RowFilter = "[Primary] = " + RowItemDetail.primary_transaksi_asal.ToString
                If TBLBundling.DefaultView.Count > 0 Then
                    RowBundling = TBLBundling.DefaultView(0).Row
                    RetVal = "Refer to stock bundling: " + RowBundling.No_Transaksi + " - " + Format(RowBundling.Tanggal_Transaksi, "dd MMM yyyy")
                End If
            ElseIf TipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                If Me.SudahLoadAOP = False Then
                    Me.SudahLoadAOP = True
                    Me.DTAAOP.Fill(Me.TBLAOP, Me.TglMax, Me.TglMin)
                End If
                Me.TBLAOP.DefaultView.RowFilter = "[Primary] = " + RowItemDetail.primary_transaksi_asal.ToString
                If TBLAOP.DefaultView.Count > 0 Then
                    RowAOP = TBLAOP.DefaultView(0).Row
                    RetVal = "Refer to purchase: " + RowAOP.No_Transaksi + " - " + Format(RowAOP.Tanggal_Transaksi, "dd MMM yyyy") + " - " + RowAOP.Supplier_Name
                End If
            ElseIf TipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                If Me.SudahLoadSR = False Then
                    Me.SudahLoadSR = True
                    Me.DTASR.Fill(Me.TBLSR, Me.TglMax, Me.TglMin, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn)
                End If
                Me.TBLSR.DefaultView.RowFilter = "[Primary] = " + RowItemDetail.primary_transaksi_asal.ToString
                If TBLSR.DefaultView.Count > 0 Then
                    RowSR = TBLSR.DefaultView(0).Row
                    RetVal = "Refer to stock in sales return: " + RowSR.No_Transaksi + " - " + Format(RowSR.Tanggal_Transaksi, "dd MMM yyyy")
                End If
            ElseIf TipeTrans = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                If Me.SudahLoadStockTaking = False Then
                    Me.SudahLoadStockTaking = True
                    Me.DTAStockTaking.Fill(Me.TBLStockTaking)
                End If
                Me.TBLStockTaking.DefaultView.RowFilter = "[Primary] = " + RowItemDetail.primary_transaksi_asal.ToString
                If TBLStockTaking.DefaultView.Count > 0 Then
                    RowStockTaking = TBLStockTaking.DefaultView(0).Row
                    RetVal = "Refer to stock taking: " + RowStockTaking.No_Transaksi + " - " + Format(RowStockTaking.Start_Date, "dd MMM yyyy")
                End If
            End If

            If RetVal.Length > 250 Then
                RetVal = RetVal.Substring(0, 250)
            End If
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub bgw2_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgw2.ProgressChanged
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw2.RunWorkerCompleted
        Try
            Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItemAdjust.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Me.AktifkanForm()
            Me.Panel1.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class