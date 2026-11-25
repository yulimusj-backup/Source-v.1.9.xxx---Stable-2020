Option Explicit On
Option Strict Off

Public Class FrmLaporanPurchaseRequisition_Epl_AVG
    Public PrimaryProduct As Double
    Public StartPeriod, EndPeriod As Date
    Public objCD As New Core_Module.ClsCekData
    Public Unit As String = ""

    Dim ObjProduct As New Core_Module.clsMasterProduct

    Dim DefPrec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency

    Private Sub FrmPurchaseReq_ExplainAvg_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim TotalDO, TotalHS, Periode As Decimal
        Dim rowdgv As DataGridViewRow

        Try
            Me.DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency

            Me.ExplainDO()
            Me.ExplainHistoricalSales(Me.StartPeriod, Me.objCD.GetLasDateOfTheMonth(Me.EndPeriod))

            'hitung total
            Me.LblUnit1.Text = Me.Unit
            Me.LblUnit2.Text = Me.Unit
            Me.LblUnit3.Text = Me.Unit
            Me.LblUnit4.Text = Me.Unit
            Periode = Me.GetJmlSelisihBulan(Me.StartPeriod, Me.EndPeriod) + 1
            Me.TxtTotalPeriode.Text = Periode
            For Each rowdgv In Me.dgvDO.Rows
                TotalDO += CDbl(rowdgv.Cells(Me.ColQty.Index).Value)
            Next
            Me.TxtTotalDO.Text = Me.objCD.DigitGrouping(TotalDO)

            For Each rowdgv In Me.dgvHS.Rows
                TotalHS += CDbl(rowdgv.Cells(Me.ColQtySales.Index).Value)
            Next
            Me.TxtTotalHS.Text = Me.objCD.DigitGrouping(TotalHS)
            Me.TxtTotal.Text = Me.objCD.DigitGrouping(TotalDO + TotalHS)
            Me.TxtAvgPerMonth.Text = Math.Round((TotalDO + TotalHS) / Periode, Me.DefPrec, MidpointRounding.AwayFromZero)

            Me.Text = "Average per month detail: " + ObjProduct.GetProductCode(Me.PrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ExplainDO()
        Dim Dummy, Dummy1 As Date
        Dim DTA As New DseLaporanPurchaseRequisitionTableAdapters.DataTable3TableAdapter
        Dim TBL As New DseLaporanPurchaseRequisition.DataTable3DataTable
        Dim Period As String = ""
        Dim rowDGV As DataGridViewRow
        Dim row As DseLaporanPurchaseRequisition.DataTable3Row

        Try
            Dummy = Me.StartPeriod
            While Dummy <= Me.EndPeriod
                Dummy1 = Me.objCD.GetLasDateOfTheMonth(Dummy)
                Period = Format(Dummy, "MMMMM yyyy")
                Me.dgvDO.Rows.Add()
                rowDGV = Me.dgvDO.Rows(Me.dgvDO.Rows.Count - 1)
                rowDGV.Cells(Me.ColPeriod.Index).Value = Period
                rowDGV.Cells(Me.ColUnits.Index).Value = Me.Unit

                If DTA.Fill(TBL, Dummy1, Dummy, Me.PrimaryProduct) > 0 Then
                    row = TBL.Rows(0)
                    rowDGV.Cells(Me.ColQty.Index).Value = Me.objCD.DigitGrouping(row.Total_Qty)
                Else
                    rowDGV.Cells(Me.ColQty.Index).Value = 0
                End If

                Dummy = Dummy1 + TimeSpan.FromDays(1)
            End While
            Me.objCD.SetNumber(Me.dgvDO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ExplainHistoricalSales(ByVal StartDate As Date, ByVal EndDate As Date)
        Dim ObjHP As New ClsTransaksiHistoricalPurchase
        Dim RefDate As Date = StartDate
        Dim HPStartDate, HPEndDate As Date
        Dim a, b As Double
        Dim LoopGo As Boolean = True
        Dim PrimaryHP As Double = -1
        Dim TBL As New DseLaporanPurchaseRequisition.tbltrans_historicalpurchase_itemDataTable
        Dim row As DseLaporanPurchaseRequisition.tbltrans_historicalpurchase_itemRow

        Dim DTA As New DseLaporanPurchaseRequisitionTableAdapters.tbltrans_historicalpurchase_itemTableAdapter
        DTA.ClearBeforeFill = True

        Dim NoTransaksi As String = ""
        Dim PeriodeHS As String = ""
        Dim TotalQty As String = ""
        Dim Formula As String = ""
        Dim QtySales As String = ""
        Dim rowDGVHS As DataGridViewRow

        Try
            'cari purchase history yang pertama
            PrimaryHP = ObjHP.GetPrimary(RefDate)
            If PrimaryHP = -1 Then
                RefDate = objCD.GetLasDateOfTheMonth(RefDate) + TimeSpan.FromDays(1)
                While RefDate < objCD.GetFirstDateofTheMonth(EndDate + TimeSpan.FromDays(1))
                    PrimaryHP = ObjHP.GetPrimary(RefDate)
                    If PrimaryHP = -1 Then
                        RefDate = objCD.GetLasDateOfTheMonth(RefDate) + TimeSpan.FromDays(1)
                    Else
                        Exit While
                    End If
                End While
            End If

            If PrimaryHP = -1 Then
                'tidak ada historical purchase pada periode yang diberikan
                Exit Sub
            Else
                While LoopGo = True
                    NoTransaksi = ObjHP.GetNoTransaksi(PrimaryHP)
                    'cari data historical purchase yang bersangkutan
                    DTA.FillBy_PrimaryProduct(TBL, PrimaryHP, Me.PrimaryProduct)

                    'hitung faktor yang digunakan untuk memindahkan data ke tabel purchase requisition
                    ObjHP.GetStartNEndDate(PrimaryHP, HPStartDate, HPEndDate)
                    If HPStartDate.Month = HPEndDate.Month And HPStartDate.Year = HPEndDate.Year Then
                        PeriodeHS = Format(HPStartDate, "MMMMM yyyy")
                    Else
                        PeriodeHS = Format(HPStartDate, "MMMMM yyyy") + " - " + _
                        Format(HPEndDate, "MMMMM yyyy")
                    End If
                    b = Me.GetJmlSelisihBulan(HPStartDate, HPEndDate) + 1
                    If HPStartDate <= StartDate And HPEndDate >= EndDate Then
                        'seluruh periode pr masuk ke dalam periode hp
                        a = Me.GetJmlSelisihBulan(RefDate, EndDate) + 1
                    ElseIf HPStartDate <= RefDate And HPEndDate <= objCD.GetFirstDateofTheMonth(EndDate) Then
                        'periode hp berada diakhir periode pr
                        a = Me.GetJmlSelisihBulan(RefDate, HPEndDate) + 1
                    ElseIf HPStartDate >= RefDate And HPEndDate >= objCD.GetFirstDateofTheMonth(EndDate) Then
                        'periode hp berada di awal periode pr
                        a = Me.GetJmlSelisihBulan(RefDate, EndDate) + 1
                    ElseIf HPStartDate >= StartDate And HPEndDate <= EndDate Then
                        '100 % masuk
                        a = b
                    End If

                    'pindahkan data ke tabel purchase requisition
                    Dim view As New Data.DataView
                    Dim cmd As String = ""
                    view = TBL.DefaultView

                    For Each row In TBL.Rows
                        cmd = "[Primary Product] = '" + CStr(row.Primary_Product) + "'"
                        view.RowFilter = cmd
                        Dim rowView As Data.DataRowView
                        For Each rowView In view
                            TotalQty = objCD.DigitGrouping(CDbl(rowView("Qty")))
                            QtySales = Math.Round(CDbl(rowView("Qty")) * a / b, _
                            0, MidpointRounding.AwayFromZero)
                            Formula = TotalQty + " * " + a.ToString + " / " + b.ToString
                        Next
                        Me.dgvHS.Rows.Add()
                        rowDGVHS = Me.dgvHS.Rows(Me.dgvHS.Rows.Count - 1)
                        rowDGVHS.Cells(Me.ColNoHS.Index).Value = Me.dgvHS.Rows.Count.ToString
                        rowDGVHS.Cells(Me.ColHSNumber.Index).Value = NoTransaksi
                        rowDGVHS.Cells(Me.ColPeriodHS.Index).Value = PeriodeHS
                        rowDGVHS.Cells(Me.ColFormula.Index).Value = Formula
                        rowDGVHS.Cells(Me.ColQtyTotalHS.Index).Value = objCD.DigitGrouping(TotalQty)
                        rowDGVHS.Cells(Me.ColQtySales.Index).Value = objCD.DigitGrouping(QtySales)
                    Next

                    'ganti reference date, untuk mendapatkan data dari pr berikutnya
                    RefDate = objCD.GetLasDateOfTheMonth(HPEndDate) + TimeSpan.FromDays(1)

                    'cari historical purchase yang selanjutnya
                    If RefDate > EndDate Then
                        LoopGo = False
                    Else
                        PrimaryHP = ObjHP.GetPrimary(RefDate)
                        If PrimaryHP = -1 Then
                            RefDate = objCD.GetLasDateOfTheMonth(RefDate) + TimeSpan.FromDays(1)
                            While RefDate <> objCD.GetFirstDateofTheMonth(EndDate + TimeSpan.FromDays(1))
                                PrimaryHP = ObjHP.GetPrimary(RefDate)

                                If PrimaryHP = -1 Then
                                    RefDate = objCD.GetLasDateOfTheMonth(RefDate) + TimeSpan.FromDays(1)
                                Else
                                    Exit While
                                End If
                            End While
                        End If

                        If PrimaryHP <> -1 Then
                            LoopGo = True
                        Else
                            LoopGo = False
                        End If
                    End If
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjHP = Nothing
            DTA.Dispose()
        End Try
    End Sub

    Private Sub TambahPenjualanHistorical(ByVal StartDate As Date, ByVal EndDate As Date, _
    ByVal DS As DseLaporanPurchaseRequisition)

    End Sub

    Private Function GetJmlSelisihBulan(ByVal StartDate As Date, ByVal EndDate As Date) As Integer
        Try
            Dim StartYear, EndYear, StartMonth, EndMonth As Integer
            StartYear = StartDate.Year
            EndYear = EndDate.Year
            StartMonth = StartDate.Month
            EndMonth = EndDate.Month

            'pastikan data start date lebih kecil dari enddate (dalam hitungan bulan)
            If StartYear > EndYear Then
                Return -1
            ElseIf StartYear = EndYear And StartMonth > EndMonth Then
                Return -1
            End If

            'hitung selisih bulan
            Return ((EndYear - StartYear) * 12) + EndMonth - StartMonth
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class