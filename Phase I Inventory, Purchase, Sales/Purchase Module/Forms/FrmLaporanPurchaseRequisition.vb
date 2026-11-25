Option Explicit On
Option Strict Off

Public Class FrmLaporanPurchaseRequisition
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCd As New Core_Module.ClsCekData

    Dim ObjProd As New Core_Module.clsMasterProduct

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Public Function GetDescProductGroup(ByVal vPrimary As Double) As String
        Dim DTA As New DseLaporanPurchaseRequisitionTableAdapters.tblmaster_product_groupTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_getDescByPrimary(CStr(vPrimary))) = False Then
                Return DTA.ScalarQuery_getDescByPrimary(CStr(vPrimary))
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub txtBrandCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBrandCode.Leave
        Try
            If Me.txtBrandCode.Text <> "" Then
                Dim obj As New Core_Module.ClsMasterProductGroup
                Dim primary As Int16 = obj.GetPrimary(Me.txtBrandCode.Text, 2)
                obj = Nothing
                If primary = -1 Then
                    Me.btnFindBrand.PerformClick()
                Else
                    Me.txtBrandName.Text = Me.GetDescProductGroup(primary)
                    Me.txtPrimaryBrand.Text = CStr(primary)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindBrand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindBrand.Click
        Try
            Dim i As New Core_Module.FrmCariMasterProductGroup
            i.ProductGroupLevel = 2
            i.BackColorAll = Me.BackColorAll

            If Me.txtBrandCode.Text <> "" Then
                i.txtCari.Text = Me.txtBrandCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryProductGroup > 0 Then
                Me.txtPrimaryBrand.Text = i.PrimaryProductGroup
                Dim objproductGroup As New Core_Module.ClsMasterProductGroup
                Me.txtBrandCode.Text = objproductGroup.GetGroupCode(Me.txtPrimaryBrand.Text)
                Me.txtBrandName.Text = Me.GetDescProductGroup(Me.txtPrimaryBrand.Text)
                objproductGroup = Nothing
                Me.dtpFrom.Focus()
            Else
                Me.txtPrimaryBrand.Text = "-1"
                Me.txtBrandName.Text = ""
                Me.txtBrandCode.Text = ""
                Me.txtBrandCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBrandCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBrandCode.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Dim obj As New Core_Module.ClsMasterProductGroup
                Dim primary As Int16 = obj.GetPrimary(Me.txtBrandCode.Text, 2)
                obj = Nothing
                If primary = -1 Then
                    Me.btnFindBrand.PerformClick()
                Else
                    Me.txtPrimaryBrand.Text = CStr(primary)
                    Me.txtBrandName.Text = Me.GetDescProductGroup(Me.txtPrimaryBrand.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBrandCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBrandCode.GotFocus
        txtBrandCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtBrandCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBrandCode.LostFocus
        txtBrandCode.BackColor = Me.txtPrimaryBrand.BackColor
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Me.Cetak()
    End Sub

    Private Function HitungJumlahBulan() As Integer
        Try
            Dim Sisabulan As Integer = 0
            If Me.dtpTo.Value.Year > Me.dtpFrom.Value.Year Then
                Dim SelisihThn As Integer = CInt(Me.dtpTo.Value.Year) - CInt(Me.dtpFrom.Value.Year)
                Dim SisaBulanThnIni As Integer = 12 * SelisihThn - CInt(Me.dtpFrom.Value.Month) + 1
                Sisabulan = SisaBulanThnIni + CInt(Me.dtpTo.Value.Month)
            ElseIf Me.dtpTo.Value.Year < Me.dtpFrom.Value.Year Then
                MsgBox("Please check your date.", MsgBoxStyle.Information, "Information")
                Me.dtpTo.Focus()
            Else
                Sisabulan = CInt(Me.dtpTo.Value.Month) - CInt(Me.dtpFrom.Value.Month) + 1
            End If
            Return Sisabulan
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Cetak()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim DS As New DseLaporanPurchaseRequisition
        Dim DTA As New DseLaporanPurchaseRequisitionTableAdapters.laporan_purchase_requisition_view2TableAdapter
        Dim DTADO As New DseLaporanPurchaseRequisitionTableAdapters.TblItemDOTableAdapter
        Dim row As DseLaporanPurchaseRequisition.laporan_purchase_requisition_view2Row
        Dim JmlBulan As Integer = 0
        Dim report As New CrpLaporanPurchaseRequisition
        Dim counter As Double = 0
        Dim DgvCol As DataGridViewColumn

        Try
            Dim objCD As New Core_Module.ClsCekData
            Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpTo.Value)
            Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
            objCD = Nothing
            Dim view As New Data.DataView
            Dim cmd As String = ""
            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            DTADO.Fill(DS.TblItemDO, TglSampai, TglDari, Me.txtPrimaryBrand.Text)
            view = DS.TblItemDO.DefaultView
            JmlBulan = Me.HitungJumlahBulan
            DTA.FillBy_PrimaryBrand(DS.laporan_purchase_requisition_view2, Me.txtPrimaryBrand.Text)

            If Me.CheckBox1.Checked = True Then
                While counter < DS.laporan_purchase_requisition_view2.Rows.Count
                    row = DS.laporan_purchase_requisition_view2.Rows(counter)
                    If row.PR_Estimation >= 0 Then
                        DS.laporan_purchase_requisition_view2.Rows.Remove(row)
                    Else
                        counter += 1
                    End If
                End While
            End If

            'tambahkan penjualan dari historical sales
            Me.TambahPenjualanHistorical(TglDari, TglSampai, DS)

            If DS.laporan_purchase_requisition_view2.Rows.Count > 0 Then
                For Each row In DS.laporan_purchase_requisition_view2.Rows
                    'tambahkan penjualan dari do
                    cmd = "[Primary Product] = '" + CStr(row.Primary_Product) + "'"
                    view.RowFilter = cmd
                    Dim rowView As Data.DataRowView
                    Dim TotalQty As Int64 = 0
                    For Each rowView In view
                        TotalQty += CDbl(rowView("Total Qty"))
                    Next
                    row.Avg_Per_Month = row.Avg_Per_Month + TotalQty

                    row.Avg_Per_Month = Math.Round(row.Avg_Per_Month / JmlBulan, CInt(Me.txtRoundingAvgPerMonth.Text), MidpointRounding.AwayFromZero)
                    row.PR_Estimation = Math.Round((row.Qty_On_Hand + row.Total_GIT + row.Purchase_Order_Balance) - (row.Avg_Per_Month * row.Lead_Time / 30), CInt(Me.txtRoundingPR.Text), MidpointRounding.AwayFromZero)
                Next
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_BrandCode.ParameterFieldName, Me.txtBrandCode.Text)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.txtFormatDay.Text)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.txtFormatDecimal.Text)
            report.SetParameterValue(report.Parameter_FormatLeadTime.ParameterFieldName, Me.txtFormatLeadTime.Text)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.txtFormatMonth.Text)
            report.SetParameterValue(report.Parameter_FormatPR.ParameterFieldName, Me.txtRoundingPR.Text)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.txtFormatYears.Text)
            report.SetParameterValue(report.Parameter_TglDari.ParameterFieldName, Me.dtpFrom.Value)
            report.SetParameterValue(report.Parameter_TglSampai.ParameterFieldName, Me.dtpTo.Value)
            report.SetParameterValue(report.Parameter_FormatAverage.ParameterFieldName, Me.txtRoundingAvgPerMonth.Text)

            Me.DataGridView1.DataSource = DS.laporan_purchase_requisition_view2
            Me.DataGridView1.Refresh()
            If Me.DataGridView1.ColumnCount > 0 Then
                Me.DataGridView1.Columns(0).Visible = False
                Me.DataGridView1.Columns(8).HeaderText = "Lead Time (days)"
                Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "N0"
                Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "N0"
                Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "N0"
                Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "N0"
                Me.DataGridView1.Columns(5).DefaultCellStyle.Format = "N0"
                Me.DataGridView1.Columns(4).DefaultCellStyle.Format = "N0"
                Me.DataGridView1.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.DataGridView1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                For Each DgvCol In Me.DataGridView1.Columns
                    DgvCol.SortMode = DataGridViewColumnSortMode.NotSortable
                Next
            End If

            'For Each row In DS.laporan_purchase_requisition_view2.Rows
            'Me.DataGridView1.Rows.Add(row.Primary_Product, row.Product_Code, _
            'row.Product_Description, row.Unit_Code, row.Qty_On_Hand, row.Qty_GIT, _
            'row.Purchase_Order_Balance, row.Avg_Per_Month, row.Lead_Time, row.PR_Estimation)
            'Next

            Me.crvReport.ReportSource = report
            Me.crvReport.DisplayStatusBar = True
            Me.crvReport.DisplayToolbar = True
            Me.crvReport.DisplayGroupTree = False
            Me.crvReport.ShowGroupTreeButton = False
            Me.crvReport.Show()
            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTADO.Dispose()
            DS.Dispose()
            '  report.Dispose()
        End Try
    End Sub

    Private Sub TambahPenjualanHistorical(ByVal StartDate As Date, ByVal EndDate As Date, _
    ByVal DS As DseLaporanPurchaseRequisition)
        Dim ObjHP As New ClsTransaksiHistoricalPurchase
        Dim ObjCD As New Core_Module.ClsCekData
        Dim RefDate As Date = StartDate
        Dim HPStartDate, HPEndDate As Date
        Dim a, b As Double
        Dim LoopGo As Boolean = True
        Dim PrimaryHP As Double = -1
        Dim TBL As DseLaporanPurchaseRequisition.tbltrans_historicalpurchase_itemDataTable = DS.tbltrans_historicalpurchase_item
        Dim row As DseLaporanPurchaseRequisition.laporan_purchase_requisition_view2Row

        Dim DTA As New DseLaporanPurchaseRequisitionTableAdapters.tbltrans_historicalpurchase_itemTableAdapter
        DTA.ClearBeforeFill = True

        Try
            'cari purchase history yang pertama
            PrimaryHP = ObjHP.GetPrimary(RefDate)
            If PrimaryHP = -1 Then
                RefDate = ObjCD.GetLasDateOfTheMonth(RefDate) + TimeSpan.FromDays(1)
                While RefDate <> ObjCD.GetFirstDateofTheMonth(EndDate + TimeSpan.FromDays(1))
                    PrimaryHP = ObjHP.GetPrimary(RefDate)
                    If PrimaryHP = -1 Then
                        RefDate = ObjCD.GetLasDateOfTheMonth(RefDate) + TimeSpan.FromDays(1)
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
                    'cari data historical purchase yang bersangkutan
                    DTA.FillBy_PrimaryHP(TBL, PrimaryHP)

                    'hitung faktor yang digunakan untuk memindahkan data ke tabel purchase requisition
                    ObjHP.GetStartNEndDate(PrimaryHP, HPStartDate, HPEndDate)
                    b = Me.GetJmlSelisihBulan(HPStartDate, HPEndDate) + 1
                    If HPStartDate <= StartDate And HPEndDate >= EndDate Then
                        'seluruh periode pr masuk ke dalam periode hp
                        a = Me.GetJmlSelisihBulan(RefDate, EndDate) + 1
                    ElseIf HPStartDate <= RefDate And HPEndDate <= ObjCD.GetFirstDateofTheMonth(EndDate) Then
                        'periode hp berada diakhir periode pr
                        a = Me.GetJmlSelisihBulan(RefDate, HPEndDate) + 1
                    ElseIf HPStartDate >= RefDate And HPEndDate >= ObjCD.GetFirstDateofTheMonth(EndDate) Then
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

                    For Each row In DS.laporan_purchase_requisition_view2.Rows
                        cmd = "[Primary Product] = '" + CStr(row.Primary_Product) + "'"
                        view.RowFilter = cmd
                        Dim rowView As Data.DataRowView
                        For Each rowView In view
                            row.Avg_Per_Month += Math.Round(CDbl(rowView("Qty")) * a / b, _
                            0, MidpointRounding.AwayFromZero)
                        Next
                    Next

                    'ganti reference date, untuk mendapatkan data dari pr berikutnya
                    RefDate = ObjCD.GetLasDateOfTheMonth(HPEndDate) + TimeSpan.FromDays(1)

                    'cari historical purchase yang selanjutnya
                    If RefDate > EndDate Then
                        LoopGo = False
                    Else
                        PrimaryHP = ObjHP.GetPrimary(RefDate)
                        If PrimaryHP = -1 Then
                            RefDate = ObjCD.GetLasDateOfTheMonth(RefDate) + TimeSpan.FromDays(1)
                            While RefDate <> ObjCD.GetFirstDateofTheMonth(EndDate + TimeSpan.FromDays(1))
                                PrimaryHP = ObjHP.GetPrimary(RefDate)

                                If PrimaryHP = -1 Then
                                    RefDate = ObjCD.GetLasDateOfTheMonth(RefDate) + TimeSpan.FromDays(1)
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
            ObjCD = Nothing
            DTA.Dispose()
        End Try
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

    Private Sub FrmLaporanPurchaseRequisition_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dtpTo.Value = Me.ObjCd.GetFirstDateofTheMonth(Me.dtpTo.Value)
            Me.dtpFrom.Value = Me.ObjCd.GetFirstDateofTheMonth(Me.dtpFrom.Value)
            Me.DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCodeNo.Click
        Dim i As New Core_Module.frmCariMasterProduct
        Dim PrimaryProduct As Double = -1
        Dim rowdgv As DataGridViewRow

        Try
            'cari product yang akan diset posisinya
            i.cboCriteria.SelectedIndex = 0
            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                PrimaryProduct = i.PrimaryProduct
                'cari baris yang berikan product tsb
                For Each rowdgv In Me.DataGridView1.Rows
                    If rowdgv.Cells(0).Value = PrimaryProduct Then
                        'fokus ke baris yang bersangkutan
                        Me.DataGridView1.FirstDisplayedScrollingRowIndex = rowdgv.Index
                        Me.DataGridView1.ClearSelection()
                        rowdgv.Selected = True
                    End If
                Next

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Try
            If e.ColumnIndex = 5 Or e.ColumnIndex = 4 Then
                Me.ExplainStock(Me.DataGridView1.Item(0, e.RowIndex).Value)
            ElseIf e.ColumnIndex = 6 Then
                Me.EplainPOB(Me.DataGridView1.Item(0, e.RowIndex).Value)
            ElseIf e.ColumnIndex = 7 Then
                Me.Explain_AveragePerMonth(Me.DataGridView1.Item(0, e.RowIndex).Value, _
                Me.DataGridView1.Item(3, e.RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Explain_AveragePerMonth(ByVal vPrimaryProduct As Double, ByVal UnitCode As String)
        Dim ij As New FrmLaporanPurchaseRequisition_Epl_AVG

        Try
            ij.PrimaryProduct = vPrimaryProduct
            ij.Unit = UnitCode
            ij.StartPeriod = Me.dtpFrom.Value.Date
            ij.EndPeriod = Me.dtpTo.Value.Date
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub EplainPOB(ByVal vPrimaryProduct As Double)
        Dim ij As New FrmLaporanPurchaseRequisition_Epl_POB

        Try
            ij.PrimaryProduct = vPrimaryProduct
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub ExplainStock(ByVal vPrimaryProduct As Double)
        Dim i As New Inventory_Module.FrmMasterKondisiStock

        Try
            i.txtCodeNo.Text = Me.ObjProd.GetProductCode(vPrimaryProduct)
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtBrandCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBrandCode.TextChanged

    End Sub
End Class