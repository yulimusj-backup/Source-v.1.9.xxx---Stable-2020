Option Explicit On
Option Strict Off

Public Class FrmLaporanStockTaking_Cost1
    Dim TBLBrand As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable
    Dim report As New CrpLaporanStockTaking_Cost
    Dim DTA As New DseLaporanStockTakingBaruTableAdapters.DataTable1TableAdapter
    Dim DS As New DseLaporanStockTakingBaru
    Dim TBL As DseLaporanStockTakingBaru.DataTable1DataTable = DS.DataTable1
    Dim row As DseLaporanStockTakingBaru.DataTable1Row
    Dim ObjST As New clsTransaksiStockTaking
    Dim ObjCD As New Core_Module.ClsCekData
    Dim DTA_Taking As New DseLaporanStockTakingBaruTableAdapters.DataTable4TableAdapter
    Dim TBL_Taking As New DseLaporanStockTakingBaru.DataTable4DataTable
    Dim row_Taking As DseLaporanStockTakingBaru.DataTable4Row
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim DefPrec As Integer = 0
    Dim ObjLicense As New Core_Module.ClsLicense

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    ''Dim DTA1 As New dseLaporanStockTakingTableAdapters.tbltrans_stock_taking_item_analisis3TableAdapter
    ''Dim TBL1 As New dseLaporanStockTaking.tbltrans_stock_taking_item_analisis3DataTable
    ''Dim row1 As dseLaporanStockTaking.tbltrans_stock_taking_item_analisis3Row

    'Dim DTAClosing2 As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing_itemTableAdapter
    'Dim TBLClosing2 As New DseTransaksiMonthlyClosing.tbltrans_purchase_closing_itemDataTable
    'Dim rowClosing2 As DseTransaksiMonthlyClosing.tbltrans_purchase_closing_itemRow
    'Dim DTAClosing As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter

    Private Sub FrmLaporanStockTaking_Cost_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim OBjBrand As New Core_Module.ClsMasterProductGroup
        Try
            Me.cboFormat.SelectedIndex = 0
            OBjBrand.IsiListBoxBrand(Me.clbBrand, Me.TBLBrand)
            Me.TxtPrimaryStockTaking.Text = "-1"

            Me.Crv1.DisplayGroupTree = False
            Me.Crv1.DisplayStatusBar = True
            Me.Crv1.DisplayToolbar = True
            Me.Crv1.ShowGroupTreeButton = False
            Me.Crv1.ShowExportButton = True
            Me.Crv1.ShowRefreshButton = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            OBjBrand = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' fungsi ini melakukan pengisian data yang akan diberikan ke dalam crp. Asal data bisa 2, setelah stock taking diproses maka pengisian data dilakukan
    ''' berdasarkan item closing, jika stock belum di proses, maka asal data dari item analisis. Keduanya cocok pada qty on hand, qty taking, dan difference
    ''' hanya saja akan berbeda pada amount karena akan dipengaruhi closing
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadData()
        Dim DV As Data.DataRowView
        Dim PrimaryBrand As Double = 0
        Dim Filter As String = ""

        Try
            Me.TBL.Rows.Clear()
            Me.DTA.ClearBeforeFill = False

            'isikan data
            If Me.clbBrand.CheckedItems.Count = Me.clbBrand.Items.Count Then
                Me.DTA.Fill(Me.TBL, Me.TxtPrimaryStockTaking.Text)
            Else
                For Each DV In Me.clbBrand.CheckedItems
                    PrimaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
                    DTA.ClearBeforeFill = False
                    Me.DTA.FillBy_PrimaryBrand(TBL, Me.TxtPrimaryStockTaking.Text, PrimaryBrand)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakRekap()
        Dim Rpt As New CrpLaporanStockTaking_Cost_Rekap
        Dim Data As New DseLaporanStockTakingBaru
        Dim TblRK As DseLaporanStockTakingBaru.DataTable2DataTable = Data.DataTable2
        Dim RowRK, RowRK2 As DseLaporanStockTakingBaru.DataTable2Row
        Dim TBLDummy As New DseLaporanStockTakingBaru.DataTable2DataTable
        Dim Filter As String = ""
        Dim DV As Data.DataRowView

        Try
            Me.LoadData()
            ' Pindah in
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    Filter = "[Brand Code] = '" + row.Brand_Code + "'"
                    TblRK.DefaultView.RowFilter = Filter

                    If TblRK.DefaultView.Count > 0 Then
                        RowRK = TblRK.DefaultView(0).Row
                        RowRK.Sub_Total += row.sub_total
                    Else
                        RowRK = TblRK.NewRow
                        With RowRK
                            .Brand_Code = row.Brand_Code
                            .Brand_Name = row.Brand_Name
                            .Sub_Total = row.sub_total
                        End With
                        TblRK.AddDataTable2Row(RowRK)
                    End If
                Next
            End If

            'urutan, abjad ascending
            TblRK.DefaultView.RowFilter = ""
            TblRK.DefaultView.Sort = "[Brand Code]"
            For Each DV In TblRK.DefaultView
                RowRK = DV.Row
                RowRK2 = TBLDummy.NewDataTable2Row
                For i As Integer = 0 To TBLDummy.Columns.Count - 1
                    RowRK2.Item(i) = RowRK.Item(i)
                Next
                TBLDummy.Rows.Add(RowRK2)
            Next
            TblRK.Rows.Clear()
            For Each RowRK In TBLDummy.Rows
                RowRK2 = TblRK.NewDataTable2Row
                For i As Integer = 0 To TBLDummy.Columns.Count - 1
                    RowRK2.Item(i) = RowRK.Item(i)
                Next
                TblRK.Rows.Add(RowRK2)
            Next

            Dim Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer = Me.Crv1
            'tampilkan report
            Rpt.SetDataSource(Data)
            Rpt.SetParameterValue(Rpt.Parameter_Tanggal.ParameterFieldName, Me.dtpStockTakingDate.Value)

            Crv.ReportSource = Rpt
            Crv.ShowRefreshButton = False
            Crv.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakDetail()
        Dim counter As Double = 0
        Dim Filter As String = ""

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.LoadData()

            'untuk setiap item, tambahkan data analisis dan counter measures
            If Me.chkAnalysis.Checked = True Or Me.chkCounterMeasure.Checked = True Then
                For Each row In Me.TBL.Rows
                    row.analisis = row.analisis.Trim
                    row.counter_measures = row.counter_measures.Trim

                    If Me.chkAnalysis.Checked = True And row.analisis <> "" Then
                        row.Description = row.Description + vbCrLf + "Analysis: " + row.analisis
                    End If
                    If Me.chkCounterMeasure.Checked = True And row.counter_measures <> "" Then
                        row.Description = row.Description + vbCrLf + "Counter Measure: " + row.counter_measures
                    End If
                Next
            End If

            Dim Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer = Me.Crv1
            'tampilkan report
            report.SetDataSource(DS)
            report.SetParameterValue(Me.report.Parameter_StockTakingDate.ParameterFieldName, Me.dtpStockTakingDate.Value)

            Crv.ReportSource = report
            Crv.ShowRefreshButton = False
            Crv.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim counter As Double = 0
        Dim Filter As String = ""
        'Dim DV As Data.DataRowView
        Dim PrimaryBrand As Double = 0
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpStockTakingDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpStockTakingDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.TxtPrimaryStockTaking.Text = "-1" Or Me.TxtPrimaryStockTaking.Text = "" Then
                MsgBox("Please specify stock taking number.", MsgBoxStyle.Information, "Information")
                Me.TxtStockTaking.Focus()
                Exit Sub
            End If
            If Me.clbBrand.CheckedItems.Count = 0 Then
                MsgBox("Please specify brand.", MsgBoxStyle.Information, "Information")
                Me.clbBrand.Focus()
                Exit Sub
            End If

            If Me.cboFormat.SelectedIndex = 0 Then
                Me.CetakDetail()
            Else
                Me.CetakRekap()
            End If

            Me.Crv1.DisplayGroupTree = False
            Me.Crv1.DisplayStatusBar = True
            Me.Crv1.DisplayToolbar = True
            Me.Crv1.ShowGroupTreeButton = False
            Me.Crv1.ShowExportButton = True
            Me.Crv1.ShowRefreshButton = False

            If Me.TabControl1.SelectedIndex = 1 Then
                Me.LoadDataAnalisis()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtStockTaking_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockTaking.GotFocus
        Try
            If Me.TxtStockTaking.Enabled = True Then
                If Me.TxtStockTaking.ReadOnly = False Then
                    TxtStockTaking.BackColor = Me.BackColorAll
                Else
                    TxtStockTaking.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtStockTaking_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtStockTaking.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtStockTaking.Text <> "" Then
                    If Me.ObjST.ApakahSudahAda(Me.TxtStockTaking.Text) = True Then
                        Me.TxtPrimaryStockTaking.Text = Me.ObjST.GetPrimary(Me.TxtStockTaking.Text)
                        Me.dtpStockTakingDate.Value = Me.ObjST.GetStartDate(Me.TxtPrimaryStockTaking.Text)
                        Me.btnFilter.Focus()
                    Else
                        'ada 2 unit dengan kode yang sama atau tidak ada unit yang dimaksud
                        Me.btnFindStockTaking.PerformClick()
                    End If
                Else
                    Me.btnFindStockTaking.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtStockTaking_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockTaking.Leave
        Dim i As New FrmCariTransaksiStockTaking
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtStockTaking.Text <> "" Then
                If Me.ObjST.ApakahSudahAda(Me.TxtStockTaking.Text) = True Then
                    Me.TxtPrimaryStockTaking.Text = Me.ObjST.GetPrimary(Me.TxtStockTaking.Text)
                    Me.TxtStockTaking.Text = Me.ObjST.GetNoTransaksi(Me.TxtPrimaryStockTaking.Text)
                    Me.dtpStockTakingDate.Value = Me.ObjST.GetStartDate(Me.TxtPrimaryStockTaking.Text)
                Else
                    If Me.TxtStockTaking.Text <> "" Then
                        i.txtCari.Text = Me.TxtStockTaking.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksiTaking > 0 Then
                        Me.TxtPrimaryStockTaking.Text = i.PrimaryTransaksiTaking
                        Me.TxtStockTaking.Text = Me.ObjST.GetNoTransaksi(Me.TxtPrimaryStockTaking.Text)
                        Me.dtpStockTakingDate.Value = Me.ObjST.GetStartDate(Me.TxtPrimaryStockTaking.Text)
                    Else
                        Me.TxtStockTaking.Clear()
                        Me.TxtPrimaryStockTaking.Text = "-1"
                        Me.dtpStockTakingDate.Value = Today
                        Me.TxtStockTaking.Focus()
                    End If
                End If
            Else
                Me.TxtStockTaking.Clear()
                Me.TxtPrimaryStockTaking.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtStockTaking_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockTaking.LostFocus
        Try
            If Me.TxtStockTaking.Enabled = True Then
                If Me.TxtStockTaking.ReadOnly = False Then
                    TxtStockTaking.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtStockTaking.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindStockTaking_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindStockTaking.Click
        Dim i As New FrmCariTransaksiStockTaking
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtStockTaking.Text <> "" Then
                i.txtCari.Text = Me.TxtStockTaking.Text
            End If

            i.ShowDialog()
            If i.PrimaryTransaksiTaking > 0 Then
                Me.TxtPrimaryStockTaking.Text = i.PrimaryTransaksiTaking
                Me.TxtStockTaking.Text = Me.ObjST.GetNoTransaksi(Me.TxtPrimaryStockTaking.Text)
                Me.dtpStockTakingDate.Value = Me.ObjST.GetStartDate(Me.TxtPrimaryStockTaking.Text)
                Me.btnFilter.Focus()
            Else
                Me.TxtStockTaking.Clear()
                Me.TxtPrimaryStockTaking.Text = "-1"
                Me.dtpStockTakingDate.Value = Today
                Me.TxtStockTaking.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub LoadDataAnalisis()
        Dim Pindah As Boolean = True
        Dim rowDGV As DataGridViewRow

        Try
            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Me.dgvItem.Rows.Clear()
            Me.SetLblTotalAmount()

            Me.DTA_Taking.Connection.ConnectionString = Core_Module.GetConstring
            If Me.TxtPrimaryStockTaking.Text = "" Then
                Me.TxtPrimaryStockTaking.Text = "-1"
            End If
            If Me.TxtPrimaryStockTaking.Text <> "-1" Then
                Me.DTA_Taking.Fill(Me.TBL_Taking, Me.TxtPrimaryStockTaking.Text)

                For Each Me.row_Taking In Me.TBL_Taking.Rows
                    If Me.chk_Diff_Only.Checked = True And Me.row_Taking.difference = 0 Then
                        Pindah = False
                    Else
                        Pindah = True
                    End If

                    If Pindah = True Then
                        Me.dgvItem.Rows.Add()
                        rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        With rowDGV
                            .Cells(Me.Col_no.Index).Value = Me.dgvItem.Rows.Count
                            .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.row_Taking.sub_total, Me.DefPrec)
                            .Cells(Me.ColAnalisis.Index).Value = Me.row_Taking.analisis
                            .Cells(Me.ColCounterMeasure.Index).Value = Me.row_Taking.counter_measures
                            .Cells(Me.ColDesc.Index).Value = Me.row_Taking.Description
                            .Cells(Me.ColDifference.Index).Value = Me.ObjCD.DigitGrouping(Me.row_Taking.difference, Me.row_Taking._Decimal)
                            .Cells(Me.ColEdit.Index).Value = "Edit"
                            .Cells(Me.ColPrimary.Index).Value = Me.row_Taking.primary
                            .Cells(Me.ColPrimaryProduct.Index).Value = Me.row_Taking.primary_product
                            .Cells(Me.ColProductCode.Index).Value = Me.row_Taking.Product_Code
                            .Cells(Me.ColQtyOnHand.Index).Value = Me.ObjCD.DigitGrouping(Me.row_Taking.qty_on_hand, Me.row_Taking._Decimal)
                            .Cells(Me.ColQtyTaking.Index).Value = Me.ObjCD.DigitGrouping(Me.row_Taking.qty_taking, Me.row_Taking._Decimal)
                        End With
                    End If
                Next
            End If

            'hitung total amount
            Me.SetLblTotalAmount()

            'cek visibilitas
            If Me.chkAnalisis1.Checked = True Then
                Me.ColAnalisis.Visible = True
            Else
                Me.ColAnalisis.Visible = False
            End If
            If Me.chkCounterMeasure1.Checked = True Then
                Me.ColCounterMeasure.Visible = True
            Else
                Me.ColCounterMeasure.Visible = False
            End If
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetLblTotalAmount()
        Dim Diff As Nullable(Of Decimal)

        Try
            If Me.TxtPrimaryStockTaking.Text = "" Or Me.TxtPrimaryStockTaking.Text = "-1" Then
                Me.LblTotalAmount.Text = "Total Amount : " + Me.ObjCur.GetCurrencyCode(1) + " " + Me.ObjCD.DigitGrouping(0, DefPrec)
            Else
                Diff = Me.DTA_Taking.ScalarQuery_GetDifferenceAmount(Me.TxtPrimaryStockTaking.Text)
                If Diff.HasValue = True Then
                    Me.LblTotalAmount.Text = "Total Amount : " + Me.ObjCur.GetCurrencyCode(1) + " " + Me.ObjCD.DigitGrouping(Diff, DefPrec)
                Else
                    Me.LblTotalAmount.Text = "Total Amount : " + Me.ObjCur.GetCurrencyCode(1) + " " + Me.ObjCD.DigitGrouping(0, DefPrec)
                End If
            End If
            Me.LblTotalAmount.Left = (Me.dgvItem.Left + Me.dgvItem.Width - Me.LblTotalAmount.Width)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            If Me.cboFormat.SelectedIndex = 0 Then
                Me.chkAnalysis.Enabled = True
                Me.chkCounterMeasure.Enabled = True
            Else
                Me.chkAnalysis.Checked = False
                Me.chkAnalysis.Enabled = False
                Me.chkCounterMeasure.Checked = False
                Me.chkCounterMeasure.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAnalisis1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAnalisis1.CheckedChanged
        Try
            If Me.chkAnalisis1.Checked = True Then
                Me.ColAnalisis.Visible = True
            Else
                Me.ColAnalisis.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkCounterMeasure1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCounterMeasure1.CheckedChanged
        Try
            If Me.chkCounterMeasure1.Checked = True Then
                Me.ColCounterMeasure.Visible = True
            Else
                Me.ColCounterMeasure.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_Diff_Only_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_Diff_Only.CheckedChanged
        Me.LoadDataAnalisis()
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If Me.TabControl1.SelectedIndex = 1 Then
            Me.LoadDataAnalisis()
        End If
    End Sub

    Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
        Dim rowDGV As DataGridViewRow
        Dim frm As New FrmLaporanStockTaking_Cost_AnalisisDetails

        Try
            If e.ColumnIndex = Me.ColEdit.Index Then
                rowDGV = Me.dgvItem.Rows(e.RowIndex)

                frm.TxtAnalysis.Text = rowDGV.Cells(Me.ColAnalisis.Index).Value
                frm.TxtCounterMeasures.Text = rowDGV.Cells(Me.ColCounterMeasure.Index).Value
                If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    rowDGV.Cells(Me.ColAnalisis.Index).Value = frm.TxtAnalysis.Text
                    rowDGV.Cells(Me.ColCounterMeasure.Index).Value = frm.TxtCounterMeasures.Text

                    Me.DTA_Taking.UpdateQuery(frm.TxtAnalysis.Text, frm.TxtCounterMeasures.Text, rowDGV.Cells(Me.ColPrimary.Index).Value)

                    'MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm = Nothing
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ij As New FrmTransaksiStockTakingProcess

        If Me.TxtPrimaryStockTaking.Text <> "-1" And Me.TxtPrimaryStockTaking.Text <> "" Then
            ij.Simpan_ItemCost_UsingScript(Me.TxtPrimaryStockTaking.Text)
        End If
    End Sub
End Class