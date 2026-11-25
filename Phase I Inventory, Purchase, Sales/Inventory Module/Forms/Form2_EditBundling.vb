Option Explicit On
Option Strict Off

Imports Core_Module

Public Class Form2_EditBundling
    Dim Buffer As String = ""
    Dim rowIndex As Integer = -1
    Dim ObjAkun As New Core_Module.ClsMasterAkun1
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjClosing As ClsTransaksiMonthlyClosingItem
    Dim ObjCD As New ClsCekData

    Private Sub Form2_EditBundling_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim rowDGV As DataGridViewRow
        Dim DTA As DataSet2_BundlingTableAdapters.DataTable2TableAdapter

        Try
            Me.dgvItem.AllowUserToAddRows = False
            DTA.DeleteQuery()
            For Each rowDGV In Me.dgvItem.Rows
                DTA.InsertQuery(rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Form2_EditBundling_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.FillDGVAkun()
            Me.LblProgress.Text = "Ready"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FillDGVAkun()
        Dim DTA As New DataSet2_BundlingTableAdapters.DataTable2TableAdapter
        Dim TBL As New DataSet2_Bundling.DataTable2DataTable
        Dim row As DataSet2_Bundling.DataTable2Row
        Dim rowDGV As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.Rows.Clear()

            DTA.Fill(TBL)
            For Each row In TBL.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count)
                rowDGV.Cells(Me.ColAcc_AccDesc.Index).Value = row.Description
                rowDGV.Cells(Me.ColAcc_AccountCode.Index).Value = row.Account_Code
                rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value = row.primary_akun
            Next
            Me.dgvItem.AllowUserToAddRows = True
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.RowIndex = Me.dgvItem.RowCount - 1 Then
                If e.ColumnIndex <> Me.ColAcc_AccountCode.Index Then
                    e.Cancel = True
                    Exit Sub
                Else
                    'col account
                    If e.ColumnIndex = Me.ColAcc_AccountCode.Index Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            Else
                If e.ColumnIndex <> Me.ColAcc_AccountCode.Index Then
                    e.Cancel = True
                    Exit Sub
                Else
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColAcc_AccountCode.Index Then
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                Or Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    If e.RowIndex <> Me.dgvItem.RowCount - 1 Then
                        Me.rowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                        '  Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                    End If
                    Exit Sub
                Else
                    ''data account number
                    If e.ColumnIndex = Me.ColAcc_AccountCode.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Me.dgvItem.Item(Me.ColAcc_PrimaryAkun.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvItem.Item(Me.ColAcc_AccDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.dgvItem.Item(Me.ColAcc_PrimaryAkun.Index, e.RowIndex).Value)
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New Core_Module.FrmMasterAkun_Cari1
                            ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            If ij.mPrimaryAkun <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                                Me.dgvItem.Item(Me.ColAcc_AccDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                                Me.dgvItem.Item(Me.ColAcc_PrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                            ElseIf ij.mPrimaryAkun = -1 Then
                                '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                    Me.rowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                    ' Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                    End If
                End If 
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0

        Try
            'set nilai default baris baru
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value = -1
            rowDGV.Cells(Me.ColAcc_AccountCode.Index).Value = ""
            rowDGV.Cells(Me.ColAcc_AccDesc.Index).Value = "" 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvItem.Rows.RemoveAt(Me.rowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNoBefore_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNoBefore.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNoBefore.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNoBefore.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNoBefore.Text)
                    Me.TxtProductDescBefore.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoBefore.Text)
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNoBefore.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNoBefore.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNoBefore.Text = i.PrimaryProduct
                        Me.TxtCodeNoBefore.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNoBefore.Text)
                        Me.TxtProductDescBefore.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoBefore.Text)
                    Else
                        Me.TxtCodeNoBefore.Clear()
                        Me.TxtProductDescBefore.Clear()
                        Me.TxtPrimaryCodeNoBefore.Text = "-1"
                        Me.TxtCodeNoBefore.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNoBefore_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNoBefore.Leave
        Dim i As New frmCariMasterProduct

        Try
            If Me.TxtCodeNoBefore.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtCodeNoBefore.Text) = 1 Then
                    Me.TxtPrimaryCodeNoBefore.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNoBefore.Text)
                    Me.TxtProductDescBefore.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoBefore.Text)
                Else
                    If Me.TxtCodeNoBefore.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNoBefore.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNoBefore.Text = i.PrimaryProduct
                        Me.TxtCodeNoBefore.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNoBefore.Text)
                        Me.TxtProductDescBefore.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoBefore.Text)
                    Else
                        Me.TxtCodeNoBefore.Clear()
                        Me.TxtPrimaryCodeNoBefore.Text = "-1"
                        Me.TxtCodeNoBefore.Focus()
                        Me.TxtProductDescBefore.Text = ""
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNoBefore.Text = "-1"
                Me.TxtProductDescBefore.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCodeNoBefore.Click
        Dim i As New frmCariMasterProduct

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNoBefore.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNoBefore.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNoBefore.Text = i.PrimaryProduct
                Me.TxtCodeNoBefore.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNoBefore.Text)
                Me.TxtProductDescBefore.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoBefore.Text)
            Else
                Me.TxtCodeNoBefore.Clear()
                Me.TxtProductDescBefore.Clear()
                Me.TxtPrimaryCodeNoBefore.Text = "-1"
                Me.TxtCodeNoBefore.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNoAfter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNoAfter.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNoAfter.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNoAfter.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNoAfter.Text)
                    Me.TxtProductDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNoAfter.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNoAfter.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNoAfter.Text = i.PrimaryProduct
                        Me.TxtCodeNoAfter.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.TxtProductDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
                    Else
                        Me.TxtCodeNoAfter.Clear()
                        Me.TxtProductDescAfter.Clear()
                        Me.TxtPrimaryCodeNoAfter.Text = "-1"
                        Me.TxtCodeNoAfter.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNoAfter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNoAfter.Leave
        Dim i As New frmCariMasterProduct

        Try
            If Me.TxtCodeNoAfter.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtCodeNoAfter.Text) = 1 Then
                    Me.TxtPrimaryCodeNoAfter.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNoAfter.Text) 
                    Me.TxtProductDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text) 
                Else
                    If Me.TxtCodeNoAfter.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNoAfter.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNoAfter.Text = i.PrimaryProduct
                        Me.TxtCodeNoAfter.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.TxtProductDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text) 
                    Else
                        Me.TxtCodeNoAfter.Clear()
                        Me.TxtPrimaryCodeNoAfter.Text = "-1" 
                        Me.TxtCodeNoAfter.Focus()
                        Me.TxtProductDescAfter.Text = ""
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNoAfter.Text = "-1"
                Me.TxtProductDescAfter.Clear() 
            End If 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCodeNoAfter.Click
        Dim i As New frmCariMasterProduct

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNoAfter.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNoAfter.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNoAfter.Text = i.PrimaryProduct
                Me.TxtCodeNoAfter.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNoAfter.Text)
                Me.TxtProductDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
            Else
                Me.TxtCodeNoAfter.Clear()
                Me.TxtProductDescAfter.Clear()
                Me.TxtPrimaryCodeNoAfter.Text = "-1"
                Me.TxtCodeNoAfter.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Try
            MsgBox("Pastikan code no before dan after tidak kontrol lokasi, batch dan expiry", MsgBoxStyle.Information, "Information")

            If Me.ObjProduct.ApakahGoods(Me.TxtPrimaryCodeNoAfter.Text) = False Then
                MsgBox("Please specify goods on after data", MsgBoxStyle.Information, "Information")
                Me.TxtProductDescAfter.Focus()
                Exit Sub
            End If
            If Me.ObjProduct.ApakahGoods(Me.TxtPrimaryCodeNoBefore.Text) = False Then
                MsgBox("Please specify goods on before data", MsgBoxStyle.Information, "Information")
                Me.TxtProductDescBefore.Focus()
                Exit Sub
            End If

            'pastikan tgl before < tgl after
            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Invalid period. Please fix", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'pastikan tidak ada closing pada period yang akan di edit
            If Me.ObjCD.GetLasDateOfTheMonth(Me.ObjClosing.GetLastClosingPeriod).Date >= Me.dtpFrom.Value.Date Then
                MsgBox("Monthly closing detected on: " + Format(Me.ObjClosing.GetLastClosingPeriod, "MMM yyyy"), MsgBoxStyle.Information, "Information")
            End If

            If MsgBox("Start edit bundling data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.dgvItem.AllowUserToAddRows = False
                Me.bgw.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw.DoWork
        Dim DTA As New DataSet2_BundlingTableAdapters.tbltransaksi_stockbundling1TableAdapter
        Dim TBL As New DataSet2_Bundling.tbltransaksi_stockbundling1DataTable
        Dim row As DataSet2_Bundling.tbltransaksi_stockbundling1Row
        Dim DTAItemBefore As New DataSet2_BundlingTableAdapters.tbltransaksi_stockbundling_item_before1TableAdapter
        Dim TBLItemBefore As New DataSet2_Bundling.tbltransaksi_stockbundling_item_before1DataTable
        Dim DTAItemAfter As New DataSet2_BundlingTableAdapters.tbltransaksi_stockbundling_item_after1TableAdapter
        Dim TBLItemAfter As New DataSet2_Bundling.tbltransaksi_stockbundling_item_after1DataTable
        Dim DTAItem As New DataSet2_BundlingTableAdapters.tblmaster_productqtyTableAdapter
        Dim TBLItemDummy As New DataSet2_Bundling.tblmaster_productqtyDataTable
        Dim TBLItem As New DataSet2_Bundling.tblmaster_productqtyDataTable
        Dim ListPrimary As New List(Of Double)

        Try
            ObjCD.UpdateThreading()

            'kumpulkan semua bundling pada periode yang disebutkan
            DTA.Fill(TBL, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)

            'untuk setiap baris bundling:
            For Each row In TBL.Rows
                'fill data item
                TBLItemBefore.Rows.Clear()
                DTAItemBefore.Fill(TBLItemBefore, row.Primary)
                TBLItemAfter.Rows.Clear()
                DTAItemAfter.Fill(TBLItemAfter, row.Primary)

                '0. hapuskan sc
                DTAItem.DeleteQuery(row.Primary, ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling)

                '1. hapus item before
                Me.HapusItemBefore(TBLItemBefore, DTAItem, ListPrimary, TBLItemDummy, TBLItem)

                '2. hapus item after
                Me.HapusItemAfter(TBLItemAfter, DTAItem, ListPrimary, TBLItemDummy, TBLItem)

                '3. modif item before + simpan ke db
                Me.ModifItemBefore(TBLItemBefore, DTAItem, ListPrimary, TBLItemDummy, TBLItem, row.No_Transaksi, row.Primary, row.Tanggal_Transaksi)

                '4. modif item after + simpan ke db
                Me.ModifItemAfter(TBLItemAfter, DTAItem, ListPrimary, TBLItemDummy, TBLItem, row.No_Transaksi, row.Primary, row.Tanggal_Transaksi)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ModifItemAfter(ByVal TBLItemAfter As DataSet2_Bundling.tbltransaksi_stockbundling_item_after1DataTable, _
    ByVal DTAItem As DataSet2_BundlingTableAdapters.tblmaster_productqtyTableAdapter, ByVal ListPrimary As List(Of Double), _
    ByVal TBLItemDummy As DataSet2_Bundling.tblmaster_productqtyDataTable, ByVal TBLItem As DataSet2_Bundling.tblmaster_productqtyDataTable, _
    ByVal vDocNo As String, ByVal vPrimaryTransaksiAsal As Double, ByVal vTransactionDate As Date)
        Dim rowItemAfter, rowItemAfter1 As DataSet2_Bundling.tbltransaksi_stockbundling_item_after1Row
        Dim RowItem1, RowItem2 As DataSet2_Bundling.tblmaster_productqtyRow
        Dim SudahStart As Boolean = False
        Dim RefIndex As Integer = -1
        Dim CurIndex As Integer = 0
        Dim Modif As Boolean = False
        Dim rowDGV As DataGridViewRow

        Try
            ListPrimary.Clear()
            TBLItem.Rows.Clear()
            TBLItemDummy.Rows.Clear()

            'modif item After
            While CurIndex <= TBLItemAfter.Rows.Count - 1
                Modif = True
                rowItemAfter = TBLItemAfter.Rows(CurIndex)

                'tentukan apakah modif
                For Each rowDGV In Me.dgvItem.Rows
                    If rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value = rowItemAfter.Primary_Inventory_Account Then
                        Modif = False
                        Exit For
                    End If
                Next

                'klo modif, start proses modif
                If Modif = True Then
                    If RefIndex = -1 Then
                        'baris ybs yang di ubah dan selanjutnya menjadi referensi
                        'code no after tidak kontrol lokasi, expiry atau batch
                        RefIndex = CurIndex
                        rowItemAfter.Primary_Product = Me.TxtPrimaryCodeNoAfter.Text
                        rowItemAfter.Expiry_Date = #1/1/1950#
                        rowItemAfter.Batch_No = ""
                        rowItemAfter.Primary_Storage_Area = -1
                        CurIndex += 1
                    Else
                        'tambahkan qty dan cost factor
                        rowItemAfter1 = TBLItemAfter.Rows(RefIndex)
                        rowItemAfter1.Qty += rowItemAfter.Qty
                        rowItemAfter1.Faktor_Pemberat += rowItemAfter.Faktor_Pemberat
                        TBLItemAfter.Rows.RemoveAt(CurIndex)
                    End If
                Else
                    CurIndex += 1
                End If
            End While

            'kumpulkan list primary product
            For Each rowItemAfter In TBLItemAfter.Rows
                If ListPrimary.Contains(rowItemAfter.Primary_Product) = False Then
                    ListPrimary.Add(rowItemAfter.Primary_Product)
                End If
            Next

            'untuk setiap item primary product:
            For Each PrimaryProduct As Double In ListPrimary
                TBLItemDummy.Rows.Clear()

                '1. ambil product qty dari db
                DTAItem.Fill(TBLItemDummy, PrimaryProduct)

                '2. untuk setiap item After, tambahkan ke table
                For Each rowItemAfter In TBLItemAfter.Rows
                    If rowItemAfter.Primary_Product = PrimaryProduct Then
                        TBLItemDummy.DefaultView.RowFilter = "[Lot No] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(rowItemAfter.Batch_No) + _
                        "' AND [Primary Storage Area] = " + rowItemAfter.Primary_Storage_Area.ToString

                        If TBLItemDummy.DefaultView.Count = 0 Then
                            'tambahkan qty
                            RowItem1 = TBLItem.DefaultView(0).Row
                            RowItem1.Qty += rowItemAfter.Qty
                        Else
                            'buat baris baru
                            RowItem1 = TBLItemDummy.Newtblmaster_productqtyRow
                            With rowItemAfter
                                RowItem1.Expiry_Date = .Expiry_Date
                                RowItem1.Lot_No = .Batch_No
                                RowItem1.Primary = -1
                                RowItem1.Primary_Product = .Primary_Product
                                RowItem1.Primary_Storage_Area = .Primary_Storage_Area
                                RowItem1.Qty = .Qty
                                RowItem1.Qty_Secondary = -1
                            End With
                            TBLItemDummy.Rows.Add(RowItem1)
                        End If
                    End If
                Next

                '3. pindahkan ke table dummy
                For Each RowItem1 In TBLItemDummy.Rows
                    If RowItem1.Qty <> 0 Then
                        RowItem2 = TBLItem.Newtblmaster_productqtyRow
                        For jk As Integer = 0 To TBLItem.Columns.Count - 1
                            RowItem2.Item(jk) = RowItem1.Item(jk)
                        Next
                        TBLItem.Rows.Add(RowItem2)
                    End If
                Next
            Next

            'hapus + simpan item product qty
            Me.Simpan_ItemProdQty_UsingScript(TBLItem, ListPrimary)

            'simpan sc
            Me.Simpan_SC_UsingScript(TBLItem, vDocNo, ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling, vPrimaryTransaksiAsal, True, vTransactionDate)

            ListPrimary.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub ModifItemBefore(ByVal TBLItemBefore As DataSet2_Bundling.tbltransaksi_stockbundling_item_before1DataTable, _
    ByVal DTAItem As DataSet2_BundlingTableAdapters.tblmaster_productqtyTableAdapter, ByVal ListPrimary As List(Of Double), _
    ByVal TBLItemDummy As DataSet2_Bundling.tblmaster_productqtyDataTable, ByVal TBLItem As DataSet2_Bundling.tblmaster_productqtyDataTable, _
    ByVal vDocNo As String, ByVal vPrimaryTransaksiAsal As Double, ByVal vTransactionDate As Date)
        Dim rowItemBefore, rowItemBefore1 As DataSet2_Bundling.tbltransaksi_stockbundling_item_before1Row
        Dim RowItem1, RowItem2 As DataSet2_Bundling.tblmaster_productqtyRow
        Dim SudahStart As Boolean = False
        Dim RefIndex As Integer = -1
        Dim CurIndex As Integer = 0
        Dim Modif As Boolean = False
        Dim rowDGV As DataGridViewRow

        Try
            ListPrimary.Clear()
            TBLItem.Rows.Clear()
            TBLItemDummy.Rows.Clear()

            'modif item before
            While CurIndex <= TBLItemBefore.Rows.Count - 1
                Modif = True
                rowItemBefore = TBLItemBefore.Rows(CurIndex)

                'tentukan apakah modif
                For Each rowDGV In Me.dgvItem.Rows
                    If rowDGV.Cells(Me.ColAcc_PrimaryAkun.Index).Value = rowItemBefore.Primary_Inventory_Account Then
                        Modif = False
                        Exit For
                    End If
                Next

                'klo modif, start proses modif
                If Modif = True Then
                    If RefIndex = -1 Then
                        RefIndex = CurIndex
                        rowItemBefore.Primary_Product = Me.TxtPrimaryCodeNoBefore.Text
                        rowItemBefore.Expiry_Date = #1/1/1950#
                        rowItemBefore.Batch_No = ""
                        rowItemBefore.Primary_Storage_Area = -1
                        CurIndex += 1
                    Else
                        rowItemBefore1 = TBLItemBefore.Rows(RefIndex)
                        rowItemBefore1.Qty += rowItemBefore.Qty
                        TBLItemBefore.Rows.RemoveAt(CurIndex)
                    End If
                Else
                    CurIndex += 1
                End If
            End While

            'kumpulkan list primary product
            For Each rowItemBefore In TBLItemBefore.Rows
                If ListPrimary.Contains(rowItemBefore.Primary_Product) = False Then
                    ListPrimary.Add(rowItemBefore.Primary_Product)
                End If
            Next

            'untuk setiap item primary product:
            For Each PrimaryProduct As Double In ListPrimary
                TBLItemDummy.Rows.Clear()

                '1. ambil product qty dari db
                DTAItem.Fill(TBLItemDummy, PrimaryProduct)

                '2. untuk setiap item before, tambahkan ke table
                For Each rowItemBefore In TBLItemBefore.Rows
                    If rowItemBefore.Primary_Product = PrimaryProduct Then
                        TBLItemDummy.DefaultView.RowFilter = "[Lot No] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(rowItemBefore.Batch_No) + _
                        "' AND [Primary Storage Area] = " + rowItemBefore.Primary_Storage_Area.ToString
                        If TBLItemDummy.DefaultView.Count = 0 Then
                            'tambahkan qty
                            RowItem1 = TBLItem.DefaultView(0).Row
                            RowItem1.Qty -= rowItemBefore.Qty
                        Else
                            'buat baris baru
                            RowItem1 = TBLItemDummy.Newtblmaster_productqtyRow
                            With rowItemBefore
                                RowItem1.Expiry_Date = .Expiry_Date
                                RowItem1.Lot_No = .Batch_No
                                RowItem1.Primary = -1
                                RowItem1.Primary_Product = .Primary_Product
                                RowItem1.Primary_Storage_Area = .Primary_Storage_Area
                                RowItem1.Qty = .Qty * -1
                                RowItem1.Qty_Secondary = -1
                            End With
                            TBLItemDummy.Rows.Add(RowItem1)
                        End If
                    End If
                Next

                '3. pindahkan ke table dummy
                For Each RowItem1 In TBLItemDummy.Rows
                    If RowItem1.Qty <> 0 Then
                        RowItem2 = TBLItem.Newtblmaster_productqtyRow
                        For jk As Integer = 0 To TBLItem.Columns.Count - 1
                            RowItem2.Item(jk) = RowItem1.Item(jk)
                        Next
                        TBLItem.Rows.Add(RowItem2)
                    End If
                Next
            Next

            'hapus + simpan item product qty
            Me.Simpan_ItemProdQty_UsingScript(TBLItem, ListPrimary)

            'simpan sc
            Me.Simpan_SC_UsingScript(TBLItem, vDocNo, ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling, vPrimaryTransaksiAsal, False, vTransactionDate)

            ListPrimary.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub HapusItemAfter(ByVal TBLItemAfter As DataSet2_Bundling.tbltransaksi_stockbundling_item_after1DataTable, _
   ByVal DTAItem As DataSet2_BundlingTableAdapters.tblmaster_productqtyTableAdapter, ByVal ListPrimary As List(Of Double), _
   ByVal TBLItemDummy As DataSet2_Bundling.tblmaster_productqtyDataTable, ByVal TBLItem As DataSet2_Bundling.tblmaster_productqtyDataTable)
        Dim rowItemAfter As DataSet2_Bundling.tbltransaksi_stockbundling_item_after1Row
        Dim RowItem1, RowItem2 As DataSet2_Bundling.tblmaster_productqtyRow

        Try
            ListPrimary.Clear()
            TBLItem.Rows.Clear()
            TBLItemDummy.Rows.Clear()

            'kumpulkan list primary product
            For Each rowItemAfter In TBLItemAfter.Rows
                If ListPrimary.Contains(rowItemAfter.Primary_Product) = False Then
                    ListPrimary.Add(rowItemAfter.Primary_Product)
                End If
            Next

            'untuk setiap item primary product:
            For Each PrimaryProduct As Double In ListPrimary
                TBLItemDummy.Rows.Clear()

                '1. ambil product qty dari db
                DTAItem.Fill(TBLItemDummy, PrimaryProduct)

                '2. untuk setiap item After, kurangkan ke table
                For Each rowItemAfter In TBLItemAfter.Rows
                    If rowItemAfter.Primary_Product = PrimaryProduct Then
                        TBLItemDummy.DefaultView.RowFilter = "[Lot No] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(rowItemAfter.Batch_No) + _
                        "' AND [Primary Storage Area] = " + rowItemAfter.Primary_Storage_Area.ToString
                        If TBLItemDummy.DefaultView.Count = 0 Then
                            'tambahkan qty
                            RowItem1 = TBLItem.DefaultView(0).Row
                            RowItem1.Qty -= rowItemAfter.Qty
                        Else
                            'buat baris baru
                            RowItem1 = TBLItemDummy.Newtblmaster_productqtyRow
                            With rowItemAfter
                                RowItem1.Expiry_Date = .Expiry_Date
                                RowItem1.Lot_No = .Batch_No
                                RowItem1.Primary = -1
                                RowItem1.Primary_Product = .Primary_Product
                                RowItem1.Primary_Storage_Area = .Primary_Storage_Area
                                RowItem1.Qty = .Qty * -1
                                RowItem1.Qty_Secondary = -1
                            End With
                            TBLItemDummy.Rows.Add(RowItem1)
                        End If
                    End If
                Next

                '3. pindahkan ke table dummy
                For Each RowItem1 In TBLItemDummy.Rows
                    If RowItem1.Qty <> 0 Then
                        RowItem2 = TBLItem.Newtblmaster_productqtyRow
                        For jk As Integer = 0 To TBLItem.Columns.Count - 1
                            RowItem2.Item(jk) = RowItem1.Item(jk)
                        Next
                        TBLItem.Rows.Add(RowItem2)
                    End If
                Next
            Next

            'hapus + simpan item product qty
            Me.Simpan_ItemProdQty_UsingScript(TBLItem, ListPrimary)

            ListPrimary.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub HapusItemBefore(ByVal TBLItemBefore As DataSet2_Bundling.tbltransaksi_stockbundling_item_before1DataTable, _
    ByVal DTAItem As DataSet2_BundlingTableAdapters.tblmaster_productqtyTableAdapter, ByVal ListPrimary As List(Of Double), _
    ByVal TBLItemDummy As DataSet2_Bundling.tblmaster_productqtyDataTable, ByVal TBLItem As DataSet2_Bundling.tblmaster_productqtyDataTable )
        Dim rowItemBefore As DataSet2_Bundling.tbltransaksi_stockbundling_item_before1Row
        Dim RowItem1, RowItem2 As DataSet2_Bundling.tblmaster_productqtyRow

        Try
            ListPrimary.Clear()
            TBLItem.Rows.Clear()
            TBLItemDummy.Rows.Clear()

            'kumpulkan list primary product
            For Each rowItemBefore In TBLItemBefore.Rows
                If ListPrimary.Contains(rowItemBefore.Primary_Product) = False Then
                    ListPrimary.Add(rowItemBefore.Primary_Product)
                End If
            Next

            'untuk setiap item primary product:
            For Each PrimaryProduct As Double In ListPrimary
                TBLItemDummy.Rows.Clear()

                '1. ambil product qty dari db
                DTAItem.Fill(TBLItemDummy, PrimaryProduct)

                '2. untuk setiap item before, tambahkan ke table
                For Each rowItemBefore In TBLItemBefore.Rows
                    If rowItemBefore.Primary_Product = PrimaryProduct Then
                        TBLItemDummy.DefaultView.RowFilter = "[Lot No] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(rowItemBefore.Batch_No) + _
                        "' AND [Primary Storage Area] = " + rowItemBefore.Primary_Storage_Area.ToString
                        If TBLItemDummy.DefaultView.Count = 0 Then
                            'tambahkan qty
                            RowItem1 = TBLItem.DefaultView(0).Row
                            RowItem1.Qty += rowItemBefore.Qty
                        Else
                            'buat baris baru
                            RowItem1 = TBLItemDummy.Newtblmaster_productqtyRow
                            With rowItemBefore
                                RowItem1.Expiry_Date = .Expiry_Date
                                RowItem1.Lot_No = .Batch_No
                                RowItem1.Primary = -1
                                RowItem1.Primary_Product = .Primary_Product
                                RowItem1.Primary_Storage_Area = .Primary_Storage_Area
                                RowItem1.Qty = .Qty
                                RowItem1.Qty_Secondary = -1
                            End With
                            TBLItemDummy.Rows.Add(RowItem1)
                        End If
                    End If
                Next

                '3. pindahkan ke table dummy
                For Each RowItem1 In TBLItemDummy.Rows
                    If RowItem1.Qty <> 0 Then
                        RowItem2 = TBLItem.Newtblmaster_productqtyRow
                        For jk As Integer = 0 To TBLItem.Columns.Count - 1
                            RowItem2.Item(jk) = RowItem1.Item(jk)
                        Next
                        TBLItem.Rows.Add(RowItem2)
                    End If
                Next
            Next

            'hapus + simpan item product qty
            Me.Simpan_ItemProdQty_UsingScript(TBLItem, ListPrimary)

            ListPrimary.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Function Simpan_SC_UsingScript(ByVal vTBL As DataSet2_Bundling.tblmaster_productqtyDataTable, ByVal vDocNo As String, _
    ByVal vTipeTransaksiAsal As ClsTransaksiStockCard.EnumJenisTransaksiAsal, ByVal vPrimaryTransaksiAsal As Double, _
    ByVal vIn As Boolean, ByVal vTransactionDate As Date) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            'simpan item
            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_SC(StartRowIndex, BatchQty, vTBL, vDocNo, vTipeTransaksiAsal, vPrimaryTransaksiAsal, vIn , vTransactionDate )
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_SC(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DataSet2_Bundling.tblmaster_productqtyDataTable, ByVal vDocNo As String, _
    ByVal vTipeTransaksiAsal As ClsTransaksiStockCard.EnumJenisTransaksiAsal, ByVal vPrimaryTransaksiAsal As Double, _
    ByVal vIn As Boolean, ByVal vTransactionDate As Date) As String
        Dim retval As String = ""
        Dim row As DataSet2_Bundling.tblmaster_productqtyRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData
        Dim TipeSC As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In

        Try
            If vIn = True Then
                TipeSC = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
            Else
                TipeSC = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
            End If

            retval = "INSERT INTO `tbltransaksi_stock_card` (`Primary Transaksi Asal`,`Type Transaksi Asal`,`Primary Product`,`Document No`," + _
            "`In`,`Out`,`Description`,`InOut`,`Date`,`Expiry`,`primary location`,`Batch Number`,`Remarks`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If vIn = True Then
                            retval = retval + vbCrLf + "(" + vPrimaryTransaksiAsal.ToString + "," + CInt(vTipeTransaksiAsal).ToString + "," + CDbl(.Primary_Product).ToString + "," _
                            + "'" + vDocNo + "'," + CDec(.Qty).ToString + ",0,'Stock Bundling'," + CInt(TipeSC).ToString + "," + _
                            "'" + Format(vTransactionDate, "yyyy-MM-dd") + "'," + "'" + Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," + _
                            CDbl(.Primary_Storage_Area).ToString + ",'" + Me.ObjCD.FormatSpecialChar_MySQL(.Lot_No) + "',''," + ")" + Delimiter
                        Else
                            retval = retval + vbCrLf + "(" + vPrimaryTransaksiAsal.ToString + "," + CInt(vTipeTransaksiAsal).ToString + "," + CDbl(.Primary_Product).ToString + "," _
                            + "'" + vDocNo + "',0," + CDec(.Qty).ToString + ",'Stock Bundling'," + CInt(TipeSC).ToString + "," + _
                            "'" + Format(vTransactionDate, "yyyy-MM-dd") + "'," + "'" + Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," + _
                            CDbl(.Primary_Storage_Area).ToString + ",'" + Me.ObjCD.FormatSpecialChar_MySQL(.Lot_No) + "',''," + ")" + Delimiter
                        End If
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function Simpan_ItemProdQty_UsingScript(ByVal vTBL As DataSet2_Bundling.tblmaster_productqtyDataTable, ByVal vListPrimary As List(Of Double)) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open() 

            'hapus item
            sql = Me.GetDeleteString_ItemProdQty(vListPrimary)
            script.Query = sql
            'execute script
            script.Execute()

            'simpan item
            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemProdQty(StartRowIndex, BatchQty, vTBL)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close() 

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemProdQty(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DataSet2_Bundling.tblmaster_productqtyDataTable) As String
        Dim retval As String = ""
        Dim row As DataSet2_Bundling.tblmaster_productqtyRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tblmaster_productqty` (`Primary Product`,`Qty`,`Qty Secondary`,`Lot No`,`Expiry Date`,`Primary Storage Area`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(.Primary_Product).ToString + "," _
                        + CDec(.Qty).ToString + "," + CDec(.Qty_Secondary).ToString + ",'" + Me.ObjCD.FormatSpecialChar_MySQL(.Lot_No) + "'," _
                        + "'" + Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," + CDbl(.Primary_Storage_Area).ToString + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function GetDeleteString_ItemProdQty(ByVal vListPrimaryProduct As List(Of Double)) As String
        Dim retval As String = ""
        Dim row As DataSet2_Bundling.tblmaster_productqtyRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            For Each PrimaryProduct As Double In vListPrimaryProduct
                retval = retval + vbCrLf + "delete from `tblmaster_productqty` where `tblmaster_productqty`.`Primary Product` = " + PrimaryProduct.ToString + ";"
            Next
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Sub bgw_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw.RunWorkerCompleted
        Try
            Me.dgvItem.AllowUserToAddRows = True
            Me.PrgBar.Value = 0
            Me.LblProgress.Text = "Ready"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class