Option Explicit On
Option Strict Off
Imports MySql.Data.MySqlClient
Imports MySql.Data
'~ab~
Public Class FrmTransaksiDO_Quick
    '     ORIGINAL CODING
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public UserCode As String = ""

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim ObjSettingProd As New Core_Module.ClsTransaksiSetting_ProdGroup
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjDO As New ClsTransaksiDO
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjBookInv As New ClsTransaksiBookingInv
    Dim ObjLoc As New Core_Module.clsMasterStorageArea
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim objItem As New Inventory_Module.ClsTransaksiMaterialPlanningGroupItem
    Dim ObjWR As New Core_Module.clsMasterWarehouse
    Dim DummyIndex As Integer = -1

    Dim Buffer As String = ""
    Dim EraseRowIndex As Integer = -1
    Dim EraseRowIndex_Supp As Integer = -1
    Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1 '

    Dim DTA As New DseTransaksiQuick_DOTableAdapters.DataTable1TableAdapter
    Dim Tbl As New DseTransaksiQuick_DO.DataTable1DataTable
    Dim Row As DseTransaksiQuick_DO.DataTable1Row

    Dim DTABook As New DseTransaksiQuick_DOTableAdapters.DataTable3TableAdapter
    Dim TblBook As New DseTransaksiQuick_DO.DataTable3DataTable
    Dim RowBook As DseTransaksiQuick_DO.DataTable3Row

    Dim DTAStock As New DseTransaksiQuick_DO2TableAdapters.stock_detail_viewTableAdapter
    Dim TBLStock As New DseTransaksiQuick_DO2.stock_detail_viewDataTable
    Dim RowStock As DseTransaksiQuick_DO2.stock_detail_viewRow

    Dim TblStockTemp As New DseTransaksiQuick_DO.TempStockDetailDataTable
    Dim rowStockTemp As DseTransaksiQuick_DO.TempStockDetailRow

    Dim TblPendingDO As New DseTransaksiQuick_DO2.Pending_DO_ViewDataTable
    Dim RowPendingDO As DseTransaksiQuick_DO2.Pending_DO_ViewRow

    Dim TblSimpan As New DseTransaksiQuick_DO.TblSimpan1DataTable
    Dim RowSimpan As DseTransaksiQuick_DO.TblSimpan1Row

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian
    Dim SudahHapusAntrian As Boolean = True
    Dim isLocationChanged As Boolean = False

    Private Sub FrmTransaksiQuick_StockInPurchase_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ShowIcon = False

        Try
            Me.dgvItem.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCreatedBy.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvCustomer.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.cboCustomerFormat.SelectedIndex = 0
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.MultiSelect = False

            Me.rbInventoryBooking.Checked = True
            Me.lblFind.Text = "Booking No."
            Me.GBPeriod.Text = "Inventory Booking Period"
            Me.cboCustomerFormat.SelectedIndex = 1
            Me.TxtLocCode.Enabled = False
            Me.BtnFindLoc.Enabled = False

            'Tambah untuk filter hak akses
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiDO_Quick_Otorisasi") Then
                Me.txtUserCode.Text = ""
                Me.txtUserCode.Enabled = True
                Me.txtUserCode.TabStop = True
                Me.btnFindUser.Enabled = True
            Else
                Me.txtUserCode.Text = Me.UserCode
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindPO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindTransaction.Click
        Dim i As New FrmCariTransaksiSO
        i.BackColorAll = Me.BackColorAll

        Dim j As New FrmCariTransaksiBookingInv
        j.BackColorAll = Me.BackColorAll
        Try
            If Me.rbInventoryBooking.Checked Then
                'i.cboCriteria.SelectedIndex = 0
                If Me.TxtTransactionNumber.Text <> "" Then
                    j.txtCari.Text = Me.TxtTransactionNumber.Text
                End If

                j.ShowDialog()
                If j.PrimaryTransaksi > 0 Then
                    Me.TxtPrimaryTransaction.Text = j.PrimaryTransaksi
                    Me.TxtTransactionNumber.Text = Me.ObjBookInv.GetNoTransaksi(Me.TxtPrimaryTransaction.Text)
                Else
                    Me.TxtPrimaryTransaction.Text = "-1"
                    Me.TxtTransactionNumber.Clear()
                End If
            End If

            If Me.rbSO.Checked Then
                'i.cboCriteria.SelectedIndex = 0
                If Me.TxtTransactionNumber.Text <> "" Then
                    i.txtCari.Text = Me.TxtTransactionNumber.Text
                End If

                i.ShowDialog()
                If i.PrimaryTransaksi > 0 Then
                    Me.TxtPrimaryTransaction.Text = i.PrimaryTransaksi
                    Me.TxtTransactionNumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimaryTransaction.Text)
                Else
                    Me.TxtPrimaryTransaction.Text = "-1"
                    Me.TxtTransactionNumber.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            j.Dispose()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub TxtPONumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTransactionNumber.GotFocus
        Try
            Me.TxtTransactionNumber.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPONumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTransactionNumber.LostFocus
        Try
            Me.TxtTransactionNumber.BackColor = Me.TxtPrimaryTransaction.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPONumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTransactionNumber.KeyDown
        Try
            If Me.rbInventoryBooking.Checked Then
                If e.KeyValue = 13 Then
                    Dim Primary As Double = ObjBookInv.GetPrimary(Me.TxtTransactionNumber.Text)
                    If Primary <> -1 Then
                        Me.TxtPrimaryTransaction.Text = Primary
                        Me.TxtTransactionNumber.Text = Me.ObjBookInv.GetNoTransaksi(Me.TxtPrimaryTransaction.Text)
                    Else
                        Me.btnFindTransaction.PerformClick()
                    End If
                End If
            End If

            If Me.rbSO.Checked Then
                If e.KeyValue = 13 Then
                    Dim Primary As Double = ObjSO.GetPrimary(Me.TxtTransactionNumber.Text)
                    If Primary <> -1 Then
                        Me.TxtPrimaryTransaction.Text = Primary
                        Me.TxtTransactionNumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimaryTransaction.Text)
                    Else
                        Me.btnFindTransaction.PerformClick()
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPONumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTransactionNumber.Leave
        Try
            If Me.rbInventoryBooking.Checked Then
                If Me.TxtTransactionNumber.Text <> "" Then
                    Dim Primary As Double = ObjBookInv.GetPrimary(Me.TxtTransactionNumber.Text)
                    If Primary <> -1 Then
                        Me.TxtPrimaryTransaction.Text = Primary
                        Me.TxtTransactionNumber.Text = Me.ObjBookInv.GetNoTransaksi(Me.TxtPrimaryTransaction.Text)
                    Else
                        Me.btnFindTransaction.PerformClick()
                    End If
                Else
                    Me.TxtPrimaryTransaction.Text = "-1"
                    Me.TxtTransactionNumber.Clear()
                End If
            End If

            If Me.rbSO.Checked Then
                If Me.TxtTransactionNumber.Text <> "" Then
                    Dim Primary As Double = ObjSO.GetPrimary(Me.TxtTransactionNumber.Text)
                    If Primary <> -1 Then
                        Me.TxtPrimaryTransaction.Text = Primary
                        Me.TxtTransactionNumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimaryTransaction.Text)
                    Else
                        Me.btnFindTransaction.PerformClick()
                    End If
                Else
                    Me.TxtPrimaryTransaction.Text = "-1"
                    Me.TxtTransactionNumber.Clear()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCreatedBy_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvCreatedBy.DataError

    End Sub

    Private Sub dgvCreatedBy_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvCreatedBy.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            rowDGV = Me.dgvCreatedBy.Rows(Me.dgvCreatedBy.RowCount - 1)
            With rowDGV
                .Cells(Me.ColUserNo.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColUserCode.Index).Value = ""
                .Cells(Me.ColUserName.Index).Value = ""
                .Cells(Me.ColUser_Primary.Index).Value = "-1"
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCreatedBy_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvCreatedBy.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvCreatedBy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCreatedBy_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCreatedBy.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvCreatedBy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow
        Dim rowDGV1 As DataGridViewRow
        Try
            If Me.rbSO.Checked Then
                If MsgBox("Check all data with no batch controlled?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    For Each rowDGV In Me.dgvItem.Rows
                        If rowDGV.Cells(Me.ColBatchControlled.Index).Value = 0 Then
                            If Me.CekStock(rowDGV) Then
                                rowDGV.Cells(Me.ColChkDO.Index).Value = Boolean.TrueString
                            End If
                            'rowDGV.Cells(Me.ColQtyDO.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQty.Index).Value, rowDGV.Cells(Me.ColDecimalMulti.Index).Value)
                        End If
                    Next
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                End If
            Else
                If MsgBox("Check all data ?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.dgvItemReject.Rows.Clear()
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    For Each rowDGV In Me.dgvItem.Rows
                        If Me.CekStock(rowDGV) Then
                            rowDGV.Cells(Me.ColChkDO.Index).Value = Boolean.TrueString
                            rowDGV.Cells(Me.ColQtyDO.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQty.Index).Value, rowDGV.Cells(Me.ColDecimalMulti.Index).Value)
                            If rowDGV.Cells(Me.ColBatchControlled.Index).Value = 1 Then
                                Me.SalinBarisBatch(rowDGV.Cells(Me.ColPrimaryPoduct.Index).Value, rowDGV.Cells(Me.ColPrimarySO.Index).Value, rowDGV.Index)
                            End If
                        Else
                            rowDGV1 = CType(rowDGV.Clone(), DataGridViewRow)
                            For i As Integer = 0 To rowDGV.Cells.Count - 1
                                rowDGV1.Cells(i).Value = rowDGV.Cells(i).Value
                            Next
                            Me.dgvItemReject.Rows.Add(rowDGV1)
                        End If
                    Next
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            If Me.rbSO.Checked Then
                If MsgBox("Uncheck all data with no batch controlled?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    For Each rowDGV In Me.dgvItem.Rows
                        If rowDGV.Cells(Me.ColBatchControlled.Index).Value = 0 Then
                            rowDGV.Cells(Me.ColChkDO.Index).Value = Boolean.FalseString
                            rowDGV.Cells(Me.ColQtyDO.Index).Value = 0
                            'clear stock temp 1 1
                            Me.HapusStockTemp(rowDGV)
                        End If
                    Next
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                End If
            Else
                If MsgBox("Uncheck all data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    For Each rowDGV In Me.dgvItem.Rows
                        'If rowDGV.Cells(Me.ColBatchControlled.Index).Value = 0 Then
                        rowDGV.Cells(Me.ColChkDO.Index).Value = Boolean.FalseString
                        rowDGV.Cells(Me.ColQtyDO.Index).Value = 0
                        'End If
                    Next
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    'hapus stock sementara
                    Me.TblStockTemp.Clear()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub dgvCreatedBy_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCreatedBy.CellBeginEdit
        If e.ColumnIndex = Me.ColUserCode.Index Then
            Me.Buffer = Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value
        Else
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub dgvCreatedBy_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCreatedBy.CellEndEdit
        Dim KodeUser As String = ""
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim RowIdx As Integer = e.RowIndex
        Try
            If IsNothing(Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColUserCode.Index Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    If e.RowIndex <> Me.dgvCreatedBy.RowCount - 1 Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                    Exit Sub
                End If
            Else
                ''data account number
                If e.ColumnIndex = Me.ColUserCode.Index Then
                    Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjUser.ApakahSudahAda(Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                        KodeUser = Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value
                        index = Me.GetIndexItem(KodeUser, e.RowIndex)
                        If index < 0 Then
                            Me.dgvCreatedBy.Item(Me.ColUserCode.Index, e.RowIndex).Value = Me.dgvCreatedBy.Item(Me.ColUserCode.Index, e.RowIndex).Value
                            Me.dgvCreatedBy.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.dgvCreatedBy.Item(Me.ColUserCode.Index, e.RowIndex).Value)
                        Else
                            If e.RowIndex <> index Then
                                MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                Me.dgvCreatedBy.Item(Me.ColUserCode.Index, e.RowIndex).Value = ""
                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                                Exit Sub
                            End If
                        End If
                    Else
                        '2. kode akun yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmCariMasterUser
                        ij.backColorAll = Me.BackColorAll
                        ij.TxtCari.Text = Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.Kodeuser <> "" Then
                            index = Me.GetIndexItem(ij.Kodeuser, e.RowIndex)
                            If index < 0 Then
                                Me.dgvCreatedBy.Item(Me.ColUserCode.Index, e.RowIndex).Value = ij.Kodeuser
                                Me.dgvCreatedBy.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Exit Sub
                                Else
                                    Me.dgvCreatedBy.Item(Me.ColUserCode.Index, e.RowIndex).Value = ij.Kodeuser
                                    Me.dgvCreatedBy.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                End If
                            End If
                            ''2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                            'Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = ij.Kodeuser.ToUpper
                            'Me.dgvCreatedBy.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value)
                        ElseIf ij.Kodeuser = "" Then
                            If e.ColumnIndex = Me.ColUserCode.Index Then
                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                    If Me.dgvCreatedBy.Rows.Count - 1 <> e.RowIndex Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                Else
                                    Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                            Else
                                Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                        End If
                        ij = Nothing
                    End If
                End If
            End If
            Me.ObjCD.SetNumber(Me.dgvCreatedBy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vKodeUser As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvCreatedBy.Rows.Count - 2
                RowDGV = Me.dgvCreatedBy.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColUserCode.Index).Value = vKodeUser Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(200)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.EraseRowIndex <> -1 Then
                Me.dgvCreatedBy.Rows.RemoveAt(Me.EraseRowIndex)
                Me.EraseRowIndex = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#Region "Datagridview Filtering"
    Private Sub dgvcustomer_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvCustomer.CellBeginEdit
        Try
            If Me.DgvCustomer.Item(Me.ColCust_Primary.Index, e.RowIndex).Value <> -1 Then
                If e.ColumnIndex = Me.ColCust_Code.Index Then
                    Me.Buffer = Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColCust_Code.Index Then
                    Me.Buffer = Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvcustomer_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCustomer.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvCustomer.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim Customer As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCust_Code.Index Then
                        Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjCustomer.ApakahSudahAda(Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Customer = Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
                            index = Me.GetIndexItem_Cus(Customer, e.RowIndex)
                            If index < 0 Then
                                Me.DgvCustomer.Item(Me.ColCust_Primary.Index, e.RowIndex).Value = Me.ObjCustomer.GetPrimary(Me.DgvCustomer.Item(Me.ColCust_Code.Index, e.RowIndex).Value)
                                Me.DgvCustomer.Item(Me.ColCust_Code.Index, e.RowIndex).Value = Me.DgvCustomer.Item(Me.ColCust_Code.Index, e.RowIndex).Value
                                Me.DgvCustomer.Item(Me.ColCust_Name.Index, e.RowIndex).Value = Me.ObjCustomer.GetCustName(Me.DgvCustomer.Item(Me.ColCust_Primary.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvCustomer.Item(Me.ColCust_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex_Supp = e.RowIndex
                                        Me.BgwCustomer.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterCustomer
                            ij.txtCari.Text = Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomer <> -1 Then
                                index = Me.GetIndexItem_Cus(ij.PrimaryCustomer, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvCustomer.Item(Me.ColCust_Code.Index, e.RowIndex).Value = Me.ObjCustomer.GetCustomerCode(ij.PrimaryCustomer)
                                    Me.DgvCustomer.Item(Me.ColCust_Name.Index, e.RowIndex).Value = Me.ObjCustomer.GetCustName(ij.PrimaryCustomer)
                                    Me.DgvCustomer.Item(Me.ColCust_Primary.Index, e.RowIndex).Value = ij.PrimaryCustomer
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_Supp = e.RowIndex
                                            Me.BgwCustomer.RunWorkerAsync()
                                        Else
                                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Exit Sub
                                    Else
                                        Me.DgvCustomer.Item(Me.ColCust_Code.Index, e.RowIndex).Value = Me.ObjCustomer.GetCustomerCode(ij.PrimaryCustomer)
                                        Me.DgvCustomer.Item(Me.ColCust_Name.Index, e.RowIndex).Value = Me.ObjCustomer.GetCustName(ij.PrimaryCustomer)
                                        Me.DgvCustomer.Item(Me.ColCust_Primary.Index, e.RowIndex).Value = ij.PrimaryCustomer
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomer = -1 Then
                                If e.ColumnIndex = Me.ColCust_Code.Index Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        If Me.DgvCustomer.Rows.Count - 1 <> e.RowIndex Then
                                            Me.EraseRowIndex_Supp = e.RowIndex
                                            Me.BgwCustomer.RunWorkerAsync()
                                        End If
                                    Else
                                        Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                Else
                                    Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvCustomer)
                    End If
                Else
                    Me.EraseRowIndex_Supp = e.RowIndex
                    Me.BgwCustomer.RunWorkerAsync()
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_Cus(ByVal vKodeSupplier As String, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvCustomer.Rows.Count - 2
                RowDGV = Me.DgvCustomer.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCust_Code.Index).Value = vKodeSupplier Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BgwSupplier_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwCustomer.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwSupplier_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwCustomer.RunWorkerCompleted
        Try
            Me.DgvCustomer.Rows.RemoveAt(Me.EraseRowIndex_Supp)
            Me.ObjCD.SetNumber(Me.DgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvcustomer_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvCustomer.DataError

    End Sub

    Private Sub dgvcustomer_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvCustomer.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            rowDGV = Me.DgvCustomer.Rows(Me.DgvCustomer.RowCount - 1)
            With rowDGV
                .Cells(Me.ColCust_No.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColCust_Code.Index).Value = ""
                .Cells(Me.ColCust_Code.Index).Value = ""
                .Cells(Me.ColCust_Primary.Index).Value = "-1"
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvcustomer_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvCustomer.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.DgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvcustomer_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCustomer.Sorted
        Me.ObjCD.SetNumber(Me.DgvCustomer)
    End Sub

    Private Sub dgvcustomer_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvCustomer.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    ''
    Private Sub dgvItemProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemProductGroup.CellBeginEdit
        Dim rowDGV As DataGridViewRow

        Try
            'yang bisa diedit:
            '1. jika baris baru
            If e.RowIndex = Me.dgvItemProductGroup.Rows.Count - 1 Then
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'pastikan baris diatasnya sudah beres dahulu
                    If e.RowIndex > 0 Then
                        rowDGV = Me.dgvItemProductGroup.Rows(e.RowIndex - 1)
                        If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value.ToString.Trim = "" Then
                            MsgBox("Please specify " + rowDGV.Cells(Me.ColProductGroup_Level.Index).Value.ToString + _
                            " code on the previous row.", MsgBoxStyle.Information, "Information")
                            Me.dgvItemProductGroup.ClearSelection()
                            rowDGV.Selected = True
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = -1 Then
                    '2. jika bukan baris baru dan primary =-1
                    'yang bisa diedit hanya code product group dan level
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    '3. jika bukan baris baru dan primary <> -1
                    'yang bisa diedit: code, remarks
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or _
                    e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProductGroup.CellEndEdit
        Dim Level As Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand

        Try
            If IsNothing(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'Remarks
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    'kalo kode product group dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                    Exit Sub
                End If
            Else
                'level
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'do nothing

                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If

                    'cek ke master
                    Level = Me.objItem.GetProductLevel(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
                    If Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                        Me.EndEditProductGroup(e, Me.ProductGroup_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                        Me.EndEditProductGroup(e, Me.Brand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                        Me.EndEditProductGroup(e, Me.SubBrand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                        Me.EndEditCategory(e)
                    Else
                        Me.EndEditProduct(e)
                    End If
                ElseIf e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'remarks
                    'pastikan ditrim dan  max 250 karakter
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'tidak boleh ada karakter aneh2
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                        MsgBox("Please avoid using the following character: ', *, %, #.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex_ItemProductGroup(ByVal vCode As String, ByVal vRefIndex As Integer, ByVal vProductGroupLevelStr As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = vCode And _
                rowDGV.Index <> vRefIndex And _
                rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = vProductGroupLevelStr Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub EndEditProductGroup(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal vProductGroup_Level As Integer)
        Try
            If Me.ObjProductGroup.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level) = True Then
                '1. kode product group ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product group yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.FrmCariMasterProductGroup
                ij.ProductGroupLevel = vProductGroup_Level
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProductGroup <> -1 Then
                    '2.a. product group yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProductGroup.GetGroupCode(ij.PrimaryProductGroup)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(ij.PrimaryProductGroup)
                ElseIf ij.PrimaryProductGroup = -1 Then
                    '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Trim(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If
            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.Buffer, vProductGroup_Level)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditCategory(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.objCategory.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode category ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.objCategory.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.objCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode category yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterCategory
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryCategory <> -1 Then
                    '2.a. category yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryCategory
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCategory.GetCategoryCode(ij.PrimaryCategory)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.objCategory.GetDesc(ij.PrimaryCategory)
                ElseIf ij.PrimaryCategory = -1 Then
                    '2.b. kode category yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.objCategory.GetPrimary(Me.Buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.objCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditProduct(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjProduct.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode product ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterProduct
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProduct <> -1 Then
                    '2.a. product yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProduct
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                ElseIf ij.PrimaryProduct = -1 Then
                    '2.b. kode product yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.Buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemProductGroup.DataError

    End Sub

    Private Sub dgvItemProductGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemProductGroup.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemProductGroup.Rows(Me.dgvItemProductGroup.Rows.Count - 1)
            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = Me.objItem.GetProductLevelStr(Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup)
            rowDGV.Cells(Me.ColProductGroup_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value = -1
            rowDGV.Cells(Me.ColProductGroup_Remarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemProductGroup.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemProductGroup.Sorted
        Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
    End Sub

    Private Sub dgvItemProductGroup_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemProductGroup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            If Me.DummyIndex <> -1 Then
                Me.dgvItemProductGroup.Rows.RemoveAt(Me.DummyIndex)
                DummyIndex = -1
                Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    Private Sub BtnFindLoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindLoc.Click
        Dim i As New Core_Module.FrmCariMasterStorageArea
        i.BackColorAll = Me.BackColorAll

        Try
            'i.cboCriteria.SelectedIndex = 0
            If Me.TxtLocCode.Text <> "" Then
                i.txtCari.Text = Me.TxtLocCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.TxtPrimaryLoc.Text = i.PrimaryStorageArea
                Me.TxtLocCode.Text = Me.ObjLoc.GetStorageAreaCode(i.PrimaryStorageArea)
                Me.txtPrimaryWarehouse.Text = Me.ObjLoc.GetPrimaryWarehouse(Me.TxtPrimaryLoc.Text)
                If Me.dgvItem.Rows.Count > 0 Then
                    isLocationChanged = True
                End If
            Else
                Me.TxtPrimaryLoc.Text = "-1"
                Me.txtPrimaryWarehouse.Text = -1
                Me.TxtLocCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtLocCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocCode.GotFocus
        Try
            Me.TxtLocCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                Dim Primary As Double = ObjLoc.GetPrimary(Me.TxtPrimaryLoc.Text)
                If Primary <> -1 Then
                    Me.TxtPrimaryLoc.Text = Primary
                    Me.TxtLocCode.Text = Me.ObjLoc.GetStorageAreaCode(Me.TxtPrimaryLoc.Text)
                    Me.txtPrimaryWarehouse.Text = Me.ObjLoc.GetPrimaryWarehouse(Me.TxtPrimaryLoc.Text)
                    If Me.dgvItem.Rows.Count > 0 Then
                        isLocationChanged = True
                    End If
                Else
                    Me.BtnFindLoc.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocCode.Leave
        Try
            If Me.TxtLocCode.Text <> "" Then
                Dim Primary As Double = ObjLoc.GetPrimary(Me.TxtLocCode.Text)
                If Primary <> -1 Then
                    Me.TxtPrimaryLoc.Text = Primary
                    Me.TxtLocCode.Text = Me.ObjLoc.GetStorageAreaCode(Me.TxtPrimaryLoc.Text)
                    Me.txtPrimaryWarehouse.Text = Me.ObjLoc.GetPrimaryWarehouse(Me.TxtPrimaryLoc.Text)
                    If Me.dgvItem.Rows.Count > 0 Then
                        isLocationChanged = True
                    End If
                Else
                    Me.BtnFindLoc.PerformClick()
                End If
            Else
                Me.TxtPrimaryLoc.Text = "-1"
                Me.txtPrimaryWarehouse.Text = -1
                Me.TxtLocCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocCode.LostFocus
        Try
            Me.TxtLocCode.BackColor = Me.TxtPrimaryLoc.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Try
            If Me.rbSO.Checked AndAlso (Me.TxtPrimaryLoc.Text = -1 Or Me.TxtPrimaryLoc.Text = "") Then
                MsgBox("Please enter location first.", MsgBoxStyle.Information, "Information")
                Me.TxtLocCode.Focus()
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim UserCodeList As New List(Of String)
        Dim rowDGV As DataGridViewRow
        'Dim ExpiredDateStr As String = ""
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If Me.ApakahDataBenar() = False Then
                Exit Sub
            End If

            Me.dgvCreatedBy.AllowUserToAddRows = False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'Me.ChkIgnoreLowTolerance.Checked = False
            'kosongkan dgv item
            Me.dgvItem.Rows.Clear()
            Me.Tbl.Rows.Clear()
            Me.TblSimpan.Rows.Clear()

            'load data-filter customer
            If Me.rbSO.Checked Then
                If Me.TxtPrimaryTransaction.Text = "" Or Me.TxtPrimaryTransaction.Text = "-1" Then
                    If Me.ChkIgnoreLowTolerance.Checked = True Then
                        DTA.Fill_Tolerance(Tbl, Me.dtpFrom.Value, Me.dtpTo.Value, "%" & Me.txtUserCode.Text.Trim & "%")
                    Else
                        DTA.Fill(Tbl, Me.dtpFrom.Value, Me.dtpTo.Value, "%" & Me.txtUserCode.Text.Trim & "%")
                    End If
                Else
                    If Me.ChkIgnoreLowTolerance.Checked = True Then
                        DTA.FillBy_PrimarySO_Torence(Tbl, Me.dtpFrom.Value, Me.dtpTo.Value, Me.TxtPrimaryTransaction.Text, "%" & Me.txtUserCode.Text.Trim & "%")
                    Else
                        DTA.FillBy_PrimarySO(Tbl, Me.dtpFrom.Value, Me.dtpTo.Value, Me.TxtPrimaryTransaction.Text, "%" & Me.txtUserCode.Text.Trim & "%")
                    End If
                End If
            Else
                If Me.TxtPrimaryTransaction.Text = "" Or Me.TxtPrimaryTransaction.Text = "-1" Then
                    If Me.ChkIgnoreLowTolerance.Checked = True Then
                        DTABook.Fill_Tolerance(TblBook, Me.dtpFrom.Value, Me.dtpTo.Value, "%" & Me.txtUserCode.Text.Trim & "%")
                    Else
                        DTABook.Fill(TblBook, Me.dtpFrom.Value, Me.dtpTo.Value, "%" & Me.txtUserCode.Text.Trim & "%")
                    End If
                Else
                    If Me.ChkIgnoreLowTolerance.Checked = True Then
                        DTABook.FillBy_PrimaryBooking_Tolerance(TblBook, Me.dtpFrom.Value, Me.dtpTo.Value, Me.TxtPrimaryTransaction.Text, "%" & Me.txtUserCode.Text.Trim & "%")
                    Else
                        DTABook.Fillby_PrimaryBooking(TblBook, Me.dtpFrom.Value, Me.dtpTo.Value, Me.TxtPrimaryTransaction.Text, "%" & Me.txtUserCode.Text.Trim & "%")
                    End If
                End If
                'Me.Tbl = TblBook.Copy
                Me.Tbl.Clear()

                For Each row As DataRow In TblBook.Rows
                    Me.Tbl.ImportRow(row)
                Next
            End If

            'Me.TBLUnit.Rows.Clear()
            'Me.TBLUnit1.Rows.Clear()
            'Me.DTAUnit.Fill(TBLUnit)
            'Me.DTAUnit1.Fill(TBLUnit1)
            'Me.DTAUnit2.Fill(Me.TBLUnit2)

            Me.GenerateFiltering()
            Dim cust As String = ""
            If Me.cboCustomerFormat.SelectedIndex = 0 Then
                cust = "Customer Code"
            ElseIf Me.cboCustomerFormat.SelectedIndex = 1 Then
                cust = "Customer Name"
            Else
                cust = "Customer Code"
            End If
            'pindahkan ke dgv
            For Each Row In Tbl.Select("", "[" & cust & "] ASC, [No Transaksi] ASC")
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Row.Amount, Row.Decimal_Curr)
                rowDGV.Cells(Me.ColChkDO.Index).Value = Boolean.FalseString
                rowDGV.Cells(Me.ColCodeNo.Index).Value = Row.Product_Code
                rowDGV.Cells(Me.ColCurrency.Index).Value = Row.Currency_Code
                rowDGV.Cells(Me.ColDate.Index).Value = Row.Tanggal_Transaksi
                rowDGV.Cells(Me.ColDesc.Index).Value = Row.Description
                rowDGV.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Row.Disc_Amount, Row.Decimal_Curr)
                rowDGV.Cells(Me.ColDiscPerct.Index).Value = Me.ObjCD.DigitGrouping(Row.Discount, 2)
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                rowDGV.Cells(Me.ColDocumentNO.Index).Value = Row.No_Transaksi
                If Me.rbSO.Checked Then
                    rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(Row.SOB_Multi, Row.Decimal_Multi)
                Else
                    If Row.Primary_Unit_Default <> Row.Primary_Units_Multi Then
                        rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(Row.Qty_Multi, Row.Decimal_Multi)
                    Else
                        rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(Row.Qty, Row.Decimal_Multi)
                    End If

                End If
                rowDGV.Cells(Me.ColTolerance.Index).Value = Me.ObjCD.DigitGrouping(Row.Tolerance_Up, Row.Decimal_Multi)
                rowDGV.Cells(Me.ColDecimalMulti.Index).Value = Row.Decimal_Multi
                rowDGV.Cells(Me.ColQtyDO.Index).Value = 0
                If Me.cboCustomerFormat.SelectedIndex = 0 Then
                    rowDGV.Cells(Me.ColCustomer.Index).Value = Row.Customer_Code
                ElseIf Me.cboCustomerFormat.SelectedIndex = 1 Then
                    rowDGV.Cells(Me.ColCustomer.Index).Value = Row.Customer_Name
                Else
                    rowDGV.Cells(Me.ColCustomer.Index).Value = Row.Customer_Code + " - " + Row.Customer_Name
                End If

                If Row.Primary_Unit_Default <> Row.Primary_Units_Multi Then
                    rowDGV.Cells(Me.ColUnit.Index).Value = Row.Unit_Code_Multi
                Else
                    rowDGV.Cells(Me.ColUnit.Index).Value = Row.Unit_Code
                End If

                rowDGV.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(Row.Unit_Price, Row.Decimal_Default)
                rowDGV.Cells(Me.ColPrimaryCustomer.Index).Value = Row.Primary_Customer
                rowDGV.Cells(Me.ColPrimaryPoduct.Index).Value = Row.Primary_Product
                rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value = Row.Primary_Units_Multi
                rowDGV.Cells(Me.ColPrimaryTransaction.Index).Value = Row.Primary
                If Me.rbSO.Checked Then
                    rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = Me.TxtPrimaryLoc.Text
                    rowDGV.Cells(Me.ColPrimaryWarehouse.Index).Value = Me.txtPrimaryWarehouse.Text
                Else
                    rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = Row.Primary_Location
                    rowDGV.Cells(Me.ColPrimaryWarehouse.Index).Value = Row.Primary_Warehouse
                End If
                rowDGV.Cells(Me.ColRemarksItem.Index).Value = ""
                rowDGV.Cells(Me.ColPrimaruUnitDefault.Index).Value = Row.Primary_Unit_Default
                rowDGV.Cells(Me.ColQtyDefault.Index).Value = Me.ObjCD.DigitGrouping(Row.SOB_Multi * Row.Ratio_Default / Row.Ratio_Reference, Row.Decimal_Default)
                rowDGV.Cells(Me.ColRatioDefault.Index).Value = Row.Ratio_Default
                rowDGV.Cells(Me.ColRatioReference.Index).Value = Row.Ratio_Reference
                rowDGV.Cells(Me.ColDefPrec_QtyDefault.Index).Value = Row.Decimal_Default
                rowDGV.Cells(Me.ColUnitDefault.Index).Value = Row.Unit_Code
                rowDGV.Cells(Me.ColBatchControlled.Index).Value = Row.Batch_No_Controlled
                rowDGV.Cells(Me.ColLocationControlled.Index).Value = Row.Location_Controlled
                If Row.Expiry_Date_Controlled = False Then
                    rowDGV.Cells(Me.ColExpiryControlled.Index).Value = 0
                Else
                    rowDGV.Cells(Me.ColExpiryControlled.Index).Value = 1
                End If
                rowDGV.Cells(Me.ColIdentifier_CustAddress.Index).Value = Row.Indentifier_Alamat_Customer 'bill to
                rowDGV.Cells(Me.ColTax.Index).Value = Row.Tax
                rowDGV.Cells(Me.ColFakturPajak.Index).Value = Row.Jenis_Faktur_Pajak
                rowDGV.Cells(Me.Colidentifier_address_ship_to.Index).Value = Row.identifier_address_ship_to
                rowDGV.Cells(Me.ColPrimary_Customer_Ship_To.Index).Value = Row.primary_customer_ship_to
                rowDGV.Cells(Me.Colidentifier_address_customer_ship.Index).Value = Row.identifier_address_customer_ship_to
                rowDGV.Cells(Me.ColPrimarySO.Index).Value = Row.Primary_SO
                rowDGV.Cells(Me.ColSONo.Index).Value = Row.SO_No
                'rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = Row.Primary_Location
                rowDGV.Cells(Me.ColBatchNo.Index).Value = Row.Batch_No
                rowDGV.Cells(Me.ColExpiryDate.Index).Value = Format(Row.Expiry_Date, "dd MMM yyyy")
            Next
            Me.IsikanTbl(True)
            Me.IsikanStock(True)
            Me.isLocationChanged = False
            Me.dgvCreatedBy.AllowUserToAddRows = True
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
            UserCodeList = Nothing
        End Try
    End Sub

    Private Sub GenerateFiltering()
        Dim FilterProduct As String = ""
        Dim FilterCustomer As String = ""
        Dim FilterGroup As String = ""
        Dim FilterUser As String = ""
        Dim rowDGV As New DataGridViewRow
        Dim Filter As String = ""
        Dim row As DseTransaksiQuick_DO.DataTable1Row
        Dim TBLCopyan As New DseTransaksiQuick_DO.DataTable1DataTable
        Dim rowCopyan As DseTransaksiQuick_DO.DataTable1Row

        Try
            If Me.dgvCreatedBy.Rows.Count > 0 Then
                For Each rowDGV In Me.dgvCreatedBy.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColUserCode.Index).Value) = False Then
                        If FilterUser = "" Then
                            FilterUser = "[Created By] = '" & rowDGV.Cells(Me.ColUserCode.Index).Value & "'"
                        Else
                            FilterUser &= " Or [Created By] = '" & rowDGV.Cells(Me.ColUserCode.Index).Value & "'"
                        End If
                    End If
                Next
            End If

            If Me.DgvCustomer.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvCustomer.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColCust_Code.Index).Value) = False Then
                        If FilterCustomer = "" Then
                            FilterCustomer = "[Primary Customer] = " & rowDGV.Cells(Me.ColCust_Primary.Index).Value
                        Else
                            FilterCustomer &= " Or [Primary Customer] = " & rowDGV.Cells(Me.ColCust_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
                    TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)
                    If FilterGroup = "" Then
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " [Primary Product Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            FilterGroup &= " [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            FilterGroup &= " [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            FilterGroup &= " [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        Else
                            FilterGroup &= " [Primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        End If
                    Else
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " OR [Primary Product Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            FilterGroup &= " OR [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            FilterGroup &= " OR [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            FilterGroup &= " OR [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        Else
                            FilterGroup &= " OR [Primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        End If
                    End If
                End If
            Next

            If FilterUser <> "" Then
                Filter = FilterUser
            End If

            If FilterCustomer <> "" Then
                If Filter = "" Then
                    Filter = FilterCustomer
                Else
                    Filter = "(" & Filter & ") AND (" & FilterCustomer & ") "
                End If
            End If

            If FilterGroup <> "" Then
                If Filter = "" Then
                    Filter = FilterGroup
                Else
                    Filter = "(" & Filter & ") AND (" & FilterGroup & ") "
                End If
            End If

            If Filter <> "" Then
                Me.Tbl.DefaultView.RowFilter = Filter
                For i As Integer = 0 To Me.Tbl.DefaultView.Count - 1
                    rowCopyan = Me.Tbl.DefaultView(i).Row
                    TBLCopyan.ImportRow(rowCopyan)
                Next

                Me.Tbl.Rows.Clear()
                For Each row In TBLCopyan.Rows
                    Tbl.ImportRow(row)
                Next
            End If

            ''yang masuk hanya yang sudah diautorisasi
            'For i As Integer = Me.Tbl.Rows.Count - 1 To 0 Step -1
            '    row = Me.Tbl.Rows(i)
            '    If Me.ObjDO.GetAuthorizationLevel(row.Primary) <> Me.ObjDO.GetAuthorizationLevelNeeded(row.Primary) Then
            '        Me.Tbl.Rows.Remove(row)
            '    End If
            'Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yg bisa diedit:
            If e.ColumnIndex = Me.ColChkDO.Index Then
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            ElseIf e.ColumnIndex = Me.ColQtyDO.Index Or e.ColumnIndex = Me.ColRemarksItem.Index Then
                If e.ColumnIndex = Me.ColQtyDO.Index Then
                    If Me.dgvItem.Item(Me.ColBatchControlled.Index, e.RowIndex).Value = 1 And Me.dgvItem.Item(Me.ColChkDO.Index, e.RowIndex).Value = Boolean.TrueString And _
                    Me.rbSO.Checked Then
                        Dim ij As New FrmTransaksiDO_Quick_PerBatch
                        ij.BackColorAll = Me.BackColorAll
                        ij.IsBatchControlled = Me.dgvItem.Item(Me.ColBatchControlled.Index, e.RowIndex).Value
                        ij.IsExpiryControlled = Me.dgvItem.Item(Me.ColExpiryControlled.Index, e.RowIndex).Value
                        ij.PrimaryCode = Me.dgvItem.Item(Me.ColPrimaryPoduct.Index, e.RowIndex).Value
                        ij.PrimaryTransaction = Me.dgvItem.Item(Me.ColPrimaryTransaction.Index, e.RowIndex).Value
                        ij.Tbl = Me.TblSimpan.Copy
                        If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                            TblSimpan = ij.Tbl.Copy
                            Me.SalinBarisKosong(Me.dgvItem.Item(Me.ColPrimaryPoduct.Index, e.RowIndex).Value, _
                             Me.dgvItem.Item(Me.ColPrimaryTransaction.Index, e.RowIndex).Value, e.RowIndex)
                        Else
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString
                        End If

                        Dim QtyNow As Decimal = 0
                        Me.TblSimpan.DefaultView.RowFilter = "[Primary Product]=" & Me.dgvItem.Item(Me.ColPrimaryPoduct.Index, e.RowIndex).Value & _
                        " AND [Primary Transaction]=" & Me.dgvItem.Item(Me.ColPrimaryTransaction.Index, e.RowIndex).Value
                        If Me.TblSimpan.DefaultView.Count > 0 Then
                            For i As Integer = 0 To Me.TblSimpan.DefaultView.Count - 1
                                Me.RowSimpan = Me.TblSimpan.DefaultView(i).Row
                                QtyNow += Me.RowSimpan.Qty_In_Multi
                            Next
                        End If
                        Me.dgvItem.Item(Me.ColQtyDO.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(QtyNow, Me.dgvItem.Item(Me.ColDecimalMulti.Index, e.RowIndex).Value)
                        e.Cancel = True
                        Exit Sub
                    End If
                End If

                If Me.dgvItem.Item(Me.ColChkDO.Index, e.RowIndex).Value = Boolean.TrueString Then
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim vPrimary As Double = -1

        Try
            If e.ColumnIndex = -1 Or e.RowIndex = -1 Then
                Exit Sub
            End If

            If e.ColumnIndex = Me.ColDocumentNO.Index Then
                If Me.rbSO.Checked Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO") = False Then
                        MsgBox("Insufficient acces.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    Else
                        vPrimary = Me.dgvItem.Item(Me.ColPrimaryTransaction.Index, e.RowIndex).Value
                        Dim ij As Sales_Module.FrmTransaksiSO
                        Dim j As Windows.Forms.Form

                        For Each j In Me.ParentForm.MdiChildren
                            ij = TryCast(j, Sales_Module.FrmTransaksiSO)
                            If IsNothing(ij) = False Then
                                ij.BringToFront()
                                ij.Focus()
                                ij.ShowTransaction(vPrimary)
                                Exit Sub
                            End If
                        Next

                        Dim i As New Sales_Module.FrmTransaksiSO
                        ' Increment the caption counter.
                        i.UserCode = Me.UserCode.ToString
                        ' Set the caption to be unique.
                        i.BackColorAll = Core_Module.GetSMESSBackColorAll
                        i.MdiParent = Me.ParentForm
                        i.Show()
                        i.ShowTransaction(vPrimary)
                    End If
                End If

                If Me.rbInventoryBooking.Checked Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiBookingInv") = False Then
                        MsgBox("Insufficient acces.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    Else
                        vPrimary = Me.dgvItem.Item(Me.ColPrimaryTransaction.Index, e.RowIndex).Value
                        Dim ij As Sales_Module.FrmTransaksiBookingInv
                        Dim j As Windows.Forms.Form

                        For Each j In Me.ParentForm.MdiChildren
                            ij = TryCast(j, Sales_Module.FrmTransaksiBookingInv)
                            If IsNothing(ij) = False Then
                                ij.BringToFront()
                                ij.Focus()
                                ij.ShowTransaction(vPrimary)
                                Exit Sub
                            End If
                        Next

                        Dim i As New Sales_Module.FrmTransaksiBookingInv
                        ' Increment the caption counter.
                        i.UserCode = Me.UserCode.ToString
                        ' Set the caption to be unique.
                        i.BackColorAll = Core_Module.GetSMESSBackColorAll
                        i.MdiParent = Me.ParentForm
                        i.Show()
                        i.ShowTransaction(vPrimary)
                    End If
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim rowDGV As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColQtyDO.Index Then
                    'kembalikan ke nilai buffer
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                ElseIf e.ColumnIndex = Me.ColRemarksItem.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Exit Sub
                End If
            Else
                If e.ColumnIndex = Me.ColChkDO.Index Then
                    If Me.dgvItem.Item(Me.ColChkDO.Index, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If

                    'kemungkinan
                    If Me.dgvItem.Item(Me.ColBatchControlled.Index, e.RowIndex).Value = 1 Then
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString And Me.rbSO.Checked Then
                            Dim ij As New FrmTransaksiDO_Quick_PerBatch
                            ij.BackColorAll = Me.BackColorAll
                            ij.IsBatchControlled = Me.dgvItem.Item(Me.ColBatchControlled.Index, e.RowIndex).Value
                            ij.IsExpiryControlled = Me.dgvItem.Item(Me.ColExpiryControlled.Index, e.RowIndex).Value
                            ij.PrimaryCode = Me.dgvItem.Item(Me.ColPrimaryPoduct.Index, e.RowIndex).Value
                            ij.PrimaryTransaction = Me.dgvItem.Item(Me.ColPrimaryTransaction.Index, e.RowIndex).Value
                            ij.MultiUnit = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                            ij.PrimaryUnit = Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value
                            ij.PrimaryLocation = Me.dgvItem.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value
                            ij.PrimaryWarehouse = Me.dgvItem.Item(Me.ColPrimaryWarehouse.Index, e.RowIndex).Value
                            ij.txtProductCode.Text = Me.dgvItem.Item(Me.ColCodeNo.Index, e.RowIndex).Value
                            ij.txtDesciption.Text = Me.dgvItem.Item(Me.ColDesc.Index, e.RowIndex).Value
                            ij.txtTransactionNo.Text = Me.dgvItem.Item(Me.ColDocumentNO.Index, e.RowIndex).Value
                            If Me.rbSO.Checked Then
                                ij.lblTransactionNo.Text = "SO No."
                                ij.TxtQtyTotal.Text = rowDGV.Cells(Me.ColQty.Index).Value
                            Else
                                ij.lblTransactionNo.Text = "Inv. Book No."
                            End If

                            ij.Tbl = Me.TblSimpan.Copy
                            ij.TblStock = Me.TBLStock.Copy
                            ij.TblStockTemp = Me.TblStockTemp.Copy
                            ij.TblPendingDO = Me.TblPendingDO.Copy
                            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                                TblSimpan = ij.Tbl.Copy
                                Me.TBLStock = ij.TblStock.Copy
                                Me.TblStockTemp = ij.TblStockTemp.Copy
                                Me.TblPendingDO = ij.TblPendingDO.Copy
                                'Me.SalinBarisKosong(Me.dgvItem.Item(Me.ColPrimaryPoduct.Index, e.RowIndex).Value, _
                                ' Me.dgvItem.Item(Me.ColPrimaryTransaction.Index, e.RowIndex).Value, e.RowIndex)
                            Else
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString
                            End If

                            Dim QtyNow As Decimal = 0
                            Me.TblSimpan.DefaultView.RowFilter = "[Primary Product]=" & Me.dgvItem.Item(Me.ColPrimaryPoduct.Index, e.RowIndex).Value & " AND [Primary Transaction]=" & Me.dgvItem.Item(Me.ColPrimaryTransaction.Index, e.RowIndex).Value
                            If Me.TblSimpan.DefaultView.Count > 0 Then
                                For i As Integer = 0 To Me.TblSimpan.DefaultView.Count - 1
                                    Me.RowSimpan = Me.TblSimpan.DefaultView(i).Row
                                    QtyNow += Me.RowSimpan.Qty_In_Multi
                                Next
                            End If
                            Me.dgvItem.Item(Me.ColQtyDO.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(QtyNow, Me.dgvItem.Item(Me.ColDecimalMulti.Index, e.RowIndex).Value)

                        ElseIf Me.rbInventoryBooking.Checked And Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString Then
                            If Me.CekStock(rowDGV) Then
                                rowDGV.Cells(Me.ColQtyDO.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQty.Index).Value, rowDGV.Cells(Me.ColDecimalMulti.Index).Value)
                                Me.SalinBarisBatch(rowDGV.Cells(Me.ColPrimaryPoduct.Index).Value, rowDGV.Cells(Me.ColPrimarySO.Index).Value, rowDGV.Index)
                                Me.simpanStockTemp(rowDGV, rowDGV.Cells(Me.ColQtyDO.Index).Value)
                            Else
                                Me.dgvItem.Item(Me.ColQtyDO.Index, e.RowIndex).Value = 0
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString
                                MsgBox("Qty is not available.", MsgBoxStyle.Information, "Information")
                            End If
                        Else
                            Me.TblSimpan.DefaultView.RowFilter = "[Primary Product]=" & Me.dgvItem.Item(Me.ColPrimaryPoduct.Index, e.RowIndex).Value & " AND [Primary Transaction]=" & Me.dgvItem.Item(Me.ColPrimaryTransaction.Index, e.RowIndex).Value
                            For i As Integer = Me.TblSimpan.DefaultView.Count - 1 To 0 Step -1
                                Me.RowSimpan = Me.TblSimpan.DefaultView(i).Row
                                Me.TblSimpan.Rows.Remove(Me.RowSimpan)
                            Next
                            Me.SalinBarisKosong(Me.dgvItem.Item(Me.ColPrimaryPoduct.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryTransaction.Index, e.RowIndex).Value, e.RowIndex)
                            Me.HapusStockTempDgv(rowDGV)
                            Dim QtyNow As Decimal = 0
                            Me.TblSimpan.DefaultView.RowFilter = "[Primary Product]=" & Me.dgvItem.Item(Me.ColPrimaryPoduct.Index, e.RowIndex).Value & " AND [Primary Transaction]=" & Me.dgvItem.Item(Me.ColPrimaryTransaction.Index, e.RowIndex).Value
                            If Me.TblSimpan.DefaultView.Count > 0 Then
                                For i As Integer = 0 To Me.TblSimpan.DefaultView.Count - 1
                                    Me.RowSimpan = Me.TblSimpan.DefaultView(i).Row
                                    QtyNow += Me.RowSimpan.Qty_In_Multi
                                Next
                            End If
                            Me.dgvItem.Item(Me.ColQtyDO.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(QtyNow, Me.dgvItem.Item(Me.ColDecimalMulti.Index, e.RowIndex).Value)
                        End If
                    Else
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString Then
                            'di cek
                            Me.dgvItem.Item(Me.ColQtyDO.Index, e.RowIndex).Value = CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) + CDec(Me.dgvItem.Item(Me.ColTolerance.Index, e.RowIndex).Value)
                            If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(Me.ColQtyDO.Index, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString
                            End If
                            If Me.CekStock(rowDGV) Then
                                rowDGV.Cells(Me.ColQtyDO.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQty.Index).Value, rowDGV.Cells(Me.ColDecimalMulti.Index).Value)
                                If rowDGV.Cells(Me.ColBatchControlled.Index).Value = 1 Then
                                    Me.SalinBarisBatch(rowDGV.Cells(Me.ColPrimaryPoduct.Index).Value, rowDGV.Cells(Me.ColPrimarySO.Index).Value, rowDGV.Index)
                                    Me.simpanStockTemp(rowDGV, rowDGV.Cells(Me.ColQtyDO.Index).Value)
                                End If
                            Else
                                Me.dgvItem.Item(Me.ColQtyDO.Index, e.RowIndex).Value = 0
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString
                                MsgBox("Qty is not available.", MsgBoxStyle.Information, "Information")
                            End If
                        Else
                            'di un check
                            Me.dgvItem.Item(Me.ColQtyDO.Index, e.RowIndex).Value = 0
                            Me.dgvItem.Item(Me.ColRemarksItem.Index, e.RowIndex).Value = ""
                            Me.HapusStockTempDgv(rowDGV)
                        End If
                    End If

                ElseIf e.ColumnIndex = Me.ColQtyDO.Index Then
                    'pastikan non negatif
                    If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.dgvItem.Item(Me.ColBatchControlled.Index, e.RowIndex).Value = 1 And Me.rbInventoryBooking.Checked Then
                        'pastikan nilainya lebih kecil atau sama dari qty
                        If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > (CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) + CDec(Me.dgvItem.Item(Me.ColTolerance.Index, e.RowIndex).Value)) Then
                            MsgBox("Maximum value your can enter is: " + Me.ObjCD.DigitGrouping((CDec(Me.dgvItem.Item(Me.ColTolerance.Index, e.RowIndex).Value) + CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value)), Me.dgvItem.Item(Me.ColDecimalMulti.Index, e.RowIndex).Value) _
                            + " " + Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value, _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If Me.CekStock(rowDGV, CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)) Then
                            rowDGV.Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(e.ColumnIndex).Value, rowDGV.Cells(Me.ColDecimalMulti.Index).Value)
                            Me.SalinBarisBatch(rowDGV.Cells(Me.ColPrimaryPoduct.Index).Value, rowDGV.Cells(Me.ColPrimarySO.Index).Value, rowDGV.Index)
                            Me.simpanStockTemp(rowDGV, rowDGV.Cells(Me.ColQtyDO.Index).Value)
                        Else
                            Me.dgvItem.Item(Me.ColQtyDO.Index, e.RowIndex).Value = 0
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString
                            MsgBox("Qty is not available.", MsgBoxStyle.Information, "Information")
                        End If

                    Else
                        'pastikan nilainya lebih kecil atau sama dari qty
                        If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > (CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) + CDec(Me.dgvItem.Item(Me.ColTolerance.Index, e.RowIndex).Value)) Then
                            MsgBox("Maximum value your can enter is: " + Me.ObjCD.DigitGrouping((CDec(Me.dgvItem.Item(Me.ColTolerance.Index, e.RowIndex).Value) + CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value)), Me.dgvItem.Item(Me.ColDecimalMulti.Index, e.RowIndex).Value) _
                            + " " + Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value, _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If Me.CekStock(rowDGV, CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)) Then
                            'digit grouping
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColDecimalMulti.Index, e.RowIndex).Value)
                            'hitung ulang qty default
                            Me.dgvItem.Item(Me.ColQtyDefault.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value * _
                            Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value / Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColDefPrec_QtyDefault.Index, e.RowIndex).Value)
                            Me.simpanStockTemp(rowDGV, rowDGV.Cells(Me.ColQtyDO.Index).Value)
                        Else
                            Me.dgvItem.Item(Me.ColQtyDO.Index, e.RowIndex).Value = 0
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString
                            MsgBox("Qty is not available.", MsgBoxStyle.Information, "Information")
                        End If
                    End If

                ElseIf e.ColumnIndex = Me.ColRemarksItem.Index Then
                    'trim
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    'truncate jika lebih panjang dr 150
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 150 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 150)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinBarisKosong(ByVal vPrimaryCode As Double, ByVal vPrimarySO As Double, ByVal veRow As Integer)
        Dim rowDGV As DataGridViewRow = Me.dgvItem.Rows(veRow)
        Try
            Me.TblSimpan.DefaultView.RowFilter = "[Primary Product]=" & vPrimaryCode & " AND [Primary Transaction]=" & vPrimarySO
            If Me.TblSimpan.DefaultView.Count > 0 Then
                For i As Integer = 0 To Me.TblSimpan.DefaultView.Count - 1
                    Me.RowSimpan = Me.TblSimpan.DefaultView(i).Row
                    With rowDGV
                        If Me.rbSO.Checked Then
                            Me.RowSimpan.Primary_Transaction = .Cells(Me.ColPrimaryTransaction.Index).Value
                            Me.RowSimpan.Transaction_No = .Cells(Me.ColDocumentNO.Index).Value
                        Else
                            Me.RowSimpan.Primary_Transaction = .Cells(Me.ColPrimarySO.Index).Value
                            Me.RowSimpan.Transaction_No = .Cells(Me.ColSONo.Index).Value
                        End If
                        Me.RowSimpan.Primary_Customer = .Cells(Me.ColPrimaryCustomer.Index).Value
                        Me.RowSimpan.Primary_Currency = .Cells(Me.ColPrimaryCurrency.Index).Value
                        Me.RowSimpan.Primary_Product = .Cells(Me.ColPrimaryPoduct.Index).Value
                        Me.RowSimpan.Primary_Unit = .Cells(Me.ColPrimaruUnitDefault.Index).Value
                        Me.RowSimpan.Primary_Multi_Unit = .Cells(Me.ColPrimaryUnitMulti.Index).Value
                        Me.RowSimpan.Primary_Location = .Cells(Me.ColPrimaryLocation.Index).Value
                        Me.RowSimpan.Primary_Warehouse = .Cells(Me.ColPrimaryWarehouse.Index).Value
                        Me.RowSimpan.Customer_Code = .Cells(Me.ColCust_Code.Index).Value
                        Me.RowSimpan.Customer_Name = .Cells(Me.ColCust_Code.Index).Value
                        Me.RowSimpan.Unit_Code_Default = .Cells(Me.ColUnitDefault.Index).Value
                        Me.RowSimpan.Unit_Code_Multi = .Cells(Me.ColUnit.Index).Value
                        Me.RowSimpan.Product_Code = .Cells(Me.ColCodeNo.Index).Value
                        Me.RowSimpan.Product_Desc = .Cells(Me.ColDesc.Index).Value
                        Me.RowSimpan.SOB_Multi = .Cells(Me.ColQty.Index).Value
                        Me.RowSimpan.Tolerance_Up = .Cells(Me.ColTolerance.Index).Value
                        Me.RowSimpan.SOB = .Cells(Me.ColQtyDefault.Index).Value
                        Me.RowSimpan.Qty_In = .Cells(Me.ColQtyDO.Index).Value * CDec(.Cells(Me.ColRatioDefault.Index).Value) / CDec(.Cells(Me.ColRatioReference.Index).Value)  'isi qty default
                        'Me.RowSimpan.Qty_In_Multi = .Cells(Me.ColQtyDO.Index).Value
                        Me.RowSimpan.Ratio_Default = .Cells(Me.ColRatioDefault.Index).Value
                        Me.RowSimpan.Ratio_References = .Cells(Me.ColRatioReference.Index).Value
                        Me.RowSimpan.Location_Code = Me.ObjLoc.GetStorageAreaCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                        Me.RowSimpan.Location_Desc = Me.ObjLoc.GetDescription(.Cells(Me.ColPrimaryLocation.Index).Value)
                        Me.RowSimpan.Warehouse_Code = Me.ObjLoc.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                        Me.RowSimpan.Warehosue_Desc = Me.ObjLoc.GetWarehouseDesc(.Cells(Me.ColPrimaryLocation.Index).Value)
                        Me.RowSimpan.Identifier_CustAddress = .Cells(Me.ColIdentifier_CustAddress.Index).Value
                        Me.RowSimpan.Tax = .Cells(Me.ColTax.Index).Value
                        Me.RowSimpan.FakturPajak = .Cells(Me.ColFakturPajak.Index).Value
                        Me.RowSimpan.identifier_address_ship_to = .Cells(Me.Colidentifier_address_ship_to.Index).Value
                        Me.RowSimpan.Primary_Customer_ship_to = .Cells(Me.ColPrimary_Customer_Ship_To.Index).Value
                        Me.RowSimpan.identifier_address_customer_ship_to = .Cells(Me.Colidentifier_address_customer_ship.Index).Value
                        Me.RowSimpan.Batch_No = .Cells(Me.ColBatchNo.Index).Value
                        Me.RowSimpan.Expiry_Date = .Cells(Me.ColExpiryDate.Index).Value
                        Me.RowSimpan.Remarks = .Cells(Me.ColRemarksItem.Index).Value
                        Me.RowSimpan.Transaction_Date = CDate(.Cells(Me.ColDate.Index).Value)
                    End With
                Next
            Else
                Me.RowSimpan = Me.TblSimpan.NewTblSimpan1Row
                With rowDGV
                    If Me.rbSO.Checked Then
                        Me.RowSimpan.Primary_Transaction = .Cells(Me.ColPrimaryTransaction.Index).Value
                        Me.RowSimpan.Transaction_No = .Cells(Me.ColDocumentNO.Index).Value
                    Else
                        Me.RowSimpan.Primary_Transaction = .Cells(Me.ColPrimarySO.Index).Value
                        Me.RowSimpan.Transaction_No = .Cells(Me.ColSONo.Index).Value
                    End If
                    Me.RowSimpan.Primary_Customer = .Cells(Me.ColPrimaryCustomer.Index).Value
                    Me.RowSimpan.Primary_Currency = .Cells(Me.ColPrimaryCurrency.Index).Value
                    Me.RowSimpan.Primary_Product = .Cells(Me.ColPrimaryPoduct.Index).Value
                    Me.RowSimpan.Primary_Unit = .Cells(Me.ColPrimaruUnitDefault.Index).Value
                    Me.RowSimpan.Primary_Multi_Unit = .Cells(Me.ColPrimaryUnitMulti.Index).Value
                    Me.RowSimpan.Primary_Location = .Cells(Me.ColPrimaryLocation.Index).Value
                    Me.RowSimpan.Primary_Warehouse = .Cells(Me.ColPrimaryWarehouse.Index).Value
                    Me.RowSimpan.Customer_Code = .Cells(Me.ColCust_Code.Index).Value
                    Me.RowSimpan.Customer_Name = .Cells(Me.ColCust_Code.Index).Value
                    Me.RowSimpan.Unit_Code_Default = .Cells(Me.ColUnitDefault.Index).Value
                    Me.RowSimpan.Unit_Code_Multi = .Cells(Me.ColUnit.Index).Value
                    Me.RowSimpan.Product_Code = .Cells(Me.ColCodeNo.Index).Value
                    Me.RowSimpan.Product_Desc = .Cells(Me.ColDesc.Index).Value
                    Me.RowSimpan.SOB_Multi = .Cells(Me.ColQty.Index).Value
                    Me.RowSimpan.Tolerance_Up = .Cells(Me.ColTolerance.Index).Value
                    Me.RowSimpan.SOB = .Cells(Me.ColQtyDefault.Index).Value
                    Me.RowSimpan.Qty_In = .Cells(Me.ColQtyDO.Index).Value * CDec(.Cells(Me.ColRatioDefault.Index).Value) / CDec(.Cells(Me.ColRatioReference.Index).Value)  'isi qty default
                    'Me.RowSimpan.Qty_In_Multi = .Cells(Me.ColQtyDO.Index).Value
                    Me.RowSimpan.Ratio_Default = .Cells(Me.ColRatioDefault.Index).Value
                    Me.RowSimpan.Ratio_References = .Cells(Me.ColRatioReference.Index).Value
                    Me.RowSimpan.Location_Code = Me.ObjLoc.GetStorageAreaCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                    Me.RowSimpan.Location_Desc = Me.ObjLoc.GetDescription(.Cells(Me.ColPrimaryLocation.Index).Value)
                    Me.RowSimpan.Warehouse_Code = Me.ObjLoc.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                    Me.RowSimpan.Warehosue_Desc = Me.ObjLoc.GetWarehouseDesc(.Cells(Me.ColPrimaryLocation.Index).Value)
                    Me.RowSimpan.Identifier_CustAddress = .Cells(Me.ColIdentifier_CustAddress.Index).Value
                    Me.RowSimpan.Tax = .Cells(Me.ColTax.Index).Value
                    Me.RowSimpan.FakturPajak = .Cells(Me.ColFakturPajak.Index).Value
                    Me.RowSimpan.identifier_address_ship_to = .Cells(Me.Colidentifier_address_ship_to.Index).Value
                    Me.RowSimpan.Primary_Customer_ship_to = .Cells(Me.ColPrimary_Customer_Ship_To.Index).Value
                    Me.RowSimpan.identifier_address_customer_ship_to = .Cells(Me.Colidentifier_address_customer_ship.Index).Value
                    Me.RowSimpan.Batch_No = .Cells(Me.ColBatchNo.Index).Value
                    Me.RowSimpan.Expiry_Date = .Cells(Me.ColExpiryDate.Index).Value
                    Me.RowSimpan.Remarks = .Cells(Me.ColRemarksItem.Index).Value
                    Me.RowSimpan.Transaction_Date = CDate(.Cells(Me.ColDate.Index).Value)
                End With
                Me.TblSimpan.AddTblSimpan1Row(Me.RowSimpan)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinBarisBatch(ByVal vPrimaryCode As Double, ByVal vPrimarySO As Double, ByVal veRow As Integer)
        Dim rowDGV As DataGridViewRow = Me.dgvItem.Rows(veRow)
        Try
            Me.TblSimpan.DefaultView.RowFilter = "[Primary Product]=" & vPrimaryCode & " AND [Primary Transaction]=" & vPrimarySO
            If Me.TblSimpan.DefaultView.Count > 0 Then
                For i As Integer = 0 To Me.TblSimpan.DefaultView.Count - 1
                    Me.RowSimpan = Me.TblSimpan.DefaultView(i).Row
                    With rowDGV
                        If Me.rbSO.Checked Then
                            Me.RowSimpan.Primary_Transaction = .Cells(Me.ColPrimaryTransaction.Index).Value
                            Me.RowSimpan.Transaction_No = .Cells(Me.ColDocumentNO.Index).Value
                        Else
                            Me.RowSimpan.Primary_Transaction = .Cells(Me.ColPrimarySO.Index).Value
                            Me.RowSimpan.Transaction_No = .Cells(Me.ColSONo.Index).Value
                        End If
                        Me.RowSimpan.Primary_Customer = .Cells(Me.ColPrimaryCustomer.Index).Value
                        Me.RowSimpan.Primary_Currency = .Cells(Me.ColPrimaryCurrency.Index).Value
                        Me.RowSimpan.Primary_Product = .Cells(Me.ColPrimaryPoduct.Index).Value
                        Me.RowSimpan.Primary_Unit = .Cells(Me.ColPrimaruUnitDefault.Index).Value
                        Me.RowSimpan.Primary_Multi_Unit = .Cells(Me.ColPrimaryUnitMulti.Index).Value
                        Me.RowSimpan.Primary_Location = .Cells(Me.ColPrimaryLocation.Index).Value
                        Me.RowSimpan.Primary_Warehouse = .Cells(Me.ColPrimaryWarehouse.Index).Value
                        Me.RowSimpan.Customer_Code = .Cells(Me.ColCust_Code.Index).Value
                        Me.RowSimpan.Customer_Name = .Cells(Me.ColCust_Code.Index).Value
                        Me.RowSimpan.Unit_Code_Default = .Cells(Me.ColUnitDefault.Index).Value
                        Me.RowSimpan.Unit_Code_Multi = .Cells(Me.ColUnit.Index).Value
                        Me.RowSimpan.Product_Code = .Cells(Me.ColCodeNo.Index).Value
                        Me.RowSimpan.Product_Desc = .Cells(Me.ColDesc.Index).Value
                        Me.RowSimpan.SOB_Multi = .Cells(Me.ColQty.Index).Value
                        Me.RowSimpan.Tolerance_Up = .Cells(Me.ColTolerance.Index).Value
                        Me.RowSimpan.SOB = .Cells(Me.ColQtyDefault.Index).Value
                        Me.RowSimpan.Qty_In = .Cells(Me.ColQtyDO.Index).Value * CDec(.Cells(Me.ColRatioDefault.Index).Value) / CDec(.Cells(Me.ColRatioReference.Index).Value)  'isi qty default
                        Me.RowSimpan.Qty_In_Multi = .Cells(Me.ColQtyDO.Index).Value
                        Me.RowSimpan.Ratio_Default = .Cells(Me.ColRatioDefault.Index).Value
                        Me.RowSimpan.Ratio_References = .Cells(Me.ColRatioReference.Index).Value
                        Me.RowSimpan.Location_Code = Me.ObjLoc.GetStorageAreaCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                        Me.RowSimpan.Location_Desc = Me.ObjLoc.GetDescription(.Cells(Me.ColPrimaryLocation.Index).Value)
                        Me.RowSimpan.Warehouse_Code = Me.ObjLoc.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                        Me.RowSimpan.Warehosue_Desc = Me.ObjLoc.GetWarehouseDesc(.Cells(Me.ColPrimaryLocation.Index).Value)
                        Me.RowSimpan.Identifier_CustAddress = .Cells(Me.ColIdentifier_CustAddress.Index).Value
                        Me.RowSimpan.Tax = .Cells(Me.ColTax.Index).Value
                        Me.RowSimpan.FakturPajak = .Cells(Me.ColFakturPajak.Index).Value
                        Me.RowSimpan.identifier_address_ship_to = .Cells(Me.Colidentifier_address_ship_to.Index).Value
                        Me.RowSimpan.Primary_Customer_ship_to = .Cells(Me.ColPrimary_Customer_Ship_To.Index).Value
                        Me.RowSimpan.identifier_address_customer_ship_to = .Cells(Me.Colidentifier_address_customer_ship.Index).Value
                        Me.RowSimpan.Batch_No = .Cells(Me.ColBatchNo.Index).Value
                        Me.RowSimpan.Expiry_Date = .Cells(Me.ColExpiryDate.Index).Value
                        Me.RowSimpan.Remarks = .Cells(Me.ColRemarksItem.Index).Value
                        Me.RowSimpan.Transaction_Date = CDate(.Cells(Me.ColDate.Index).Value)
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub IsikanTbl(Optional ByVal vPakeBatch As Boolean = True)
        Dim rowDGV As DataGridViewRow
        Dim ObjSuplier As New Core_Module.clsMasterSupplier
        Try
            For i As Integer = 0 To Me.dgvItem.Rows.Count - 1
                rowDGV = Me.dgvItem.Rows(i)
                With rowDGV
                    If vPakeBatch = True Then
                        If .Cells(Me.ColBatchControlled.Index).Value = 1 Then
                            Me.RowSimpan = Me.TblSimpan.NewTblSimpan1Row
                            If Me.rbSO.Checked Then
                                Me.RowSimpan.Primary_Transaction = .Cells(Me.ColPrimaryTransaction.Index).Value
                                Me.RowSimpan.Transaction_No = .Cells(Me.ColDocumentNO.Index).Value
                            Else
                                Me.RowSimpan.Primary_Transaction = .Cells(Me.ColPrimarySO.Index).Value
                                Me.RowSimpan.Transaction_No = .Cells(Me.ColSONo.Index).Value
                            End If
                            Me.RowSimpan.Result_No = .Cells(Me.ColDocumentNO.Index).Value '2026-11-25 IM, untuk Result nanti
                            Me.RowSimpan.Primary_Customer = .Cells(Me.ColPrimaryCustomer.Index).Value
                            Me.RowSimpan.Primary_Currency = .Cells(Me.ColPrimaryCurrency.Index).Value
                            Me.RowSimpan.Primary_Product = .Cells(Me.ColPrimaryPoduct.Index).Value
                            Me.RowSimpan.Primary_Unit = .Cells(Me.ColPrimaruUnitDefault.Index).Value
                            Me.RowSimpan.Primary_Multi_Unit = .Cells(Me.ColPrimaryUnitMulti.Index).Value
                            Me.RowSimpan.Primary_Location = .Cells(Me.ColPrimaryLocation.Index).Value
                            Me.RowSimpan.Primary_Warehouse = .Cells(Me.ColPrimaryWarehouse.Index).Value
                            Me.RowSimpan.Remarks = .Cells(Me.ColRemarksItem.Index).Value
                            Me.RowSimpan.Customer_Code = .Cells(Me.ColCust_Code.Index).Value
                            Me.RowSimpan.Customer_Name = ObjSuplier.GetSupplierName(Me.RowSimpan.Primary_Customer)
                            Me.RowSimpan.Unit_Code_Default = .Cells(Me.ColUnitDefault.Index).Value
                            Me.RowSimpan.Unit_Code_Multi = .Cells(Me.ColUnit.Index).Value
                            Me.RowSimpan.Product_Code = .Cells(Me.ColCodeNo.Index).Value
                            Me.RowSimpan.Product_Desc = .Cells(Me.ColDesc.Index).Value
                            Me.RowSimpan.SOB_Multi = .Cells(Me.ColQty.Index).Value
                            Me.RowSimpan.Tolerance_Up = .Cells(Me.ColTolerance.Index).Value
                            Me.RowSimpan.SOB = .Cells(Me.ColQtyDefault.Index).Value
                            Me.RowSimpan.Qty_In = .Cells(Me.ColQtyDO.Index).Value
                            Me.RowSimpan.Ratio_Default = .Cells(Me.ColRatioDefault.Index).Value
                            Me.RowSimpan.Ratio_References = .Cells(Me.ColRatioReference.Index).Value
                            Me.RowSimpan.Location_Code = Me.ObjLoc.GetStorageAreaCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                            Me.RowSimpan.Location_Desc = Me.ObjLoc.GetDescription(.Cells(Me.ColPrimaryLocation.Index).Value)
                            Me.RowSimpan.Warehouse_Code = Me.ObjLoc.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                            Me.RowSimpan.Warehosue_Desc = Me.ObjLoc.GetWarehouseDesc(.Cells(Me.ColPrimaryLocation.Index).Value)
                            Me.RowSimpan.Identifier_CustAddress = .Cells(Me.ColIdentifier_CustAddress.Index).Value
                            Me.RowSimpan.Tax = .Cells(Me.ColTax.Index).Value
                            Me.RowSimpan.FakturPajak = .Cells(Me.ColFakturPajak.Index).Value
                            Me.RowSimpan.identifier_address_ship_to = .Cells(Me.Colidentifier_address_ship_to.Index).Value
                            Me.RowSimpan.Primary_Customer_ship_to = .Cells(Me.ColPrimary_Customer_Ship_To.Index).Value
                            Me.RowSimpan.identifier_address_customer_ship_to = .Cells(Me.Colidentifier_address_customer_ship.Index).Value
                            Me.RowSimpan.Batch_No = .Cells(Me.ColBatchNo.Index).Value
                            Me.RowSimpan.Expiry_Date = .Cells(Me.ColExpiryDate.Index).Value
                            Me.RowSimpan.Transaction_Date = CDate(.Cells(Me.ColDate.Index).Value)
                            Me.TblSimpan.AddTblSimpan1Row(Me.RowSimpan)
                        End If
                    Else
                        If CDec(.Cells(Me.ColQtyDO.Index).Value) > 0 Then
                            If .Cells(Me.ColBatchControlled.Index).Value = 0 Then
                                Me.RowSimpan = Me.TblSimpan.NewTblSimpan1Row
                                If Me.rbSO.Checked Then
                                    Me.RowSimpan.Primary_Transaction = .Cells(Me.ColPrimaryTransaction.Index).Value
                                    Me.RowSimpan.Transaction_No = .Cells(Me.ColDocumentNO.Index).Value
                                Else
                                    Me.RowSimpan.Primary_Transaction = .Cells(Me.ColPrimarySO.Index).Value
                                    Me.RowSimpan.Transaction_No = .Cells(Me.ColSONo.Index).Value
                                End If
                                Me.RowSimpan.Result_No = .Cells(Me.ColDocumentNO.Index).Value '2026-11-25 IM, untuk Result nanti
                                Me.RowSimpan.Primary_Customer = .Cells(Me.ColPrimaryCustomer.Index).Value
                                Me.RowSimpan.Primary_Currency = .Cells(Me.ColPrimaryCurrency.Index).Value
                                Me.RowSimpan.Primary_Product = .Cells(Me.ColPrimaryPoduct.Index).Value
                                Me.RowSimpan.Primary_Unit = .Cells(Me.ColPrimaruUnitDefault.Index).Value
                                Me.RowSimpan.Primary_Multi_Unit = .Cells(Me.ColPrimaryUnitMulti.Index).Value
                                Me.RowSimpan.Primary_Location = .Cells(Me.ColPrimaryLocation.Index).Value
                                Me.RowSimpan.Primary_Warehouse = .Cells(Me.ColPrimaryWarehouse.Index).Value
                                Me.RowSimpan.Remarks = .Cells(Me.ColRemarksItem.Index).Value
                                Me.RowSimpan.Customer_Code = .Cells(Me.ColCust_Code.Index).Value
                                Me.RowSimpan.Customer_Name = ObjSuplier.GetSupplierName(Me.RowSimpan.Primary_Customer)
                                Me.RowSimpan.Unit_Code_Default = .Cells(Me.ColUnitDefault.Index).Value
                                Me.RowSimpan.Unit_Code_Multi = .Cells(Me.ColUnit.Index).Value
                                Me.RowSimpan.Product_Code = .Cells(Me.ColCodeNo.Index).Value
                                Me.RowSimpan.Product_Desc = .Cells(Me.ColDesc.Index).Value
                                Me.RowSimpan.SOB_Multi = .Cells(Me.ColQty.Index).Value
                                Me.RowSimpan.Tolerance_Up = .Cells(Me.ColTolerance.Index).Value
                                Me.RowSimpan.SOB = .Cells(Me.ColQtyDefault.Index).Value
                                Me.RowSimpan.Qty_In = (CDec(.Cells(Me.ColQtyDO.Index).Value) * CDec(.Cells(Me.ColRatioDefault.Index).Value)) / CDec(.Cells(Me.ColRatioReference.Index).Value)
                                Me.RowSimpan.Qty_In_Multi = .Cells(Me.ColQtyDO.Index).Value
                                Me.RowSimpan.Ratio_Default = .Cells(Me.ColRatioDefault.Index).Value
                                Me.RowSimpan.Ratio_References = .Cells(Me.ColRatioReference.Index).Value
                                Me.RowSimpan.Location_Code = Me.ObjLoc.GetStorageAreaCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                                Me.RowSimpan.Location_Desc = Me.ObjLoc.GetDescription(.Cells(Me.ColPrimaryLocation.Index).Value)
                                Me.RowSimpan.Warehouse_Code = Me.ObjLoc.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                                Me.RowSimpan.Warehosue_Desc = Me.ObjLoc.GetWarehouseDesc(.Cells(Me.ColPrimaryLocation.Index).Value)
                                Me.RowSimpan.Identifier_CustAddress = .Cells(Me.ColIdentifier_CustAddress.Index).Value
                                Me.RowSimpan.Tax = .Cells(Me.ColTax.Index).Value
                                Me.RowSimpan.FakturPajak = .Cells(Me.ColFakturPajak.Index).Value
                                Me.RowSimpan.identifier_address_ship_to = .Cells(Me.Colidentifier_address_ship_to.Index).Value
                                Me.RowSimpan.Primary_Customer_ship_to = .Cells(Me.ColPrimary_Customer_Ship_To.Index).Value
                                Me.RowSimpan.identifier_address_customer_ship_to = .Cells(Me.Colidentifier_address_customer_ship.Index).Value
                                Me.RowSimpan.Batch_No = .Cells(Me.ColBatchNo.Index).Value
                                Me.RowSimpan.Expiry_Date = .Cells(Me.ColExpiryDate.Index).Value
                                Me.RowSimpan.Transaction_Date = CDate(.Cells(Me.ColDate.Index).Value)
                                Me.TblSimpan.AddTblSimpan1Row(Me.RowSimpan)
                            End If
                        End If
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanStock(ByVal vResetData As Boolean)
        Dim ListPrimaryProduct As New List(Of Double)
        Dim Tbl As New DataTable
        Dim TblDO As New DataTable
        Try
            'get list primary
            If Me.dgvItem.Rows.Count > 0 Then
                For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                    If ListPrimaryProduct.Contains(rowdgv.Cells(Me.ColPrimaryPoduct.Index).Value) = False Then
                        ListPrimaryProduct.Add(rowdgv.Cells(Me.ColPrimaryPoduct.Index).Value)
                    End If
                Next
            Else
                Exit Sub
            End If

            'If ListPrimaryProduct.Count > 0 Then
            '    For Each primaryProduct As Double In ListPrimaryProduct
            '        If String.IsNullOrEmpty(inPrimaryProduct) Then
            '            inPrimaryProduct = primaryProduct
            '        Else
            '            inPrimaryProduct = inPrimaryProduct & ", " & primaryProduct
            '        End If
            '    Next
            'Else
            '    Exit Sub
            'End If

            'DTAStock.Connection.ConnectionString = Core_Module.GetConstring
            'DTAStock.ClearBeforeFill = vResetData
            'DTAStock.FillBy_ListPrimaryProduct(TBLStock, inPrimaryProduct)
            If vResetData Then
                Me.TBLStock.Clear()
                Me.TblPendingDO.Clear()
                Me.TblStockTemp.Clear()
            End If
            Tbl = Me.GetDataStockByPrimary(ListPrimaryProduct)
            For Each row As DataRow In Tbl.Rows
                Me.RowStock = Me.TBLStock.Newstock_detail_viewRow
                Me.RowStock.ItemArray = row.ItemArray
                Me.TBLStock.Rows.Add(Me.RowStock)
            Next
            'Me.TBLStock = Me.GetDataStockByPrimary(ListPrimaryProduct)

            TblDO = Me.GetDataPendingDOByPrimary(ListPrimaryProduct)
            For Each row As DataRow In TblDO.Rows
                Me.RowPendingDO = Me.TblPendingDO.NewPending_DO_ViewRow
                Me.RowPendingDO.ItemArray = row.ItemArray
                Me.TblPendingDO.Rows.Add(Me.RowPendingDO)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Tbl.Dispose()
            TblDO.Dispose()
        End Try
    End Sub

    Private Function GetDataStockByPrimary(ByVal ListPrimary As List(Of Double)) As DataTable
        Dim Tbl As New DataTable
        If ListPrimary Is Nothing OrElse ListPrimary.Count = 0 Then
            Return Tbl
        End If
        Dim ConnectionString As String = Core_Module.GetConstring
        Dim Conn As New MySqlConnection(ConnectionString)
        Try
            Conn.Open()
            Dim ListParameter As New List(Of String)
            For i As Integer = 0 To ListPrimary.Count - 1
                ListParameter.Add("@P" & i.ToString())
            Next
            Dim SQL As String = "SELECT `Primary Product`, Qty, `Lot No`, " & _
            "`Expiry Date`, " & _
            "IFNULL(Location, 'No Location') AS Location, " & _
            "IFNULL(`Warehouse Code`, '') AS Warehouse, " & _
            "`Decimal`, `Primary Storage Area`, " & _
            "`Primary Unit`, `Primary Warehouse` " & _
            "FROM stock_detail_view " & _
            "WHERE `Primary Product` IN (" & _
            String.Join(",", ListParameter.ToArray()) & ")"

            Dim Cmd As New MySqlCommand(SQL, Conn)

            For i As Integer = 0 To ListPrimary.Count - 1
                Cmd.Parameters.AddWithValue( _
                "@P" & i.ToString(), _
                ListPrimary(i))
            Next

            Dim DA As New MySqlDataAdapter(Cmd)
            DA.Fill(Tbl)
        Catch ex As Exception
            Throw New Exception( _
                "Gagal mengambil data Product." & vbCrLf & _
                ex.Message, ex)
        Finally
            If Conn IsNot Nothing Then
                If Conn.State <> ConnectionState.Closed Then
                    Conn.Close()
                End If
                Conn.Dispose()
            End If
        End Try

        Return Tbl
    End Function

    Private Function GetDataPendingDOByPrimary(ByVal ListPrimary As List(Of Double)) As DataTable
        Dim Tbl As New DataTable
        If ListPrimary Is Nothing OrElse ListPrimary.Count = 0 Then
            Return Tbl
        End If
        Dim ConnectionString As String = Core_Module.GetConstring
        Dim Conn As New MySqlConnection(ConnectionString)
        Try
            Conn.Open()
            Dim ListParameter As New List(Of String)
            For i As Integer = 0 To ListPrimary.Count - 1
                ListParameter.Add("@P" & i.ToString())
            Next
            Dim SQL As String = "SELECT * FROM (" & _
            "SELECT " & _
            "doi.`Primary Product` AS `Primary Product`, " & _
            "so.`Primary` AS `Primary SO`, " & _
            "doi.`Primary Location` AS `Primary Storage Area`, " & _
            "sa.`Primary Warehouse` AS `Primary Warehouse`, " & _
            "doi.`Primary Units` AS `Primary Unit`, " & _
            "do.`No Transaksi` AS `DO Number`, " & _
            "p.`Product Code` AS `Product Code`, " & _
            "p.`Description` AS `Description`, " & _
            "so.`No Transaksi` AS `Sales Order No`, " & _
            "doi.`Batch No` AS `Batch No`, " & _
            "doi.`Expiry Date` AS `Expiry Date`, " & _
            "doi.`Total Qty` AS `Total Qty`, " & _
            "mu.`Unit Code` AS `Unit Code`, " & _
            "doi.`Secondary Qty` AS `Secondary Qty`, " & _
            "'' AS `Secondary Unit`, " & _
            "doi.`Remarks` AS `Remarks`, " & _
            "p.`Decimal` AS `Decimal` " & _
            "FROM `tbltrans_do` do " & _
            "INNER JOIN `tbltrans_do_item` doi " & _
            "ON do.`Primary` = doi.`Primary DO` " & _
            "INNER JOIN `tblmaster_product` p " & _
            "ON doi.`Primary Product` = p.`Primary` " & _
            "INNER JOIN `tbltrans_so` so " & _
            "ON doi.`Primary SO` = so.`Primary` " & _
            "INNER JOIN `tblmaster_storage_area` sa " & _
            "ON doi.`Primary Location` = sa.`Primary` " & _
            "INNER JOIN `tblmaster_master_unit` mu " & _
            "ON doi.`Primary Units` = mu.`Primary` " & _
            "WHERE do.`Primary Stock Out Sales` = -1" & _
            ") AS d " & _
            "WHERE `Primary Product` IN (" & _
            String.Join(",", ListParameter.ToArray()) & ")"

            Dim Cmd As New MySqlCommand(SQL, Conn)

            For i As Integer = 0 To ListPrimary.Count - 1
                Cmd.Parameters.AddWithValue( _
                "@P" & i.ToString(), _
                ListPrimary(i))
            Next

            Dim DA As New MySqlDataAdapter(Cmd)
            DA.Fill(Tbl)
        Catch ex As Exception
            Throw New Exception( _
                "Gagal mengambil data Product." & vbCrLf & _
                ex.Message, ex)
        Finally
            If Conn IsNot Nothing Then
                If Conn.State <> ConnectionState.Closed Then
                    Conn.Close()
                End If
                Conn.Dispose()
            End If
        End Try

        Return Tbl
    End Function


    Private Sub btnStart_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Dim AdaYgDiBuat As Boolean = False
        Try
            If isLocationChanged = True Then 'cek apakah ada perubahan location
                MsgBox("Location changed, cannot start.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColChkDO.Index).Value = Boolean.TrueString Then
                    AdaYgDiBuat = True
                    Exit For
                End If
            Next

            If AdaYgDiBuat = True Then
                'Me.StartImport()
                Me.BackgroundWorker2.RunWorkerAsync()
                'Me.btnFilter.PerformClick()
            Else
                MsgBox("Please specify at least 1 item to create.", MsgBoxStyle.Information, "Information")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsikanCustomerSONo() As Boolean
        Dim ij As New FrmTransaksiDO_Quick_InputSoNo
        Dim Dgv As DataGridView = ij.DgvItem
        Dim rowD As DataGridViewRow
        Dim ListOfPrimary As New List(Of Double)
        Try
            If Me.TblSimpan.Rows.Count > 0 Then
                For Each Me.RowSimpan In Me.TblSimpan
                    If ListOfPrimary.Contains(Me.RowSimpan.Primary_Customer) = False Then
                        ListOfPrimary.Add(Me.RowSimpan.Primary_Customer)
                        rowD = Dgv.Rows(Dgv.Rows.Add)
                        With rowD
                            .Cells(ij.ColNo.Index).Value = ""
                            .Cells(ij.ColPrimaryCustomer.Index).Value = Me.RowSimpan.Primary_Customer
                            .Cells(ij.ColCustomerCode.Index).Value = Me.RowSimpan.Customer_Code
                            .Cells(ij.ColCustomerName.Index).Value = Me.RowSimpan.Customer_Name
                            .Cells(ij.ColCustomerSoNo.Index).Value = ""
                        End With
                    End If
                Next
                Me.ObjCD.SetNumber(ij.DgvItem)
                ij.BackColorAll = Me.BackColorAll

                If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                    For Each Me.RowSimpan In Me.TblSimpan.Rows
                        For Each rowD In Dgv.Rows
                            If Me.RowSimpan.Primary_Customer = rowD.Cells(ij.ColPrimaryCustomer.Index).Value Then
                                Me.RowSimpan.Customer_SO_No = rowD.Cells(ij.ColCustomerSoNo.Index).Value
                            End If
                        Next
                    Next

                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function StartImport() As Boolean
        Dim rowDGV As DataGridViewRow
        Dim ij As New Sales_Module.FrmTransaksiDO3
        Dim ObjLoc As New Core_Module.clsMasterStorageArea
        Dim ObjWarehouse As New Core_Module.clsMasterWarehouse
        Dim objDO As New ClsTransaksiDO
        Dim ListOfCustomer As New List(Of Double)
        Dim ListOfCustomerGroup As New List(Of String)
        Dim ObjSup As New Core_Module.clsMasterSupplier
        Dim PrimaryCust As Double = -1
        Dim tax As Decimal = 0
        Dim noFaktur As String = ""
        Dim PrimaryCurr As Double = -1
        Dim identifierShipTo As String = ""
        Dim PrimaryCustomerShipTo As Double = -1
        Dim identifierCustomerShipTo As String = ""
        Dim IdentifierAddress As String = ""
        Dim GroupCustomer As String = ""
        Dim MinDate As Date
        Dim TransactionNo As String = ""
        Dim TransactionNoArray() As String
        Dim tblResult As New DseTransaksiQuick_DO2.tblResultDataTable
        Dim rowResult As DseTransaksiQuick_DO2.tblResultRow
        Try

            Me.IsikanTbl(False)
            'If Me.rbInventoryBooking.Checked Then
            '    Me.isikanQtyBatch()
            'End If
            ListOfCustomer.Clear()
            ListOfCustomerGroup.Clear()
            For Each rowDGV1 As DataGridViewRow In Me.dgvItem.Rows
                If rowDGV1.Cells(Me.ColChkDO.Index).Value = Boolean.TrueString Then
                    If ListOfCustomer.Contains(rowDGV1.Cells(Me.ColPrimaryCustomer.Index).Value) = False Then
                        ListOfCustomer.Add(rowDGV1.Cells(Me.ColPrimaryCustomer.Index).Value)
                    End If
                    PrimaryCust = rowDGV1.Cells(Me.ColPrimaryCustomer.Index).Value
                    tax = rowDGV1.Cells(Me.ColTax.Index).Value
                    noFaktur = rowDGV1.Cells(Me.ColFakturPajak.Index).Value
                    PrimaryCurr = rowDGV1.Cells(Me.ColPrimaryCurrency.Index).Value
                    identifierShipTo = rowDGV1.Cells(Me.Colidentifier_address_ship_to.Index).Value
                    PrimaryCustomerShipTo = rowDGV1.Cells(Me.ColPrimary_Customer_Ship_To.Index).Value
                    identifierCustomerShipTo = rowDGV1.Cells(Me.Colidentifier_address_customer_ship.Index).Value
                    IdentifierAddress = rowDGV1.Cells(Me.ColIdentifier_CustAddress.Index).Value
                    GroupCustomer = PrimaryCust & "|" & tax & "|" & noFaktur & "|" & PrimaryCurr & "|" & identifierShipTo & "|" & PrimaryCustomerShipTo & "|" & _
                    identifierCustomerShipTo & "|" & IdentifierAddress
                    If ListOfCustomerGroup.Contains(GroupCustomer) = False Then
                        ListOfCustomerGroup.Add(GroupCustomer)
                    End If
                End If

                'If rowDGV1.Cells(Me.ColchkDO.Index).Value = Boolean.FalseString And ListOfSupplier.Contains(rowDGV1.Cells(Me.ColPrimaryCustomer.Index).Value) = False Then
                '    Me.TblSimpan.DefaultView.RowFilter = "[Primary Product]=" & rowDGV1.Cells(Me.ColPrimaryPoduct.Index).Value
                '    For i As Integer = Me.TblSimpan.DefaultView.Count - 1 To 0 Step -1
                '        Me.RowSimpan = Me.TblSimpan.DefaultView(i).Row
                '        Me.TblSimpan.Rows.Remove(Me.RowSimpan)
                '    Next
                'End If
            Next

            For Each rowDGV1 As DataGridViewRow In Me.dgvItem.Rows
                PrimaryCust = rowDGV1.Cells(Me.ColPrimaryCustomer.Index).Value
                tax = rowDGV1.Cells(Me.ColTax.Index).Value
                noFaktur = rowDGV1.Cells(Me.ColFakturPajak.Index).Value
                PrimaryCurr = rowDGV1.Cells(Me.ColPrimaryCurrency.Index).Value
                identifierShipTo = rowDGV1.Cells(Me.Colidentifier_address_ship_to.Index).Value
                PrimaryCustomerShipTo = rowDGV1.Cells(Me.ColPrimary_Customer_Ship_To.Index).Value
                identifierCustomerShipTo = rowDGV1.Cells(Me.Colidentifier_address_customer_ship.Index).Value
                IdentifierAddress = rowDGV1.Cells(Me.ColIdentifier_CustAddress.Index).Value
                GroupCustomer = PrimaryCust & "|" & tax & "|" & noFaktur & "|" & PrimaryCurr & "|" & identifierShipTo & "|" & PrimaryCustomerShipTo & "|" & _
                identifierCustomerShipTo & "|" & IdentifierAddress
                Dim primarySO As Double = -1
                If Me.rbSO.Checked Then
                    primarySO = rowDGV1.Cells(Me.ColPrimaryTransaction.Index).Value
                Else
                    primarySO = rowDGV1.Cells(Me.ColPrimarySO.Index).Value
                End If
                If rowDGV1.Cells(Me.ColChkDO.Index).Value = Boolean.FalseString And ListOfCustomerGroup.Contains(GroupCustomer) = False Then
                    Me.TblSimpan.DefaultView.RowFilter = "[Primary Product]=" & rowDGV1.Cells(Me.ColPrimaryPoduct.Index).Value & " AND [Primary Customer]=" & rowDGV1.Cells(Me.ColPrimaryCustomer.Index).Value & _
                    " AND [Primary Transaction]= " & primarySO
                    For i As Integer = Me.TblSimpan.DefaultView.Count - 1 To 0 Step -1
                        Me.RowSimpan = Me.TblSimpan.DefaultView(i).Row
                        Me.TblSimpan.Rows.Remove(Me.RowSimpan)
                    Next
                ElseIf rowDGV1.Cells(Me.ColQtyDO.Index).Value = 0 Then 'hapus juga yg qty 0
                    Me.TblSimpan.DefaultView.RowFilter = "[Primary Product]=" & rowDGV1.Cells(Me.ColPrimaryPoduct.Index).Value & _
                    " AND [Primary Customer]=" & rowDGV1.Cells(Me.ColPrimaryCustomer.Index).Value & _
                    " AND [Primary Transaction]= " & primarySO & " AND [Qty In]= 0 And [Qty In Multi]= 0"
                    For i As Integer = Me.TblSimpan.DefaultView.Count - 1 To 0 Step -1
                        Me.RowSimpan = Me.TblSimpan.DefaultView(i).Row
                        Me.TblSimpan.Rows.Remove(Me.RowSimpan)
                    Next
                End If
            Next

            'If Me.IsikanCustomerSONo() = False Then
            '    Return False
            '    Exit Function
            'End If

            Me.TblSimpan.DefaultView.RowFilter = ""
            If Me.TblSimpan.DefaultView.Count > 0 Then

                Me.prg01.Visible = True
                Me.prg01.Value = 0
                Me.prg01.Maximum = Me.TblSimpan.Rows.Count
                ij.prgProgress.Value = 0
                ij.prgProgress.Maximum = Me.prg01.Maximum
                Me.LblProgress.Text = ""
                Me.LblProgress.Visible = True

                ij.vFromQuick = True
                ij.vDirectlySave = False
                ij.UserCode = Me.UserCode
                ij.BackColorAll = Me.BackColorAll
                'ij.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                ij.Show()
                ij.Location = New Point((ij.Width * -1) - 5, (ij.Height * -1) - 5)

                For Each GroupCustomer In ListOfCustomerGroup
                    '=========================================
                    ' Ambil kembali isi GroupCustomer
                    '=========================================
                    Dim GroupData() As String = GroupCustomer.Split("|"c)
                    TransactionNo = ""
                    PrimaryCust = CDbl(GroupData(0))
                    tax = CDbl(GroupData(1))
                    noFaktur = GroupData(2)
                    PrimaryCurr = CDbl(GroupData(3))
                    identifierShipTo = GroupData(4)
                    PrimaryCustomerShipTo = CDbl(GroupData(5))
                    identifierCustomerShipTo = GroupData(6)
                    IdentifierAddress = GroupData(7)


                    MinDate = #12/12/2500#
                    ij.SetBaru()
                    ij.TxtDONo.Text = objDO.GetNoBaru1(ij.dtpDate.Value)
                    ij.vShowDialog = False


                    ij.txtPrimaryCustomer.Text = PrimaryCust
                    ij.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(PrimaryCust)
                    ij.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(PrimaryCust)

                    'isi bill to address, ship to address
                    ij.DefaultFillCustomer(False)
                    Me.ObjCustomer.IsiCboAddress(ij.cboCustomerAddress, ij.txtPrimaryCustomer.Text)
                    ij.txtCustomerAddIdentifier.Text = IdentifierAddress
                    ij.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(ij.txtCustomerAddIdentifier.Text)
                    If PrimaryCustomerShipTo = -1 Then
                        ij.Panel1.Visible = True
                        ij.TxtIdentifierAddressShipTo.Text = identifierShipTo
                        Me.ObjCustomer.IsiCboAddress(ij.CboAddressShipTo, ij.txtPrimaryCustomer.Text)
                        If ij.CboAddressShipTo.Items.Count > 0 Then
                            ij.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(ij.TxtIdentifierAddressShipTo.Text)
                        End If
                        ij.Panel2.Visible = False
                        ij.TxtPrimaryCustomerShipTo.Text = "-1"
                        ij.TxtCustomerCodeShipTo.Text = ""
                        ij.TxtCustomerDescShipTo.Text = ""
                        ij.TxtIdentifierAddressCustomerShipTo.Text = ""
                        ij.CboAddressCustomerShipTo.Items.Clear()
                        ij.Label6.Location = New Point(ij.Label6.Location.X, ij.Label1.Location.Y)
                        ij.TxtRemark.Location = New Point(ij.TxtRemark.Location.X, ij.TxtCustomerName.Location.Y + 5)
                        ij.TxtRemark.Size = New Size(ij.TxtRemark.Size.Width, (ij.TxtCustomerName.Size.Height * 2) + (ij.cboCustomerAddress.Top - ij.TxtCustomerName.Bottom))
                    Else
                        ij.Panel2.Visible = True
                        ij.TxtPrimaryCustomerShipTo.Text = PrimaryCustomerShipTo
                        ij.TxtCustomerCodeShipTo.Text = Me.ObjCustomer.GetCustomerCode(PrimaryCustomerShipTo)
                        ij.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(PrimaryCustomerShipTo)
                        ij.TxtIdentifierAddressCustomerShipTo.Text = identifierCustomerShipTo
                        Me.ObjCustomer.IsiCboAddress(ij.CboAddressCustomerShipTo, PrimaryCustomerShipTo)
                        If ij.CboAddressCustomerShipTo.Items.Count > 0 Then
                            ij.CboAddressCustomerShipTo.Text = Me.ObjCustomer.GetKodeAddress(ij.TxtIdentifierAddressCustomerShipTo.Text)
                        End If
                        ij.Panel1.Visible = False
                        ij.TxtIdentifierAddressShipTo.Text = ""
                        ij.CboAddressShipTo.Items.Clear()
                        ij.Label6.Location = New Point(ij.Label6.Location.X, ij.cboCustomerAddress.Location.Y)
                        ij.TxtRemark.Size = New Size(ij.TxtRemark.Size.Width, ij.TxtCustomerName.Size.Height)
                        ij.TxtRemark.Location = New Point(ij.TxtRemark.Location.X, ij.cboCustomerAddress.Location.Y)
                    End If


                    ij.TxtRemark.Text = Me.TxtCancellationReason.Text
                    ij.TxtRemark.Focus()


                    Me.TblSimpan.DefaultView.RowFilter = "[Primary Customer]=" & PrimaryCust & " AND [tax]=" & tax & " AND [FakturPajak]= '" & noFaktur & _
                    "' AND [identifier_address_ship_to]= '" & identifierShipTo & "' AND [Primary Customer ship to]= '" & PrimaryCustomerShipTo & _
                    "' AND [identifier_address_customer_ship_to]= '" & identifierCustomerShipTo & "'"
                    For i As Integer = 0 To Me.TblSimpan.DefaultView.Count - 1
                        Me.RowSimpan = Me.TblSimpan.DefaultView(i).Row

                        Me.prg01.Value += 1
                        Me.LblProgress.Text = "Data " + Me.prg01.Value.ToString + " of " + Me.prg01.Maximum.ToString + " - DO Number: " + ij.TxtDONo.Text
                        ij.prgProgress.Value = Me.prg01.Value

                        With RowSimpan
                            'ij.TxtSuppDONumbe.Text = .Supplier_DO_No
                            'tambahkan baris baru
                            Dim ExpiredDateStr As String
                            If Me.RowSimpan.Expiry_Date = #1/1/1950# Then
                                ExpiredDateStr = "-"
                            Else
                                ExpiredDateStr = Format(Me.RowSimpan.Expiry_Date, "dd MMM yyyy")
                            End If

                            If RowSimpan.Transaction_Date <= MinDate Then
                                MinDate = RowSimpan.Transaction_Date.Date
                            End If

                            ij.dgvItem.Rows.Add()
                            rowDGV = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)
                            With rowDGV
                                .Cells(ij.ColNo.Index).Value = ""
                                .Cells(ij.ColSONO.Index).Value = Me.RowSimpan.Transaction_No
                                'If Me.chkDirectSave.Checked = True Then
                                '    .Cells(ij.ColSONO.Index).Value = Me.RowSimpan.Transaction_No
                                'Else
                                '    .Cells(ij.ColSONO.Index).Value = ""
                                'End If

                                .Cells(ij.ColCodeNo.Index).Value = Me.RowSimpan.Product_Code
                                .Cells(ij.ColDescription.Index).Value = Me.RowSimpan.Product_Desc
                                .Cells(ij.ColQty.Index).Value = Me.ObjCD.DigitGrouping(Me.RowSimpan.Qty_In)
                                .Cells(ij.ColTotalQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(Me.RowSimpan.Qty_In_Multi)
                                .Cells(ij.ColUnitsCodeMulti.Index).Value = Me.RowSimpan.Unit_Code_Multi
                                If Me.RowSimpan.Primary_Unit <> Me.RowSimpan.Primary_Multi_Unit Then
                                    .Cells(ij.ColRatioDesc.Index).Value = ij.ObjCD.DigitGrouping(Me.RowSimpan.Ratio_References, _
                                    ObjProduct.GetAngkaDecimal(Me.RowSimpan.Primary_Product, Me.RowSimpan.Primary_Unit)) _
                                    & " " & Me.RowSimpan.Unit_Code_Multi & " = " & ij.ObjCD.DigitGrouping(Me.RowSimpan.Ratio_Default, _
                                    Me.ObjProduct.GetAngkaDecimal(Me.RowSimpan.Primary_Product)) & " " & Me.RowSimpan.Unit_Code_Default
                                Else
                                    .Cells(ij.ColRatioDesc.Index).Value = "-"
                                End If
                                .Cells(ij.ColRatioDefault.Index).Value = Me.RowSimpan.Ratio_Default
                                .Cells(ij.ColRatioReference.Index).Value = Me.RowSimpan.Ratio_References

                                .Cells(ij.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(ij.ColQty.Index).Value) * CDec(.Cells(ij.ColRatioDefault.Index).Value) / CDec(.Cells(ij.ColRatioReference.Index).Value), Me.RowSimpan.Decimal_Default)
                                '.Cells(ij.ColQty.Index).Value = Me.RowSimpan.Qty_In
                                .Cells(ij.ColUnit.Index).Value = Me.RowSimpan.Unit_Code_Default
                                .Cells(ij.ColSecQty.Index).Value = -1
                                .Cells(ij.ColSecUnit.Index).Value = ""
                                .Cells(ij.ColBatchNo.Index).Value = Me.RowSimpan.Batch_No
                                .Cells(ij.ColExpiryDate.Index).Value = ExpiredDateStr
                                .Cells(ij.ColLocation.Index).Value = Me.RowSimpan.Location_Code
                                .Cells(ij.ColWarehouse.Index).Value = Me.RowSimpan.Warehouse_Code
                                .Cells(ij.ColRemarks.Index).Value = Me.RowSimpan.Remarks
                                .Cells(ij.ColPrimarySO.Index).Value = Me.RowSimpan.Primary_Transaction
                                .Cells(ij.ColPrimaryCodeNo.Index).Value = Me.RowSimpan.Primary_Product
                                .Cells(ij.ColPrimaryLocation.Index).Value = Me.RowSimpan.Primary_Location
                                .Cells(ij.ColPrimaryUnit.Index).Value = Me.RowSimpan.Primary_Unit
                                .Cells(ij.ColPrimarySecUnit.Index).Value = -1
                                .Cells(ij.ColQtyMaster.Index).Value = 0
                                .Cells(ij.ColQtyPerMaster.Index).Value = 0
                                .Cells(ij.ColPrimaryMasterUnits.Index).Value = -1
                                '.Cells(ij.ColPrimaryPd.Index).Value = -1
                                .Cells(ij.ColPrimaryUnitsMulti.Index).Value = Me.RowSimpan.Primary_Multi_Unit
                                .Cells(ij.ColRemarks.Index).Value = Me.RowSimpan.Remarks
                            End With
                            If TransactionNo.Contains(.Result_No) = False Then
                                If String.IsNullOrEmpty(TransactionNo) Then
                                    TransactionNo = .Result_No
                                Else
                                    TransactionNo = TransactionNo & ";" & .Result_No
                                End If
                            End If
                        End With
                    Next

                    Me.ObjCD.SetNumber(Me.dgvItem)
                    ij.vMinDate = MinDate
                    ij.FormatDGV()
                    ij.Hide()
                    ij.vDirectlySave = Me.chkDirectSave.Checked
                    ij.vShowDialog = True
                    ij.ShowDialog()
                    If ij.vPrimaryDO <> -1 And String.IsNullOrEmpty(TransactionNo) = False Then
                        'TransactionNoArray = TransactionNo.Split(";"c)
                        'For Each No As String In TransactionNoArray

                        'Next
                        rowResult = tblResult.NewtblResultRow
                        With rowResult
                            .Primary_DO = ij.vPrimaryDO
                            .No_DO = ij.vDONo
                            .No_Transaksi = TransactionNo
                        End With
                    End If
                    
                    If ij.vStopCancellation = True Then
                        Exit For
                    End If
                Next
            End If
            'Me.btnFilter.PerformClick()
            Dim jk As New FrmTransaksiDO_Quick_Result

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If Me.StartAntrian() Then
                MsgBox("Import finished.", MsgBoxStyle.Information, "Information")
            End If
            Me.prg01.Visible = False
            Me.LblProgress.Visible = False
            Me.btnFilter.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function StartAntrian() As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, -1, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiDO_SimpanBaru, Me.UserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            Me.SudahHapusAntrian = False

                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.StartImport()
                            'hapus data antri untuk transaksi ini
                            If Me.SudahHapusAntrian = False Then
                                If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                    MsgBox("Logic error(1), contact your provider when you see this message", _
                                    MsgBoxStyle.Information, "Information")
                                End If
                                Me.SudahHapusAntrian = True
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Private Sub cboCustomerFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCustomerFormat.SelectedIndexChanged

    End Sub

    Private Sub LlbClearFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbClearFilter.Click
        Try

            Me.dgvCreatedBy.Rows.Clear()
            Me.dgvItemProductGroup.Rows.Clear()
            Me.DgvCustomer.Rows.Clear()
            Me.dtpFrom.Value = Today
            Me.dtpTo.Value = Today
            Me.TxtTransactionNumber.Clear()
            Me.TxtLocCode.Clear()
            Me.ChkIgnoreLowTolerance.Checked = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    ''data demo
    'Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
    '    Dim rowdgv As DataGridViewRow
    '    Dim i As Integer = 0
    '    Try
    '        Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
    '        Me.dgvItem.Rows.Clear()
    '        Me.dgvItem.AllowUserToAddRows = False
    '        Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '        For i = 0 To 4
    '            Me.dgvItem.Rows.Add()
    '            rowdgv = Me.dgvItem.Rows(i)
    '            With rowdgv
    '                .Cells(Me.ColNo.Index).Value = i + 1
    '                .Cells(Me.ColCustomer.Index).Value = "Customer " & i + 1
    '                .Cells(Me.ColDOcumentNO.Index).Value = "00000" & i + 1 & "/BOOK/05/2026"
    '                .Cells(Me.ColDate.Index).Value = 19 + i & " May 2026"
    '                .Cells(Me.ColCodeNo.Index).Value = "PRODUCT " & i + 1
    '                .Cells(Me.ColDesc.Index).Value = "Product no " & i + 1
    '                .Cells(Me.ColQty.Index).Value = i + 1 + 5
    '                .Cells(Me.ColTolerance.Index).Value = "0"
    '                .Cells(Me.ColQtyDO.Index).Value = 0
    '                .Cells(Me.ColUnit.Index).Value = "PCS"
    '                .Cells(Me.ColChkDO.Index).Value = False
    '                .Cells(Me.ColRemarksItem.Index).Value = "DO No " & i + 1
    '                .Cells(Me.ColBatchControlled.Index).Value = 0
    '            End With
    '        Next

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        Me.dgvItem.AllowUserToAddRows = True
    '        Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '        Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '    End Try

    'End Sub

    'Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
    '    Try
    '        If e.ColumnIndex = Me.ColChkDO.Index Then
    '            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = False Then
    '                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = True
    '                Me.dgvItem.Item(Me.ColQtyDO.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value
    '            Else
    '                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = False
    '                Me.dgvItem.Item(Me.ColQtyDO.Index, e.RowIndex).Value = 0
    '            End If
    '            'Me.KontrolLinkLabel(e.RowIndex)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub rbInventoryBooking_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbInventoryBooking.CheckedChanged
        Try
            If Me.rbInventoryBooking.Checked Then
                Me.GBPeriod.Text = "Inventory Booking Period"
                Me.lblFind.Text = "Booking No."
                Me.TxtLocCode.Clear()
                Me.TxtPrimaryLoc.Text = -1
                Me.TxtLocCode.Enabled = False
                Me.BtnFindLoc.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub rbSO_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbSO.CheckedChanged
        Try
            If Me.rbSO.Checked Then
                Me.GBPeriod.Text = "SO Period"
                Me.lblFind.Text = "SO No."
                Me.TxtLocCode.Enabled = True
                Me.BtnFindLoc.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Function CekStock(ByVal rowDgv As DataGridViewRow, Optional ByVal vQtyUser As Decimal = 0) As Boolean
        Dim PrimaryProduct As Double = rowDgv.Cells(Me.ColPrimaryPoduct.Index).Value
        Dim PrimaryUnit As Double = rowDgv.Cells(Me.ColPrimaruUnitDefault.Index).Value
        Dim PrimarySO As Double = rowDgv.Cells(Me.ColPrimaryTransaction.Index).Value
        Dim PrimaryLocation As Double = rowDgv.Cells(Me.ColPrimaryLocation.Index).Value
        Dim PrimaryWarehouse As Double = rowDgv.Cells(Me.ColPrimaryWarehouse.Index).Value
        Dim BatchNo As String = rowDgv.Cells(Me.ColBatchNo.Index).Value
        Dim ExpiryDate As String = rowDgv.Cells(Me.ColExpiryDate.Index).Value
        Dim QtyInput As Decimal = rowDgv.Cells(Me.ColQty.Index).Value
        Dim isMulti As Boolean = False
        Dim QtyAvailable As Decimal = 0
        Dim QtyTemp As Decimal = 0
        Dim QtyPendingDO As Decimal = 0
        Dim QtyBooking As Decimal = 0
        Try
            'kalo multi ada, multi
            If rowDgv.Cells(Me.ColPrimaryUnitMulti.Index).Value <> -1 Then
                PrimaryUnit = rowDgv.Cells(Me.ColPrimaryUnitMulti.Index).Value
                isMulti = True
            End If

            'kalo dari Inventory booking
            If Me.rbInventoryBooking.Checked Then
                PrimarySO = rowDgv.Cells(Me.ColPrimarySO.Index).Value
            End If


            If rowDgv.Cells(Me.ColBatchControlled.Index).Value = 0 Then
                QtyAvailable = getAvailableQty(PrimaryProduct, PrimaryUnit, PrimaryLocation, PrimaryWarehouse)
                QtyPendingDO = getPendingDOQty(PrimaryProduct, PrimaryUnit, PrimaryLocation, PrimaryWarehouse)
                QtyTemp = Me.getTempQty(PrimaryProduct, PrimaryUnit, PrimaryLocation, PrimaryWarehouse, PrimarySO) 'tanpa transaksi yang sama
                If Me.rbSO.Checked Then
                    QtyTemp = Me.getTempQty(PrimaryProduct, PrimaryUnit, PrimaryLocation, PrimaryWarehouse, PrimarySO) 'tanpa transaksi yang sama
                End If
            Else
                QtyAvailable = getAvailableQty(PrimaryProduct, PrimaryUnit, PrimaryLocation, PrimaryWarehouse, BatchNo, ExpiryDate)
                QtyPendingDO = getPendingDOQty(PrimaryProduct, PrimaryUnit, PrimaryLocation, PrimaryWarehouse, BatchNo, ExpiryDate)
                QtyTemp = Me.getTempQty(PrimaryProduct, PrimaryUnit, PrimaryLocation, PrimaryWarehouse, PrimarySO, BatchNo, ExpiryDate) 'tanpa transaksi yang sama
                If Me.rbSO.Checked Then
                    QtyTemp = Me.getTempQty(PrimaryProduct, PrimaryUnit, PrimaryLocation, PrimaryWarehouse, PrimarySO) 'tanpa transaksi yang sama
                End If
            End If

            'cek QtyAvailabel dan dikuran qty temp, jika qty lebih kecil, isi qty terbanyak, jika 0 atau minus, return false
            QtyAvailable = QtyAvailable - QtyTemp - QtyPendingDO
            If QtyAvailable <= 0 Then
                Return False
            Else
                'simpan primarylocation dan warehouse ke dgv, jika SO
                If Me.rbSO.Checked Then
                    rowDgv.Cells(Me.ColPrimaryLocation.Index).Value = PrimaryLocation
                    rowDgv.Cells(Me.ColPrimaryWarehouse.Index).Value = PrimaryWarehouse
                End If
                If vQtyUser <= 0 Then
                    If QtyInput < QtyAvailable Then
                        rowDgv.Cells(Me.ColQtyDO.Index).Value = Me.ObjCD.DigitGrouping(QtyInput, rowDgv.Cells(Me.ColDecimalMulti.Index).Value)
                        'simpan ke temp
                        Me.simpanStockTemp(rowDgv, QtyInput)
                    Else
                        rowDgv.Cells(Me.ColQtyDO.Index).Value = Me.ObjCD.DigitGrouping(QtyAvailable, rowDgv.Cells(Me.ColDecimalMulti.Index).Value)
                        Me.simpanStockTemp(rowDgv, QtyAvailable)
                    End If
                Else
                    If vQtyUser <= QtyAvailable Then
                        rowDgv.Cells(Me.ColQtyDO.Index).Value = Me.ObjCD.DigitGrouping(vQtyUser, rowDgv.Cells(Me.ColDecimalMulti.Index).Value)
                        'simpan ke temp
                        Me.simpanStockTemp(rowDgv, vQtyUser)
                    Else
                        rowDgv.Cells(Me.ColQtyDO.Index).Value = Me.ObjCD.DigitGrouping(QtyAvailable, rowDgv.Cells(Me.ColDecimalMulti.Index).Value)
                        Me.simpanStockTemp(rowDgv, QtyAvailable)
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Function

    Private Function getAvailableQty(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByRef vPrimaryLocation As Double, _
     ByRef vPrimaryWarehouse As Double, Optional ByVal vBatchNo As String = "", Optional ByVal vExpiryDate As Date = #1/1/1950#) As Decimal
        Dim vQtyAvailable As Decimal = 0
        Try
            If Me.rbSO.Checked Then 'belum ada location, jadi tanpa location diceknya
                'Me.TBLStock.DefaultView.RowFilter = "[Primary Product]= " & vPrimaryProduct & " AND [Primary Unit]= " & vPrimaryUnit & _
                '" AND [Lot No]= '" & vBatchNo & "' AND [Expiry Date]= '" & vExpiryDate.Date & "'"
                Me.TBLStock.DefaultView.RowFilter = "[Primary Product]= " & vPrimaryProduct & " AND [Primary Unit]= " & vPrimaryUnit & _
                " AND [Primary Storage Area]= " & vPrimaryLocation & " AND [Primary Warehouse]= " & vPrimaryWarehouse & _
                " AND [Lot No]= '" & vBatchNo & "' AND [Expiry Date]= '" & vExpiryDate.Date & "'"
                If Me.TBLStock.DefaultView.Count > 0 Then
                    Me.RowStock = Me.TBLStock.DefaultView(0).Row
                    vQtyAvailable = Me.RowStock.Qty
                    'vPrimaryLocation = Me.RowStock.Primary_Storage_Area 'diisi primarylocationnya
                    'vPrimaryWarehouse = Me.RowStock.Primary_Warehouse 'diisi primarylocationnya
                Else
                    Return 0
                End If

            Else
                Me.TBLStock.DefaultView.RowFilter = "[Primary Product]= " & vPrimaryProduct & " AND [Primary Unit]= " & vPrimaryUnit & _
                " AND [Primary Storage Area]= " & vPrimaryLocation & " AND [Primary Warehouse]= " & vPrimaryWarehouse & _
                " AND [Lot No]= '" & vBatchNo & "' AND [Expiry Date]= '" & vExpiryDate.Date & "'"
                If Me.TBLStock.DefaultView.Count > 0 Then
                    Me.RowStock = Me.TBLStock.DefaultView(0).Row
                    vQtyAvailable = Me.RowStock.Qty
                Else
                    Return 0
                End If
            End If

            Return vQtyAvailable
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.TBLStock.DefaultView.RowFilter = ""
        End Try
    End Function

    Private Function getBookingInvQty(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByRef vPrimaryLocation As Double, _
 ByRef vPrimaryWarehouse As Double, Optional ByVal vBatchNo As String = "", Optional ByVal vExpiryDate As Date = #1/1/1950#) As Decimal
        Dim vQtyBooking As Decimal = 0
        Try

            Me.TblBook.DefaultView.RowFilter = "[Primary Product]= " & vPrimaryProduct & " AND [Primary Unit]= " & vPrimaryUnit & _
            " AND [Primary Storage Area]= " & vPrimaryLocation & " AND [Primary Warehouse]= " & vPrimaryWarehouse & _
            " AND [Lot No]= '" & vBatchNo & "' AND [Expiry Date]= '" & vExpiryDate.Date & "'"
            If Me.TBLStock.DefaultView.Count > 0 Then
                Me.RowStock = Me.TBLStock.DefaultView(0).Row
                vQtyBooking = Me.RowStock.Qty
            Else
                Return 0
            End If

            Return vQtyBooking
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.TBLStock.DefaultView.RowFilter = ""
        End Try
    End Function

    Private Function getPendingDOQty(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByRef vPrimaryLocation As Double, _
 ByRef vPrimaryWarehouse As Double, Optional ByVal vBatchNo As String = "", Optional ByVal vExpiryDate As Date = #1/1/1950#) As Decimal
        Dim vQtyPendingDO As Decimal = 0
        Try

            Me.TblPendingDO.DefaultView.RowFilter = "[Primary Product]= " & vPrimaryProduct & " AND [Primary Unit]= " & vPrimaryUnit & _
            " AND [Primary Storage Area]= " & vPrimaryLocation & " AND [Primary Warehouse]= " & vPrimaryWarehouse & _
            " AND [Batch No]= '" & vBatchNo & "' AND [Expiry Date]= '" & vExpiryDate.Date & "'"
            If Me.TblPendingDO.DefaultView.Count > 0 Then
                For Each rowdv As DataRowView In Me.TblPendingDO.DefaultView
                    Me.RowPendingDO = rowdv.Row
                    vQtyPendingDO += Me.RowPendingDO.Total_Qty
                Next
            Else
                Return 0
            End If

            Return vQtyPendingDO
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.TBLStock.DefaultView.RowFilter = ""
        End Try
    End Function

    Private Function getTempQty(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByVal vPrimaryLocation As Double, _
    ByVal vPrimaryWarehouse As Double, ByVal vPrimarySO As Double, Optional ByVal vBatchNo As String = "", Optional ByVal vExpiryDate As Date = #1/1/1950#) As Decimal
        Dim vQtyTemp As Decimal = 0
        Try
            Me.TblStockTemp.DefaultView.RowFilter = "[Primary Product]= " & vPrimaryProduct & " AND [Primary Unit]= " & vPrimaryUnit & _
            " AND [Primary Storage Area]= " & vPrimaryLocation & " AND [Primary Warehouse]= " & vPrimaryWarehouse & " AND [Primary SO]<> " & vPrimarySO & _
            " AND [Lot No]= '" & vBatchNo & "' AND [Expiry Date]= '" & vExpiryDate.Date & "'"
            If Me.TblStockTemp.DefaultView.Count > 0 Then
                Me.rowStockTemp = Me.TblStockTemp.DefaultView(0).Row
                vQtyTemp = Me.rowStockTemp.Qty
            Else
                Return 0
            End If

            Return vQtyTemp
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.TblStockTemp.DefaultView.RowFilter = ""
        End Try
    End Function

    Private Sub simpanStockTemp(ByVal rowDGV As DataGridViewRow, ByVal vQty As Decimal)
        Dim PrimaryProduct As Double = rowDGV.Cells(Me.ColPrimaryPoduct.Index).Value
        Dim PrimaryUnit As Double = rowDGV.Cells(Me.ColPrimaruUnitDefault.Index).Value
        Dim PrimarySO As Double = rowDGV.Cells(Me.ColPrimaryTransaction.Index).Value
        Dim PrimaryLocation As Double = rowDGV.Cells(Me.ColPrimaryLocation.Index).Value
        Dim PrimaryWarehouse As Double = rowDGV.Cells(Me.ColPrimaryWarehouse.Index).Value
        Dim BatchNo As String = rowDGV.Cells(Me.ColBatchNo.Index).Value
        Dim ExpiryDate As Date = rowDGV.Cells(Me.ColExpiryDate.Index).Value
        Dim Qty As Decimal = rowDGV.Cells(Me.ColQtyDO.Index).Value
        Try
            'kalo multi ada, multi
            If rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value <> -1 Then
                PrimaryUnit = rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value
            End If
            'kalo dari Inventory booking
            If Me.rbInventoryBooking.Checked Then
                PrimarySO = rowDGV.Cells(Me.ColPrimarySO.Index).Value
            End If

            Me.TblStockTemp.DefaultView.RowFilter = "[Primary Product]= " & PrimaryProduct & " AND [Primary Unit]= " & PrimaryUnit & _
                        " AND [Primary Storage Area]= " & PrimaryLocation & " AND [Primary Warehouse]= " & PrimaryWarehouse & " AND [Primary SO]= " & PrimarySO & _
                        " AND [Lot No]= '" & BatchNo & "' AND [Expiry Date]= '" & ExpiryDate.Date & "'"
            If Me.TblStockTemp.DefaultView.Count > 0 Then 'hanya update
                Me.rowStockTemp = Me.TblStockTemp.DefaultView(0).Row
                rowStockTemp.Qty = vQty
            Else
                Me.rowStockTemp = Me.TblStockTemp.NewTempStockDetailRow
                With rowDGV
                    rowStockTemp.Primary_Product = PrimaryProduct
                    rowStockTemp.Primary_Storage_Area = PrimaryLocation
                    rowStockTemp.Primary_SO = PrimarySO
                    rowStockTemp.Primary_Unit = PrimaryUnit
                    rowStockTemp.Primary_Warehouse = PrimaryWarehouse
                    rowStockTemp.Lot_No = BatchNo
                    rowStockTemp.Expiry_Date = ExpiryDate.Date
                    rowStockTemp.Qty = vQty
                    rowStockTemp.Location = Me.ObjLoc.GetStorageAreaCode(PrimaryLocation)
                    rowStockTemp.Warehouse = Me.ObjWR.GetWarehouseCode(PrimaryWarehouse)
                    rowStockTemp._Decimal = rowDGV.Cells(Me.ColDecimalMulti.Index).Value
                End With
                Me.TblStockTemp.Rows.Add(Me.rowStockTemp)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HapusStockTemp(ByVal rowDGV As DataGridViewRow)
        Dim PrimaryProduct As Double = rowDGV.Cells(Me.ColPrimaryPoduct.Index).Value
        Dim PrimaryUnit As Double = rowDGV.Cells(Me.ColPrimaruUnitDefault.Index).Value
        Dim PrimarySO As Double = rowDGV.Cells(Me.ColPrimaryTransaction.Index).Value
        Dim PrimaryLocation As Double = rowDGV.Cells(Me.ColPrimaryLocation.Index).Value
        Dim PrimaryWarehouse As Double = rowDGV.Cells(Me.ColPrimaryWarehouse.Index).Value
        Dim BatchNo As String = rowDGV.Cells(Me.ColBatchNo.Index).Value
        Dim ExpiryDate As Date = rowDGV.Cells(Me.ColExpiryDate.Index).Value
        Dim Qty As Decimal = rowDGV.Cells(Me.ColQtyDO.Index).Value
        Try
            'kalo multi ada, multi
            If rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value <> -1 Then
                PrimaryUnit = rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value
            End If
            'kalo dari Inventory booking
            If Me.rbInventoryBooking.Checked Then
                PrimarySO = rowDGV.Cells(Me.ColPrimarySO.Index).Value
            End If

            Me.TblStockTemp.DefaultView.RowFilter = "[Primary Product]= " & PrimaryProduct & " AND [Primary Unit]= " & PrimaryUnit & _
                        " AND [Primary Storage Area]= " & PrimaryLocation & " AND [Primary Warehouse]= " & PrimaryWarehouse & " AND [Primary SO]= " & PrimarySO & _
                        " AND [Lot No]= '" & BatchNo & "' AND [Expiry Date]= '" & ExpiryDate.Date & "'"
            If Me.TblStockTemp.DefaultView.Count > 0 Then 'delete kalau ada
                For i As Integer = Me.TblStockTemp.DefaultView.Count - 1 To 0
                    Me.rowStockTemp = Me.TblStockTemp.DefaultView(i).Row
                    Me.TblStockTemp.Rows.Remove(Me.rowStockTemp)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HapusStockTempDgv(ByVal rowDGV As DataGridViewRow)
        Dim PrimaryProduct As Double = rowDGV.Cells(Me.ColPrimaryPoduct.Index).Value
        Dim PrimaryUnit As Double = rowDGV.Cells(Me.ColPrimaruUnitDefault.Index).Value
        Dim PrimarySO As Double = rowDGV.Cells(Me.ColPrimaryTransaction.Index).Value
        Dim PrimaryLocation As Double = rowDGV.Cells(Me.ColPrimaryLocation.Index).Value
        Dim PrimaryWarehouse As Double = rowDGV.Cells(Me.ColPrimaryWarehouse.Index).Value
        Dim BatchNo As String = rowDGV.Cells(Me.ColBatchNo.Index).Value
        Dim ExpiryDate As Date = rowDGV.Cells(Me.ColExpiryDate.Index).Value
        Dim Qty As Decimal = rowDGV.Cells(Me.ColQtyDO.Index).Value
        Dim i As Integer = 0
        Try
            'kalo multi ada, multi
            If rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value <> -1 Then
                PrimaryUnit = rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value
            End If
            'kalo dari Inventory booking
            If Me.rbInventoryBooking.Checked Then
                PrimarySO = rowDGV.Cells(Me.ColPrimarySO.Index).Value
            End If

            Me.TblStockTemp.DefaultView.RowFilter = "[Primary Product]= " & PrimaryProduct & " AND [Primary Unit]= " & PrimaryUnit & _
                        " AND [Primary Storage Area]= " & PrimaryLocation & " AND [Primary Warehouse]= " & PrimaryWarehouse & " AND [Primary SO]= " & PrimarySO 
            If Me.TblStockTemp.DefaultView.Count > 0 Then 'delete kalau ada
                For i = Me.TblStockTemp.DefaultView.Count - 1 To 0 Step -1
                    Me.rowStockTemp = Me.TblStockTemp.DefaultView(i).Row
                    Me.TblStockTemp.Rows.Remove(Me.rowStockTemp)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUserCode.KeyDown
        Dim i As New Core_Module.FrmCariMasterUser
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 AndAlso Me.txtUserCode.Enabled Then
                If Me.ObjUser.ApakahSudahAda(Me.txtUserCode.Text) Then
                    Dim PrimaryBaru As Decimal = Me.ObjUser.GetPrimary(Me.txtUserCode.Text)
                    If PrimaryBaru <> Me.txtPrimaryUser.Text Then
                        Me.txtPrimaryUser.Text = PrimaryBaru
                    End If
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtUserCode.Text <> "" Then
                        i.TxtCari.Text = Me.txtUserCode.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.txtPrimaryUser.Text = Me.ObjUser.GetPrimary(Me.txtUserCode.Text)
                        Me.txtUserCode.Text = i.Kodeuser
                    Else
                        Me.txtUserCode.Clear()
                        Me.txtPrimaryUser.Text = "-1"
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Public Sub txtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserCode.Leave
        Dim i As New Core_Module.FrmCariMasterUser
        i.backColorAll = Me.BackColorAll
        Try
            If Me.txtUserCode.Text <> "" AndAlso Me.txtUserCode.Enabled Then
                If Me.ObjUser.ApakahSudahAda(Me.txtUserCode.Text) Then
                    Dim PrimaryBaru As Decimal = Me.ObjUser.GetPrimary(Me.txtUserCode.Text)
                    If PrimaryBaru <> Me.txtPrimaryUser.Text Then
                        Me.txtPrimaryUser.Text = PrimaryBaru
                    End If
                Else
                    If Me.txtUserCode.Text <> "" Then
                        i.TxtCari.Text = Me.txtUserCode.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.txtPrimaryUser.Text = Me.ObjUser.GetPrimary(Me.txtUserCode.Text)
                        Me.txtUserCode.Text = i.Kodeuser
                    Else
                        Me.txtUserCode.Clear()
                        Me.txtPrimaryUser.Text = "-1"
                    End If
                End If
            Else
                Me.txtUserCode.Clear()
                Me.txtPrimaryUser.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindUser.Click
        Dim i As New Core_Module.FrmCariMasterUser
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtUserCode.Enabled Then
                If Me.txtUserCode.Text <> "" Then
                    i.TxtCari.Text = Me.txtUserCode.Text
                End If

                i.ShowDialog()
                If i.Kodeuser <> "" Then
                    Me.txtPrimaryUser.Text = Me.ObjUser.GetPrimary(i.Kodeuser)
                    Me.txtUserCode.Text = i.Kodeuser
                Else
                    Me.txtUserCode.Clear()
                    Me.txtPrimaryUser.Text = "-1"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub
End Class