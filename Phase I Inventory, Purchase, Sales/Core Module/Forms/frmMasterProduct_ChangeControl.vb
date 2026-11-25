Option Explicit On
Option Strict Off

Public Class frmMasterProduct_ChangeControl
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""

    Public Enum ProductLevel As Integer
        ProductGroup
        Brand
        SubBrand
        Category
        Product
    End Enum

    Dim ObjMonitor As New Core_Module.ClsTransaksiMonitor
    Dim LangsungClose As Boolean = False
    Dim Buffer As String = ""
    Dim rowIndex As Integer = -1
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCategory As New Core_Module.ClsMasterCategory
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup 
    Dim AdaBerubah As Boolean = False
    Dim DTA_tblmaster_product1 As New DseMasterProduct_ChangeControlTableAdapters.tblmaster_product1TableAdapter
    Dim DTA_tblmaster_product2 As New DseMasterProduct_ChangeControlTableAdapters.tblmaster_product4TableAdapter
    Dim DTA_tblmaster_product3 As New DseMasterProduct_ChangeControlTableAdapters.tblmaster_product5TableAdapter
    Dim ScrollPosition As Integer = 0
    Dim SplitterPosition As Double = 0
    Dim UpdateSplitterPosition As Boolean = False

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Private Sub frmMasterProduct_ChangeControl_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LangsungClose = False Then
                If MsgBox("Close form?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterProduct_ChangeControl_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim JmlOnline As Integer = 0

        Try
            Me.TxtReadOnly.Visible = False
            Me.TxtReadOnly.TabStop = False

            'pastikan tidak ada user lain yang online saat mengerjakan data ini
            JmlOnline = Me.ObjMonitor.GetJmlMonitor(ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License)
            If JmlOnline > 1 Then
                MsgBox("Please make sure other user is not using SMESS when you updating data.", MsgBoxStyle.Information, "Information")
                Dim i As New Core_Module.FrmTransaksiMonitor
                i.ShowDialog()
                Me.LangsungClose = True
                i.Dispose()

                'Me.Close()
                Me.TimerClose.Enabled = True
                Exit Sub
            End If

            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvProduct.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.chkControlBatchNo.Checked = True
            Me.chkControlBatchNo.Checked = False
            Me.chkControlExpiry.Checked = True
            Me.chkControlExpiry.Checked = False
            Me.chkControlLocation.Checked = True
            Me.chkControlLocation.Checked = False

            Me.chkModifyBatch.Checked = True
            Me.chkModifyBatch.Checked = False
            Me.chkModifyExpiry.Checked = True
            Me.chkModifyExpiry.Checked = False
            Me.chkModifyLocation.Checked = True
            Me.chkModifyLocation.Checked = False

            Me.chkModifyBatch.Focus()

            Me.DTA_tblmaster_product1.Connection.ConnectionString = Core_Module.GetConstring

            Me.TimerNote.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkControlBatchNo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkControlBatchNo.CheckedChanged
        Try
            If Me.chkControlBatchNo.Checked = True Then
                Me.TxtBatchNoDefault.ReadOnly = False
                Me.TxtBatchNoDefault.BackColor = Me.TxtPrimaryLocationDefault.BackColor
                Me.TxtBatchNoDefault.TabStop = True
                If Me.chkModifyExpiry.Checked = True Then
                    Me.chkControlExpiry.Enabled = True
                End If
            Else
                If Me.chkModifyExpiry.Checked = True Then
                    Me.chkControlExpiry.Checked = False
                    Me.chkControlExpiry.Enabled = False
                    Me.dtpExpiryDefault.Enabled = False
                End If

                Me.TxtBatchNoDefault.Text = ""
                Me.TxtBatchNoDefault.ReadOnly = True
                Me.TxtBatchNoDefault.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtBatchNoDefault.TabStop = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkControlExpiry_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkControlExpiry.CheckedChanged
        Try
            If Me.chkControlExpiry.Checked = True Then
                Me.dtpExpiryDefault.Enabled = True
            Else
                Me.dtpExpiryDefault.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkControlLocation_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkControlLocation.CheckedChanged
        Try
            If Me.chkControlLocation.Checked = True Then
                Me.TxtLocationDefault.ReadOnly = False
                Me.TxtLocationDefault.BackColor = Me.TxtPrimaryLocationDefault.BackColor
                Me.btnFindLocation.Enabled = True
                Me.TxtLocationDefault.TabStop = True
            Else
                Me.TxtLocationDefault.ReadOnly = True
                Me.TxtLocationDefault.BackColor = Me.TxtReadOnly.BackColor
                Me.btnFindLocation.Enabled = False
                Me.TxtLocationDefault.TabStop = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStart_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, -1, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterProduct_UpdateControlState, Me.UserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            Me.Start1()
                            Exit Sub
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
    End Sub

    Private Sub Start1()
        Dim JmlOnline As Integer = 0
        Dim rowDGV As DataGridViewRow
        Dim AdaYgUpdate As Boolean = False
        Dim PerluHapusAntrian As Boolean = True
        Dim ObjIdentifier As New ClsIdentifier

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If Me.ParentForm.MdiChildren.Length > 2 Then
                MsgBox("Please close other form before processing data update.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.chkModifyExpiry.Checked = False And Me.chkModifyLocation.Checked = False And Me.chkModifyBatch.Checked = False Then
                MsgBox("Please specify control state to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.chkControlBatchNo.Checked = True Then
                Me.TxtBatchNoDefault.Text = Me.TxtBatchNoDefault.Text.Trim
                If Me.TxtBatchNoDefault.Text = "" Then
                    MsgBox("Please specify default batch no.", MsgBoxStyle.Information, "Information")
                    Me.TxtBatchNoDefault.Focus()
                    Exit Sub
                End If
            End If
            If Me.chkControlExpiry.Checked = True Then
                If Me.dtpExpiryDefault.Value.Date <= "1/1/1950" Then
                    MsgBox("Please specify value > 1 Jan 1950", MsgBoxStyle.Information, "Information")
                    Me.dtpExpiryDefault.Focus()
                    Exit Sub
                End If
            End If
            If Me.chkControlLocation.Checked = True Then
                Me.TxtLocationDefault.Text = Me.TxtLocationDefault.Text.Trim
                If Me.TxtLocationDefault.Text = "" Then
                    MsgBox("Please specify default location.", MsgBoxStyle.Information, "Information")
                    Me.TxtLocationDefault.Focus()
                    Exit Sub
                End If
            End If

            'pastikan tidak ada user lain yang online saat mengerjakan data ini
            JmlOnline = Me.ObjMonitor.GetJmlMonitor(ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License)
            If JmlOnline > 1 Then
                MsgBox("Please make sure other user is not using SMESS when you updating data.", MsgBoxStyle.Information, "Information")
                Dim i As New Core_Module.FrmTransaksiMonitor
                i.ShowDialog()
                Exit Sub
            End If

            If Me.dgvProduct.RowCount = 0 Then
                MsgBox("Please specify at least 1 product to modify control state.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            For Each rowDGV In Me.dgvProduct.Rows
                If rowDGV.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                    AdaYgUpdate = True
                    Exit For
                End If
            Next
            If AdaYgUpdate = False Then
                MsgBox("Please specify at least 1 product to modify control state.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.chkModifyExpiry.Checked = True And Me.chkModifyBatch.Checked = True Then
                If Me.chkControlBatchNo.Checked = False And Me.chkControlExpiry.Checked = True Then
                    MsgBox("A product must use batch no control to enable expiry date control.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.chkModifyExpiry.Checked = True Or Me.chkModifyBatch.Checked = True Then
                If Me.SimulasiKontrolExpiry = False Then
                    Exit Sub
                End If
            End If

            If MsgBox("Start data update?" + vbCrLf + _
            "WARNING: This change is permanent. We advise you to back up your data before executing this change.", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.Prg1.Value = 0
                Me.LblProgress.Text = ""
                Me.Panel1.Visible = True

                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgw1.RunWorkerAsync()
                PerluHapusAntrian = False  '-> tandai hapus antriannya di bgw process finished
            Else
                'hapus data antri untuk transaksi ini
                If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                    MsgBox("Logic error(1), contact your provider when you see this message", _
                    MsgBoxStyle.Information, "Information")
                End If
                PerluHapusAntrian = False ' pada bagian finally, tidak perlu lagi hapus antrian
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'hapus data antri untuk transaksi ini
            If PerluHapusAntrian = True Then
                If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                    MsgBox("Logic error(1), contact your provider when you see this message", _
                    MsgBoxStyle.Information, "Information")
                End If
            End If
        End Try
    End Sub 

    Private Sub IsikanTBLDummy(ByVal TBL As DseMasterProduct_ChangeControl.tblmaster_product2DataTable)
        Dim rowDGV As DataGridViewRow
        Dim row As DseMasterProduct_ChangeControl.tblmaster_product2Row

        Try
            For Each rowDGV In Me.dgvProduct.Rows
                If rowDGV.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                    row = TBL.Newtblmaster_product2Row
                    row.Primary_Product = rowDGV.Cells(Me.ColPrimaryProduct.Index).Value
                    row.Product_Code = rowDGV.Cells(Me.ColCodeNo.Index).Value
                    row.Description = rowDGV.Cells(Me.ColDesc.Index).Value

                    If rowDGV.Cells(Me.ColExpiry.Index).Value = Boolean.TrueString Then
                        row.Expiry_Date_Control = 1
                    Else
                        row.Expiry_Date_Control = 0
                    End If
                    If rowDGV.Cells(Me.ColBatchNo.Index).Value = Boolean.TrueString Then
                        row.Batch_No_Control = 1
                    Else
                        row.Batch_No_Control = 0
                    End If
                    If rowDGV.Cells(Me.ColLocation.Index).Value = Boolean.TrueString Then
                        row.Location_Control = 1
                    Else
                        row.Location_Control = 0
                    End If

                    TBL.Rows.Add(row)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function SimulasiKontrolExpiry() As Boolean
        Dim TBL As New DseMasterProduct_ChangeControl.tblmaster_product2DataTable
        Dim row As DseMasterProduct_ChangeControl.tblmaster_product2Row
        Dim Kontrol As Integer = 0
        Dim idx As Integer = 0
        Dim Dummy As String = ""

        Try
            '1. salin dari dgv ke tabel dummy
            Me.IsikanTBLDummy(TBL)

            '2. update batch no kontrol, jika di ubah
            If Me.chkModifyBatch.Checked = True Then
                If Me.chkControlBatchNo.Checked = True Then
                    Kontrol = 1
                Else
                    Kontrol = 0
                End If

                For Each row In TBL.Rows
                    row.Batch_No_Control = Kontrol
                Next
            End If
            '3. update expiry date kontrol, jika di ubah
            If Me.chkModifyExpiry.Checked = True Then
                If Me.chkControlExpiry.Checked = True Then
                    Kontrol = 1
                Else
                    Kontrol = 0
                End If
                For Each row In TBL.Rows
                    row.Expiry_Date_Control = Kontrol
                Next
            End If

            'cek setiap bari, jika ada yang expiry date kontrol tapi ga batch no kontrol
            While idx <= TBL.Rows.Count - 1
                row = TBL.Rows(idx)

                If row.Expiry_Date_Control = 1 And row.Batch_No_Control = 0 Then
                    idx += 1
                Else
                    TBL.Rows.Remove(row)
                End If
            End While

            If TBL.Rows.Count = 0 Then
                Return True
            Else
                TBL.DefaultView.Sort = "[Product Code]"
                idx = 1
                For Each DV As DataRowView In TBL.DefaultView
                    row = DV.Row
                    If Dummy = "" Then
                        Dummy = idx.ToString + ". " + row.Product_Code
                    Else
                        Dummy = Dummy + vbCrLf + idx.ToString + ". " + row.Product_Code
                    End If
                    idx += 1
                Next

                MsgBox("We can't change control state for the following product: " + vbCrLf + Dummy + vbCrLf _
                + "Please make sure every product with expiry date control use batch no control.", MsgBoxStyle.Information, "Information")

                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL.Dispose()
        End Try
    End Function

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
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
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

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            If Me.rowIndex <> -1 And Me.rowIndex <> Me.dgvItemProductGroup.RowCount - 1 Then
                Me.dgvItemProductGroup.Rows.RemoveAt(Me.rowIndex)
                rowIndex = -1
                Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
                Me.AdaBerubah = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetProductLevel(ByVal vProductGroupLevelStr As String) As ProductLevel
        Try
            If vProductGroupLevelStr = "Product Group" Then
                Return ProductLevel.ProductGroup
            ElseIf vProductGroupLevelStr = "Brand" Then
                Return ProductLevel.Brand
            ElseIf vProductGroupLevelStr = "Sub Brand" Then
                Return ProductLevel.SubBrand
            ElseIf vProductGroupLevelStr = "Category" Then
                Return ProductLevel.Category
            Else
                Return ProductLevel.Product
            End If

            Return ProductLevel.Product
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ProductLevel.Product
        End Try
    End Function

    Public Function GetProductLevelStr(ByVal vProductGroupLevel As ProductLevel) As String
        Try
            If vProductGroupLevel = ProductLevel.ProductGroup Then
                Return "Product Group"
            ElseIf vProductGroupLevel = ProductLevel.Brand Then
                Return "Brand"
            ElseIf vProductGroupLevel = ProductLevel.SubBrand Then
                Return "Sub Brand"
            ElseIf vProductGroupLevel = ProductLevel.Category Then
                Return "Category"
            Else
                Return "Product"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub dgvItemProductGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProductGroup.CellEndEdit
        Dim Level As ProductLevel = ProductLevel.Brand

        Try
            If IsNothing(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    'kalo kode product group dihapus, maka hapus satu bari semuanya
                    Me.rowIndex = e.RowIndex
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
                    Level = Me.GetProductLevel(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
                    If Level = ProductLevel.ProductGroup Then
                        Me.EndEditProductGroup(e, Me.ProductGroup_Level)
                    ElseIf Level = ProductLevel.Brand Then
                        Me.EndEditProductGroup(e, Me.Brand_Level)
                    ElseIf Level = ProductLevel.SubBrand Then
                        Me.EndEditProductGroup(e, Me.SubBrand_Level)
                    ElseIf Level = ProductLevel.Category Then
                        Me.EndEditCategory(e)
                    Else
                        Me.EndEditProduct(e)
                    End If

                    AdaBerubah = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.rowIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If
            'pastikan tidak double dengan yang lainnya
            rowIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If rowIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.rowIndex = e.RowIndex
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
            If Me.ObjCategory.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode category ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode category yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterCategory
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryCategory <> -1 Then
                    '2.a. category yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryCategory
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                ElseIf ij.PrimaryCategory = -1 Then
                    '2.b. kode category yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.rowIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            rowIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If rowIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.rowIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.Buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
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
                        Me.rowIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            rowIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If rowIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.rowIndex = e.RowIndex
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
        'do nothing -> dibuat untuk handle data error event, jangan dihapuskan
    End Sub

    Private Sub dgvItemProductGroup_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemProductGroup.Leave
        Try
            Me.LoadDGVProduct()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemProductGroup.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemProductGroup.Rows(Me.dgvItemProductGroup.Rows.Count - 1)
            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = Me.GetProductLevelStr(ProductLevel.ProductGroup)
            rowDGV.Cells(Me.ColProductGroup_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemProductGroup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvProduct.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColModify.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProduct.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColModify.Index Then
                'do nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDGV In Me.dgvProduct.Rows
                rowDGV.Cells(Me.ColModify.Index).Value = Boolean.TrueString
            Next
            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDGV In Me.dgvProduct.Rows
                rowDGV.Cells(Me.ColModify.Index).Value = Boolean.FalseString
            Next
            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGVProduct()
        Dim DTA As New DseMasterProduct_ChangeControlTableAdapters.tblmaster_productTableAdapter
        Dim TBL As New DseMasterProduct_ChangeControl.tblmaster_productDataTable
        Dim TBL1 As New DseMasterProduct_ChangeControl.tblmaster_productDataTable
        Dim row, row1 As DseMasterProduct_ChangeControl.tblmaster_productRow
        Dim rowDGV As DataGridViewRow
        Dim Level As ProductLevel = ProductLevel.Brand
        Dim DV As Data.DataRowView
        Dim Ketemu As Boolean = False

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemProductGroup.AllowUserToAddRows = False
            TBL.DefaultView.Sort = "Primary"
            TBL.DefaultView.RowFilter = ""

            'kosongkan dgv product
            Me.dgvProduct.Rows.Clear()

            'untuk setiap baris pada dgv product filter, isikan ke tabel1
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                TBL1.Rows.Clear()

                Level = Me.GetProductLevel(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, rowDGV.Index).Value)
                If Level = ProductLevel.ProductGroup Then
                    DTA.FillBy_PrimaryProductGroup(TBL1, CDec(rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value))
                ElseIf Level = ProductLevel.Brand Then
                    DTA.FillBy_PrimaryBrand(TBL1, CDec(rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value))
                ElseIf Level = ProductLevel.SubBrand Then
                    DTA.FillBy_PrimarySubBrand(TBL1, CDec(rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value))
                ElseIf Level = ProductLevel.Category Then
                    DTA.FillBy_PrimaryCategory(TBL1, CDec(rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value))
                Else
                    DTA.FillBy_PrimaryProduct(TBL1, CDec(rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value))
                End If

                While TBL1.Rows.Count > 0
                    row1 = TBL1.Rows(0)
                    Ketemu = False

                    For Each DV In TBL.DefaultView
                        row = DV.Row
                        If row.Primary = row1.Primary Then
                            Ketemu = True
                            Exit For
                        ElseIf row.Primary < row1.Primary Then
                            Exit For
                        End If
                    Next

                    If Ketemu = False Then
                        row = TBL.NewRow
                        For ij As Integer = 0 To TBL.Columns.Count - 1
                            row.Item(ij) = row1.Item(ij)
                        Next
                        TBL.Rows.Add(row)
                    End If

                    TBL1.Rows.Remove(row1)
                End While
            Next

            'pindahkan semua dari tabel utama ke dgv
            For Each row In TBL.Rows
                Me.dgvProduct.Rows.Add()
                rowDGV = Me.dgvProduct.Rows(Me.dgvProduct.Rows.Count - 1)
                If row.Batch_No_Controlled = 1 Then
                    rowDGV.Cells(Me.ColBatchNo.Index).Value = True
                Else
                    rowDGV.Cells(Me.ColBatchNo.Index).Value = False
                End If
                rowDGV.Cells(Me.ColExpiry.Index).Value = row.Expiry_Date_Controlled
                If row.Location_Controlled = 1 Then
                    rowDGV.Cells(Me.ColLocation.Index).Value = True
                Else
                    rowDGV.Cells(Me.ColLocation.Index).Value = False
                End If
                rowDGV.Cells(Me.ColCodeNo.Index).Value = row.Product_Code
                rowDGV.Cells(Me.ColDesc.Index).Value = row.Description
                rowDGV.Cells(Me.ColPrimaryProduct.Index).Value = row.Primary
                rowDGV.Cells(Me.ColModify.Index).Value = False
                rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
            Next

            Me.dgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItemProductGroup.AllowUserToAddRows = True

            Me.AdaBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Private Sub TxtLocationDefault_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationDefault.GotFocus
        Try
            If Me.TxtLocationDefault.Enabled = True Then
                If Me.TxtLocationDefault.ReadOnly = False Then
                    Me.TxtLocationDefault.BackColor = Me.BackColorAll
                Else
                    Me.TxtLocationDefault.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationDefault_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationDefault.LostFocus
        Try
            If Me.TxtLocationDefault.Enabled = True Then
                If Me.TxtLocationDefault.ReadOnly = False Then
                    Me.TxtLocationDefault.BackColor = Me.TxtPrimaryLocationDefault.BackColor
                Else
                    Me.TxtLocationDefault.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBatchNoDefault_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBatchNoDefault.GotFocus
        Try
            If Me.TxtBatchNoDefault.Enabled = True Then
                If Me.TxtBatchNoDefault.ReadOnly = False Then
                    Me.TxtBatchNoDefault.BackColor = Me.BackColorAll
                Else
                    Me.TxtBatchNoDefault.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBatchNoDefault_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBatchNoDefault.LostFocus
        Try
            If Me.TxtBatchNoDefault.Enabled = True Then
                If Me.TxtBatchNoDefault.ReadOnly = False Then
                    Me.TxtBatchNoDefault.BackColor = Me.TxtPrimaryLocationDefault.BackColor
                Else
                    Me.TxtBatchNoDefault.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindLocation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindLocation.Click
        Dim FrmCari As New FrmCariMasterStorageArea
        Dim ObjS As New clsMasterStorageArea

        Try
            FrmCari.BackColorAll = Me.BackColorAll
            If Me.TxtPrimaryLocationDefault.Text = "-1" And Me.TxtLocationDefault.Text = "" Then
                FrmCari.ShowDialog()
            Else
                FrmCari.txtCari.Text = Me.TxtLocationDefault.Text
                FrmCari.ShowDialog()
            End If
            If CDbl(FrmCari.PrimaryStorageArea) > 0 Then
                Me.TxtPrimaryLocationDefault.Text = FrmCari.PrimaryStorageArea
                Me.TxtLocationDefault.Text = ObjS.GetStorageAreaCode(FrmCari.PrimaryStorageArea)
                Me.TxtLocationDefaultDesc.Text = ObjS.GetDescription(FrmCari.PrimaryStorageArea)
                Me.btnStart.Focus()
            Else
                Me.TxtPrimaryLocationDefault.Text = -1
                Me.TxtLocationDefault.Text = ""
                Me.TxtLocationDefaultDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCari.Dispose()
        End Try
    End Sub

    Private Sub TxtLocationDefault_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocationDefault.KeyDown
        Dim ObjS As New clsMasterStorageArea
        Try
            If e.KeyValue = 13 Then
                If ObjS.ApakahSudahAda(Me.TxtLocationDefault.Text) = True Then
                    Me.TxtPrimaryLocationDefault.Text = ObjS.GetPrimary(Me.TxtLocationDefault.Text)
                    Me.TxtLocationDefaultDesc.Text = ObjS.GetDescription(Me.TxtPrimaryLocationDefault.Text)
                    Me.btnStart.Focus() ''Me.btnOK.Focus() -r9
                Else
                    Me.btnFindLocation.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationDefault_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationDefault.Leave
        Dim ObjS As New clsMasterStorageArea
        Try
            If Me.TxtLocationDefault.Text <> "" Then
                If ObjS.ApakahSudahAda(Me.TxtLocationDefault.Text) = True Then
                    Me.TxtPrimaryLocationDefault.Text = ObjS.GetPrimary(Me.TxtLocationDefault.Text)
                    Me.TxtLocationDefaultDesc.Text = ObjS.GetDescription(Me.TxtPrimaryLocationDefault.Text)
                Else
                    Me.btnFindLocation.PerformClick()
                End If
            Else
                Me.TxtPrimaryLocationDefault.Text = -1
                Me.TxtLocationDefault.Text = ""
                Me.TxtLocationDefaultDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvProduct.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            Me.UpdateControlState()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            'hapus data antri untuk transaksi ini
            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                MsgBox("Logic error(1), contact your provider when you see this message", _
                MsgBoxStyle.Information, "Information")
            End If

            Me.LoadDGVProduct()
            Me.dgvProduct.FirstDisplayedScrollingRowIndex = Me.ScrollPosition

            Me.LblProgress.Text = ""
            Me.Prg1.Value = 0
            Me.Panel1.Visible = False

            MsgBox("Update finished.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateControlState()
        Dim TBL As New DseMasterProduct_ChangeControl.tblmaster_product2DataTable
        Dim JmlUpdate As Integer = 0
        Dim JmlBaris As Double = 0
        Dim rowDGV As DataGridViewRow
        Dim CurStep As Integer = 0

        Try
            Me.ScrollPosition = Me.dgvProduct.FirstDisplayedScrollingRowIndex

            For Each rowDGV In Me.dgvProduct.Rows
                If rowDGV.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                    JmlBaris += 1
                End If
            Next

            If Me.chkModifyBatch.Checked = True Then
                JmlUpdate += 1
            End If
            If Me.chkModifyExpiry.Checked = True Then
                JmlUpdate += 1
            End If
            If Me.chkModifyLocation.Checked = True Then
                JmlUpdate += 1
            End If

            Me.IsikanTBLDummy(TBL)
            If Me.chkModifyBatch.Checked = True Then
                CurStep += 1
                Me.ModifyBatch(Me.chkControlBatchNo.Checked, TBL, CurStep, JmlUpdate)
            End If
            If Me.chkModifyExpiry.Checked = True Then
                CurStep += 1
                Me.ModifyExpiry(Me.chkControlExpiry.Checked, TBL, CurStep, JmlUpdate)
            End If
            If Me.chkModifyLocation.Checked = True Then
                CurStep += 1
                Me.ModifyLocation(Me.chkControlLocation.Checked, TBL, CurStep, JmlUpdate)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL.Dispose()
        End Try
    End Sub

    Private Sub ModifyLocation(ByVal vKontrol As Boolean, ByVal vTBL As DseMasterProduct_ChangeControl.tblmaster_product2DataTable, ByVal vStep As String, ByVal vTotalStep As String)
        Dim row As DseMasterProduct_ChangeControl.tblmaster_product2Row
        Dim PrimaryLokasiDefault As Integer = -1
        Dim Lanjut As Boolean = False

        Try
            If vKontrol = False Then
                PrimaryLokasiDefault = -1
            Else
                PrimaryLokasiDefault = Me.TxtPrimaryLocationDefault.Text
            End If

            Me.Prg1.Value = 0
            Me.Prg1.Maximum = vTBL.Rows.Count

            For Each row In vTBL.Rows
                Me.Prg1.Value += 1
                Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating location control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code

                Lanjut = False
                If vKontrol = False And row.Location_Control = 1 Then
                    Lanjut = True
                End If
                If vKontrol = True And row.Location_Control = 0 Then
                    Lanjut = True
                End If

                If Lanjut = True Then
                    If vKontrol = False Then
                        row.Location_Control = 0
                    Else
                        row.Location_Control = 1
                    End If

                    'tblmaster_product
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating location control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating master data)"
                    Me.DTA_tblmaster_product1.UpdateQuery_tblmaster_product(vKontrol, row.Primary_Product)
                    'tblmaster_productqty
                    Me.DTA_tblmaster_product1.UpdateQuery_tblmaster_productqty(PrimaryLokasiDefault, row.Primary_Product)
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_purchase_closing_item(PrimaryLokasiDefault, row.Primary_Product)
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating location control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating stock in/out transaction)"
                    'tbltransaksi_stockinout_item
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltransaksi_stockinout_item(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltransaksi_stockbundling_item_after
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltransaksi_stockbundling_item_after(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltransaksi_stockbundling_item_before(row.Primary_Product)
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltransaksi_stockbundling_item_before(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltransaksi_stocktransfer_item
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltransaksi_stocktransfer_item1(PrimaryLokasiDefault, row.Primary_Product)
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltransaksi_stocktransfer_item2(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_aop_item
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating location control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating purchase transaction)"
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_aop_item(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_booking_invitem
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_booking_invitem(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_purchasereturn_item
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_purchasereturn_item(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_do_item
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating location control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating sales transaction)"
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_do_item(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_salesinvoice_item
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_salesinvoice_item(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_salesinvoice_item_copy
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_salesinvoice_item_copy(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_salesinvoice_item_import
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_salesinvoice_item_import(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_salesreturn_item
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_salesreturn_item(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_salesreturn_item_copy
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_salesreturn_item_copy(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_salesreturn_item_import
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_salesreturn_item_import(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_stock_taking_itembuffer
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating location control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating stock taking transaction)"
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_stock_taking_itembuffer(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_stock_taking_itempenyesuaian
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_stock_taking_itempenyesuaian(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltrans_stock_taking_itemstart
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_stock_taking_itemstart(PrimaryLokasiDefault, row.Primary_Product)
                    'tbltransaksi_stock_card
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating location control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating stock card)"
                    Me.DTA_tblmaster_product1.UpdateQuery_tbltransaksi_stock_card(PrimaryLokasiDefault, row.Primary_Product)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ModifyBatch(ByVal vKontrol As Boolean, ByVal vTBL As DseMasterProduct_ChangeControl.tblmaster_product2DataTable, ByVal vStep As String, ByVal vTotalStep As String)
        Dim row As DseMasterProduct_ChangeControl.tblmaster_product2Row
        Dim BatchNoDefault As String = ""
        Dim Lanjut As Boolean = False

        Try
            If vKontrol = False Then
                BatchNoDefault = ""
            Else
                BatchNoDefault = Me.TxtBatchNoDefault.Text
            End If

            Me.Prg1.Value = 0
            Me.Prg1.Maximum = vTBL.Rows.Count

            For Each row In vTBL.Rows
                Me.Prg1.Value += 1
                Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating batch no control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code

                Lanjut = False
                If vKontrol = False And row.Batch_No_Control = 1 Then
                    Lanjut = True
                End If
                If vKontrol = True And row.Batch_No_Control = 0 Then
                    Lanjut = True
                End If

                If Lanjut = True Then
                    If vKontrol = False Then
                        row.Batch_No_Control = 0
                    Else
                        row.Batch_No_Control = 1
                    End If

                    'tblmaster_product
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating batch no control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating master data)"
                    Me.DTA_tblmaster_product2.UpdateQuery_tblmaster_product(vKontrol, row.Primary_Product)
                    'tblmaster_productqty
                    Me.DTA_tblmaster_product2.UpdateQuery_tblmaster_productqty(BatchNoDefault, row.Primary_Product)
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_purchase_closing_item(BatchNoDefault, row.Primary_Product)
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating batch no control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating stock in/out transaction)"
                    'tbltransaksi_stockinout_item
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltransaksi_stockinout_item(BatchNoDefault, row.Primary_Product)
                    'tbltransaksi_stockbundling_item_after
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltransaksi_stockbundling_item_after(BatchNoDefault, row.Primary_Product)
                    'tbltransaksi_stockbundling_item_before(row.Primary_Product)
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltransaksi_stockbundling_item_before(BatchNoDefault, row.Primary_Product)
                    'tbltransaksi_stocktransfer_item
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltransaksi_stocktransfer_item1(BatchNoDefault, row.Primary_Product)
                    'tbltrans_aop_item
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating batch no control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating purchase transaction)"
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_aop_item(BatchNoDefault, row.Primary_Product)
                    'tbltrans_booking_invitem
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_booking_invitem(BatchNoDefault, row.Primary_Product)
                    'tbltrans_purchasereturn_item
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_purchasereturn_item(BatchNoDefault, row.Primary_Product)
                    'tbltrans_do_item
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating batch no control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating sales transaction)"
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_do_item(BatchNoDefault, row.Primary_Product)
                    'tbltrans_salesinvoice_item
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_salesinvoice_item(BatchNoDefault, row.Primary_Product)
                    'tbltrans_salesinvoice_item_copy
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_salesinvoice_item_copy(BatchNoDefault, row.Primary_Product)
                    'tbltrans_salesinvoice_item_import
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_salesinvoice_item_import(BatchNoDefault, row.Primary_Product)
                    'tbltrans_salesreturn_item
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_salesreturn_item(BatchNoDefault, row.Primary_Product)
                    'tbltrans_salesreturn_item_copy
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_salesreturn_item_copy(BatchNoDefault, row.Primary_Product)
                    'tbltrans_salesreturn_item_import
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_salesreturn_item_import(BatchNoDefault, row.Primary_Product)
                    'tbltrans_stock_taking_itembuffer
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating batch no control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating stock taking transaction)"
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_stock_taking_itembuffer(BatchNoDefault, row.Primary_Product)
                    'tbltrans_stock_taking_itempenyesuaian
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_stock_taking_itempenyesuaian(BatchNoDefault, row.Primary_Product)
                    'tbltrans_stock_taking_itemstart
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_stock_taking_itemstart(BatchNoDefault, row.Primary_Product)
                    'tbltrans_stock_taking_tag_item
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltrans_stock_taking_tagitem(BatchNoDefault, row.Primary_Product)
                    'tbltransaksi_stock_card
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating batch no control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating stock card)"
                    Me.DTA_tblmaster_product2.UpdateQuery_tbltransaksi_stock_card(BatchNoDefault, row.Primary_Product)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ModifyExpiry(ByVal vKontrol As Boolean, ByVal vTBL As DseMasterProduct_ChangeControl.tblmaster_product2DataTable, ByVal vStep As String, ByVal vTotalStep As String)
        Dim row As DseMasterProduct_ChangeControl.tblmaster_product2Row
        Dim ExpiryDateDefault As Date = #1/1/1950#
        Dim Lanjut As Boolean = False

        Try
            If vKontrol = False Then
                ExpiryDateDefault = #1/1/1950#
            Else
                ExpiryDateDefault = Me.dtpExpiryDefault.Value.Date
            End If

            Me.Prg1.Value = 0
            Me.Prg1.Maximum = vTBL.Rows.Count

            For Each row In vTBL.Rows
                Me.Prg1.Value += 1
                Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating expiry date control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code

                Lanjut = False
                If vKontrol = False And row.Expiry_Date_Control = 1 Then
                    Lanjut = True
                End If
                If vKontrol = True And row.Expiry_Date_Control = 0 Then
                    Lanjut = True
                End If

                If Lanjut = True Then
                    If vKontrol = False Then
                        row.Expiry_Date_Control = 0
                    Else
                        row.Expiry_Date_Control = 1
                    End If

                    'tblmaster_product
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating expiry date control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating master data)"
                    Me.DTA_tblmaster_product3.UpdateQuery_tblmaster_product(vKontrol, row.Primary_Product)
                    'tblmaster_productqty
                    Me.DTA_tblmaster_product3.UpdateQuery_tblmaster_productqty(ExpiryDateDefault, row.Primary_Product)
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_purchase_closing_item(ExpiryDateDefault, row.Primary_Product)
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating expiry date control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating stock in/out transaction)"
                    'tbltransaksi_stockinout_item
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltransaksi_stockinout_item(ExpiryDateDefault, row.Primary_Product)
                    'tbltransaksi_stockbundling_item_after
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltransaksi_stockbundling_item_after(ExpiryDateDefault, row.Primary_Product)
                    'tbltransaksi_stockbundling_item_before(row.Primary_Product)
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltransaksi_stockbundling_item_before(ExpiryDateDefault, row.Primary_Product)
                    'tbltransaksi_stocktransfer_item
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltransaksi_stocktransfer_item1(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_aop_item
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating expiry date control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating purchase transaction)"
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_aop_item(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_booking_invitem
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_booking_invitem(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_purchasereturn_item
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_purchasereturn_item(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_do_item
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating expiry date control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating sales transaction)"
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_do_item(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_salesinvoice_item
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_salesinvoice_item(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_salesinvoice_item_copy
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_salesinvoice_item_copy(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_salesinvoice_item_import
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_salesinvoice_item_import(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_salesreturn_item
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_salesreturn_item(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_salesreturn_item_copy
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_salesreturn_item_copy(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_salesreturn_item_import
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_salesreturn_item_import(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_stock_taking_itembuffer
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating expiry date control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating stock taking transaction)"
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_stock_taking_itembuffer(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_stock_taking_itempenyesuaian
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_stock_taking_itempenyesuaian(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_stock_taking_itemstart
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_stock_taking_itemstart(ExpiryDateDefault, row.Primary_Product)
                    'tbltrans_stock_taking_tag_item
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltrans_stock_taking_tagitem(ExpiryDateDefault, row.Primary_Product)
                    'tbltransaksi_stock_card
                    Me.LblProgress.Text = "Step " + vStep + " of " + vTotalStep + " - Updating expiry date control - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " - " + row.Product_Code + " (updating stock card)"
                    Me.DTA_tblmaster_product3.UpdateQuery_tbltransaksi_stock_card(ExpiryDateDefault, row.Primary_Product)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkModifyBatch_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkModifyBatch.CheckedChanged
        Try
            If Me.chkModifyBatch.Checked = True Then
                Me.chkControlBatchNo.Enabled = True
            Else
                Me.chkControlBatchNo.Checked = False
                Me.chkControlBatchNo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkModifyExpiry_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkModifyExpiry.CheckedChanged
        Try
            If Me.chkModifyExpiry.Checked = True Then
                chkControlExpiry.Enabled = True
            Else
                chkControlExpiry.Enabled = False
                Me.chkControlExpiry.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkModifyLocation_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkModifyLocation.CheckedChanged
        Try
            If Me.chkModifyLocation.Checked = True Then
                Me.chkControlLocation.Enabled = True
            Else
                Me.chkControlLocation.Enabled = False
                Me.chkControlLocation.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TimerClose_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerClose.Tick
        Try
            Me.LangsungClose = True
            Me.Close()
            Me.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TimerNote_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerNote.Tick
        Try
            Me.SplitterPosition = Me.btnStart.Height + Me.btnStart.Top + 15
            Me.TimerNote.Enabled = False
            Me.UpdateSplitterPosition = True
            Me.SplitContainer1.SplitterDistance = Me.SplitterPosition
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterProduct_ChangeControl_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        Try
            If UpdateSplitterPosition = True Then
                Me.SplitContainer1.SplitterDistance = Me.SplitterPosition
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class