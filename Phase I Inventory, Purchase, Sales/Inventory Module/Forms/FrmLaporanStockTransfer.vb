Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class FrmLaporanStockTransfer
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""

    ''20190729
    Dim ObjUserSettings As New Core_Module.ClsTransaksiSettingUser

    Private Scope_Index As EnumScope
    Private LevelGroup As EnumLevel
    Private LevelIndex As EnumLevelIndex
    Private ShowSubTotal As EnumShowSubTotal

    Private Buffer As String = ""
    Private objGroup As New Core_Module.ClsMasterProductGroup
    Private ObjCategory As New Core_Module.ClsMasterCategory
    Private ObjProduct As New Core_Module.clsMasterProduct
    Private ObjLocation As New Core_Module.clsMasterStorageArea
    Private ObjWareHouse As New Core_Module.clsMasterWarehouse
    Private ObjCD As New Core_Module.ClsCekData

    Private EraseRowindexLocation As Integer = 0
    Private EraseRowIndexProduct As Integer = 0
    Private EraseRowIndexBatchNo As Integer = 0
    Private BufferColChk As Boolean = False
    Dim ObjLicense As New Core_Module.ClsLicense

#Region "Enum saja"
    Public Enum EnumScope
        Goods_Only
        Service_Only
        Both
    End Enum

    Public Enum EnumLevel As Integer
        Category
        SubBrand
        Brand
        ProductGroup
    End Enum

    Public Enum EnumLevelIndex As Integer
        Index01
        Index02
        Index03
    End Enum

    Public Enum EnumShowSubTotal
        No
        Yes
    End Enum

    Public Enum EnumPilih
        No
        Yes
    End Enum

    Public Enum JenisProductGroup As Integer
        Product_Group
        Brand
        Sub_Brand
        Category
        Product
    End Enum
#End Region

#Region "Dgv Location"
    Private Sub DgvLocation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvLocation.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColLocation_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.DgvLocation.Columns(Me.ColLocation_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvLocation.Columns(Me.ColLocation_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    '''simpan nilai awal di buffer
                    Me.Buffer = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    If e.ColumnIndex = Me.ColLocation_Source.Index Or e.ColumnIndex = Me.ColLocation_Destination.Index Then
                    Else
                        e.Cancel = True
                    End If
                End If
            Else
                If e.ColumnIndex = Me.ColLocation_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.DgvLocation.Columns(Me.ColLocation_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvLocation.Columns(Me.ColLocation_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.Buffer = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    If e.ColumnIndex = Me.ColLocation_Source.Index Or e.ColumnIndex = Me.ColLocation_Destination.Index Then
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvLocation_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvLocation.CellContentClick
        Try
            If e.ColumnIndex < 0 Or e.RowIndex < 0 Then
                Exit Sub
            End If
            Dim Row As DataGridViewRow = Me.DgvLocation.Rows(e.RowIndex)
            If IsNothing(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColLocation_Source.Index Or _
                e.ColumnIndex = Me.ColLocation_Destination.Index Then
                    Me.EraseRowindexLocation = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            Else
                'Batasan Tidak Boleh kedua-duanya colchk = false 
                If e.ColumnIndex = Me.ColLocation_Source.Index Then
                    If Me.BufferColChk = True Then
                        If Me.DgvLocation.Item(Me.ColLocation_Destination.Index, e.RowIndex).Value = False Then
                            Me.EraseRowindexLocation = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
                End If
                If e.ColumnIndex = Me.ColLocation_Destination.Index Then
                    If Me.BufferColChk = True Then
                        If Me.DgvLocation.Item(Me.ColLocation_Source.Index, e.RowIndex).Value = False Then
                            Me.EraseRowindexLocation = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvLocation_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvLocation.CellEnter
        Try
            If IsNothing(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColLocation_Source.Index Then
                    Me.BufferColChk = Me.DgvLocation.Item(Me.ColLocation_Source.Index, e.RowIndex).Value
                ElseIf e.ColumnIndex = Me.ColLocation_Destination.Index Then
                    Me.BufferColChk = Me.DgvLocation.Item(Me.ColLocation_Destination.Index, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvLocation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvLocation.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvLocation.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryLocation As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColLocation_Code.Index Then
                        Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ChkFilterByWareHouse.Checked = False Then
                            If Me.ObjLocation.ApakahSudahAda(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                PrimaryLocation = Me.ObjLocation.GetPrimary(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                                index = Me.GetIndexItemLocation(PrimaryLocation, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = PrimaryLocation
                                    Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(PrimaryLocation)
                                    Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(PrimaryLocation)
                                    Me.DgvLocation.Item(Me.ColLocation_Source.Index, e.RowIndex).Value = True
                                    Me.DgvLocation.Item(Me.ColLocation_Destination.Index, e.RowIndex).Value = True
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    End If
                                End If
                            Else

                                Dim ij As New Core_Module.FrmCariMasterStorageArea
                                ij.txtCari.Text = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.ShowDialog()
                                If ij.PrimaryStorageArea <> -1 Then

                                    index = Me.GetIndexItemLocation(ij.PrimaryStorageArea, e.RowIndex)
                                    If index < 0 Then
                                        Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ij.PrimaryStorageArea
                                        Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(ij.PrimaryStorageArea)
                                        Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(ij.PrimaryStorageArea)
                                        Me.DgvLocation.Item(Me.ColLocation_Source.Index, e.RowIndex).Value = True
                                        Me.DgvLocation.Item(Me.ColLocation_Destination.Index, e.RowIndex).Value = True
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowindexLocation = e.RowIndex
                                                Me.BackgroundWorker1.RunWorkerAsync()
                                            Else
                                                Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            'Me.ObjCD.SetNumber(Me.dgvLocation)
                                            Exit Sub
                                        Else
                                            Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ij.PrimaryStorageArea
                                            Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(ij.PrimaryStorageArea)
                                            Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(ij.PrimaryStorageArea)
                                            Me.DgvLocation.Item(Me.ColLocation_Source.Index, e.RowIndex).Value = True
                                            Me.DgvLocation.Item(Me.ColLocation_Destination.Index, e.RowIndex).Value = True
                                        End If
                                    End If
                                ElseIf ij.PrimaryStorageArea = -1 Then
                                    If Me.DgvLocation.Rows.Count - 1 <> e.RowIndex Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    End If
                                    If Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                            'Me.ObjCD.SetNumber(Me.dgvLocation)
                        ElseIf Me.ChkFilterByWareHouse.Checked = True Then
                            If Me.ObjWareHouse.ApakahSudahAda(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                PrimaryLocation = Me.ObjWareHouse.GetPrimary(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                                index = Me.GetIndexItemLocation(PrimaryLocation, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = PrimaryLocation
                                    Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjWareHouse.GetWarehouseCode(PrimaryLocation)
                                    Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjWareHouse.GetDescription(PrimaryLocation)
                                    Me.DgvLocation.Item(Me.ColLocation_Source.Index, e.RowIndex).Value = True
                                    Me.DgvLocation.Item(Me.ColLocation_Destination.Index, e.RowIndex).Value = True
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    End If
                                End If
                            Else

                                Dim ij As New Core_Module.frmCariMasterWarehouse
                                ij.txtCari.Text = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.ShowDialog()
                                If ij.PrimaryWarehouse <> -1 Then

                                    index = Me.GetIndexItemLocation(ij.PrimaryWarehouse, e.RowIndex)
                                    If index < 0 Then
                                        Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ij.PrimaryWarehouse
                                        Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjWareHouse.GetWarehouseCode(ij.PrimaryWarehouse)
                                        Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjWareHouse.GetDescription(ij.PrimaryWarehouse)
                                        Me.DgvLocation.Item(Me.ColLocation_Source.Index, e.RowIndex).Value = True
                                        Me.DgvLocation.Item(Me.ColLocation_Destination.Index, e.RowIndex).Value = True
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowindexLocation = e.RowIndex
                                                Me.BackgroundWorker1.RunWorkerAsync()
                                            Else
                                                Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            'Me.ObjCD.SetNumber(Me.dgvLocation)
                                            Exit Sub
                                        Else
                                            Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ij.PrimaryWarehouse
                                            Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjWareHouse.GetWarehouseCode(ij.PrimaryWarehouse)
                                            Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjWareHouse.GetDescription(ij.PrimaryWarehouse)
                                            Me.DgvLocation.Item(Me.ColLocation_Source.Index, e.RowIndex).Value = True
                                            Me.DgvLocation.Item(Me.ColLocation_Destination.Index, e.RowIndex).Value = True
                                        End If
                                    End If
                                ElseIf ij.PrimaryWarehouse = -1 Then
                                    If Me.DgvLocation.Rows.Count - 1 <> e.RowIndex Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    End If
                                    If Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                            'Me.ObjCD.SetNumber(Me.dgvLocation)
                        End If
                    End If
                Else
                    If Me.DgvLocation.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            'Me.EraseRowindexLocation = e.RowIndex
                            'Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                        Exit Sub
                    End If
                    If Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowindexLocation = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                        Exit Sub
                    End If
                End If
                'Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvLocation.Columns(Me.ColLocation_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvLocation.Columns(Me.ColLocation_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Function GetIndexItemLocation(ByVal vPrimary As Double, _
          ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.DgvLocation.Rows.Count - 2
                RowDGV = Me.DgvLocation.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColLocation_Primary.Index).Value = vPrimary Then
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
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvLocation.Rows.RemoveAt(EraseRowindexLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Dgv Batch No"
    Private Sub DgvBatchNo_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvBatchNo.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColBatch_BatchNo.Index Then
                Me.Buffer = Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvBatchNo_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvBatchNo.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvBatchNo.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim BatchNo As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColBatch_BatchNo.Index Then
                        'set ke huruf besar dan hapus spasi
                        Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim

                        index = Me.GetIndexItemBatchNo(Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value, RowIdx)
                        If index >= 0 Then
                            If e.RowIndex <> index Then
                                MsgBox("Duplication Batch No, please input another Batch No.", MsgBoxStyle.Information, "Information")
                                Me.DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, e.RowIndex).Value = ""
                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                    Me.EraseRowIndexBatchNo = e.RowIndex
                                    Me.BackgroundWorker3.RunWorkerAsync()
                                Else
                                    Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                Else

                    If Me.DgvBatchNo.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexBatchNo = e.RowIndex
                            Me.BackgroundWorker3.RunWorkerAsync()
                        Else
                            Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Exit Sub
                    End If

                    If Me.DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexBatchNo = e.RowIndex
                            Me.BackgroundWorker3.RunWorkerAsync()
                        Else
                            Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemBatchNo(ByVal vBatchNo As String, _
         ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.DgvBatchNo.Rows.Count - 2
                RowDGV = Me.DgvBatchNo.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColBatch_BatchNo.Index).Value = vBatchNo Then
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

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.DgvBatchNo.Rows.RemoveAt(EraseRowIndexBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Atur Tampilan"

    ''20190729
    Private Sub FrmLaporanStockTransfer_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim WarehouseSetting As String = ""
        Try
            If Me.ChkFilterByWareHouse.Checked = True Then
                WarehouseSetting = "1"
            ElseIf Me.ChkFilterByWareHouse.Checked = False Then
                WarehouseSetting = "0"
            End If
            Me.ObjUserSettings.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmLaporanStockTransfer, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmLaporanStockTransfer_FilterByWarehouse, WarehouseSetting, Me.UserCode)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanStockDetail2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CboPilihGroupLevel.SelectedIndex = 4
            Me.CboScope.SelectedIndex = 0
            Me.DtpMinExpiryDate.Value = Now
            Me.DtpMaxExpiryDate.Value = Now
            Me.DtpMaxExpiryDate.Enabled = True
            Me.DtpMinExpiryDate.Enabled = True
            Me.ChkShowAllItemsInTrans.Checked = False
            Me.ChkShowSubTotal.Checked = False
            Me.DgvLocation.AllowUserToAddRows = True
            Me.DgvGroup.AllowUserToAddRows = True
            Me.DgvBatchNo.AllowUserToAddRows = True
            Me.DgvLocation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvBatchNo.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.ChkFilterByWareHouse.Checked = True
            Me.TabPage1.Text = "Warehouse"

            ''20190729
            Dim WarehouseSetting As String = ""
            WarehouseSetting = Me.ObjUserSettings.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmLaporanStockTransfer, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmLaporanStockTransfer_FilterByWarehouse, Me.UserCode)
            If WarehouseSetting = "" Or WarehouseSetting = "1" Then
                Me.ChkFilterByWareHouse.Checked = True
            ElseIf WarehouseSetting = "0" Then
                Me.ChkFilterByWareHouse.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboPilihGroupLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboPilihGroupLevel.SelectedIndexChanged
        Try
            If Me.CboPilihGroupLevel.SelectedIndex = 0 Then
                LevelGroup = EnumLevel.ProductGroup
                LevelIndex = EnumLevelIndex.Index01
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 1 Then
                LevelGroup = EnumLevel.Brand
                LevelIndex = EnumLevelIndex.Index01
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 2 Then
                LevelGroup = EnumLevel.SubBrand
                LevelIndex = EnumLevelIndex.Index01
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 3 Then
                LevelGroup = EnumLevel.Category
                LevelIndex = EnumLevelIndex.Index02
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 4 Then
                LevelIndex = EnumLevelIndex.Index03
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboScope_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboScope.SelectedIndexChanged
        Try
            If Me.CboScope.SelectedIndex = 0 Then
                Me.Scope_Index = EnumScope.Goods_Only
            ElseIf Me.CboScope.SelectedIndex = 1 Then
                Me.Scope_Index = EnumScope.Service_Only
            ElseIf Me.CboScope.SelectedIndex = 2 Then
                Me.Scope_Index = EnumScope.Both
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowSubTotal_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkShowSubTotal.CheckedChanged
        Try
            If Me.ChkShowSubTotal.Checked = False Then
                Me.ShowSubTotal = EnumShowSubTotal.No
            ElseIf Me.ChkShowSubTotal.Checked = True Then
                Me.ShowSubTotal = EnumShowSubTotal.Yes
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
            'System.Environment.GetFolderPath(SpecialFolder.CommonApplicationData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Dgv Product Group"
    Private Sub DgvProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvGroup.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColGroup_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    'Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    '''simpan nilai awal di buffer
                    Me.Buffer = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColGroup_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    'Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.Buffer = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvGroup.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvGroup.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryGroup As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim _Group As Boolean = True
        Try
            With row
                If IsNothing(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColGroup_Code.Index Then
                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper

                        If Me.LevelIndex = EnumLevelIndex.Index01 Then
                            If Me.objGroup.ApakahSudahAda(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, Me.LevelGroup) = True Then
                                PrimaryGroup = Me.objGroup.GetPrimary(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, Me.LevelGroup)
                                index = Me.GetIndexItemGroup(PrimaryGroup, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = PrimaryGroup
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.objGroup.GetGroupCode(PrimaryGroup)
                                    Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.objGroup.GetGroupDesc(PrimaryGroup)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    End If
                                End If
                            Else
                                Dim ij As New Core_Module.FrmCariMasterProductGroup
                                ij.ProductGroupLevel = Me.LevelGroup
                                ij.txtCari.Text = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.ShowDialog()
                                If ij.PrimaryProductGroup <> -1 Then

                                    index = Me.GetIndexItemGroup(ij.PrimaryProductGroup, e.RowIndex)
                                    If index < 0 Then
                                        Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.objGroup.GetGroupCode(ij.PrimaryProductGroup)
                                        Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.objGroup.GetGroupDesc(ij.PrimaryProductGroup)
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowIndexProduct = e.RowIndex
                                                Me.BackgroundWorker2.RunWorkerAsync()
                                            Else
                                                Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            'Me.ObjCD.SetNumber(Me.dgvGroup)
                                            Exit Sub
                                        Else
                                            Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                                            Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.objGroup.GetGroupCode(ij.PrimaryProductGroup)
                                            Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.objGroup.GetGroupDesc(ij.PrimaryProductGroup)
                                        End If
                                    End If
                                ElseIf ij.PrimaryProductGroup = -1 Then
                                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    End If
                                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                            'Me.ObjCD.SetNumber(Me.dgvGroup)
                        ElseIf Me.LevelIndex = EnumLevelIndex.Index02 Then
                            EndEditCategory(sender, e)
                        ElseIf Me.LevelIndex = EnumLevelIndex.Index03 Then
                            EndEditProduct(sender, e)
                        End If
                    End If
                Else
                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                End If
                'Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            'Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub EndEditCategory(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.DgvGroup.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryGroup As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim LvlGroup As Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
        Try
            With row
                If IsNothing(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColGroup_Code.Index Then
                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjCategory.ApakahSudahAda(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, LvlGroup) = True Then
                            PrimaryGroup = Me.ObjCategory.GetPrimary(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemGroup(PrimaryGroup, e.RowIndex)
                            If index < 0 Then
                                Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = PrimaryGroup
                                Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(PrimaryGroup)
                                Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(PrimaryGroup)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterCategory
                            ij.txtCari.Text = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCategory <> -1 Then

                                index = Me.GetIndexItemGroup(ij.PrimaryCategory, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryCategory
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                                    Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    Else
                                        Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryCategory
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                                        Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                                    End If
                                End If
                            ElseIf ij.PrimaryCategory = -1 Then
                                If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                                If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                    End If
                Else
                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                End If
                'Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub EndEditProduct(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.DgvGroup.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryGroup As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim LvlGroup As Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
        Try
            With row
                If IsNothing(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColGroup_Code.Index Then
                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjProduct.ApakahSudahAda(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryGroup = Me.ObjProduct.GetPrimary(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemGroup(PrimaryGroup, e.RowIndex)
                            If index < 0 Then
                                Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = PrimaryGroup
                                Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(PrimaryGroup)
                                Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(PrimaryGroup)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterProduct
                            ij.txtCari.Text = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryProduct <> -1 Then

                                index = Me.GetIndexItemGroup(ij.PrimaryProduct, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryProduct
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                    Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    Else
                                        Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryProduct
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                        Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                    End If
                                End If
                            ElseIf ij.PrimaryProduct = -1 Then
                                If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                                If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                    End If
                Else
                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                End If
                'Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Function GetIndexItemGroup(ByVal vPrimary As Double, _
                    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                RowDGV = Me.DgvGroup.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColGroup_Primary.Index).Value = vPrimary Then
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

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.DgvGroup.Rows.RemoveAt(EraseRowIndexProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Get_Filter_Product"
    Private Sub GetPrimaryMinMax(ByRef vPrimaryMin_Result As Double, ByRef vPrimaryMax_Result As Double, _
        ByVal vPrimary_Pembanding As Double)
        Try
            If vPrimary_Pembanding <= vPrimaryMin_Result Then
                vPrimaryMin_Result = vPrimary_Pembanding
            End If
            If vPrimary_Pembanding >= vPrimaryMax_Result Then
                vPrimaryMax_Result = vPrimary_Pembanding
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetFilterDGVProduct(ByRef vFilter As String, _
    ByVal vEnumJenisProductGroup As FrmLaporanStockTransfer.JenisProductGroup, _
    ByVal vPrimaryParent As Double)
        Try
            If vEnumJenisProductGroup = JenisProductGroup.Category Then
                If String.IsNullOrEmpty(vFilter) = False Then
                    vFilter &= " OR [Primary Category] = " & vPrimaryParent.ToString & ""
                Else
                    vFilter = "[Primary Category] = " & vPrimaryParent.ToString & ""
                End If
            ElseIf vEnumJenisProductGroup = JenisProductGroup.Sub_Brand Then
                If String.IsNullOrEmpty(vFilter) = False Then
                    vFilter &= " OR [Primary Sub Brand] = " & vPrimaryParent.ToString & ""
                Else
                    vFilter = "[Primary Sub Brand] = " & vPrimaryParent.ToString & ""
                End If
            ElseIf vEnumJenisProductGroup = JenisProductGroup.Brand Then
                If String.IsNullOrEmpty(vFilter) = False Then
                    vFilter &= " OR [Primary Brand] = " & vPrimaryParent.ToString & ""
                Else
                    vFilter = "[Primary Brand] = " & vPrimaryParent.ToString & ""
                End If
            ElseIf vEnumJenisProductGroup = JenisProductGroup.Product_Group Then
                If String.IsNullOrEmpty(vFilter) = False Then
                    vFilter &= " OR [Primary Group] = " & vPrimaryParent.ToString & ""
                Else
                    vFilter = "[Primary Group] = " & vPrimaryParent.ToString & ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub GetFilterProduct(ByRef vParamString As String, _
    ByRef TBL_Stock As DseLaporanStockTransfer.DataTable1DataTable)
        Dim PrimaryMin As Double = Double.MaxValue
        Dim PrimaryMax As Double = Double.MinValue
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim Ketemu As Boolean = False
        Dim DTA_Group As New DseLaporanStockTransferTableAdapters.TblGroupTableAdapter
        Dim TBL_Group As New DseLaporanStockTransfer.TblGroupDataTable
        Dim Row_Group As DseLaporanStockTransfer.TblGroupRow
        Dim DTA_Brand As New DseLaporanStockTransferTableAdapters.TblBrandTableAdapter
        Dim TBL_Brand As New DseLaporanStockTransfer.TblBrandDataTable
        Dim Row_Brand As DseLaporanStockTransfer.TblBrandRow
        Dim DTA_SubBrand As New DseLaporanStockTransferTableAdapters.TblSubBrandTableAdapter
        Dim TBL_SubBrand As New DseLaporanStockTransfer.TblSubBrandDataTable
        Dim Row_SubBrand As DseLaporanStockTransfer.TblSubBrandRow
        Dim DTA_Category As New DseLaporanStockTransferTableAdapters.TblCategoryTableAdapter
        Dim TBL_Category As New DseLaporanStockTransfer.TblCategoryDataTable
        Dim Row_Category As DseLaporanStockTransfer.TblCategoryRow
        Dim Row_Stock As DseLaporanStockTransfer.DataTable1Row
        Dim ListPrimaryGroupParent As New List(Of Double)
        Try
            vParamString = ""
            If Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product_Group Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Product_Group, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Group.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Group.FillBy_PrimaryProductGroup(TBL_Group, PrimaryMax, PrimaryMin)
                    If TBL_Group.Rows.Count > 0 Then
                        TBL_Group.DefaultView.RowFilter = Filter
                        TBL_Group.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Group.DefaultView.Count - 1
                            Row_Group = TBL_Group.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Group.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Group = Row_Group.Primary_Group
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Group.Primary_Group) = False Then
                                    ListPrimaryGroupParent.Add(Row_Group.Primary_Group)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    Else
                                        vParamString = "[Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Group] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Brand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Brand.FillBy_PrimaryBrand(TBL_Brand, PrimaryMax, PrimaryMin)
                    If TBL_Brand.Rows.Count > 0 Then
                        TBL_Brand.DefaultView.RowFilter = Filter
                        TBL_Brand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Brand.DefaultView.Count - 1
                            Row_Brand = TBL_Brand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Brand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Brand = Row_Brand.Primary_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Brand.Primary_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_Brand.Primary_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Brand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Sub_Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Sub_Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_SubBrand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_SubBrand.FillBy_PrimarySubBrand(TBL_SubBrand, PrimaryMax, PrimaryMin)
                    If TBL_SubBrand.Rows.Count > 0 Then
                        TBL_SubBrand.DefaultView.RowFilter = Filter
                        TBL_SubBrand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_SubBrand.DefaultView.Count - 1
                            Row_SubBrand = TBL_SubBrand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_SubBrand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_SubBrand = Row_SubBrand.Primary_Sub_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_SubBrand.Primary_Sub_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_SubBrand.Primary_Sub_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary SubBrand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Category Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Category, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Category.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Category.FillBy_PrimaryCategory(TBL_Category, PrimaryMax, PrimaryMin)
                    'Input Primary Parent Group
                    If TBL_Category.Rows.Count > 0 Then
                        TBL_Category.DefaultView.RowFilter = Filter
                        TBL_Category.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Category.DefaultView.Count - 1
                            Row_Category = TBL_Category.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Category.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Category = Row_Category.Primary_Category
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Category.Primary_Category) = False Then
                                    ListPrimaryGroupParent.Add(Row_Category.Primary_Category)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    Else
                                        vParamString = "[Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Category] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        If String.IsNullOrEmpty(vParamString) = False Then
                            vParamString &= " OR [Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        Else
                            vParamString = "[Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        End If
                    Next
                End If
            End If

            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
                TBL_Stock.DefaultView.RowFilter = Nothing
                TBL_Stock.DefaultView.Sort = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Group.Dispose()
            TBL_Group.Dispose()
            DTA_Brand.Dispose()
            TBL_Brand.Dispose()
            DTA_SubBrand.Dispose()
            TBL_SubBrand.Dispose()
            DTA_Category.Dispose()
            TBL_Category.Dispose()
            ListPrimaryGroupParent = Nothing
        End Try
    End Sub
#End Region

#Region "Get_Filter_BatchNo"
    Public Sub GetFilterBatchNo(ByRef vParamString As String)
        Dim RowDGV As DataGridViewRow
        Try
            vParamString = ""
            If Me.DgvBatchNo.Rows.Count > 1 Then
                For ij As Integer = 0 To Me.DgvBatchNo.Rows.Count - 2
                    RowDGV = Me.DgvBatchNo.Rows(ij)
                    If String.IsNullOrEmpty(vParamString) = False Then
                        vParamString &= " OR [Batch No] LIKE '" & RowDGV.Cells(Me.ColBatch_BatchNo.Index).Value.ToString & "%'"
                    Else
                        vParamString = "[Batch No] LIKE '" & RowDGV.Cells(Me.ColBatch_BatchNo.Index).Value.ToString & "%'"
                    End If
                Next
            End If

            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Get_Filter_Location"
    Private Function GetLocationAndDestination( _
       ByVal vLocationBool As Boolean, ByVal vDestinationBool As Boolean, _
       ByVal vPrimaryLocation As Double) As String
        Dim Retval As String = ""
        Try
            If vLocationBool = True And vDestinationBool = True Then
                Retval = "([Primary Storage Area] = " & vPrimaryLocation.ToString & " OR [primary_destination_area] = " & vPrimaryLocation.ToString & ")"
            ElseIf vLocationBool = True And vDestinationBool = False Then
                Retval = "[Primary Storage Area] = " & vPrimaryLocation.ToString & ""
            ElseIf vLocationBool = False And vDestinationBool = True Then
                Retval = "[primary_destination_area] = " & vPrimaryLocation.ToString & ""
            ElseIf vLocationBool = False And vDestinationBool = False Then
                Retval = "([Primary Storage Area] <> " & vPrimaryLocation.ToString & " AND [primary_destination_area] <> " & vPrimaryLocation.ToString & ")"
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub GetFilterLocation(ByRef vParamString As String)
        Dim RowDGV As DataGridViewRow
        Try
            vParamString = ""
            If Me.DgvLocation.Rows.Count > 1 Then
                For ij As Integer = 0 To Me.DgvLocation.Rows.Count - 2
                    RowDGV = Me.DgvLocation.Rows(ij)
                    If String.IsNullOrEmpty(vParamString) = False Then
                        vParamString &= " OR " & Me.GetLocationAndDestination(CBool(RowDGV.Cells(Me.ColLocation_Source.Index).Value), _
                        CBool(RowDGV.Cells(Me.ColLocation_Destination.Index).Value), CDbl(RowDGV.Cells(Me.ColLocation_Primary.Index).Value))
                    Else
                        vParamString = Me.GetLocationAndDestination(CBool(RowDGV.Cells(Me.ColLocation_Source.Index).Value), _
                        CBool(RowDGV.Cells(Me.ColLocation_Destination.Index).Value), CDbl(RowDGV.Cells(Me.ColLocation_Primary.Index).Value))
                    End If
                Next
            End If

            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Get_Filter_GoodORService"
    Public Sub GetFilterGoodsORService(ByRef vParamString As String)
        Try
            vParamString = ""
            If Me.CboScope.SelectedIndex = FrmLaporanStockTransfer.EnumScope.Goods_Only Then
                vParamString = "[GoodsService] = 0"
            ElseIf CboScope.SelectedIndex = FrmLaporanStockTransfer.EnumScope.Service_Only Then
                vParamString = "[GoodsService] = 1"
            ElseIf CboScope.SelectedIndex = FrmLaporanStockTransfer.EnumScope.Both Then
                vParamString = "[GoodsService] = 0 OR [GoodsService] = 1"
            End If
            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    Public Function GetOverallFilter(ByRef TBLStock As DseLaporanStockTransfer.DataTable1DataTable) As String
        Dim Retval As String = ""
        Dim ProductStr As String = ""
        Dim BatchNoStr As String = ""
        Dim LocationStr As String = ""
        Dim GoodsStr As String = ""
        Try
            Me.GetFilterProduct(ProductStr, TBLStock)
            Me.GetFilterBatchNo(BatchNoStr)
            Me.GetFilterLocation(LocationStr)
            Me.GetFilterGoodsORService(GoodsStr)
            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(ProductStr) = False Then
                    Retval &= " AND " & ProductStr
                End If
            Else
                Retval = ProductStr
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(BatchNoStr) = False Then
                    Retval &= " AND " & BatchNoStr
                End If
            Else
                Retval = BatchNoStr
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(LocationStr) = False Then
                    Retval &= " AND " & LocationStr
                End If
            Else
                Retval = LocationStr
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(GoodsStr) = False Then
                    Retval &= " AND " & GoodsStr
                End If
            Else
                Retval = GoodsStr
            End If

            'Return Value
            If Retval.StartsWith(" AND ") = True Then
                Retval = Retval.Remove(0, 4)
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub InputTBLSubTotal(ByRef TBLSubTotal As DseLaporanStockTransfer.TblTotalDataTable, _
    ByVal RowStock As DseLaporanStockTransfer.DataTable1Row)
        Dim RowSubTotal As DseLaporanStockTransfer.TblTotalRow
        Dim Filter As String = ""
        Try
            Filter = "[Primary Transaksi] = " & RowStock.Primary_Transaksi.ToString & " AND [Primary Unit] = " & RowStock.Primary_Units.ToString & ""
            TBLSubTotal.DefaultView.RowFilter = Filter
            If TBLSubTotal.DefaultView.Count > 0 Then
                RowSubTotal = TBLSubTotal.DefaultView(0).Row
                RowSubTotal.Total_Qty += RowStock.Qty
            Else
                RowSubTotal = TBLSubTotal.NewRow
                RowSubTotal.Primary_Transaksi = RowStock.Primary_Transaksi
                RowSubTotal.Total_Qty = RowStock.Qty
                RowSubTotal._Decimal = RowStock._Decimal
                RowSubTotal.Primary_Unit = RowStock.Primary_Units
                RowSubTotal.Unit = RowStock.Unit
                TBLSubTotal.Rows.Add(RowSubTotal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock(ByRef TBLSort As DseLaporanStockTransfer.TblCetakDataTable, _
    ByVal TBLDummy As DseLaporanStockTransfer.DataTable1DataTable, ByVal vPrimaryStock As Double, _
    ByRef TBLSubTotal As DseLaporanStockTransfer.TblTotalDataTable)
        Dim RowSort As DseLaporanStockTransfer.TblCetakRow
        Dim RowDummy As DseLaporanStockTransfer.DataTable1Row
        Dim Filter As String = ""
        Dim ParamStr As String = ""
        Try
            'Masukan sisa item yang ada di TBLStock Per Primary Stock
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = ""
            '1. Dapatkan primary auto number dari tbl cetak
            TBLSort.DefaultView.RowFilter = "[Primary Stock] = " & vPrimaryStock.ToString & ""
            If TBLSort.DefaultView.Count > 0 Then
                For ij As Integer = 0 To TBLSort.DefaultView.Count - 1
                    RowSort = TBLSort.DefaultView(ij).Row
                    If String.IsNullOrEmpty(Filter) = False Then
                        Filter &= " AND [Primary] <> " & RowSort.Primary_AutoNumber.ToString & ""
                    Else
                        Filter = "[Primary] <> " & RowSort.Primary_AutoNumber.ToString & ""
                    End If
                Next
                '2. Input sisa data item ke TBL Sort (Cetak)
                If String.IsNullOrEmpty(Filter) = False Then
                    TBLDummy.DefaultView.RowFilter = ""
                    TBLDummy.DefaultView.Sort = ""
                    ParamStr = "[Primary Transaksi] = " & vPrimaryStock.ToString & " AND " & Filter & ""
                    TBLDummy.DefaultView.RowFilter = ParamStr
                    If TBLDummy.DefaultView.Count > 0 Then
                        For ij As Integer = 0 To TBLDummy.DefaultView.Count - 1
                            RowDummy = TBLDummy.DefaultView(ij).Row
                            RowSort = TBLSort.NewTblCetakRow
                            With RowSort
                                .Primary_AutoNumber = RowDummy.Primary
                                .Primary_Stock = RowDummy.Primary_Transaksi
                                .Primary_Product = RowDummy.Primary_Product
                                ._Date = RowDummy._Date
                                .Document_No = RowDummy.Document_No
                                .Remarks = RowDummy.Remarks
                                If String.IsNullOrEmpty(.Remarks) = True Then
                                    .Remarks = "-"
                                End If
                                .Code_No = RowDummy.Product_Code
                                .Description = RowDummy.Description
                                ._Decimal = RowDummy._Decimal
                                .Primary_Unit = RowDummy.Primary_Units
                                .Unit = RowDummy.Unit
                                If String.IsNullOrEmpty(.Unit) = True Then
                                    .Unit = "-"
                                End If
                                .Qty = RowDummy.Qty
                                .Batch_No = RowDummy.Batch_No
                                If String.IsNullOrEmpty(.Batch_No) = True Then
                                    .Batch_No = "-"
                                End If
                                .Expiry_Date = Format(RowDummy.Expiry_Date, "dd MMM yyyy")
                                If CDate(.Expiry_Date) = #1/1/1950# Then
                                    .Expiry_Date = "-"
                                End If
                                .Location = RowDummy.Storage_Code
                                If String.IsNullOrEmpty(.Location) = True Then
                                    .Location = "-"
                                End If
                                .Destination = RowDummy.Destination_Code
                                If String.IsNullOrEmpty(.Destination) = True Then
                                    .Destination = "-"
                                End If
                                .Rounding = .Rounding - RowDummy._Decimal
                                .SubTotal = ""
                            End With
                            TBLSort.Rows.Add(RowSort)
                            'Input Sub Total
                            Me.InputTBLSubTotal(TBLSubTotal, RowDummy)
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinDataWarehouse(ByRef TBLStock As DseLaporanStockTransfer.DataTable1DataTable, _
    ByRef TBLDummy As DseLaporanStockTransfer.DataTable1DataTable, _
    ByVal TBLWH As DseLaporanStockTransfer.DataTable2DataTable)
        Dim RowStock As DseLaporanStockTransfer.DataTable1Row
        Dim RowWH As DseLaporanStockTransfer.DataTable2Row
        Try
            TBLStock.Rows.Clear()
            TBLDummy.Rows.Clear()
            If TBLWH.Rows.Count > 0 Then
                For Each RowWH In TBLWH.Rows
                    'Stock Data
                    RowStock = TBLStock.NewDataTable1Row
                    RowStock.Primary = Me.GetPrimaryMax_TBLQtySimulasi(TBLStock) + 1
                    RowStock.Primary_Transaksi = RowWH.Primary_Transaksi
                    RowStock.Batch_No = RowWH.Batch_No
                    RowStock.Description = RowWH.Description
                    RowStock.Destination_Code = RowWH.Destination_Code
                    RowStock.Document_No = RowWH.Document_No
                    RowStock.Expiry_Date = RowWH.Expiry_Date
                    RowStock.GoodsService = RowWH.GoodsService
                    RowStock.Primary_Brand = RowWH.Primary_Brand
                    RowStock.Primary_Category = RowWH.Primary_Category
                    RowStock.primary_destination_area = RowWH.primary_destination_area
                    RowStock.Primary_Group = RowWH.Primary_Group
                    RowStock.Primary_Product = RowWH.Primary_Product
                    RowStock.Primary_Storage_Area = RowWH.Primary_Storage_Area
                    RowStock.Primary_SubBrand = RowWH.Primary_SubBrand
                    RowStock.Primary_Units = RowWH.Primary_Units
                    RowStock.Product_Code = RowWH.Product_Code
                    RowStock.Qty = RowWH.Qty
                    RowStock.Remarks = RowWH.Remarks
                    RowStock.Storage_Code = RowWH.Storage_Code
                    RowStock.Unit = RowWH.Unit
                    RowStock._Date = RowWH._Date
                    RowStock._Decimal = RowWH._Decimal
                    TBLStock.Rows.Add(RowStock)
                    'Dummy Data
                    RowStock = TBLDummy.NewDataTable1Row
                    RowStock.Primary = Me.GetPrimaryMax_TBLQtySimulasi(TBLDummy) + 1
                    RowStock.Primary_Transaksi = RowWH.Primary_Transaksi
                    RowStock.Batch_No = RowWH.Batch_No
                    RowStock.Description = RowWH.Description
                    RowStock.Destination_Code = RowWH.Destination_Code
                    RowStock.Document_No = RowWH.Document_No
                    RowStock.Expiry_Date = RowWH.Expiry_Date
                    RowStock.GoodsService = RowWH.GoodsService
                    RowStock.Primary_Brand = RowWH.Primary_Brand
                    RowStock.Primary_Category = RowWH.Primary_Category
                    RowStock.primary_destination_area = RowWH.primary_destination_area
                    RowStock.Primary_Group = RowWH.Primary_Group
                    RowStock.Primary_Product = RowWH.Primary_Product
                    RowStock.Primary_Storage_Area = RowWH.Primary_Storage_Area
                    RowStock.Primary_SubBrand = RowWH.Primary_SubBrand
                    RowStock.Primary_Units = RowWH.Primary_Units
                    RowStock.Product_Code = RowWH.Product_Code
                    RowStock.Qty = RowWH.Qty
                    RowStock.Remarks = RowWH.Remarks
                    RowStock.Storage_Code = RowWH.Storage_Code
                    RowStock.Unit = RowWH.Unit
                    RowStock._Date = RowWH._Date
                    RowStock._Decimal = RowWH._Decimal
                    TBLDummy.Rows.Add(RowStock)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax_TBLQtySimulasi(ByVal TBL As _
    DseLaporanStockTransfer.DataTable1DataTable) As Double
        Dim Sort As String = "[Primary] DESC"
        Dim row As DseLaporanStockTransfer.DataTable1Row
        Try
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = Sort
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                Return row.Primary
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub Cetak(ByVal vEndDate As Date, ByVal vStartDate As Date)
        Dim Report As New CrpLaporanStockTransfer_SMESS
        Dim TampilkanSubTotal As Double = 0
        Dim ListPrimaryStock As New List(Of Double)
        Dim PrimaryStock As Double = -1
        Dim SubTotalStr As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DS As New DseLaporanStockTransfer
        Dim TBL As DseLaporanStockTransfer.TblCetakDataTable = DS.TblCetak
        Dim TBLSort As New DseLaporanStockTransfer.TblCetakDataTable
        Dim Row As DseLaporanStockTransfer.TblCetakRow
        Dim RowSort As DseLaporanStockTransfer.TblCetakRow
        'Data Stock pake Location
        Dim DTAStock As New DseLaporanStockTransferTableAdapters.DataTable1TableAdapter
        Dim TBLStock As New DseLaporanStockTransfer.DataTable1DataTable
        Dim RowStock As DseLaporanStockTransfer.DataTable1Row
        Dim TBLDummy As New DseLaporanStockTransfer.DataTable1DataTable
        'Data Stock pake Warehouse
        Dim DTAStock_WH As New DseLaporanStockTransferTableAdapters.DataTable2TableAdapter
        Dim TBLStock_WH As New DseLaporanStockTransfer.DataTable2DataTable
        '---
        Dim TBLSubTotal As New DseLaporanStockTransfer.TblTotalDataTable
        Dim RowSubTotal As DseLaporanStockTransfer.TblTotalRow
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTAStock.Connection.ConnectionString = Core_Module.GetConstring
            DTAStock_WH.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ChkFilterByWareHouse.Checked = False Then
                DTAStock.FillBy_TanggalDariSampai(TBLStock, vEndDate, vStartDate)
                DTAStock.FillBy_TanggalDariSampai(TBLDummy, vEndDate, vStartDate)
            Else
                DTAStock_WH.FillBy_TanggalDariSampai(TBLStock_WH, vEndDate, vStartDate)
                Me.SalinDataWarehouse(TBLStock, TBLDummy, TBLStock_WH)
            End If
            'TampilkanSubTotal
            If Me.ChkShowSubTotal.Checked = True Then
                TampilkanSubTotal = 1
            End If
            'Hitung Stock
            If TBLStock.Rows.Count > 0 Then
                TBLStock.DefaultView.RowFilter = Nothing
                TBLStock.DefaultView.Sort = Nothing
                TBLStock.DefaultView.RowFilter = Me.GetOverallFilter(TBLStock)
                TBLStock.DefaultView.Sort = "[Date] ASC, [Document No] ASC, [Product Code] ASC"
                If TBLStock.DefaultView.Count > 0 Then
                    For ij As Integer = 0 To TBLStock.DefaultView.Count - 1
                        RowStock = TBLStock.DefaultView(ij).Row
                        'List Primary Transaksi Stock Transfer
                        If ListPrimaryStock.Contains(RowStock.Primary_Transaksi) = False Then
                            ListPrimaryStock.Add(RowStock.Primary_Transaksi)
                        End If
                        'add DS_Cetak
                        RowSort = TBLSort.NewTblCetakRow
                        With RowSort
                            .Primary_AutoNumber = RowStock.Primary
                            .Primary_Stock = RowStock.Primary_Transaksi
                            .Primary_Product = RowStock.Primary_Product
                            ._Date = RowStock._Date
                            .Document_No = RowStock.Document_No
                            .Remarks = RowStock.Remarks
                            If String.IsNullOrEmpty(.Remarks) = True Then
                                .Remarks = "-"
                            End If
                            .Code_No = RowStock.Product_Code
                            .Description = RowStock.Description
                            ._Decimal = RowStock._Decimal
                            .Primary_Unit = RowStock.Primary_Units
                            .Unit = RowStock.Unit
                            If String.IsNullOrEmpty(.Unit) = True Then
                                .Unit = "-"
                            End If
                            .Qty = RowStock.Qty
                            .Batch_No = RowStock.Batch_No
                            If String.IsNullOrEmpty(.Batch_No) = True Then
                                .Batch_No = "-"
                            End If
                            .Expiry_Date = Format(RowStock.Expiry_Date, "dd MMM yyyy")
                            If CDate(.Expiry_Date) = #1/1/1950# Then
                                .Expiry_Date = "-"
                            End If
                            .Location = RowStock.Storage_Code
                            If String.IsNullOrEmpty(.Location) = True Then
                                .Location = "-"
                            End If
                            .Destination = RowStock.Destination_Code
                            If String.IsNullOrEmpty(.Destination) = True Then
                                .Destination = "-"
                            End If
                            .Rounding = .Rounding - RowStock._Decimal
                            .SubTotal = ""
                        End With
                        TBLSort.Rows.Add(RowSort)
                        'Input Sub Total
                        Me.InputTBLSubTotal(TBLSubTotal, RowStock)
                    Next
                End If
                'All Items
                If Me.ChkShowAllItemsInTrans.Checked = True Then
                    TBLDummy.DefaultView.RowFilter = ""
                    TBLDummy.DefaultView.Sort = ""
                    For Each PrimaryStock In ListPrimaryStock
                        Me.GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock(TBLSort, TBLDummy, PrimaryStock, TBLSubTotal)
                    Next
                End If
                'Sub Total
                If TampilkanSubTotal = 1 Then
                    If TBLSubTotal.Rows.Count > 0 Then
                        For Each RowSort In TBLSort.Rows
                            SubTotalStr = ""
                            TBLSubTotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSort.Primary_Stock.ToString & ""
                            If TBLSubTotal.DefaultView.Count > 0 Then
                                For ij As Integer = 0 To TBLSubTotal.DefaultView.Count - 1
                                    RowSubTotal = TBLSubTotal.DefaultView(ij).Row
                                    If String.IsNullOrEmpty(SubTotalStr) = False Then
                                        SubTotalStr &= ", " & ObjCD.DigitGrouping(RowSubTotal.Total_Qty, RowSubTotal._Decimal) & " " & RowSubTotal.Unit
                                    Else
                                        SubTotalStr = ObjCD.DigitGrouping(RowSubTotal.Total_Qty, RowSubTotal._Decimal) & " " & RowSubTotal.Unit
                                    End If
                                Next
                                RowSort.SubTotal = SubTotalStr
                            End If
                        Next
                    End If
                End If
                'Sorting 
                TBLSort.DefaultView.RowFilter = ""
                TBLSort.DefaultView.Sort = "[Date] ASC, [Document No] ASC, [Code No] ASC"
                If TBLSort.DefaultView.Count > 0 Then
                    For ij As Integer = 0 To TBLSort.DefaultView.Count - 1
                        RowSort = TBLSort.DefaultView(ij).Row
                        Row = TBL.NewTblCetakRow
                        For ik As Integer = 0 To TBL.Columns.Count - 1
                            Row.Item(ik) = RowSort.Item(ik)
                        Next
                        TBL.Rows.Add(Row)
                    Next
                End If
            End If

            'Jika Report Kosong berikan Message Box
            If TBL.Rows.Count <= 0 Then
                Windows.Forms.Cursor.Current = Cursors.Arrow
                MsgBox("There is no data.", MsgBoxStyle.Information, "Information")
                Me.crv1.ReportSource = Nothing
                Exit Sub
            End If

            'Report Source
            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_TampilkanSubTotal.ParameterFieldName, TampilkanSubTotal)
            Report.SetParameterValue(Report.Parameter_StartDate.ParameterFieldName, Format(vStartDate, "dd MMM yyyy"))
            Report.SetParameterValue(Report.Parameter_EndDate.ParameterFieldName, Format(vEndDate, "dd MMM yyyy"))
            If Me.ChkFilterByWareHouse.Checked = False Then
                Report.SetParameterValue(Report.Parameter_Location_Warehouse.ParameterFieldName, "Location")
            Else
                Report.SetParameterValue(Report.Parameter_Location_Warehouse.ParameterFieldName, "Warehouse")
            End If
            Me.crv1.ReportSource = Report
            Windows.Forms.Cursor.Current = Cursors.Arrow
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAStock.Dispose()
            TBLStock.Dispose()
            TBLDummy.Dispose()
            ListPrimaryStock = Nothing
            ObjCD = Nothing
            TBLSort.Dispose()
            TBLSubTotal.Dispose()
            DTAStock_WH.Dispose()
            TBLStock_WH.Dispose()
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.DtpMaxExpiryDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpMaxExpiryDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.DtpMinExpiryDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpMinExpiryDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If


            If Me.DtpMinExpiryDate.Value > Me.DtpMaxExpiryDate.Value Then
                MsgBox("Period is not valid, please check again.", MsgBoxStyle.Information, "Information")
                Me.crv1.ReportSource = Nothing
                Exit Sub
            End If
            Me.Cetak(Me.DtpMaxExpiryDate.Value.Date, Me.DtpMinExpiryDate.Value.Date)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkFilterByWareHouse_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkFilterByWareHouse.CheckedChanged
        Try
            If Me.ChkFilterByWareHouse.Checked = True Then
                Me.TabPage1.Text = "Warehouse"
                Me.DgvLocation.Rows.Clear()
            Else
                Me.TabPage1.Text = "Location"
                Me.DgvLocation.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class