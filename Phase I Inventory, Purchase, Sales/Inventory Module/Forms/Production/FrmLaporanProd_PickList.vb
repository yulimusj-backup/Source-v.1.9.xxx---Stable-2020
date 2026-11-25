Option Explicit On
Option Strict Off
Public Class FrmLaporanProd_PickList
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TblDummyList As New DseLaporanProd_PickList.Tbl_DummyListDataTable
    Public TblDetail As New DseLaporanProd_PickList.tbldetailDataTable
    Public TblProductQty As New DseLaporanProd_PickList.tblmaster_productqtyDataTable
    Public TblPriority As New DseLaporanProd_PickList.tblmaster_storage_areaDataTable
    Public TblExclude As New DseLaporanProd_PickList.tblmaster_storage_areaDataTable

    Dim DS As New DseLaporanProd_PickList
    Dim TblCetak As DseLaporanProd_PickList.tblCetakDataTable = DS.tblCetak
    Dim rowCetak As DseLaporanProd_PickList.tblCetakRow
    Dim TblCetakFilter As DseLaporanProd_PickList.tblCetak_FilterDataTable = DS.tblCetak_Filter
    Dim RowCetakFilter As DseLaporanProd_PickList.tblCetak_FilterRow

    Dim ObjCD As New Core_Module.ClsCekData
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim ObjProcess As New Core_Module.ClsMasterProses
    Dim ObjWU As New Core_Module.ClsMasterWorkUnit
    Dim ObjWUGroup As New Core_Module.ClsMasterWorkUnitGroup
    Dim ObjLoc As New Core_Module.clsMasterStorageArea
    Dim objWar As New Core_Module.clsMasterWarehouse
    Dim ObjJo As New Inventory_Module.ClsTrans_ManufacturingOrder
    Dim Obj As New ClsLaporanProd_PickList

    Dim Buffer As String = ""
    Dim Index As Integer = -1
    Dim IsTransOK As Boolean = False

    Dim listOfTransfer As New DseLaporanProd_PickList.DataTable4DataTable
    Dim rowListOfTransfer As DseLaporanProd_PickList.DataTable4Row

    'aul tambahan
    Dim vGetData As ClsLaporanProd_PickList.EnumGetData

    Private Sub FrmLaporanProd_PickList_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.Obj.Simpan_FilterSetting(Me.UserCode, Me.dgvProduct, Me.dgvProcess, Me.dgvWU, Me.dgvLocation, Me.dgvJO, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub FrmLaporanProd_PickList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvItem.Columns(Me.ColItem_CodeNoRM.Index).Frozen = True
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvProduct.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvProcess.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvWU.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvLocation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvJO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.cboData.SelectedIndex = 0
            Me.cboFormat.SelectedIndex = 0
            Me.dtpFrom.Value = Now
            Me.dtpTo.Value = Now
            Me.ColItem_Identifier.Visible = False
            Me.llbColumnSetting.Visible = False
            Me.ColItem_QtyNeed.HeaderText = "Need Qty"
            Me.Obj.IsiDgv_FilterSetting(Me.UserCode, Me.dgvProduct, Me.dgvProcess, Me.dgvWU, Me.dgvLocation, Me.dgvJO, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbLocationPriority_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbLocationPriority.LinkClicked
        Dim IJ As New FrmLaporanProd_PickList_LocPrio
        Try
            IJ.UserCode = Me.UserCode
            IJ.ShowDialog()
            If IJ.DialogResult = Windows.Forms.DialogResult.OK Then
                Me.btnShow.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            IJ.Dispose()
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim ReportDetail As New CrpLaporanProd_PickList_Detail
        Dim ReportRecap As New CrpLaporanProd_PickList_Recap
        Dim hj As New FrmCetakLaporanProd_PickList
        Dim Tmp As Integer = 0
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.dgvItem.RowCount <= 0 Then
                MsgBox("Rows count in datagrid = 0, Please click button show first.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.cboFormat.SelectedIndex = -1 Then
                MsgBox("Please change format report.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.IsikanTblCetak()
            If Me.TblCetakFilter.Rows.Count > 0 Then
                Tmp = 1
            End If

            If Me.cboFormat.SelectedIndex = 0 Then
                ReportRecap.SetDataSource(DS)
                ReportRecap.SetParameterValue(ReportRecap.Parameter_From.ParameterFieldName, Me.dtpFrom.Value)
                ReportRecap.SetParameterValue(ReportRecap.Parameter_To.ParameterFieldName, Me.dtpTo.Value)
                ReportRecap.SetParameterValue(ReportRecap.Parameter_TampilkanFilter.ParameterFieldName, Tmp)
                hj.crv1.ReportSource = ReportRecap
            Else
                ReportDetail.SetDataSource(DS)
                ReportDetail.SetParameterValue(ReportDetail.Parameter_From.ParameterFieldName, Me.dtpFrom.Value)
                ReportDetail.SetParameterValue(ReportDetail.Parameter_To.ParameterFieldName, Me.dtpTo.Value)
                ReportDetail.SetParameterValue(ReportDetail.Parameter_TampilkanFilter.ParameterFieldName, Tmp)
                hj.crv1.ReportSource = ReportDetail
            End If
            Windows.Forms.Cursor.Current = Cursors.Default
            hj.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnTransfer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTransfer.Click
        Dim FTransfer As New Inventory_Module.FrmTransaksiStockTransfer
        Dim ListTransafer As New List(Of String)
        Dim row As DseLaporanProd_PickList.Tbl_DummyListRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            ListTransafer.Clear()
            Me.listOfTransfer.Clear()

            If Me.dgvItem.RowCount <= 0 Then
                MsgBox("Rows count in datagrid = 0, Please click button show first.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            For i As Integer = 0 To Me.dgvItem.RowCount - 1
                If Me.dgvItem.Item(Me.ColItem_Transfer.Index, i).Value = True Then
                    If ListTransafer.Contains(Me.dgvItem.Item(Me.ColItem_Transfer.Index, i).Value) = False Then
                        ListTransafer.Add(Me.dgvItem.Item(Me.ColItem_Identifier.Index, i).Value)
                        'jika format nya rekap
                        If Me.cboFormat.SelectedIndex = 0 Then
                            Me.TblDummyList.DefaultView.RowFilter = "[Primary RM]=" & Me.dgvItem.Item(Me.ColItem_PrimaryProduct.Index, i).Value & _
                            " AND [Location]='" & Me.dgvItem.Item(Me.ColItem_Location.Index, i).Value & "'"
                            If Me.TblDummyList.DefaultView.Count > 0 Then
                                For Each dv As Data.DataRowView In Me.TblDummyList.DefaultView
                                    row = dv.Row
                                    If String.IsNullOrEmpty(row.Identifer) = False Then
                                        If ListTransafer.Contains(row.Identifer) = False Then
                                            ListTransafer.Add(row.Identifer)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                    Me.listOfTransfer.DefaultView.RowFilter = "[Primary RM] = " & Me.dgvItem.Item(Me.ColItem_PrimaryProduct.Index, i).Value.ToString _
                    & " and [Location To Code] = '" & Me.dgvItem.Item(Me.ColItem_Location.Index, i).Value & "'"
                    If Me.listOfTransfer.DefaultView.Count = 0 Then
                        Me.rowListOfTransfer = Me.listOfTransfer.NewDataTable4Row
                        Me.rowListOfTransfer.Primary_RM = Me.dgvItem.Item(Me.ColItem_PrimaryProduct.Index, i).Value
                        Me.rowListOfTransfer.Location_To_Code = Me.dgvItem.Item(Me.ColItem_Location.Index, i).Value
                        Me.listOfTransfer.Rows.Add(Me.rowListOfTransfer)
                    End If
                End If
            Next

            If ListTransafer.Count = 0 Then
                MsgBox("Please enter your data transfer.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            ''MsgBox("Kelompokin berdasarkan from-to yang sama")

            FTransfer.UserCode = Me.UserCode
            FTransfer.BackColorAll = Me.BackColorAll
            FTransfer.IsTransfer_PickList = True
            FTransfer.InputData_TransferPickList(sender, e, Me.IsTransOK, Me.TblDetail, ListTransafer, Me.listOfTransfer, Me.cboFormat.SelectedIndex)
            If IsTransOK = True Then
                Me.btnShow.PerformClick()
            End If
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwItem_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwItem.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwItem_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwItem.RunWorkerCompleted
        Try
            If Me.dgvItem.RowCount > 0 Then
                Me.dgvItem.Rows.RemoveAt(Me.Index)
                Me.ObjCD.SetNumber(Me.dgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwJO_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwJO.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwJO_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwJO.RunWorkerCompleted
        Try
            If Me.dgvJO.RowCount > 0 Then
                Me.dgvJO.Rows.RemoveAt(Me.Index)
                Me.ObjCD.SetNumber(Me.dgvJO)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwLocation_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwLocation.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwLocation_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwLocation.RunWorkerCompleted
        Try
            If dgvLocation.RowCount > 0 Then
                Me.dgvLocation.Rows.RemoveAt(Me.Index)
                Me.ObjCD.SetNumber(Me.dgvLocation)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwProcess_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwProcess.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwProcess_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwProcess.RunWorkerCompleted
        Try
            If Me.dgvProcess.RowCount > 0 Then
                Me.dgvProcess.Rows.RemoveAt(Me.Index)
                Me.ObjCD.SetNumber(Me.dgvProcess)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwProduct_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwProduct.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwProduct_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwProduct.RunWorkerCompleted
        Try
            If Me.dgvProduct.RowCount > 1 Then
                Me.dgvProduct.Rows.RemoveAt(Me.Index)
                Me.ObjCD.SetNumber(Me.dgvProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwWu_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwWu.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwWu_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwWu.RunWorkerCompleted
        Try
            If Me.dgvWU.RowCount > 0 Then
                Me.dgvWU.Rows.RemoveAt(Me.Index)
                Me.ObjCD.SetNumber(Me.dgvWU)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'product RM Filter--------------------------------------#
    Private Sub dgvProduct_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvProduct.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colProduct_GroupLevel.Index Or e.ColumnIndex = Me.colProduct_Code.Index Then
                If e.ColumnIndex = Me.colProduct_Code.Index Then
                    If IsNothing(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
                Me.Buffer = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProduct.CellEndEdit
        Try
            If IsNothing(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If IsNothing(Me.Buffer) = True Then
                        Me.Index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                        Exit Sub
                    End If
                    If Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colProduct_GroupLevel.Index Then
                        cellEndEditGroupLevel(sender, e)
                    End If
                    If e.ColumnIndex = Me.colProduct_Code.Index Then
                        cellEndEditCode(sender, e)
                    End If
                Else
                    If IsNothing(Me.Buffer) = False Then
                        If String.IsNullOrEmpty(Buffer) = False Then
                            Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    Me.Index = e.RowIndex
                    Me.BgwProduct.RunWorkerAsync()
                End If
            Else
                If e.ColumnIndex = Me.colProduct_Code.Index Then
                    If IsNothing(Me.Buffer) = False Then
                        Me.Index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditGroupLevel(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindGroup As New Core_Module.FrmCariMasterProductGroup
        Dim frmFindCategory As New Core_Module.frmCariMasterCategory
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            Select Case Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value
                Case "Group Level"
                    GoTo Group_
                Case "Group"
group_:
                    frmFindGroup.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 3
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.Index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                    End If
                Case "Brand"
                    frmFindGroup.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 2
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.Index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                    End If
                Case "Sub Brand"
                    frmFindGroup.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 1
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.Index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                    End If
                Case "Category"
                    frmFindCategory.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindCategory.ShowDialog()
                    If frmFindCategory.PrimaryCategory > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindCategory.PrimaryCategory
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objCategory.GetCategoryCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                Case "Product"
                    frmFindProduct.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindProduct.ShowDialog()
                    If frmFindProduct.PrimaryProduct > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindProduct.PrimaryProduct
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objProduct.GetProductCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
            End Select
            If Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexProduct(e.RowIndex, Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value, Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value & " " & Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.Index = e.RowIndex
                    Me.BgwProduct.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditCode(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindGroup As New Core_Module.FrmCariMasterProductGroup
        Dim frmFindCategory As New Core_Module.frmCariMasterCategory
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
Back_:
            Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
            Select Case Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value
                Case "Group Level"
                    GoTo Group_
                Case "Group Level"
group_:
                    If Me.objGroup.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 3) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 3)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 3
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.Index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
                Case "Brand"
                    If Me.objGroup.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 2) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 2
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.Index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
                Case "Sub Brand"
                    If Me.objGroup.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 1) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 1)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 1
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.Index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
                Case "Category"
                    If Me.objCategory.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objCategory.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindCategory.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindCategory.ShowDialog()
                        If frmFindCategory.PrimaryCategory > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindCategory.PrimaryCategory
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objCategory.GetCategoryCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.Index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
                Case "Product"
                    If Me.objProduct.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objProduct.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindProduct.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindProduct.ShowDialog()
                        If frmFindProduct.PrimaryProduct > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindProduct.PrimaryProduct
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objProduct.GetProductCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.Index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
            End Select
            If Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexProduct(e.RowIndex, Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value, Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value & " " & Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.Index = e.RowIndex
                    Me.BgwProduct.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getIndexProduct(ByVal vRowException As Integer, ByVal vGroupLevel As String, ByVal vPrimary As Double) As Integer

        Try
            For Each dr As DataGridViewRow In Me.dgvProduct.Rows
                With dr
                    If Me.dgvProduct.Rows.IndexOf(dr) <> vRowException Then
                        If dr.Cells(Me.colProduct_Primary.Index).Value <> -1 Then
                            If (dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Group Level" Or dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Group") And dr.Cells(Me.colProduct_Primary.Index).Value = vPrimary Then
                                Return Me.dgvProduct.Rows.IndexOf(dr)
                            ElseIf dr.Cells(Me.colProduct_GroupLevel.Index).Value = vGroupLevel And dr.Cells(Me.colProduct_Primary.Index).Value = vPrimary Then
                                Return Me.dgvProduct.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvProduct_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvProduct.DataError

    End Sub

    Private Sub dgvProduct_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvProduct.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvProduct.Rows(Me.dgvProduct.Rows.Count - 1)
            dr.Cells(Me.colProduct_No.Index).Value = Me.dgvProduct.Rows.Count
            dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Group Level"
            dr.Cells(Me.colProduct_Desc.Index).Value = ""
            dr.Cells(Me.colProduct_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvProduct.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvProduct)
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

    'process RM Filter--------------------------------------#
    Private Sub dgvProcess_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvProcess.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvProcess.Item(Me.colProcess_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.colProcess_Code.Index Then
                    '''simpan nilai awal di buffer
                    Me.Buffer = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.colProcess_Code.Index Then
                    Me.Buffer = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemProcess(ByVal vPrimary As Double, _
           ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvProcess.Rows.Count - 2
                RowDGV = Me.dgvProcess.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.colProcess_Primary.Index).Value = vPrimary Then
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

    Private Sub dgvProcess_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProcess.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvProcess.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryProcess As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.colProcess_Code.Index Then
                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjProcess.ApakahSudahAda(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryProcess = Me.ObjProcess.GetPrimary(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemProcess(PrimaryProcess, e.RowIndex)
                            If index < 0 Then
                                Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value = PrimaryProcess
                                Me.dgvProcess.Item(Me.colProcess_Code.Index, e.RowIndex).Value = Me.ObjProcess.GetCode(PrimaryProcess)
                                Me.dgvProcess.Item(Me.colProcess_Desc.Index, e.RowIndex).Value = Me.ObjProcess.GetDesc(PrimaryProcess)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvProcess.Item(Me.colProcess_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.Index = e.RowIndex
                                        Me.BgwProcess.RunWorkerAsync()
                                    Else
                                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvProcess)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterProsesCari
                            ij.txtCari.Text = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.mPrimary <> -1 Then
                                index = Me.GetIndexItemProcess(ij.mPrimary, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value = ij.mPrimary
                                    Me.dgvProcess.Item(Me.colProcess_Code.Index, e.RowIndex).Value = Me.ObjProcess.GetCode(ij.mPrimary)
                                    Me.dgvProcess.Item(Me.colProcess_Desc.Index, e.RowIndex).Value = Me.ObjProcess.GetDesc(ij.mPrimary)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.Index = e.RowIndex
                                            Me.BgwProcess.RunWorkerAsync()
                                        Else
                                            Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvProcess)
                                        Exit Sub
                                    Else
                                        Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value = ij.mPrimary
                                        Me.dgvProcess.Item(Me.colProcess_Code.Index, e.RowIndex).Value = Me.ObjProcess.GetCode(ij.mPrimary)
                                        Me.dgvProcess.Item(Me.colProcess_Desc.Index, e.RowIndex).Value = Me.ObjProcess.GetDesc(ij.mPrimary)
                                    End If
                                End If
                            ElseIf ij.mPrimary = -1 Then
                                If Me.dgvProcess.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.Index = e.RowIndex
                                        Me.BgwProcess.RunWorkerAsync()
                                    Else
                                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvProcess)
                                    Exit Sub
                                End If
                                If Me.dgvProcess.Item(Me.colProcess_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.Index = e.RowIndex
                                        Me.BgwProcess.RunWorkerAsync()
                                    Else
                                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvProcess)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvProcess)
                    End If
                Else
                    If e.ColumnIndex = Me.colProcess_Code.Index Then
                        If Me.dgvProcess.Rows.Count - 1 <> e.RowIndex Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.Index = e.RowIndex
                                Me.BgwProcess.RunWorkerAsync()
                            Else
                                Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Exit Sub
                        End If
                        If Me.dgvProcess.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.Index = e.RowIndex
                                Me.BgwProcess.RunWorkerAsync()
                            Else
                                Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvProcess.Columns(Me.colProcess_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.dgvProcess.Columns(Me.colProcess_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub dgvProcess_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvProcess.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvProcess.Rows(Me.dgvProcess.Rows.Count - 1)
            dr.Cells(Me.colProcess_No.Index).Value = Me.dgvProcess.Rows.Count
            dr.Cells(Me.colProcess_Code.Index).Value = ""
            dr.Cells(Me.colProcess_Desc.Index).Value = ""
            dr.Cells(Me.colProcess_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvProcess.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvProcess)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvProcess.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvProcess)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'workunit RM Filter--------------------------------------#
    Private Sub dgvWU_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvWU.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColWU_Group.Index Or e.ColumnIndex = Me.ColWU_Code.Index Then
                If e.ColumnIndex = Me.ColWU_Code.Index Then
                    If IsNothing(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
                Me.Buffer = Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWU_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWU.CellEndEdit
        Try
            If IsNothing(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If IsNothing(Me.Buffer) = True Then
                        Me.Index = e.RowIndex
                        Me.BgwWu.RunWorkerAsync()
                        Exit Sub
                    End If
                    If Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colProduct_GroupLevel.Index Then
                        cellEndEditGroupLevel_WU(sender, e)
                    End If
                    If e.ColumnIndex = Me.colProduct_Code.Index Then
                        cellEndEditCode_WU(sender, e)
                    End If
                Else
                    If IsNothing(Me.Buffer) = False Then
                        If String.IsNullOrEmpty(Buffer) = False Then
                            Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    Me.Index = e.RowIndex
                    Me.BgwWu.RunWorkerAsync()
                End If
            Else
                If e.ColumnIndex = Me.ColWU_Code.Index Then
                    If IsNothing(Me.Buffer) = False Then
                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.Index = e.RowIndex
                    Me.BgwWu.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditGroupLevel_WU(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindWUGroup As New Core_Module.FrmMasterWorkUnitGroup_Cari
        Dim frmFindWU As New Core_Module.FrmMasterWorkUnit_Cari
        Try
            Select Case Me.dgvWU.Item(Me.ColWU_Group.Index, e.RowIndex).Value
                Case "Work Unit Group Level"
                    GoTo WUGroup
                Case "Work Unit Group"
wugroup:
                    frmFindWUGroup.txtCari.Text = Me.dgvWU.Item(Me.ColWU_Code.Index, e.RowIndex).Value
                    frmFindWUGroup.ShowDialog()
                    If frmFindWUGroup.PrimaryWorkUnitGroup > 0 Then
                        Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value = frmFindWUGroup.PrimaryWorkUnitGroup
                        Me.dgvWU.Item(Me.ColWU_Code.Index, e.RowIndex).Value = ObjWUGroup.GetCode(Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value)
                        Me.dgvWU.Item(Me.ColWU_Desc.Index, e.RowIndex).Value = ObjWUGroup.GetDesc(Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.Index = e.RowIndex
                        Me.BgwWu.RunWorkerAsync()
                    End If
                Case "Work Unit"
                    frmFindWU.txtCari.Text = Me.dgvWU.Item(Me.ColWU_Code.Index, e.RowIndex).Value
                    frmFindWU.ShowDialog()
                    If frmFindWU.PrimaryWorkUnit > 0 Then
                        Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value = frmFindWU.PrimaryWorkUnit
                        Me.dgvWU.Item(Me.ColWU_Code.Index, e.RowIndex).Value = ObjWU.GetCode(Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value)
                        Me.dgvWU.Item(Me.ColWU_Desc.Index, e.RowIndex).Value = ObjWU.GetDesc(Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
            End Select
            If Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexWU(e.RowIndex, Me.dgvWU.Item(Me.ColWU_Group.Index, e.RowIndex).Value, Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvWU.Item(Me.ColWU_Group.Index, e.RowIndex).Value & " " & Me.dgvWU.Item(Me.ColWU_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.Index = e.RowIndex
                    Me.BgwWu.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditCode_WU(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindWUGroup As New Core_Module.FrmMasterWorkUnitGroup_Cari
        Dim frmFindWU As New Core_Module.FrmMasterWorkUnit_Cari
        Try
            Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
            Select Case Me.dgvWU.Item(Me.ColWU_Group.Index, e.RowIndex).Value
                Case "Work Unit Group Level"
                    GoTo WUGroup
                Case "Work Unit Group"
WUGroup:
                    If Me.ObjWUGroup.Apakahsudahada(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value = ObjWUGroup.GetPrimary(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvWU.Item(Me.ColWU_Desc.Index, e.RowIndex).Value = ObjWUGroup.GetDesc(Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindWUGroup.txtCari.Text = Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindWUGroup.ShowDialog()
                        If frmFindWUGroup.PrimaryWorkUnitGroup > 0 Then
                            Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value = frmFindWUGroup.PrimaryWorkUnitGroup
                            Me.dgvWU.Item(Me.ColWU_Code.Index, e.RowIndex).Value = ObjWUGroup.GetCode(Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value)
                            Me.dgvWU.Item(Me.ColWU_Desc.Index, e.RowIndex).Value = ObjWUGroup.GetDesc(Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.Index = e.RowIndex
                            Me.BgwWu.RunWorkerAsync()
                        End If
                    End If
                Case "Work Unit"
                    If Me.ObjWU.ApakahSudahAda(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value = ObjWU.GetPrimary(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvWU.Item(Me.ColWU_Desc.Index, e.RowIndex).Value = ObjWU.GetDesc(Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindWU.txtCari.Text = Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindWU.ShowDialog()
                        If frmFindWU.PrimaryWorkUnit > 0 Then
                            Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value = frmFindWU.PrimaryWorkUnit
                            Me.dgvWU.Item(Me.ColWU_Code.Index, e.RowIndex).Value = ObjWU.GetCode(Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value)
                            Me.dgvWU.Item(Me.ColWU_Desc.Index, e.RowIndex).Value = ObjWU.GetDesc(Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.Index = e.RowIndex
                            Me.BgwWu.RunWorkerAsync()
                        End If
                    End If
            End Select
            If Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexWU(e.RowIndex, Me.dgvWU.Item(Me.ColWU_Group.Index, e.RowIndex).Value, Me.dgvWU.Item(Me.ColWU_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvWU.Item(Me.ColWU_Group.Index, e.RowIndex).Value & " " & Me.dgvWU.Item(Me.ColWU_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.Index = e.RowIndex
                    Me.BgwWu.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getIndexWU(ByVal vRowException As Integer, ByVal vGroupLevel As String, ByVal vPrimary As Double) As Integer

        Try
            For Each dr As DataGridViewRow In Me.dgvWU.Rows
                With dr
                    If Me.dgvWU.Rows.IndexOf(dr) <> vRowException Then
                        If dr.Cells(Me.ColWU_Primary.Index).Value <> -1 Then
                            If (dr.Cells(Me.ColWU_Group.Index).Value = "Work Unit Group Level" Or dr.Cells(Me.ColWU_Group.Index).Value = "Work Unit Group") And dr.Cells(Me.ColWU_Primary.Index).Value = vPrimary Then
                                Return Me.dgvWU.Rows.IndexOf(dr)
                            ElseIf dr.Cells(Me.ColWU_Group.Index).Value = vGroupLevel And dr.Cells(Me.ColWU_Primary.Index).Value = vPrimary Then
                                Return Me.dgvWU.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvWU_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvWU.DataError

    End Sub

    Private Sub dgvWU_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvWU.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvWU.Rows(Me.dgvWU.Rows.Count - 1)
            dr.Cells(Me.ColWU_No.Index).Value = Me.dgvWU.Rows.Count
            dr.Cells(Me.ColWU_Group.Index).Value = "Work Unit Group Level"
            'dr.Cells(Me.ColWU_Code.Index).Value = ""
            dr.Cells(Me.ColWU_Desc.Index).Value = ""
            dr.Cells(Me.ColWU_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWU_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvWU.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvWU)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWU_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvWU.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvWU)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'input location RM Filter--------------------------------------#
    Private Sub dgvLocation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvLocation.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColLocation_Code.Index Or e.ColumnIndex = Me.ColLocation_Group.Index Then
                If e.ColumnIndex = Me.ColLocation_Code.Index Then
                    If IsNothing(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
                Me.Buffer = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvLocation.CellEndEdit
        Try
            If IsNothing(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If IsNothing(Me.Buffer) = True Then
                        Me.Index = e.RowIndex
                        Me.BgwLocation.RunWorkerAsync()
                        Exit Sub
                    End If
                    If Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColLocation_Group.Index Then
                        cellEndEditGroupLevel_Loc(sender, e)
                    End If
                    If e.ColumnIndex = Me.ColLocation_Code.Index Then
                        cellEndEditCode_Loc(sender, e)
                    End If
                Else
                    If Me.dgvLocation.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.Index = e.RowIndex
                            Me.BgwLocation.RunWorkerAsync()
                        Else
                            Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Exit Sub
                    End If
                    If Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.Index = e.RowIndex
                            Me.BgwLocation.RunWorkerAsync()
                        Else
                            Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    End If
                End If
            Else
                If Me.dgvLocation.Rows.Count - 1 <> e.RowIndex Then
                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                        Me.Index = e.RowIndex
                        Me.BgwLocation.RunWorkerAsync()
                    Else
                        Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                    Exit Sub
                End If
                If Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                        Me.Index = e.RowIndex
                        Me.BgwLocation.RunWorkerAsync()
                    Else
                        Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditGroupLevel_Loc(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindLocation As New Core_Module.FrmCariMasterStorageArea
        Dim frmFindWarehouse As New Core_Module.frmCariMasterWarehouse
        Try
            Select Case Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value
                Case "Location Level"
                    GoTo Loc_
                Case "Location"
loc_:
                    frmFindLocation.txtCari.Text = Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value
                    frmFindLocation.ShowDialog()
                    If frmFindLocation.PrimaryStorageArea > 0 Then
                        Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = frmFindLocation.PrimaryStorageArea
                        Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = ObjLoc.GetStorageAreaCode(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                        Me.dgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = ObjLoc.GetDescription(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.Index = e.RowIndex
                        Me.BgwLocation.RunWorkerAsync()
                        Exit Sub
                    End If
                Case "Warehouse"
                    frmFindWarehouse.txtCari.Text = Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value
                    frmFindWarehouse.ShowDialog()
                    If frmFindWarehouse.PrimaryWarehouse > 0 Then
                        Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = frmFindWarehouse.PrimaryWarehouse
                        Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = objWar.GetWarehouseCode(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                        Me.dgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = objWar.GetDescription(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.Index = e.RowIndex
                        Me.BgwLocation.RunWorkerAsync()
                        Exit Sub
                    End If
            End Select

            If Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndex(e.RowIndex, Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value, Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value & " " & Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.Index = e.RowIndex
                    Me.BgwLocation.RunWorkerAsync()
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getIndex(ByVal vRowException As Integer, ByVal vGroupLevel As String, ByVal vPrimary As Double) As Integer

        Try
            For Each dr As DataGridViewRow In Me.dgvLocation.Rows
                With dr
                    If Me.dgvLocation.Rows.IndexOf(dr) <> vRowException Then
                        If dr.Cells(Me.ColLocation_Primary.Index).Value <> -1 Then
                            If (dr.Cells(Me.ColLocation_Group.Index).Value = "Location Level" Or dr.Cells(Me.ColLocation_Group.Index).Value = "Location") And dr.Cells(Me.ColLocation_Primary.Index).Value = vPrimary Then
                                Return Me.dgvLocation.Rows.IndexOf(dr)
                            ElseIf dr.Cells(Me.ColLocation_Group.Index).Value = vGroupLevel And dr.Cells(Me.ColLocation_Primary.Index).Value = vPrimary Then
                                Return Me.dgvLocation.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub cellEndEditCode_Loc(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindLocation As New Core_Module.FrmCariMasterStorageArea
        Dim frmFindWarehouse As New Core_Module.frmCariMasterWarehouse
        Try
            Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
            Select Case Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value
                Case "Location Level"
                    GoTo Location_
                Case "Location"
Location_:
                    If Me.ObjLoc.ApakahSudahAda(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ObjLoc.GetPrimary(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = ObjLoc.GetDescription(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindLocation.txtCari.Text = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindLocation.ShowDialog()
                        If frmFindLocation.PrimaryStorageArea > 0 Then
                            Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = frmFindLocation.PrimaryStorageArea
                            Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = ObjLoc.GetStorageAreaCode(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                            Me.dgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = ObjLoc.GetDescription(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.Index = e.RowIndex
                            Me.BgwLocation.RunWorkerAsync()
                        End If
                    End If
                Case "Warehouse"
                    If Me.objWar.ApakahSudahAda(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = objWar.GetPrimary(Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = objWar.GetDescription(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindWarehouse.txtCari.Text = Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindWarehouse.ShowDialog()
                        If frmFindWarehouse.PrimaryWarehouse > 0 Then
                            Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = frmFindWarehouse.PrimaryWarehouse
                            Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = objWar.GetWarehouseCode(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                            Me.dgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = objWar.GetDescription(Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.Index = e.RowIndex
                            Me.BgwLocation.RunWorkerAsync()
                        End If
                    End If
            End Select

            If Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndex(e.RowIndex, Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value, Me.dgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvLocation.Item(Me.ColLocation_Group.Index, e.RowIndex).Value & " " & Me.dgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.Index = e.RowIndex
                    Me.BgwLocation.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvLocation.DataError

    End Sub

    Private Sub dgvLocation_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvLocation.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvLocation.Rows(Me.dgvLocation.Rows.Count - 1)
            dr.Cells(Me.ColLocation_No.Index).Value = Me.dgvLocation.Rows.Count
            dr.Cells(Me.ColLocation_Group.Index).Value = "Location Level"
            dr.Cells(Me.ColLocation_Desc.Index).Value = ""
            dr.Cells(Me.ColLocation_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvLocation.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvLocation_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvLocation.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'JO Filter------------------------------------------------------#
    Private Sub dgvJO_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvJO.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColJo_JONo.Index Then
                Me.Buffer = Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvJO_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvJO.CellEndEdit
        Dim frm As New FrmTrans_ManufacturingOrder_PrintJO_Cari
        Dim frm_PrintJO As New FrmTrans_ManufacturingOrder_PrintJO_PCK

        Try
            If IsNothing(Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.Buffer = Me.dgvJO(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColJo_JONo.Index Then
                        If Me.ObjJo.Apakahsudahada(Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Me.dgvJO.Item(Me.ColJo_Identififer.Index, e.RowIndex).Value = frm_PrintJO.GetIdentifierJo(Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvJO.Item(Me.ColJo_Date.Index, e.RowIndex).Value = Format(frm.GetDateMOByJoNo(Me.dgvJO.Item(Me.ColJo_JONo.Index, e.RowIndex).Value), "dd MMM yyyy")
                            If GetIndexItemJO(Me.dgvJO.Item(Me.ColJo_Identififer.Index, e.RowIndex).Value, e.RowIndex) <> -1 Then
                                MsgBox("Job Order " & Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                If IsNothing(Me.Buffer) = False Then
                                    If String.IsNullOrEmpty(Buffer) = False Then
                                        Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        Me.dgvJO.Item(Me.ColJo_Identififer.Index, e.RowIndex).Value = frm_PrintJO.GetIdentifierJo(Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value)
                                        Me.dgvJO.Item(Me.ColJo_Date.Index, e.RowIndex).Value = Format(frm.GetDateMOByJoNo(Me.dgvJO.Item(Me.ColJo_JONo.Index, e.RowIndex).Value), "dd MMM yyyy")
                                        Exit Sub
                                    End If
                                End If
                                Me.Index = e.RowIndex
                                Me.BgwJO.RunWorkerAsync()
                            End If
                        Else
                            frm.txtCari.Text = Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value
                            frm.ShowDialog()
                            If frm.IdentifierJO <> "" Then
                                Me.dgvJO.Item(Me.ColJo_Identififer.Index, e.RowIndex).Value = frm.IdentifierJO
                                Me.dgvJO.Item(Me.ColJo_JONo.Index, e.RowIndex).Value = Me.ObjJo.GetJONumber(Me.dgvJO.Item(Me.ColJo_Identififer.Index, e.RowIndex).Value)
                                Me.dgvJO.Item(Me.ColJo_Date.Index, e.RowIndex).Value = Format(frm.GetDateMOByJoNo(Me.dgvJO.Item(Me.ColJo_JONo.Index, e.RowIndex).Value), "dd MMM yyyy")
                                If GetIndexItemJO(Me.dgvJO.Item(Me.ColJo_Identififer.Index, e.RowIndex).Value, e.RowIndex) <> -1 Then
                                    MsgBox("Job Order " & Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                    If IsNothing(Me.Buffer) = False Then
                                        If String.IsNullOrEmpty(Buffer) = False Then
                                            Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            Me.dgvJO.Item(Me.ColJo_Identififer.Index, e.RowIndex).Value = frm_PrintJO.GetIdentifierJo(Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value)
                                            Me.dgvJO.Item(Me.ColJo_Date.Index, e.RowIndex).Value = Format(frm.GetDateMOByJoNo(Me.dgvJO.Item(Me.ColJo_JONo.Index, e.RowIndex).Value), "dd MMM yyyy")
                                            Exit Sub
                                        End If
                                    End If
                                    Me.Index = e.RowIndex
                                    Me.BgwJO.RunWorkerAsync()
                                End If
                            Else
                                If IsNothing(Me.Buffer) = False Then
                                    If String.IsNullOrEmpty(Buffer) = False Then
                                        Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        Exit Sub
                                    End If
                                End If
                                Me.Index = e.RowIndex
                                Me.BgwJO.RunWorkerAsync()
                            End If
                        End If
                    End If
                Else
                    If Me.dgvJO.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.Index = e.RowIndex
                            Me.BgwJO.RunWorkerAsync()
                        Else
                            Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Exit Sub
                    End If
                    If Me.dgvJO.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.Index = e.RowIndex
                            Me.BgwJO.RunWorkerAsync()
                        Else
                            Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    End If
                End If
            Else
                If Me.dgvJO.Rows.Count - 1 <> e.RowIndex Then
                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                        Me.Index = e.RowIndex
                        Me.BgwJO.RunWorkerAsync()
                    Else
                        Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                    Exit Sub
                End If
                If Me.dgvJO.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                        Me.Index = e.RowIndex
                        Me.BgwJO.RunWorkerAsync()
                    Else
                        Me.dgvJO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemJO(ByVal vIdentifierJO As String, ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvJO.Rows.Count - 2
                RowDGV = Me.dgvJO.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColJo_Identififer.Index).Value = vIdentifierJO Then
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

    Private Sub dgvJO_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvJO.DataError

    End Sub

    Private Sub dgvJO_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvJO.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvJO.Rows(Me.dgvJO.Rows.Count - 1)
            dr.Cells(Me.ColJo_No.Index).Value = Me.dgvJO.Rows.Count
            dr.Cells(Me.ColJo_JONo.Index).Value = ""
            dr.Cells(Me.ColJo_Date.Index).Value = ""
            dr.Cells(Me.ColJo_Identififer.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvJO_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvJO.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvJO_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvJO.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Try
            For i As Integer = 0 To Me.dgvItem.RowCount - 1
                If CDec(Me.dgvItem.Item(Me.ColItem_QtyTransfer.Index, i).Value) > 0 Then
                    Me.dgvItem.Item(Me.ColItem_Transfer.Index, i).Value = True
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Try
            For i As Integer = 0 To Me.dgvItem.RowCount - 1
                If CDec(Me.dgvItem.Item(Me.ColItem_QtyTransfer.Index, i).Value) > 0 Then
                    Me.dgvItem.Item(Me.ColItem_Transfer.Index, i).Value = False
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.dtpTo.Value < Me.dtpFrom.Value Then
                MsgBox("To date must be > with from date.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.dtpFrom.Value = CDate(Me.dtpFrom.Value.Month & "/" & Me.dtpFrom.Value.Day & "/" & Me.dtpFrom.Value.Year & " " & Me.dtpFrom.Value.Hour & ":" & Me.dtpFrom.Value.Minute & ":00")
            Me.dtpTo.Value = CDate(Me.dtpTo.Value.Month & "/" & Me.dtpTo.Value.Day & "/" & Me.dtpTo.Value.Year & " " & Me.dtpTo.Value.Hour & ":" & Me.dtpTo.Value.Minute & ":00")

            'aul tambahan
            If Me.cboData.SelectedIndex = 0 Then
                Me.vGetData = ClsLaporanProd_PickList.EnumGetData.All_Data
            ElseIf Me.cboData.SelectedIndex = 1 Then
                Me.vGetData = ClsLaporanProd_PickList.EnumGetData.RM_Only
            ElseIf Me.cboData.SelectedIndex = 2 Then
                Me.vGetData = ClsLaporanProd_PickList.EnumGetData.WIP_Only
            End If
            'aul tambahan tutup

            Me.Obj.LoadData_ToDatagrid(Me.TblDummyList, Me.TblDetail, Me.TblProductQty, Me.TblPriority, Me.TblExclude, _
            Me.dtpFrom.Value, Me.dtpTo.Value, Me.dgvItem, Me.dgvProduct, Me.dgvProcess, Me.dgvWU, Me.dgvLocation, _
            Me.dgvJO, Me, Me.ObjCD, True, True, Me.vGetData, Me.UserCode)

            If Me.dgvItem.Rows.Count = 0 Then
                MsgBox("Data not available", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColItem_Transfer.Index Then
                If CDec(Me.dgvItem.Item(Me.ColItem_QtyTransfer.Index, e.RowIndex).Value) > 0 And CDec(Me.dgvItem.Item(Me.ColItem_QtyNeed.Index, e.RowIndex).Value) > 0 Then
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    MsgBox("Product code can't transfer, value qty transfer is greater than > 0", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.ColItem_QtyPackage.Index Then
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim Frm As New FrmLaporanProd_PickList_TransDetail

        Try
            If e.ColumnIndex = Me.ColItem_TransferDetail.Index Then
                If CDec(Me.dgvItem.Item(Me.ColItem_QtyTransfer.Index, e.RowIndex).Value) = 0 Then
                    MsgBox("No transfer qty specified.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Frm.UserCode = Me.UserCode
                Frm.BackColorAll = Me.BackColorAll
                Frm.PrimaryProduct = Me.dgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value
                Frm.PrimaryLocTo = Me.dgvItem.Item(Me.ColItem_PrimaryLoc.Index, e.RowIndex).Value
                Frm.IdentifierTrans = Me.dgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value
                Frm.TblDummyList = Me.TblDummyList
                Frm.TblDetail = Me.TblDetail
                Frm.TblProductQty = Me.TblProductQty
                Frm.TblPriority = Me.TblPriority
                Frm.TblExclude = Me.TblExclude
                If Me.cboFormat.SelectedIndex = 0 Then
                    Frm.ApakahRekap = True
                End If

                If Frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.Obj.LoadData_ToDatagrid(Me.TblDummyList, Me.TblDetail, Me.TblProductQty, Me.TblPriority, _
                    Me.TblExclude, Me.dtpFrom.Value, Me.dtpTo.Value, Me.dgvItem, Me.dgvProduct, Me.dgvProcess, _
                    Me.dgvWU, Me.dgvLocation, Me.dgvJO, Me, Me.ObjCD, False, False, Me.vGetData, Me.UserCode)
                    Windows.Forms.Cursor.Current = Cursors.Default
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim QtyPackage As String = ""
        Dim Dta As New DseLaporanProd_PickListTableAdapters.tblmaster_productqtyTableAdapter
        Dim row As DseLaporanProd_PickList.Tbl_DummyListRow

        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColItem_QtyPackage.Index Then
                    If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify positive number.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    QtyPackage = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, _
                    Me.objProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value))

                    'jika qty package di ubah maka ubah semua baris atas product tsb
                    For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                        If rowD.Cells(Me.ColItem_PrimaryProduct.Index).Value = Me.dgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value Then
                            rowD.Cells(Me.ColItem_QtyPackage.Index).Value = QtyPackage
                        End If
                    Next

                    'update qty package pada master product
                    Dta.UpdateQuery_PackageQtyByPrimaryProduct(CDec(QtyPackage), Me.dgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value)
                    'update juga pada tbl_dummylist
                    If Me.TblDummyList.Rows.Count > 0 Then
                        Me.TblDummyList.DefaultView.RowFilter = "[Primary RM] = " & Me.dgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value
                        If Me.TblDummyList.DefaultView.Count > 0 Then
                            For Each dv As Data.DataRowView In Me.TblDummyList.DefaultView
                                row = dv.Row
                                row.Qty_Package = CDec(QtyPackage)
                            Next
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.cboFormat.SelectedIndex = 0 Then
                'recap
                Me.ColItem_Warehouse.Visible = False
                Me.ColItem_WU.Visible = False
                Me.ColItem_WONO.Visible = False
                Me.ColItem_JONo.Visible = False
                Me.ColItem_StartTime.Visible = False
            Else
                Me.ColItem_Warehouse.Visible = True
                Me.ColItem_WU.Visible = True
                Me.ColItem_WONO.Visible = True
                Me.ColItem_JONo.Visible = True
                Me.ColItem_StartTime.Visible = True
            End If

            If Me.dgvItem.RowCount > 0 Then
                Me.Obj.LoadData_ToDatagrid(Me.TblDummyList, Me.TblDetail, Me.TblProductQty, Me.TblPriority, _
                Me.TblExclude, Me.dtpFrom.Value, Me.dtpTo.Value, Me.dgvItem, Me.dgvProduct, Me.dgvProcess, _
                Me.dgvWU, Me.dgvLocation, Me.dgvJO, Me, Me.ObjCD, False, False, Me.vGetData, Me.UserCode)
            End If
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowNotReadyOnly_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowNotReadyOnly.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.dtpTo.Value < Me.dtpFrom.Value Then
                MsgBox("To date must be > with from date.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.Obj.LoadData_ToDatagrid(Me.TblDummyList, Me.TblDetail, Me.TblProductQty, Me.TblPriority, _
            Me.TblExclude, Me.dtpFrom.Value, Me.dtpTo.Value, Me.dgvItem, Me.dgvProduct, Me.dgvProcess, _
            Me.dgvWU, Me.dgvLocation, Me.dgvJO, Me, Me.ObjCD, False, True, Me.vGetData, Me.UserCode)
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsikanTblCetak()
        Me.TblCetak.Rows.Clear()
        Me.TblCetakFilter.Rows.Clear()
        Try
            For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                With rowDGV
                    rowCetak = TblCetak.NewtblCetakRow
                    rowCetak.RM_Code = .Cells(Me.ColItem_CodeNoRM.Index).Value
                    rowCetak.Description = .Cells(Me.ColItem_DescRM.Index).Value
                    rowCetak.Input_Location = .Cells(Me.ColItem_Location.Index).Value
                    rowCetak.Work_Unit = .Cells(Me.ColItem_WU.Index).Value
                    rowCetak.JO_No = .Cells(Me.ColItem_JONo.Index).Value
                    rowCetak.WO_No = .Cells(Me.ColItem_WONO.Index).Value
                    rowCetak.Schedule_Star = .Cells(Me.ColItem_StartTime.Index).Value
                    rowCetak.Need_Qty = .Cells(Me.ColItem_QtyNeed.Index).Value
                    rowCetak.OnSite_Qty = .Cells(Me.ColItem_QtyOnSite.Index).Value
                    rowCetak.Available_To_Transfer = .Cells(Me.ColItem_QtyTransfer.Index).Value
                    rowCetak.Minus = .Cells(Me.ColItem_QtyMinus.Index).Value
                    rowCetak.Unit = .Cells(Me.ColItem_Unit.Index).Value
                    TblCetak.AddtblCetakRow(rowCetak)
                End With
            Next

            For ij As Integer = 0 To dgvProduct.Rows.Count - 1
                If String.IsNullOrEmpty(Me.dgvProduct.Item(Me.colProduct_Code.Index, ij).Value) = False Then
                    RowCetakFilter = TblCetakFilter.NewtblCetak_FilterRow
                    RowCetakFilter.RM_Level = Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, ij).Value
                    RowCetakFilter.RM_Level = Strings.Replace(RowCetakFilter.RM_Level, "Level", "")
                    RowCetakFilter.RM_Code = Me.dgvProduct.Item(Me.colProduct_Code.Index, ij).Value
                    RowCetakFilter.RM_Desc = Me.dgvProduct.Item(Me.colProduct_Desc.Index, ij).Value
                    TblCetakFilter.AddtblCetak_FilterRow(RowCetakFilter)
                End If
            Next

            For ij As Integer = 0 To dgvProcess.Rows.Count - 1
                If String.IsNullOrEmpty(Me.dgvProcess.Item(Me.colProcess_Code.Index, ij).Value) = False Then
                    If TblCetakFilter.Rows.Count > 0 Then
                        TblCetakFilter.DefaultView.RowFilter = "[Process Code]=''"
                        If TblCetakFilter.DefaultView.Count > 0 Then
                            RowCetakFilter = TblCetakFilter.DefaultView(0).Row
                            RowCetakFilter.Process_Code = Me.dgvProcess.Item(Me.colProcess_Code.Index, ij).Value
                            RowCetakFilter.Process_Desc = Me.dgvProcess.Item(Me.colProcess_Desc.Index, ij).Value
                        Else
                            RowCetakFilter = TblCetakFilter.NewtblCetak_FilterRow
                            RowCetakFilter.Process_Code = Me.dgvProcess.Item(Me.colProcess_Code.Index, ij).Value
                            RowCetakFilter.Process_Desc = Me.dgvProcess.Item(Me.colProcess_Desc.Index, ij).Value
                            TblCetakFilter.AddtblCetak_FilterRow(RowCetakFilter)
                        End If
                    Else
                        RowCetakFilter = TblCetakFilter.NewtblCetak_FilterRow
                        RowCetakFilter.Process_Code = Me.dgvProcess.Item(Me.colProcess_Code.Index, ij).Value
                        RowCetakFilter.Process_Desc = Me.dgvProcess.Item(Me.colProcess_Desc.Index, ij).Value
                        TblCetakFilter.AddtblCetak_FilterRow(RowCetakFilter)
                    End If
                End If
            Next

            For ij As Integer = 0 To dgvWU.Rows.Count - 1
                If String.IsNullOrEmpty(Me.dgvWU.Item(Me.ColWU_Code.Index, ij).Value) = False Then
                    If TblCetakFilter.Rows.Count > 0 Then
                        TblCetakFilter.DefaultView.RowFilter = "[WU Code]=''"
                        If TblCetakFilter.DefaultView.Count > 0 Then
                            RowCetakFilter = TblCetakFilter.DefaultView(0).Row
                            RowCetakFilter.WU_Level = Me.dgvWU.Item(Me.ColWU_Group.Index, ij).Value
                            RowCetakFilter.WU_Code = Me.dgvWU.Item(Me.ColWU_Code.Index, ij).Value
                            RowCetakFilter.Wu_Desc = Me.dgvWU.Item(Me.ColWU_Desc.Index, ij).Value
                        Else
                            RowCetakFilter = TblCetakFilter.NewtblCetak_FilterRow
                            RowCetakFilter.WU_Level = Me.dgvWU.Item(Me.ColWU_Group.Index, ij).Value
                            RowCetakFilter.WU_Code = Me.dgvWU.Item(Me.ColWU_Code.Index, ij).Value
                            RowCetakFilter.Wu_Desc = Me.dgvWU.Item(Me.ColWU_Desc.Index, ij).Value
                            TblCetakFilter.AddtblCetak_FilterRow(RowCetakFilter)
                        End If
                    Else
                        RowCetakFilter = TblCetakFilter.NewtblCetak_FilterRow
                        RowCetakFilter.WU_Level = Me.dgvWU.Item(Me.ColWU_Group.Index, ij).Value
                        RowCetakFilter.WU_Code = Me.dgvWU.Item(Me.ColWU_Code.Index, ij).Value
                        RowCetakFilter.Wu_Desc = Me.dgvWU.Item(Me.ColWU_Desc.Index, ij).Value
                        TblCetakFilter.AddtblCetak_FilterRow(RowCetakFilter)
                    End If
                End If
            Next

            For ij As Integer = 0 To dgvLocation.Rows.Count - 1
                If String.IsNullOrEmpty(Me.dgvLocation.Item(Me.ColLocation_Code.Index, ij).Value) = False Then
                    If TblCetakFilter.Rows.Count > 0 Then
                        TblCetakFilter.DefaultView.RowFilter = "[Location Code]=''"
                        If TblCetakFilter.DefaultView.Count > 0 Then
                            RowCetakFilter = TblCetakFilter.DefaultView(0).Row
                            RowCetakFilter.Location_Level = Me.dgvLocation.Item(Me.ColLocation_Group.Index, ij).Value
                            RowCetakFilter.Location_Level = Strings.Replace(RowCetakFilter.Location_Level, "Level", "")
                            RowCetakFilter.Location_Code = Me.dgvLocation.Item(Me.ColLocation_Code.Index, ij).Value
                            RowCetakFilter.Location_Desc_ = Me.dgvLocation.Item(Me.ColLocation_Desc.Index, ij).Value
                        Else
                            RowCetakFilter = TblCetakFilter.NewtblCetak_FilterRow
                            RowCetakFilter.Location_Level = Me.dgvLocation.Item(Me.ColLocation_Group.Index, ij).Value
                            RowCetakFilter.Location_Code = Me.dgvLocation.Item(Me.ColLocation_Code.Index, ij).Value
                            RowCetakFilter.Location_Desc_ = Me.dgvLocation.Item(Me.ColLocation_Desc.Index, ij).Value
                            TblCetakFilter.AddtblCetak_FilterRow(RowCetakFilter)
                        End If
                    Else
                        RowCetakFilter = TblCetakFilter.NewtblCetak_FilterRow
                        RowCetakFilter.Location_Level = Me.dgvLocation.Item(Me.ColLocation_Group.Index, ij).Value
                        RowCetakFilter.Location_Code = Me.dgvLocation.Item(Me.ColLocation_Code.Index, ij).Value
                        RowCetakFilter.Location_Desc_ = Me.dgvLocation.Item(Me.ColLocation_Desc.Index, ij).Value
                        TblCetakFilter.AddtblCetak_FilterRow(RowCetakFilter)
                    End If
                End If
            Next

            For ij As Integer = 0 To dgvJO.Rows.Count - 1
                If String.IsNullOrEmpty(Me.dgvJO.Item(Me.ColJo_JONo.Index, ij).Value) = False Then
                    If TblCetakFilter.Rows.Count > 0 Then
                        TblCetakFilter.DefaultView.RowFilter = "[JO No]=''"
                        If TblCetakFilter.DefaultView.Count > 0 Then
                            RowCetakFilter = TblCetakFilter.DefaultView(0).Row
                            RowCetakFilter.JO_No = Me.dgvJO.Item(Me.ColJo_JONo.Index, ij).Value
                            RowCetakFilter.JO_Date = Me.dgvJO.Item(Me.ColJo_Date.Index, ij).Value
                        Else
                            RowCetakFilter = TblCetakFilter.NewtblCetak_FilterRow
                            RowCetakFilter.JO_No = Me.dgvJO.Item(Me.ColJo_JONo.Index, ij).Value
                            RowCetakFilter.JO_Date = Me.dgvJO.Item(Me.ColJo_Date.Index, ij).Value
                            TblCetakFilter.AddtblCetak_FilterRow(RowCetakFilter)
                        End If
                    Else
                        RowCetakFilter = TblCetakFilter.NewtblCetak_FilterRow
                        RowCetakFilter.JO_No = Me.dgvJO.Item(Me.ColJo_JONo.Index, ij).Value
                        RowCetakFilter.JO_Date = Me.dgvJO.Item(Me.ColJo_Date.Index, ij).Value
                        TblCetakFilter.AddtblCetak_FilterRow(RowCetakFilter)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim DTA As New DataSet1TableAdapters.DataTable1TableAdapter
        Dim TBL As New DataSet1.DataTable1DataTable
        Dim row As DataSet1.DataTable1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    DTA.UpdateQuery(row.primary_location_in, row.primary_mo, row.primary_work_unit)
                Next
                MsgBox("Update Selesai", MsgBoxStyle.Information, "Information")
            End If
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class