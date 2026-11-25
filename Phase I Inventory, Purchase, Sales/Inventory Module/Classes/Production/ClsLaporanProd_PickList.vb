Option Explicit On
Option Strict Off

Public Class ClsLaporanProd_PickList
    Dim TblSCH As New DseLaporanProd_PickList.tbltrans_mo_item_wo_item_schDataTable
    Dim TblBOM As New DseLaporanProd_PickList.DataTable1DataTable
    'Dim TblRM As New DseLaporanProd_PickList.tbltrans_mo_item_rmDataTable

    Public Enum EnumGetData As Integer
        All_Data
        RM_Only
        WIP_Only
    End Enum

    Public Enum EnumGetQty_Transfer As Integer
        Proses_1
        Proses_2
        Proses_3
        Proses_4 'by transfer detail
        Proses_5 'by transfer detail
    End Enum

    Public Function LoadData1(ByRef TblDummyList As DseLaporanProd_PickList.Tbl_DummyListDataTable, _
    ByRef TblDetail As DseLaporanProd_PickList.tbldetailDataTable, _
    ByRef vTblProductQty As DseLaporanProd_PickList.tblmaster_productqtyDataTable, _
    ByRef TblPriority As DseLaporanProd_PickList.tblmaster_storage_areaDataTable, _
    ByRef TblExclude As DseLaporanProd_PickList.tblmaster_storage_areaDataTable, _
    ByVal vFrom As Date, ByVal vTo As Date, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal hj As FrmLaporanProd_PickList, ByVal dgvWU As DataGridView, _
    ByVal dgvJO As DataGridView, ByVal dgvProduct As DataGridView, _
    ByVal dgvProcess As DataGridView, ByVal dgvLocation As DataGridView, _
    ByVal vShowPopUp As Boolean, ByVal vGetData As EnumGetData, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseLaporanProd_PickListTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim DTA1 As New DseLaporanProd_PickListTableAdapters.DataTable1TableAdapter
        Dim RowSCH As DseLaporanProd_PickList.tbltrans_mo_item_wo_item_schRow
        Dim RowDummyList As DseLaporanProd_PickList.Tbl_DummyListRow
        Dim DtaProdQty As New DseLaporanProd_PickListTableAdapters.tblmaster_productqtyTableAdapter
        Dim TblProductQty_dum As New DseLaporanProd_PickList.tblmaster_productqtyDataTable
        Dim TblDummyList_2 As New DseLaporanProd_PickList.Tbl_DummyListDataTable

        'aul tambahan
        'get data WIP
        Dim DtaWip As New DseLaporanProd_PickListTableAdapters.get_wip_dataTableAdapter
        Dim TblWip As New DseLaporanProd_PickList.get_wip_dataDataTable
        Dim RowWip As DseLaporanProd_PickList.get_wip_dataRow
        'get data RM
        Dim DtaRm As New DseLaporanProd_PickListTableAdapters.get_rm_dataTableAdapter
        Dim TblRm As New DseLaporanProd_PickList.get_rm_dataDataTable
        Dim RowRm As DseLaporanProd_PickList.get_rm_dataRow

        Dim ListCodeNo As New List(Of String)
        Dim PrimaryMin As Double = -1
        Dim PrimaryMax As Double = -1
        Dim vQty As Decimal = 0

        Try
            TblDummyList.Rows.Clear()
            TblDetail.Clear()
            vTblProductQty.Clear()
            TblPriority.Clear()
            TblExclude.Clear()

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Fill(TblSCH, vFrom, vTo) > 0 Then
                TblSCH.DefaultView.RowFilter = Nothing
                TblSCH.DefaultView.Sort = "[primary mo] ASC"
                RowSCH = TblSCH.DefaultView(0).Row
                PrimaryMin = RowSCH.primary_mo
                TblSCH.DefaultView.Sort = "[primary mo] DESC"
                RowSCH = TblSCH.DefaultView(0).Row
                PrimaryMax = RowSCH.primary_mo

                'aul tambahan
                If vGetData = EnumGetData.All_Data Or vGetData = EnumGetData.WIP_Only Then
                    DtaWip.Connection.ConnectionString = Core_Module.GetConstring
                    DtaWip.FillBy_PrimaryMaxMin(TblWip, PrimaryMax, PrimaryMin)
                End If
                If vGetData = EnumGetData.All_Data Or vGetData = EnumGetData.RM_Only Then
                    DtaRm.Connection.ConnectionString = Core_Module.GetConstring
                    DtaRm.Fill(TblRm, PrimaryMax, PrimaryMin)
                End If

                'Isi tbldummylist
                For i As Integer = 0 To Me.TblSCH.Rows.Count - 1
                    RowSCH = Me.TblSCH.Rows(i)
                    With RowSCH
                        If TblWip.Count > 0 Then
                            TblWip.DefaultView.Sort = "[Primary Product] ASC"
                            TblWip.DefaultView.RowFilter = "[Identifier JO]='" & .identifier_jo.ToString & "' AND [Identifier WO]='" & .identifier_wo.ToString & "' AND [Primary WU] = " & .primary_work_unit
                            If TblWip.DefaultView.Count > 0 Then
                                For ij As Integer = 0 To TblWip.DefaultView.Count - 1
                                    RowWip = TblWip.DefaultView(ij).Row
                                    ''get primary product min max 
                                    'If PrimaryMin = -1 Or PrimaryMax = -1 Then
                                    '    PrimaryMin = RowWip.Primary_Product
                                    '    PrimaryMax = RowWip.Primary_Product
                                    'Else
                                    '    If RowWip.Primary_Product < PrimaryMin Then
                                    '        PrimaryMin = RowWip.Primary_Product
                                    '    End If
                                    '    If RowWip.Primary_Product > PrimaryMax Then
                                    '        PrimaryMax = RowWip.Primary_Product
                                    '    End If
                                    'End If

                                    RowDummyList = TblDummyList.NewTbl_DummyListRow
                                    With RowDummyList
                                        .No = TblDummyList.Rows.Count + 1
                                        .Primary_RM = RowWip.Primary_Product
                                        .Primary_Unit = RowWip.Primary_Unit
                                        .Primary_Location = RowWip.Primary_Input_Location
                                        .Primary_Warehouse = RowWip.Primary_Warehouse
                                        .Code_No = RowWip.Product_Code
                                        .Description = RowWip.Description
                                        .Unit_Code = RowWip.Unit_Code
                                        .Location = RowWip.Storage_Area_Code
                                        .Warehouse = RowWip.Warehouse_Code
                                        .Start_Time = RowSCH.plan_start_time    'sch
                                        .Primary_Work_Unit = RowWip.Primary_WU
                                        .Work_Unit_Code = RowWip.work_unit_code
                                        .Identifier_JO = RowSCH.identifier_jo
                                        .Identifier_WO = RowSCH.identifier_wo
                                        .JO_Number = RowSCH.JO_Number    'sch
                                        .Wo_Number = RowSCH.WO_Number    'sch
                                        ._Decimal = RowWip._Decimal
                                        .Primary_Work_Unit_Group = RowWip.primary_work_unit_group
                                        .Primary_Process = RowSCH.primary_process    'sch
                                        .Primary_Category = RowWip.Primary_Category
                                        .Primary_Brand = RowWip.Primary_Brand
                                        .Primary_Sub_Brand = RowWip.Primary_Sub_Brand
                                        .Primary_Product_Group = RowWip.Primary_Product_Group
                                        .Created_Date_MO = RowWip.Created_Date_MO
                                        .Location_Controlled = RowWip.Location_Controlled
                                        .Expiry_Date_Controlled = RowWip.Expiry_Date_Controlled
                                        If RowWip.Target_Finish_Date_R2 <> #1/1/1950# Then
                                            .Target_Finish_Date = RowWip.Target_Finish_Date_R2
                                        Else
                                            If RowWip.Target_Finish_Date_R1 <> #1/1/1950# Then
                                                .Target_Finish_Date = RowWip.Target_Finish_Date_R1
                                            Else
                                                If RowWip.Target_Finish_Date <> #1/1/1950# Then
                                                    .Target_Finish_Date = RowWip.Target_Finish_Date
                                                End If
                                            End If
                                        End If
                                        .Qty_Need = RowWip.Balance_Qty
                                        .Qty_Sisa = .Qty_Need
                                        .Primary_JO = RowWip.Primary_JO
                                        .Is_RM = 0
                                        .Qty_Package = RowWip.Package_Qty
                                        .Batch_Controlled = RowWip.Batch_No_Controlled
                                    End With
                                    TblDummyList.AddTbl_DummyListRow(RowDummyList)
                                Next
                            End If
                        End If

                        If TblRm.Count > 0 Then
                            TblRm.DefaultView.Sort = "[Primary Product Code RM] ASC"
                            TblRm.DefaultView.RowFilter = "[Identifier JO]='" & .identifier_jo.ToString & "' AND [Identifier WO]='" & .identifier_wo.ToString & "' AND [Primary WU] = " & .primary_work_unit
                            If TblRm.DefaultView.Count > 0 Then
                                For ij As Integer = 0 To TblRm.DefaultView.Count - 1
                                    RowRm = TblRm.DefaultView(ij).Row
                                    ''get primary product min max 
                                    'If PrimaryMin = -1 Or PrimaryMax = -1 Then
                                    '    PrimaryMin = RowRm.Primary_Product_Code_RM
                                    '    PrimaryMax = RowRm.Primary_Product_Code_RM
                                    'Else
                                    '    If RowRm.Primary_Product_Code_RM < PrimaryMin Then
                                    '        PrimaryMin = RowRm.Primary_Product_Code_RM
                                    '    End If
                                    '    If RowRm.Primary_Product_Code_RM > PrimaryMax Then
                                    '        PrimaryMax = RowRm.Primary_Product_Code_RM
                                    '    End If
                                    'End If

                                    RowDummyList = TblDummyList.NewTbl_DummyListRow
                                    With RowDummyList
                                        .No = TblDummyList.Rows.Count + 1
                                        .Primary_RM = RowRm.Primary_Product_Code_RM
                                        .Primary_Unit = RowRm.Primary_Unit
                                        .Primary_Location = RowRm.Primary_Input_Location
                                        .Primary_Warehouse = RowRm.Primary_Warehouse
                                        .Code_No = RowRm.Product_Code
                                        .Description = RowRm.Description
                                        .Unit_Code = RowRm.Unit_Code
                                        .Location = RowRm.Storage_Area_Code
                                        .Warehouse = RowRm.Warehouse_Code
                                        .Start_Time = RowSCH.plan_start_time    'sch
                                        .Primary_Work_Unit = RowRm.Primary_WU
                                        .Work_Unit_Code = RowRm.work_unit_code
                                        .Identifier_JO = RowSCH.identifier_jo
                                        .Identifier_WO = RowSCH.identifier_wo
                                        .JO_Number = RowSCH.JO_Number    'sch
                                        .Wo_Number = RowSCH.WO_Number    'sch
                                        ._Decimal = RowRm._Decimal
                                        .Primary_Work_Unit_Group = RowRm.primary_work_unit_group
                                        .Primary_Process = RowSCH.primary_process    'sch
                                        .Primary_Category = RowRm.Primary_Category
                                        .Primary_Brand = RowRm.Primary_Brand
                                        .Primary_Sub_Brand = RowRm.Primary_Sub_Brand
                                        .Primary_Product_Group = RowRm.Primary_Product_Group
                                        .Created_Date_MO = RowRm.Created_Date_MO
                                        .Location_Controlled = RowRm.Location_Controlled
                                        .Expiry_Date_Controlled = RowRm.Expiry_Date_Controlled
                                        If RowRm.Target_Finish_Date_R2 <> #1/1/1950# Then
                                            .Target_Finish_Date = RowRm.Target_Finish_Date_R2
                                        Else
                                            If RowRm.Target_Finish_Date_R1 <> #1/1/1950# Then
                                                .Target_Finish_Date = RowRm.Target_Finish_Date_R1
                                            Else
                                                If RowRm.Target_Finish_Date <> #1/1/1950# Then
                                                    .Target_Finish_Date = RowRm.Target_Finish_Date
                                                End If
                                            End If
                                        End If
                                        .Qty_Need = RowRm.Balance_Qty
                                        .Qty_Sisa = .Qty_Need
                                        .Primary_JO = RowRm.Primary_JO
                                        .Is_RM = 1
                                        .Qty_Package = RowRm.Package_Qty
                                        .Batch_Controlled = RowRm.Batch_No_Controlled
                                        .Qty_Minus = .Qty_Transfer - (.Qty_Need - .Qty_On_Site)
                                    End With
                                    TblDummyList.AddTbl_DummyListRow(RowDummyList)
                                Next
                            End If
                        End If
                    End With
                Next
            End If

            Me.DeleteItem_CannotAvailableQtyNeed(TblDummyList)

            Me.ApakahProductRMLocationControlled(TblDummyList, vShowPopUp)

            Me.FilterByProduct(TblDummyList, dgvProduct, hj)
            Me.FilterByWorkunit(TblDummyList, dgvWU, hj)
            Me.FilterByJO(TblDummyList, dgvJO, hj)
            Me.FilterByProcess(TblDummyList, dgvProcess, hj)
            Me.FilterByLocation(TblDummyList, dgvLocation, hj)

            'get product qty
            PrimaryMin = -1
            PrimaryMax = -1
            TblDummyList_2 = TblDummyList.Copy
            If TblDummyList_2.Count > 0 Then
                TblDummyList_2.DefaultView.RowFilter = Nothing
                TblDummyList_2.DefaultView.Sort = "[Primary RM] ASC"
                RowDummyList = TblDummyList_2.DefaultView(0).Row
                PrimaryMin = RowDummyList.Primary_RM
                RowDummyList = TblDummyList_2.DefaultView(TblDummyList_2.DefaultView.Count - 1).Row
                PrimaryMax = RowDummyList.Primary_RM
            End If

            DtaProdQty.Connection.ConnectionString = Core_Module.GetConstring
            DtaProdQty.Fill(TblProductQty_dum, PrimaryMin, PrimaryMax)
            vTblProductQty.Clear()
            For Each RowDummyList In TblDummyList
                If ListCodeNo.Contains(RowDummyList.Code_No) = False Then
                    TblProductQty_dum.DefaultView.RowFilter = "[Primary Product] = " & RowDummyList.Primary_RM
                    If TblProductQty_dum.DefaultView.Count > 0 Then
                        vTblProductQty.Merge(TblProductQty_dum.DefaultView.ToTable)
                    End If
                    ListCodeNo.Add(RowDummyList.Code_No)
                End If
            Next

            Me.SortingBy(TblDummyList)
            Me.PrioritasStrictLocation(vTblProductQty, TblPriority, TblExclude, vUserCode)
            Me.GetQty_OnSite(TblDummyList, vTblProductQty, vUserCode)
            Me.GetQty_Transfer(TblDummyList, TblDetail, vTblProductQty, vUserCode)
            Me.FilterByNotReadOnly(TblDummyList, hj)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DtaProdQty.Dispose()
            DtaWip.Dispose()
            TblWip.Dispose()
            DtaRm.Dispose()
            TblRm.Dispose()
            TblProductQty_dum.Dispose()
        End Try
    End Function

    Public Function LoadData_ToDatagrid(ByRef TblDummyList As DseLaporanProd_PickList.Tbl_DummyListDataTable, _
   ByRef TblDetail As DseLaporanProd_PickList.tbldetailDataTable, ByRef vTblProductQty As DseLaporanProd_PickList.tblmaster_productqtyDataTable, _
   ByRef TblPriority As DseLaporanProd_PickList.tblmaster_storage_areaDataTable, ByRef TblExclude As DseLaporanProd_PickList.tblmaster_storage_areaDataTable, _
   ByVal vFrom As Date, ByVal vTo As Date, ByVal DgvItem As DataGridView, ByVal dgvProduct As DataGridView, _
   ByVal dgvProcess As DataGridView, ByVal dgvWorkUnit As DataGridView, ByVal dgvLocation As DataGridView, _
   ByVal dgvJO As DataGridView, ByVal i As FrmLaporanProd_PickList, ByVal vObjCD As Core_Module.ClsCekData, _
   ByVal vShowPopUp As Boolean, ByVal vIsGetData As Boolean, ByVal vGetData As EnumGetData, _
   ByVal vUserCode As String) As Boolean
        Dim rowDGV As New DataGridViewRow
        Dim RowDummyList As DseLaporanProd_PickList.Tbl_DummyListRow
        Dim vTblDuplicate As New DseLaporanProd_PickList.Tbl_DummyListDataTable
        Dim vRowDuplicate As DseLaporanProd_PickList.Tbl_DummyListRow
        Try
            If vIsGetData = True Then
                Me.LoadData1(TblDummyList, TblDetail, vTblProductQty, TblPriority, TblExclude, _
                vFrom, vTo, vObjCD, i, dgvWorkUnit, dgvJO, dgvProduct, dgvProcess, dgvLocation, _
                vShowPopUp, vGetData, vUserCode)
            End If

            DgvItem.Rows.Clear()
            DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If TblDummyList.Rows.Count > 0 Then
                If i.cboFormat.SelectedIndex = 0 Then   'format rekap
                    vTblDuplicate = TblDummyList.Copy
                    TblDummyList.Rows.Clear()
                    vTblDuplicate.DefaultView.Sort = "[Code No], [Start Time], [Target Finish Date], [Created Date MO], [Primary JO]"
                    For ij As Integer = 0 To vTblDuplicate.DefaultView.Count - 1
                        vRowDuplicate = vTblDuplicate.DefaultView(ij).Row
                        TblDummyList.DefaultView.RowFilter = "[Primary RM]=" & vRowDuplicate.Primary_RM & _
                        " AND [Primary Location]=" & vRowDuplicate.Primary_Location
                        If TblDummyList.DefaultView.Count > 0 Then
                            RowDummyList = TblDummyList.DefaultView(0).Row
                            RowDummyList.Qty_Need += vRowDuplicate.Qty_Need
                            RowDummyList.Qty_On_Site += vRowDuplicate.Qty_On_Site
                            RowDummyList.Qty_Sisa += vRowDuplicate.Qty_Sisa
                            RowDummyList.Qty_Transfer += vRowDuplicate.Qty_Transfer
                            RowDummyList.Qty_Minus += vRowDuplicate.Qty_Minus
                        Else
                            RowDummyList = TblDummyList.NewTbl_DummyListRow
                            For ik As Integer = 0 To vTblDuplicate.Columns.Count - 1
                                RowDummyList.Item(ik) = vRowDuplicate.Item(ik)
                            Next
                            TblDummyList.AddTbl_DummyListRow(RowDummyList)
                        End If
                    Next
                End If

                For ij As Integer = 0 To TblDummyList.Rows.Count - 1
                    RowDummyList = TblDummyList.Rows(ij)
                    DgvItem.Rows.Add()
                    rowDGV = DgvItem.Rows(DgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(i.ColItem_No.Index).Value = DgvItem.Rows.Count
                        .Cells(i.ColItem_PrimaryProduct.Index).Value = RowDummyList.Primary_RM
                        .Cells(i.ColItem_PrimaryWU.Index).Value = RowDummyList.Primary_Work_Unit
                        .Cells(i.ColItem_PrimaryLoc.Index).Value = RowDummyList.Primary_Location
                        .Cells(i.ColItem_CodeNoRM.Index).Value = RowDummyList.Code_No
                        .Cells(i.ColItem_DescRM.Index).Value = RowDummyList.Description
                        .Cells(i.ColItem_Unit.Index).Value = RowDummyList.Unit_Code
                        .Cells(i.ColItem_Location.Index).Value = RowDummyList.Location
                        .Cells(i.ColItem_Warehouse.Index).Value = RowDummyList.Warehouse
                        .Cells(i.ColItem_JONo.Index).Value = RowDummyList.JO_Number
                        .Cells(i.ColItem_WONO.Index).Value = RowDummyList.Wo_Number
                        .Cells(i.ColItem_WU.Index).Value = RowDummyList.Work_Unit_Code
                        .Cells(i.ColItem_StartTime.Index).Value = Format(RowDummyList.Start_Time, "dd MMM yyyy HH.mm")
                        .Cells(i.ColItem_QtyNeed.Index).Value = vObjCD.DigitGrouping(RowDummyList.Qty_Need, RowDummyList._Decimal)
                        .Cells(i.ColItem_QtyOnSite.Index).Value = vObjCD.DigitGrouping(RowDummyList.Qty_On_Site, RowDummyList._Decimal)
                        .Cells(i.ColItem_QtyTransfer.Index).Value = vObjCD.DigitGrouping(RowDummyList.Qty_Transfer, RowDummyList._Decimal)
                        If RowDummyList.Qty_Minus <= 0 Then
                            .Cells(i.ColItem_QtyMinus.Index).Value = vObjCD.DigitGrouping(RowDummyList.Qty_Minus * -1, RowDummyList._Decimal)
                        Else
                            .Cells(i.ColItem_QtyMinus.Index).Value = "+" & vObjCD.DigitGrouping(RowDummyList.Qty_Minus, RowDummyList._Decimal)
                        End If
                        .Cells(i.ColItem_QtyPackage.Index).Value = vObjCD.DigitGrouping(RowDummyList.Qty_Package, RowDummyList._Decimal)
                        .Cells(i.ColItem_Transfer.Index).Value = False
                        .Cells(i.ColItem_Identifier.Index).Value = RowDummyList.Identifer
                    End With
                Next

                If i.cboFormat.SelectedIndex = 0 Then
                    TblDummyList.Rows.Clear()
                    TblDummyList = vTblDuplicate.Copy
                End If
            End If

            DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function DeleteItem_CannotAvailableQtyNeed(ByVal TblDummyList As DseLaporanProd_PickList.Tbl_DummyListDataTable) As Boolean
        Dim RowDummyList As DseLaporanProd_PickList.Tbl_DummyListRow
        Try
            TblDummyList.DefaultView.RowFilter = ""
            TblDummyList.DefaultView.RowFilter = "[Qty Need]=0"
            If TblDummyList.DefaultView.Count > 0 Then
                RowDummyList = TblDummyList.DefaultView(0).Row
                TblDummyList.Rows.Remove(RowDummyList)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahProductRMLocationControlled(ByVal TblDummyList As DseLaporanProd_PickList.Tbl_DummyListDataTable, Optional ByVal vMunculkanPopUp As Boolean = True) As Boolean
        Dim RowDummyList As DseLaporanProd_PickList.Tbl_DummyListRow
        Dim NO As Integer = 1
        Dim Dummy As String = ""
        Dim vListNotLocation As New List(Of String)
        Try
            TblDummyList.DefaultView.RowFilter = ""
            TblDummyList.DefaultView.RowFilter = "[Location Controlled]=0"
            If TblDummyList.DefaultView.Count > 0 Then
                For i As Integer = TblDummyList.DefaultView.Count - 1 To 0 Step -1
                    RowDummyList = TblDummyList.DefaultView(i).Row
                    If vListNotLocation.Contains(RowDummyList.Code_No) = False Then
                        vListNotLocation.Add(RowDummyList.Code_No)
                    End If
                    TblDummyList.Rows.Remove(RowDummyList)
                Next
            End If
            TblDummyList.DefaultView.RowFilter = ""

            vListNotLocation.Sort()
            If vMunculkanPopUp = True Then
                For Each ij As String In vListNotLocation
                    If NO <= 10 Then
                        If Dummy = "" Then
                            Dummy = NO & ". " & ij
                        Else
                            Dummy &= vbNewLine & NO & ". " & ij
                        End If
                        NO += 1
                    Else
                        Dummy &= " " & " any more..."
                    End If
                Next

                If vListNotLocation.Count > 0 Then
                    MsgBox("In list product code not location controlled, cannot view in datagrid :" & vbNewLine & Dummy, MsgBoxStyle.Information, "Information")
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function FilterByNotReadOnly(ByRef TblDummyList As DseLaporanProd_PickList.Tbl_DummyListDataTable, ByVal i As FrmLaporanProd_PickList) As Boolean
        Dim RowDummyList As DseLaporanProd_PickList.Tbl_DummyListRow
        Dim tbl As New DseLaporanProd_PickList.Tbl_DummyListDataTable
        Dim row As DseLaporanProd_PickList.Tbl_DummyListRow
        Try
            tbl.Rows.Clear()
            tbl = TblDummyList.Copy


            If i.chkShowNotReadyOnly.Checked = True Then
                TblDummyList.Rows.Clear()

                For Each row In tbl.Rows
                    If row.Qty_Need - row.Qty_On_Site > 0 Then
                        RowDummyList = TblDummyList.NewRow
                        For ij As Integer = 0 To TblDummyList.Columns.Count - 1
                            RowDummyList.Item(ij) = row.Item(ij)
                        Next
                        TblDummyList.Rows.Add(RowDummyList)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Public Function FilterData_OnlyIsRM() As Boolean
    '    Dim DTA As New DseLaporanProd_PickListTableAdapters.tbltrans_mo_item_rmTableAdapter
    '    Dim row As DseLaporanProd_PickList.tbltrans_mo_item_rmRow
    '    Dim RowBOM As DseLaporanProd_PickList.DataTable1Row
    '    Dim vPrimaryRM_Min As Double = -1
    '    Dim vPrimaryRM_Max As Double = -1
    '    Try
    '        If TblBOM.Rows.Count > 0 Then
    '            TblBOM.DefaultView.RowFilter = Nothing
    '            TblBOM.DefaultView.Sort = "[Primary Product RM] ASC"
    '            RowBOM = TblBOM.DefaultView(0).Row
    '            vPrimaryRM_Min = RowBOM.Primary_Product
    '            TblBOM.DefaultView.Sort = "[Primary Product RM] DESC"
    '            RowBOM = TblBOM.DefaultView(0).Row
    '            vPrimaryRM_Max = RowBOM.Primary_Product
    '        End If

    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA.Fill(TblRM, vPrimaryRM_Min, vPrimaryRM_Max)

    '        TblRM.DefaultView.RowFilter = "[Status RM]=" & 0
    '        For ij As Integer = 0 To TblRM.DefaultView.Count - 1
    '            row = TblRM.DefaultView(ij).Row
    '            TblBOM.DefaultView.RowFilter = "[Primary Product RM]=" & row.Primary_Product_Code_RM
    '            If TblBOM.DefaultView.Count > 0 Then
    '                RowBOM = TblBOM.DefaultView(0).Row
    '                TblBOM.Rows.Remove(RowBOM)
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Sub PrioritasStrictLocation(ByRef vTblProductQty As DseLaporanProd_PickList.tblmaster_productqtyDataTable, _
    ByRef TblPriority As DseLaporanProd_PickList.tblmaster_storage_areaDataTable, _
    ByRef TblExclude As DseLaporanProd_PickList.tblmaster_storage_areaDataTable, _
    ByVal vUserCode As String, Optional ByVal vLoadLocPriority As Boolean = True)
        Dim TblDummy As New DseLaporanProd_PickList.tblmaster_productqtyDataTable
        Dim rowProductQty As DseLaporanProd_PickList.tblmaster_productqtyRow
        Dim RowLocation As DseLaporanProd_PickList.tblmaster_storage_areaRow
        Dim rowExclude As DseLaporanProd_PickList.tblmaster_storage_areaRow
        Dim DV As Data.DataRowView

        Try
            If vLoadLocPriority = True Then
                Me.LoadLocationPriority(TblPriority, TblExclude, vUserCode)
            End If

            '1. urutkan berdasarkan lokasi prioritas
            'pindahkan ke tabel dummy
            For Each RowLocation In TblPriority.Rows
                vTblProductQty.DefaultView.RowFilter = "[Primary Storage Area] = " & RowLocation.Primary_Location
                vTblProductQty.DefaultView.Sort = "[Expiry Date], [Lot No]"
                TblDummy.Merge(vTblProductQty.DefaultView.ToTable)

                While vTblProductQty.DefaultView.Count > 0
                    rowProductQty = vTblProductQty.DefaultView(0).Row
                    vTblProductQty.Rows.Remove(rowProductQty)
                End While
            Next

            vTblProductQty.DefaultView.RowFilter = ""
            vTblProductQty.DefaultView.Sort = "[Expiry Date], [Lot No]"
            'pindahkan sisanya ke tabel dummy
            TblDummy.Merge(vTblProductQty.DefaultView.ToTable)

            'kembali pindahkan ke tabel utama
            vTblProductQty.Rows.Clear()
            vTblProductQty = TblDummy.Copy

            '2. tandai lokasi product yang exclude transfer
            For Each rowExclude In TblExclude
                vTblProductQty.DefaultView.RowFilter = "[Primary Storage Area] = " & rowExclude.Primary_Location
                For Each DV In vTblProductQty.DefaultView
                    rowProductQty = DV.Row
                    rowProductQty.Is_Exclude_Qty = 1
                Next
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblDummy.Dispose()
            vTblProductQty.DefaultView.Sort = ""
            vTblProductQty.DefaultView.RowFilter = ""
        End Try
    End Sub

    Public Sub LoadLocationPriority(ByRef TblPriority As DseLaporanProd_PickList.tblmaster_storage_areaDataTable, _
    ByRef TblExclude As DseLaporanProd_PickList.tblmaster_storage_areaDataTable, ByVal vUserCode As String)
        Dim DTA As New DseLaporanProd_PickListTableAdapters.tbltrans_custom_locationpriorityTableAdapter
        Dim TblLocationPriority1 As New DseLaporanProd_PickList.tbltrans_custom_locationpriorityDataTable
        Dim row As DseLaporanProd_PickList.tbltrans_custom_locationpriorityRow
        Dim DTAS As New DseLaporanProd_PickListTableAdapters.tblmaster_storage_areaTableAdapter
        Dim TblS As New DseLaporanProd_PickList.tblmaster_storage_areaDataTable
        Dim rowS, rowPriority, rowExclude As DseLaporanProd_PickList.tblmaster_storage_areaRow

        Try
            TblPriority.Rows.Clear()
            TblExclude.Rows.Clear()

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAS.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.FillBy(TblLocationPriority1, vUserCode) > 0 Then
                If DTAS.Fill(TblS) > 0 Then
                    TblLocationPriority1.DefaultView.RowFilter = "[Tipe]=" & FrmLaporanProd_PickList_LocPrio.EnumTipe._PriorityLocation
                    For i As Integer = 0 To TblLocationPriority1.DefaultView.Count - 1
                        row = TblLocationPriority1.DefaultView(i).Row
                        If row.Location_Level = FrmLaporanProd_PickList_LocPrio.EnumLevel._Lvl_Warehouse Then
                            TblS.DefaultView.RowFilter = "[Primary Warehouse]=" & row.Primary_Warehouse
                            If TblS.DefaultView.Count > 0 Then
                                TblPriority.Merge(TblS.DefaultView.ToTable)
                            End If
                        ElseIf row.Location_Level = FrmLaporanProd_PickList_LocPrio.EnumLevel._Lvl_Location Then
                            TblS.DefaultView.RowFilter = "[Primary Location]=" & row.Primary_Storage_Area
                            If TblS.DefaultView.Count > 0 Then
                                TblPriority.Merge(TblS.DefaultView.ToTable)
                            End If
                        End If
                    Next

                    TblLocationPriority1.DefaultView.RowFilter = "[Tipe]=" & FrmLaporanProd_PickList_LocPrio.EnumTipe._ExcludeLocation
                    For i As Integer = 0 To TblLocationPriority1.DefaultView.Count - 1
                        row = TblLocationPriority1.DefaultView(i).Row
                        If row.Location_Level = FrmLaporanProd_PickList_LocPrio.EnumLevel._Lvl_Warehouse Then
                            TblS.DefaultView.RowFilter = "[Primary Warehouse]=" & row.Primary_Warehouse
                            If TblS.DefaultView.Count > 0 Then
                                TblExclude.Merge(TblS.DefaultView.ToTable)
                            End If
                        ElseIf row.Location_Level = FrmLaporanProd_PickList_LocPrio.EnumLevel._Lvl_Location Then
                            TblS.DefaultView.RowFilter = "[Primary Location]=" & row.Primary_Storage_Area
                            If TblS.DefaultView.Count > 0 Then
                                TblExclude.Merge(TblS.DefaultView.ToTable)
                            End If
                        End If
                    Next
                End If
            End If

            If TblPriority.Rows.Count > 0 Then
                Dim ListPrimaryLoc As New List(Of Double)
                Dim UrutanPriority As Double = 0

                For Each rowPriority In TblPriority
                    If ListPrimaryLoc.Contains(rowPriority.Primary_Location) = False Then
                        ListPrimaryLoc.Add(rowPriority.Primary_Location)
                        UrutanPriority += 1
                        rowPriority.Urutan = UrutanPriority
                    Else
                        rowPriority.Urutan = -1
                    End If
                Next

                TblPriority.DefaultView.RowFilter = "[Urutan] = -1"
                While TblPriority.DefaultView.Count > 0
                    rowPriority = TblPriority.DefaultView(0).Row
                    TblPriority.Rows.Remove(rowPriority)
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAS.Dispose()
            TblLocationPriority1 = Nothing
            TblS = Nothing
        End Try
    End Sub

    Public Function GetQty_OnSite(ByRef TblDummyList As DseLaporanProd_PickList.Tbl_DummyListDataTable, _
    ByRef vTblProductQty As DseLaporanProd_PickList.tblmaster_productqtyDataTable, _
    ByVal vUserCode As String) As Boolean
        Dim DTA As New DseLaporanProd_PickListTableAdapters.tblmaster_productqtyTableAdapter
        Dim RowDummyList As DseLaporanProd_PickList.Tbl_DummyListRow
        Dim rowProductQty As DseLaporanProd_PickList.tblmaster_productqtyRow

        Dim PrimaryCodeNo_Bef As Double = -1
        Dim UrutanCodeNo As Integer = 0

        Try
            For i As Integer = 0 To TblDummyList.Rows.Count - 1
                RowDummyList = TblDummyList.Rows(i)
                With RowDummyList
                    If .Primary_RM <> PrimaryCodeNo_Bef Then
                        PrimaryCodeNo_Bef = .Primary_RM
                        UrutanCodeNo = 0
                    End If
                    vTblProductQty.DefaultView.RowFilter = "[Primary Product]=" & RowDummyList.Primary_RM & "AND [Primary Storage Area]=" & RowDummyList.Primary_Location
                    If vTblProductQty.DefaultView.Count > 0 Then
                        For ij As Integer = 0 To vTblProductQty.DefaultView.Count - 1
                            rowProductQty = vTblProductQty.DefaultView(ij).Row
                            UrutanCodeNo += 1
                            If RowDummyList.Qty_Need <> RowDummyList.Qty_On_Site Then
                                If RowDummyList.Qty_Sisa >= rowProductQty.Qty_Sisa Then
                                    RowDummyList.Qty_On_Site += rowProductQty.Qty_Sisa
                                    rowProductQty.Qty_Sisa = 0
                                    RowDummyList.Qty_Sisa = RowDummyList.Qty_Need - RowDummyList.Qty_On_Site
                                Else
                                    RowDummyList.Qty_On_Site += RowDummyList.Qty_Sisa
                                    rowProductQty.Qty_Sisa = rowProductQty.Qty_Sisa - RowDummyList.Qty_Sisa
                                    RowDummyList.Qty_Sisa = 0
                                End If
                            Else
                                Exit For
                            End If
                            rowProductQty.Is_Qty_On_Site = 1
                            rowProductQty.Urutan_Product_On_Site = UrutanCodeNo
                            RowDummyList.Is_Exclude_Qty = rowProductQty.Is_Exclude_Qty
                        Next
                    End If
                    .Qty_Minus = .Qty_Transfer - (.Qty_Need - .Qty_On_Site)
                End With
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub TblDetail_AddRow(ByRef RowDummyList As DseLaporanProd_PickList.Tbl_DummyListRow, _
    ByRef TblDetail As DseLaporanProd_PickList.tbldetailDataTable, _
    ByVal rowProductQty As DseLaporanProd_PickList.tblmaster_productqtyRow, _
    ByVal vQtyTransfer As Decimal, ByVal vEnumProsesTrans As EnumGetQty_Transfer, _
    Optional ByVal RowDummyList_2 As DseLaporanProd_PickList.Tbl_DummyListRow = Nothing)
        Dim rowDetail As DseLaporanProd_PickList.tbldetailRow
        Dim Filter As String = ""

        Try
            If vEnumProsesTrans <> EnumGetQty_Transfer.Proses_3 Then
                Filter = "[Primary Product]=" & RowDummyList.Primary_RM & _
                           " AND [Primary From Location]=" & rowProductQty.Primary_Storage_Area & " AND [Batch]='" & _
                           rowProductQty.Lot_No & "' AND [Primary Work Unit]=" & RowDummyList.Primary_Work_Unit & _
                           " AND [WO No]='" & RowDummyList.Wo_Number & "' AND [JO No]='" & RowDummyList.JO_Number & _
                           "' AND [Batch] = '" & rowProductQty.Lot_No & "' AND [Expiry Date] = #" & rowProductQty.Expiry_Date & "#"
            Else
                Filter = "[Primary Product]=" & RowDummyList_2.Primary_RM & " AND [Primary From Location]=" & _
                            RowDummyList_2.Primary_Location & " AND [Batch]='" & rowProductQty.Lot_No & _
                            "' AND [Primary Work Unit]=" & RowDummyList_2.Primary_Work_Unit & " AND [WO No]='" & _
                            RowDummyList_2.Wo_Number & "' AND [JO No]='" & RowDummyList_2.JO_Number & _
                            "' AND [Batch] = '" & rowProductQty.Lot_No & "' AND [Expiry Date] = #" & rowProductQty.Expiry_Date & "#"
            End If

            TblDetail.DefaultView.RowFilter = ""
            TblDetail.DefaultView.RowFilter = Filter
            If TblDetail.DefaultView.Count > 0 Then
                rowDetail = TblDetail.DefaultView(0).Row
                rowDetail.Qty += vQtyTransfer
            Else
                If RowDummyList.Qty_Transfer > 0 Then
                    rowDetail = TblDetail.NewtbldetailRow
                    With rowDetail
                        If vEnumProsesTrans <> EnumGetQty_Transfer.Proses_3 Then
                            .Primary_Product = rowProductQty.Primary_Product
                            .Primary_From_Location = rowProductQty.Primary_Storage_Area
                            .From_Location = rowProductQty.Storage_Area_Code
                        Else
                            .Primary_Product = RowDummyList_2.Primary_RM
                            .Primary_From_Location = RowDummyList_2.Primary_Location
                            .From_Location = RowDummyList_2.Location
                        End If

                        .Batch = rowProductQty.Lot_No
                        .Qty = vQtyTransfer
                        .Identfier = RowDummyList.Identifer
                        .Primary_To_Location = RowDummyList.Primary_Location
                        .To_Location = RowDummyList.Location
                        .Primary_Work_Unit = RowDummyList.Primary_Work_Unit
                        .Expiry_Date = rowProductQty.Expiry_Date
                        .Code_No = RowDummyList.Code_No
                        .WO_No = RowDummyList.Wo_Number
                        .JO_No = RowDummyList.JO_Number
                        .Expiry_Date_Controlled = RowDummyList.Expiry_Date_Controlled
                        ._Decimal = RowDummyList._Decimal
                        .Batch_Controlled = RowDummyList.Batch_Controlled
                        TblDetail.AddtbldetailRow(rowDetail)
                    End With
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub GetQty_InProductQty(ByRef TblDummyList As DseLaporanProd_PickList.Tbl_DummyListDataTable, _
    ByRef TblDetail As DseLaporanProd_PickList.tbldetailDataTable, _
    ByRef vTblProductQty As DseLaporanProd_PickList.tblmaster_productqtyDataTable, _
    ByRef RowDummyList As DseLaporanProd_PickList.Tbl_DummyListRow, _
    ByVal vEnumProsesTrans As EnumGetQty_Transfer, _
    Optional ByVal RowDummyList_2 As DseLaporanProd_PickList.Tbl_DummyListRow = Nothing)
        Dim rowProductQty As DseLaporanProd_PickList.tblmaster_productqtyRow
        Dim QtyTransfer As Decimal = 0
        Dim Filter As String = ""

        Try
            If vEnumProsesTrans = EnumGetQty_Transfer.Proses_1 Then
                'Proses 1 cari dulu qty sisa yang bukan on site di product qty 
                Filter = "[Primary Product]=" & RowDummyList.Primary_RM & " AND [Is Qty On Site] = 0 AND [Is Exclude Qty] = 0"
            ElseIf vEnumProsesTrans = EnumGetQty_Transfer.Proses_2 Then
                'Proses 2 bila tidak ada maka cari qty sisa ke lokasi on site
                Filter = "[Primary Product]=" & RowDummyList.Primary_RM & " AND [Is Qty On Site] = 1" & _
                " AND [Primary Storage Area] = " & RowDummyList_2.Primary_Location & " AND [Is Exclude Qty] = 0"
            ElseIf vEnumProsesTrans = EnumGetQty_Transfer.Proses_3 Then
                'Proses 3 bila masih tidak ada maka ambil qty on site dari baris paling bawah
                Filter = "[Primary Product]=" & RowDummyList.Primary_RM & " AND [Is Qty On Site] = 1" & _
                " AND [Primary Storage Area] = " & RowDummyList_2.Primary_Location & " AND [Is Exclude Qty] = 0"
            End If

            vTblProductQty.DefaultView.RowFilter = Filter
            If vTblProductQty.DefaultView.Count > 0 Then
                For ij As Integer = 0 To vTblProductQty.DefaultView.Count - 1
                    rowProductQty = vTblProductQty.DefaultView(ij).Row
                    QtyTransfer = 0
                    If vEnumProsesTrans = EnumGetQty_Transfer.Proses_1 Or vEnumProsesTrans = EnumGetQty_Transfer.Proses_2 Then
                        If rowProductQty.Qty_Sisa > 0 Then
                            If RowDummyList.Qty_Need <> RowDummyList.Qty_Transfer Then
                                If rowProductQty.Qty_Sisa >= RowDummyList.Qty_Sisa Then
                                    QtyTransfer = RowDummyList.Qty_Sisa
                                    rowProductQty.Qty_Sisa -= RowDummyList.Qty_Sisa
                                    RowDummyList.Qty_Sisa = 0
                                Else
                                    QtyTransfer = rowProductQty.Qty_Sisa
                                    RowDummyList.Qty_Sisa -= rowProductQty.Qty_Sisa
                                    rowProductQty.Qty_Sisa = 0
                                End If

                                RowDummyList.Qty_Transfer += QtyTransfer
                                RowDummyList.Qty_Minus = RowDummyList.Qty_Transfer - (RowDummyList.Qty_Need - RowDummyList.Qty_On_Site)
                            Else
                                Exit For
                            End If
                        End If
                    ElseIf vEnumProsesTrans = EnumGetQty_Transfer.Proses_3 Then
                        If RowDummyList_2.Qty_On_Site > 0 Then
                            If RowDummyList.Qty_Sisa >= RowDummyList_2.Qty_On_Site Then
                                QtyTransfer = RowDummyList_2.Qty_On_Site
                                RowDummyList.Qty_Transfer += QtyTransfer
                                RowDummyList.Qty_Sisa -= QtyTransfer
                                RowDummyList.Qty_Minus = RowDummyList.Qty_Transfer - (RowDummyList.Qty_Need - RowDummyList.Qty_On_Site)

                                RowDummyList_2.Qty_Sisa += RowDummyList_2.Qty_On_Site
                                RowDummyList_2.Qty_On_Site = 0
                                RowDummyList_2.Qty_Minus = RowDummyList_2.Qty_Transfer - (RowDummyList_2.Qty_Need - RowDummyList_2.Qty_On_Site)
                            Else
                                QtyTransfer = RowDummyList.Qty_Sisa
                                RowDummyList_2.Qty_On_Site -= QtyTransfer
                                RowDummyList_2.Qty_Sisa += QtyTransfer
                                RowDummyList_2.Qty_Minus = RowDummyList_2.Qty_Sisa * -1

                                RowDummyList.Qty_Transfer += RowDummyList.Qty_Sisa
                                RowDummyList.Qty_Sisa = 0
                                RowDummyList.Qty_Minus = RowDummyList.Qty_Transfer - (RowDummyList.Qty_Need - RowDummyList.Qty_On_Site)
                            End If
                        End If
                    End If

                    'add tbldetail
                    If QtyTransfer > 0 Then
                        Me.TblDetail_AddRow(RowDummyList, TblDetail, rowProductQty, QtyTransfer, vEnumProsesTrans, RowDummyList_2)
                    End If
                    If RowDummyList.Qty_Sisa = 0 Then
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub CalculateQty_Transfer(ByVal Code_No As String, ByRef TblDummyList As DseLaporanProd_PickList.Tbl_DummyListDataTable, _
    ByRef TblDetail As DseLaporanProd_PickList.tbldetailDataTable, _
    ByRef vTblProductQty As DseLaporanProd_PickList.tblmaster_productqtyDataTable)
        Dim RowDummyList, RowDummyList_2 As DseLaporanProd_PickList.Tbl_DummyListRow
        Dim ListPrimaryLoc As New List(Of Double)

        Try
            TblDummyList.DefaultView.RowFilter = "[Code No] = '" & Code_No.ToString & "'"
            For i As Integer = 0 To TblDummyList.DefaultView.Count - 1
                RowDummyList = TblDummyList.DefaultView(i).Row
                RowDummyList.Identifer = RowDummyList.Primary_RM & "-" & TblDummyList.Rows.IndexOf(RowDummyList)
                If RowDummyList.Qty_On_Site < RowDummyList.Qty_Need Then
                    'Proses 1 cari dulu qty sisa yang bukan on site di product qty 
                    Me.GetQty_InProductQty(TblDummyList, TblDetail, vTblProductQty, RowDummyList, _
                    EnumGetQty_Transfer.Proses_1, Nothing)

                    'Proses 2 bila tidak ada maka cari qty sisa ke lokasi on site
                    If RowDummyList.Qty_Sisa > 0 Then
                        ListPrimaryLoc.Clear()
                        For i2 As Integer = TblDummyList.DefaultView.Count - 1 To 0 Step -1
                            RowDummyList_2 = TblDummyList.DefaultView(i2).Row
                            'tidak ambil dari on site yang lokasinya termasuk ke dalam exclude lokasi
                            If RowDummyList_2.Is_Exclude_Qty = 0 Then
                                'pastikan tidak proses lokasi yang sama dan pastikan tidak proses dirinya sendiri
                                If RowDummyList_2.Primary_Location <> RowDummyList.Primary_Location Then
                                    If TblDummyList.Rows.IndexOf(RowDummyList) <> TblDummyList.Rows.IndexOf(RowDummyList_2) Then
                                        If ListPrimaryLoc.Contains(RowDummyList_2.Primary_Location) = False Then
                                            ListPrimaryLoc.Add(RowDummyList_2.Primary_Location)
                                            Me.GetQty_InProductQty(TblDummyList, TblDetail, vTblProductQty, RowDummyList, _
                                            EnumGetQty_Transfer.Proses_2, RowDummyList_2)
                                            If RowDummyList.Qty_Sisa = 0 Then
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If

                    'Proses 3 bila masih tidak ada maka ambil qty on site dari baris paling bawah
                    If RowDummyList.Qty_Sisa > 0 Then
                        ListPrimaryLoc.Clear()
                        If TblDummyList.DefaultView.Count - 1 > i Then
                            For i3 As Integer = TblDummyList.DefaultView.Count - 1 To i + 1 Step -1
                                RowDummyList_2 = TblDummyList.DefaultView(i3).Row
                                'tidak ambil dari on site yang lokasinya termasuk ke dalam exclude lokasi
                                If RowDummyList_2.Is_Exclude_Qty = 0 Then
                                    If RowDummyList_2.Primary_Location <> RowDummyList.Primary_Location Then
                                        If ListPrimaryLoc.Contains(RowDummyList_2.Primary_Location) = False Then
                                            ListPrimaryLoc.Add(RowDummyList_2.Primary_Location)
                                            Me.GetQty_InProductQty(TblDummyList, TblDetail, vTblProductQty, RowDummyList, _
                                            EnumGetQty_Transfer.Proses_3, RowDummyList_2)
                                        End If
                                    End If
                                    If RowDummyList.Qty_Sisa = 0 Then
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetQty_Transfer(ByRef TblDummyList As DseLaporanProd_PickList.Tbl_DummyListDataTable, _
    ByRef TblDetail As DseLaporanProd_PickList.tbldetailDataTable, _
    ByRef vTblProductQty As DseLaporanProd_PickList.tblmaster_productqtyDataTable, _
    ByVal vUserCode As String) As Boolean
        Dim RowDummyList As DseLaporanProd_PickList.Tbl_DummyListRow
        Dim ListCodeNo As New List(Of String)

        Try
            TblDetail.Rows.Clear()

            For Each RowDummyList In TblDummyList.Rows
                If ListCodeNo.Contains(RowDummyList.Code_No) = False Then
                    ListCodeNo.Add(RowDummyList.Code_No)
                End If
            Next

            For Each Code_No As String In ListCodeNo
                Me.CalculateQty_Transfer(Code_No, TblDummyList, TblDetail, vTblProductQty)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FilterByProduct(ByRef vTbl As Data.DataTable, ByVal dgvProduct As DataGridView, ByVal i As FrmLaporanProd_PickList)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
            For Each dr As DataGridViewRow In dgvProduct.Rows
                With dr
                    If dr.Cells(i.colProduct_Primary.Index).Value <> -1 Then
                        If dr.Cells(i.colProduct_GroupLevel.Index).Value = "Group" Or dr.Cells(i.colProduct_GroupLevel.Index).Value = "Group Level" Then
                            If filter = "" Then
                                filter = "[Primary Product Group] = " & .Cells(i.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Product Group] = " & .Cells(i.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(i.colProduct_GroupLevel.Index).Value = "Brand" Then
                            If filter = "" Then
                                filter = "[Primary Brand] = " & .Cells(i.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Brand] = " & .Cells(i.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(i.colProduct_GroupLevel.Index).Value = "Sub Brand" Then
                            If filter = "" Then
                                filter = "[Primary Sub Brand] = " & .Cells(i.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Sub Brand] = " & .Cells(i.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(i.colProduct_GroupLevel.Index).Value = "Category" Then
                            If filter = "" Then
                                filter = "[Primary Category] = " & .Cells(i.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Category] = " & .Cells(i.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(i.colProduct_GroupLevel.Index).Value = "Product" Then
                            If filter = "" Then
                                filter = "[Primary RM] = " & .Cells(i.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary RM] = " & .Cells(i.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                    End If
                End With
            Next
            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            For ij As Integer = 0 To tblDummy.DefaultView.Count - 1
                rowDummy = tblDummy.DefaultView(ij).Row
                rowTbl = vTbl.NewRow
                For ii As Integer = 0 To tblDummy.Columns.Count - 1
                    rowTbl.Item(ii) = rowDummy.Item(ii)
                Next
                vTbl.Rows.Add(rowTbl)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblDummy.Dispose()
        End Try
    End Sub

    Private Sub FilterByProcess(ByRef vTbl As Data.DataTable, ByVal dgvProcess As DataGridView, ByVal i As FrmLaporanProd_PickList)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
            For Each dr As DataGridViewRow In dgvProcess.Rows
                With dr
                    If dr.Cells(i.colProcess_Primary.Index).Value <> -1 Then
                        If filter = "" Then
                            filter = "[Primary Process] = " & .Cells(i.colProcess_Primary.Index).Value.ToString
                        Else
                            filter += " OR [Primary Process] = " & .Cells(i.colProcess_Primary.Index).Value.ToString
                        End If
                    End If
                End With
            Next

            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            If tblDummy.DefaultView.Count > 0 Then
                For ij As Integer = 0 To tblDummy.DefaultView.Count - 1
                    rowDummy = tblDummy.DefaultView(ij).Row
                    rowTbl = vTbl.NewRow
                    For ii As Integer = 0 To tblDummy.Columns.Count - 1
                        rowTbl.Item(ii) = rowDummy.Item(ii)
                    Next
                    vTbl.Rows.Add(rowTbl)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblDummy.Dispose()
        End Try
    End Sub

    Private Sub FilterByWorkunit(ByRef vTbl As Data.DataTable, ByVal dgvWorkunit As DataGridView, ByVal i As FrmLaporanProd_PickList)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
            For Each dr As DataGridViewRow In dgvWorkunit.Rows
                With dr
                    If dr.Cells(i.ColWU_Primary.Index).Value <> -1 Then
                        If dr.Cells(i.ColWU_Group.Index).Value = "Work Unit Group" Or dr.Cells(i.ColWU_Group.Index).Value = "Work Unit Group Level" Then
                            If filter = "" Then
                                filter = "[Primary Work Unit Group] = " & .Cells(i.ColWU_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Work Unit Group] = " & .Cells(i.ColWU_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(i.ColWU_Group.Index).Value = "Work Unit" Then
                            If filter = "" Then
                                filter = "[Primary Work Unit] = " & .Cells(i.ColWU_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Work Unit] = " & .Cells(i.ColWU_Primary.Index).Value.ToString
                            End If
                        End If
                    End If
                End With
            Next
            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            For ij As Integer = 0 To tblDummy.DefaultView.Count - 1
                rowDummy = tblDummy.DefaultView(ij).Row
                rowTbl = vTbl.NewRow
                For ii As Integer = 0 To tblDummy.Columns.Count - 1
                    rowTbl.Item(ii) = rowDummy.Item(ii)
                Next
                vTbl.Rows.Add(rowTbl)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblDummy.Dispose()
        End Try
    End Sub

    Private Sub FilterByLocation(ByRef vTbl As Data.DataTable, ByVal dgvLocation As DataGridView, ByVal i As FrmLaporanProd_PickList)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
            For Each dr As DataGridViewRow In dgvLocation.Rows
                With dr
                    If dr.Cells(i.ColWU_Primary.Index).Value <> -1 Then
                        If dr.Cells(i.ColLocation_Group.Index).Value = "Location" Or dr.Cells(i.ColLocation_Group.Index).Value = "Location Level" Then
                            If filter = "" Then
                                filter = "[Primary Location] = " & .Cells(i.ColLocation_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Location] = " & .Cells(i.ColLocation_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(i.ColLocation_Group.Index).Value = "Warehouse" Then
                            If filter = "" Then
                                filter = "[Primary Warehouse] = " & .Cells(i.ColLocation_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Warehouse] = " & .Cells(i.ColLocation_Primary.Index).Value.ToString
                            End If
                        End If
                    End If
                End With
            Next
            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            For ij As Integer = 0 To tblDummy.DefaultView.Count - 1
                rowDummy = tblDummy.DefaultView(ij).Row
                rowTbl = vTbl.NewRow
                For ii As Integer = 0 To tblDummy.Columns.Count - 1
                    rowTbl.Item(ii) = rowDummy.Item(ii)
                Next
                vTbl.Rows.Add(rowTbl)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblDummy.Dispose()
        End Try
    End Sub

    Private Sub FilterByJO(ByRef vTbl As Data.DataTable, ByVal dgvJO As DataGridView, ByVal i As FrmLaporanProd_PickList)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
            For Each dr As DataGridViewRow In dgvJO.Rows
                With dr
                    If dr.Cells(i.ColJo_Identififer.Index).Value <> "" Then
                        If filter = "" Then
                            filter = "[Identifier JO] = '" & .Cells(i.ColJo_Identififer.Index).Value.ToString & "'"
                        Else
                            filter += " OR [Identifier JO] = '" & .Cells(i.ColJo_Identififer.Index).Value.ToString & "'"
                        End If
                    End If
                End With
            Next

            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            If tblDummy.DefaultView.Count > 0 Then
                For ij As Integer = 0 To tblDummy.DefaultView.Count - 1
                    rowDummy = tblDummy.DefaultView(ij).Row
                    rowTbl = vTbl.NewRow
                    For ii As Integer = 0 To tblDummy.Columns.Count - 1
                        rowTbl.Item(ii) = rowDummy.Item(ii)
                    Next
                    vTbl.Rows.Add(rowTbl)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblDummy.Dispose()
        End Try
    End Sub

    Private Function SortingBy(ByRef tbl As Data.DataTable) As Boolean
        Dim tblSort As New DseLaporanProd_PickList.Tbl_DummyListDataTable
        Dim rowSort As DseLaporanProd_PickList.Tbl_DummyListRow
        Dim row As DseLaporanProd_PickList.Tbl_DummyListRow
        Try
            tblSort = tbl.Copy
            tbl.Rows.Clear()
            tblSort.DefaultView.Sort = "[Code No], [Start Time], [Target Finish Date], [Created Date MO], [Primary JO]"
            For Each DV As System.Data.DataRowView In tblSort.DefaultView
                rowSort = DV.Row
                row = tbl.NewRow
                For ij As Integer = 0 To tbl.Columns.Count - 1
                    row.Item(ij) = rowSort.Item(ij)
                Next
                tbl.Rows.Add(row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblSort.Dispose()
        End Try
    End Function

    Public Sub IsiDgvTransferDetail(ByVal vPrimaryProduct As Double, ByVal vIdentifierTrans As String, _
    ByVal vPrimaryLoc As Double, ByVal vDgv As DataGridView, ByVal vFrm As FrmLaporanProd_PickList_TransDetail, _
    ByRef vTblDummyList As DseLaporanProd_PickList.Tbl_DummyListDataTable, _
    ByRef vTblDetail As DseLaporanProd_PickList.tbldetailDataTable, _
    ByRef vTblProductQty As DseLaporanProd_PickList.tblmaster_productqtyDataTable, _
    ByVal vTblTransDet As DseLaporanProd_PickList.tbltransferdetailDataTable, _
    ByVal vApakahRekap As Boolean, ByRef DecProd As Integer)
        Dim TblDummyList_Dum As New DseLaporanProd_PickList.Tbl_DummyListDataTable
        Dim TblDetail_Dum As New DseLaporanProd_PickList.tbldetailDataTable
        Dim TblProductQty_Dum As New DseLaporanProd_PickList.tblmaster_productqtyDataTable

        Dim rowTransDet As DseLaporanProd_PickList.tbltransferdetailRow
        Dim rowDummyList, rowDummyList_2 As DseLaporanProd_PickList.Tbl_DummyListRow
        Dim rowDetail As DseLaporanProd_PickList.tbldetailRow
        Dim rowProductQty As DseLaporanProd_PickList.tblmaster_productqtyRow
        Dim rowD As DataGridViewRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim FilterDetail As String = ""
        Dim FilterDummy As String = ""
        Dim FilterProdQty As String = ""
        Dim LocationTo As String = ""
        Dim ApakahBatch As Boolean = False
        Dim ApakahExpDate As Boolean = False
        Dim QtyProdQtySisa As Decimal = 0
        Dim QtyOnSite_Temp As Decimal = 0
        Dim QtyMinus_Temp As Decimal = 0
        Dim IsLoop As Boolean = False
        Dim HitungQty As Boolean = False
        Dim IdxRow As Integer = -1

        Try
            vTblDetail.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct
            vTblDummyList.DefaultView.RowFilter = "[Primary RM] = " & vPrimaryProduct
            vTblProductQty.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct
            TblDummyList_Dum.Merge(vTblDummyList.DefaultView.ToTable)
            TblDetail_Dum.Merge(vTblDetail.DefaultView.ToTable)
            TblProductQty_Dum.Merge(vTblProductQty.DefaultView.ToTable)

            For i As Integer = TblDummyList_Dum.Rows.Count - 1 To 0 Step -1
                rowDummyList = TblDummyList_Dum.Rows(i)
                QtyMinus_Temp = 0
                If rowDummyList.Identifer = vIdentifierTrans Then
                    Exit For
                End If
                If rowDummyList.Qty_On_Site > 0 Then
                    IdxRow = TblDummyList_Dum.Rows.IndexOf(rowDummyList)
                    If IdxRow > 0 Then
                        For i2 As Integer = CInt(IdxRow - 1) To 0 Step -1
                            rowDummyList_2 = TblDummyList_Dum.Rows(i2)
                            If rowDummyList.Primary_Location = rowDummyList_2.Primary_Location Then
                                rowDummyList.Qty_Sisa += rowDummyList.Qty_On_Site
                                rowDummyList.Qty_On_Site = 0
                                rowDummyList.Qty_Minus = rowDummyList.Qty_Transfer - (rowDummyList.Qty_Need - rowDummyList.Qty_On_Site)

                                Exit For
                            End If
                        Next
                    End If
                End If
            Next

            For Each rowDetail In TblDetail_Dum
                If rowDetail.Primary_From_Location = vPrimaryLoc Then
                    GoTo SkipAddRow
                End If
                If vApakahRekap = True Then
                    If rowDetail.Primary_To_Location <> vPrimaryLoc Then
                        GoTo SkipAddRow
                    End If
                    HitungQty = True
                    LocationTo = rowDetail.To_Location
                    vTblTransDet.DefaultView.RowFilter = "[Primary Product] = " & rowDetail.Primary_Product _
                    & " AND [Primary Location From] = " & rowDetail.Primary_From_Location _
                    & " AND [Batch] = '" & rowDetail.Batch & "' AND [Exp Date] = #" & rowDetail.Expiry_Date & "#"
                Else
                    If rowDetail.Identfier <> vIdentifierTrans Then
                        HitungQty = False
                        If IsLoop = False Then
                            GoTo SkipAddRow
                        End If
                    Else
                        IsLoop = True
                        HitungQty = True
                        LocationTo = rowDetail.To_Location
                    End If
                    vTblTransDet.DefaultView.RowFilter = "[Primary Product] = " & rowDetail.Primary_Product _
                    & " AND [Primary Location From] = " & rowDetail.Primary_From_Location _
                    & " AND [Batch] = '" & rowDetail.Batch & "' AND [Exp Date] = #" & rowDetail.Expiry_Date & "#"
                End If

                ApakahBatch = rowDetail.Batch_Controlled
                ApakahExpDate = rowDetail.Expiry_Date_Controlled

                If vTblTransDet.DefaultView.Count > 0 Then
                    rowTransDet = vTblTransDet.DefaultView(0).Row
                    If HitungQty = True Then
                        rowTransDet.Qty_Transfer += rowDetail.Qty
                    End If
                Else
                    rowTransDet = vTblTransDet.NewtbltransferdetailRow
                    With rowTransDet
                        .Primary_Product = rowDetail.Primary_Product
                        .Primary_Location_From = rowDetail.Primary_From_Location
                        .Code_No = rowDetail.Code_No
                        .Location_From = rowDetail.From_Location
                        .Exp_Date = rowDetail.Expiry_Date
                        .Batch = rowDetail.Batch
                        .Identifier = rowDetail.Identfier
                        ._Decimal = rowDetail._Decimal
                        If HitungQty = True Then
                            .Qty_Transfer = rowDetail.Qty
                        End If
                    End With
                    vTblTransDet.Rows.Add(rowTransDet)
                End If
SkipAddRow:
            Next

            IsLoop = False
            For i As Integer = TblDetail_Dum.Count - 1 To 0 Step -1
                rowDetail = TblDetail_Dum.Rows(i)
                If IsLoop = True Then
                    If rowDetail.Identfier <> vIdentifierTrans Then
                        Exit For
                    End If
                End If

                QtyProdQtySisa = 0
                TblProductQty_Dum.DefaultView.RowFilter = "[Primary Storage Area] = " & rowDetail.Primary_From_Location _
                & " AND [Lot No] = '" & rowDetail.Batch & "' AND [Expiry Date] = #" & rowDetail.Expiry_Date & "#"
                If TblProductQty_Dum.DefaultView.Count > 0 Then
                    rowProductQty = TblProductQty_Dum.DefaultView(0).Row
                    rowProductQty.Qty_Sisa += rowDetail.Qty
                    QtyProdQtySisa = rowProductQty.Qty_Sisa
                End If

                TblProductQty_Dum.DefaultView.RowFilter = "[Primary] = -1 AND [Primary Storage Area] = " & rowDetail.Primary_To_Location
                While TblProductQty_Dum.DefaultView.Count > 0
                    rowProductQty = TblProductQty_Dum.DefaultView(0).Row
                    TblProductQty_Dum.Rows.Remove(rowProductQty)
                End While

                QtyOnSite_Temp = 0
                TblDummyList_Dum.DefaultView.RowFilter = "[Identifer] = '" & rowDetail.Identfier & "'"
                If TblDummyList_Dum.DefaultView.Count > 0 Then
                    rowDummyList = TblDummyList_Dum.DefaultView(0).Row
                    QtyOnSite_Temp = rowDummyList.Qty_On_Site
                    'If rowDummyList.Qty_Sisa < 0 Then
                    'rowDummyList.Qty_Sisa = 0
                    'End If
                    rowDummyList.Qty_Sisa += rowDetail.Qty
                    rowDummyList.Qty_Transfer -= rowDetail.Qty
                    rowDummyList.Qty_Minus = rowDummyList.Qty_Transfer - (rowDummyList.Qty_Need - rowDummyList.Qty_On_Site)
                End If

                vTblTransDet.DefaultView.RowFilter = "[Identifier] = '" & rowDetail.Identfier _
                & "' AND [Primary Location From] = " & rowDetail.Primary_From_Location _
                & " AND [Batch] = '" & rowDetail.Batch & "' AND [Exp Date] = #" & rowDetail.Expiry_Date & "#"
                If vTblTransDet.DefaultView.Count > 0 Then
                    rowTransDet = vTblTransDet.DefaultView(0).Row
                    rowTransDet.Qty_On_Hand = QtyProdQtySisa
                End If

                rowDetail.Qty = 0
                If rowDetail.Identfier = vIdentifierTrans Then
                    IsLoop = True
                End If
            Next

            vDgv.Rows.Clear()
            If ApakahBatch = False Then
                vDgv.Columns(vFrm.ColTranDet_Batch.Index).Visible = False
            End If
            If ApakahExpDate = False Then
                vDgv.Columns(vFrm.ColTranDet_ExpDate.Index).Visible = False
            End If

            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each rowTransDet In vTblTransDet
                vDgv.Rows.Add()
                rowD = vDgv.Rows(vDgv.Rows.Count - 1)
                With rowD
                    .Cells(vFrm.ColTranDet_PrimaryCodeNo.Index).Value = rowTransDet.Primary_Product
                    .Cells(vFrm.ColTranDet_PrimaryLocFrom.Index).Value = rowTransDet.Primary_Location_From
                    .Cells(vFrm.ColTranDet_PrimaryLocTo.Index).Value = vPrimaryLoc
                    .Cells(vFrm.ColTranDet_CodeNo.Index).Value = rowTransDet.Code_No
                    .Cells(vFrm.ColTranDet_LocationFrom.Index).Value = rowTransDet.Location_From
                    .Cells(vFrm.ColTransDet_LocationTo.Index).Value = LocationTo
                    .Cells(vFrm.ColTranDet_Batch.Index).Value = rowTransDet.Batch
                    .Cells(vFrm.ColTranDet_ExpDate.Index).Value = Format(rowTransDet.Exp_Date, "dd MMM yyyy")
                    .Cells(vFrm.ColTranDet_QtyOnHand.Index).Value = ObjCD.DigitGrouping(rowTransDet.Qty_On_Hand, rowTransDet._Decimal)
                    .Cells(vFrm.ColTranDet_QtyTrans.Index).Value = ObjCD.DigitGrouping(rowTransDet.Qty_Transfer, rowTransDet._Decimal)
                    DecProd = rowTransDet._Decimal
                End With
            Next
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            vTblDetail.Clear()
            vTblDummyList.Clear()
            vTblProductQty.Clear()
            vTblDetail = TblDetail_Dum.Copy
            vTblDummyList = TblDummyList_Dum.Copy
            vTblProductQty = TblProductQty_Dum.Copy
        End Try
    End Sub

    Public Enum EnumFilter As Integer
        Filter_RM
        Filter_Process
        Filter_WU
        Filter_LocInput
        Filter_JO
    End Enum

    Public Enum EnumGroupLevel As Integer
        Group
        Brand
        SubBrand
        Category
        Product
        WorkUnitGroup
        WorkUnit
        Location
        Warehouse
    End Enum

    Public Function Simpan_FilterSetting(ByVal vUserCode As String, ByVal vDgvRM As DataGridView, _
    ByVal vDgvProcess As DataGridView, ByVal vDgvWU As DataGridView, ByVal vDgvLocInp As DataGridView, _
    ByVal vDgvJO As DataGridView, ByVal vFrm As FrmLaporanProd_PickList) As Boolean
        Dim Dta As New DseLaporanProd_PickListTableAdapters.tbltrans_mo_item_picklist_settingTableAdapter
        Dim row As DataGridViewRow
        Dim GroupLevel As EnumGroupLevel

        Try
            Dta.Connection.ConnectionString = core_module.getconstring
            Dta.Delete(vUserCode)

            For i As Integer = 0 To vDgvRM.RowCount - 1
                row = vDgvRM.Rows(i)
                If String.IsNullOrEmpty(row.Cells(vFrm.colProduct_Code.Index).Value) = False Then
                    Select Case row.Cells(vFrm.colProduct_GroupLevel.Index).Value
                        Case "Group"
                            GroupLevel = EnumGroupLevel.Group
                        Case "Group Level"
                            GroupLevel = EnumGroupLevel.Group
                        Case "Brand"
                            GroupLevel = EnumGroupLevel.Brand
                        Case "Sub Brand"
                            GroupLevel = EnumGroupLevel.SubBrand
                        Case "Category"
                            GroupLevel = EnumGroupLevel.Category
                        Case "Product"
                            GroupLevel = EnumGroupLevel.Product
                    End Select

                    Dta.Insert(vUserCode, EnumFilter.Filter_RM, GroupLevel, row.Cells(vFrm.colProduct_Primary.Index).Value, "-")
                End If
            Next

            For i As Integer = 0 To vDgvProcess.RowCount - 1
                row = vDgvProcess.Rows(i)
                If String.IsNullOrEmpty(row.Cells(vFrm.colProcess_Code.Index).Value) = False Then
                    Dta.Insert(vUserCode, EnumFilter.Filter_Process, 0, row.Cells(vFrm.colProcess_Primary.Index).Value, "-")
                End If
            Next

            For i As Integer = 0 To vDgvWU.RowCount - 1
                row = vDgvWU.Rows(i)
                If String.IsNullOrEmpty(row.Cells(vFrm.ColWU_Code.Index).Value) = False Then
                    Select Case row.Cells(vFrm.ColWU_Group.Index).Value
                        Case "Work Unit Group Level"
                            GroupLevel = EnumGroupLevel.WorkUnitGroup
                        Case "Work Unit Group"
                            GroupLevel = EnumGroupLevel.WorkUnitGroup
                        Case "Work Unit"
                            GroupLevel = EnumGroupLevel.WorkUnit
                    End Select

                    Dta.Insert(vUserCode, EnumFilter.Filter_WU, GroupLevel, row.Cells(vFrm.ColWU_Primary.Index).Value, "-")
                End If
            Next

            For i As Integer = 0 To vDgvLocInp.RowCount - 1
                row = vDgvLocInp.Rows(i)
                If String.IsNullOrEmpty(row.Cells(vFrm.ColLocation_Code.Index).Value) = False Then
                    Select Case row.Cells(vFrm.ColLocation_Group.Index).Value
                        Case "Location Level"
                            GroupLevel = EnumGroupLevel.Location
                        Case "Location"
                            GroupLevel = EnumGroupLevel.Location
                        Case "Warehouse"
                            GroupLevel = EnumGroupLevel.Warehouse
                    End Select

                    Dta.Insert(vUserCode, EnumFilter.Filter_LocInput, GroupLevel, row.Cells(vFrm.ColLocation_Primary.Index).Value, "-")
                End If
            Next

            For i As Integer = 0 To vDgvJO.RowCount - 1
                row = vDgvJO.Rows(i)
                If String.IsNullOrEmpty(row.Cells(vFrm.ColJo_JONo.Index).Value) = False Then
                    Dta.Insert(vUserCode, EnumFilter.Filter_JO, 0, -1, row.Cells(vFrm.ColJo_Identififer.Index).Value)
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub IsiDgv_FilterSetting(ByVal vUserCode As String, ByVal vDgvRM As DataGridView, _
    ByVal vDgvProcess As DataGridView, ByVal vDgvWU As DataGridView, ByVal vDgvLocInp As DataGridView, _
    ByVal vDgvJO As DataGridView, ByVal vFrm As FrmLaporanProd_PickList)
        Dim Dta As New DseLaporanProd_PickListTableAdapters.tbltrans_mo_item_picklist_settingTableAdapter
        Dim Tbl As New DseLaporanProd_PickList.tbltrans_mo_item_picklist_settingDataTable
        Dim row As DseLaporanProd_PickList.tbltrans_mo_item_picklist_settingRow
        Dim rowD As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjGroup As New Core_Module.ClsMasterProductGroup
        Dim ObjCategory As New Core_Module.ClsMasterCategory
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjWU As New Core_Module.ClsMasterWorkUnit
        Dim ObjWU_Group As New Core_Module.ClsMasterWorkUnitGroup
        Dim ObjProcess As New Core_Module.ClsMasterProses
        Dim ObjLocInp As New Core_Module.clsMasterStorageArea
        Dim ObjWH As New Core_Module.clsMasterWarehouse
        Dim ObjJO As New ClsTrans_ManufacturingOrder
        Dim frm_JOCari As New FrmTrans_ManufacturingOrder_PrintJO_Cari

        Try
            vDgvRM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvProcess.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvWU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvLocInp.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvJO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            vDgvRM.AllowUserToAddRows = False
            vDgvProcess.AllowUserToAddRows = False
            vDgvWU.AllowUserToAddRows = False
            vDgvLocInp.AllowUserToAddRows = False
            vDgvJO.AllowUserToAddRows = False

            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.FillBy_UserCode(Tbl, vUserCode)
            If Tbl.Count > 0 Then
                For Each row In Tbl
                    With row
                        If .Filter_Tipe = EnumFilter.Filter_RM Then
                            vDgvRM.Rows.Add()
                            rowD = vDgvRM.Rows(vDgvRM.RowCount - 1)
                            rowD.Cells(vFrm.colProduct_Primary.Index).Value = .Primary_Acuan
                            If .Level_Acuan = EnumGroupLevel.Group Then
                                rowD.Cells(vFrm.colProduct_GroupLevel.Index).Value = "Group"
                                rowD.Cells(vFrm.colProduct_Code.Index).Value = ObjGroup.GetGroupCode(.Primary_Acuan)
                                rowD.Cells(vFrm.colProduct_Desc.Index).Value = ObjGroup.GetGroupDesc(.Primary_Acuan)
                            ElseIf .Level_Acuan = EnumGroupLevel.Brand Then
                                rowD.Cells(vFrm.colProduct_GroupLevel.Index).Value = "Brand"
                                rowD.Cells(vFrm.colProduct_Code.Index).Value = ObjGroup.GetGroupCode(.Primary_Acuan)
                                rowD.Cells(vFrm.colProduct_Desc.Index).Value = ObjGroup.GetGroupDesc(.Primary_Acuan)
                            ElseIf .Level_Acuan = EnumGroupLevel.SubBrand Then
                                rowD.Cells(vFrm.colProduct_GroupLevel.Index).Value = "Sub Brand"
                                rowD.Cells(vFrm.colProduct_Code.Index).Value = ObjGroup.GetGroupCode(.Primary_Acuan)
                                rowD.Cells(vFrm.colProduct_Desc.Index).Value = ObjGroup.GetGroupDesc(.Primary_Acuan)
                            ElseIf .Level_Acuan = EnumGroupLevel.Category Then
                                rowD.Cells(vFrm.colProduct_GroupLevel.Index).Value = "Category"
                                rowD.Cells(vFrm.colProduct_Code.Index).Value = ObjCategory.GetCategoryCode(.Primary_Acuan)
                                rowD.Cells(vFrm.colProduct_Desc.Index).Value = ObjCategory.GetDesc(.Primary_Acuan)
                            ElseIf .Level_Acuan = EnumGroupLevel.Product Then
                                rowD.Cells(vFrm.colProduct_GroupLevel.Index).Value = "Product"
                                rowD.Cells(vFrm.colProduct_Code.Index).Value = ObjProd.GetProductCode(.Primary_Acuan)
                                rowD.Cells(vFrm.colProduct_Desc.Index).Value = ObjProd.GetDescription(.Primary_Acuan)
                            End If
                        ElseIf .Filter_Tipe = EnumFilter.Filter_Process Then
                            vDgvProcess.Rows.Add()
                            rowD = vDgvProcess.Rows(vDgvProcess.RowCount - 1)
                            rowD.Cells(vFrm.colProcess_Primary.Index).Value = .Primary_Acuan
                            rowD.Cells(vFrm.colProcess_Code.Index).Value = ObjProcess.GetCode(.Primary_Acuan)
                            rowD.Cells(vFrm.colProcess_Desc.Index).Value = ObjProcess.GetDesc(.Primary_Acuan)
                        ElseIf .Filter_Tipe = EnumFilter.Filter_WU Then
                            vDgvWU.Rows.Add()
                            rowD = vDgvWU.Rows(vDgvWU.RowCount - 1)
                            rowD.Cells(vFrm.ColWU_Primary.Index).Value = .Primary_Acuan
                            If .Level_Acuan = EnumGroupLevel.WorkUnitGroup Then
                                rowD.Cells(vFrm.ColWU_Group.Index).Value = "Work Unit Group"
                                rowD.Cells(vFrm.ColWU_Code.Index).Value = ObjWU_Group.GetCode(.Primary_Acuan)
                                rowD.Cells(vFrm.ColWU_Desc.Index).Value = ObjWU_Group.GetDesc(.Primary_Acuan)
                            ElseIf .Level_Acuan = EnumGroupLevel.WorkUnit Then
                                rowD.Cells(vFrm.ColWU_Group.Index).Value = "Work Unit"
                                rowD.Cells(vFrm.ColWU_Code.Index).Value = ObjWU.GetCode(.Primary_Acuan)
                                rowD.Cells(vFrm.ColWU_Desc.Index).Value = ObjWU.GetDesc(.Primary_Acuan)
                            End If
                        ElseIf .Filter_Tipe = EnumFilter.Filter_LocInput Then
                            vDgvLocInp.Rows.Add()
                            rowD = vDgvLocInp.Rows(vDgvLocInp.RowCount - 1)
                            rowD.Cells(vFrm.ColLocation_Primary.Index).Value = .Primary_Acuan
                            If .Level_Acuan = EnumGroupLevel.Location Then
                                rowD.Cells(vFrm.ColLocation_Group.Index).Value = "Location"
                                rowD.Cells(vFrm.ColLocation_Code.Index).Value = ObjLocInp.GetStorageAreaCode(.Primary_Acuan)
                                rowD.Cells(vFrm.ColLocation_Desc.Index).Value = ObjLocInp.GetDescription(.Primary_Acuan)
                            ElseIf .Level_Acuan = EnumGroupLevel.Warehouse Then
                                rowD.Cells(vFrm.ColLocation_Group.Index).Value = "Warehouse"
                                rowD.Cells(vFrm.ColLocation_Code.Index).Value = ObjWH.GetWarehouseCode(.Primary_Acuan)
                                rowD.Cells(vFrm.ColLocation_Desc.Index).Value = ObjWH.GetDescription(.Primary_Acuan)
                            End If
                        ElseIf .Filter_Tipe = EnumFilter.Filter_JO Then
                            vDgvJO.Rows.Add()
                            rowD = vDgvJO.Rows(vDgvJO.RowCount - 1)
                            rowD.Cells(vFrm.ColJo_Identififer.Index).Value = .Identifier_JO
                            rowD.Cells(vFrm.ColJo_JONo.Index).Value = ObjJO.GetJONumber(.Identifier_JO)
                            rowD.Cells(vFrm.ColJo_Date.Index).Value = Format(frm_JOCari.GetDateMOByJoNo(rowD.Cells(vFrm.ColJo_JONo.Index).Value), "dd MMM yyyy")
                        End If
                    End With
                Next

                ObjCD.SetNumber(vDgvRM)
                ObjCD.SetNumber(vDgvProcess)
                ObjCD.SetNumber(vDgvWU)
                ObjCD.SetNumber(vDgvLocInp)
                ObjCD.SetNumber(vDgvJO)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vDgvRM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgvProcess.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgvWU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgvLocInp.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgvJO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgvRM.AllowUserToAddRows = True
            vDgvProcess.AllowUserToAddRows = True
            vDgvWU.AllowUserToAddRows = True
            vDgvLocInp.AllowUserToAddRows = True
            vDgvJO.AllowUserToAddRows = True

            ObjCD = Nothing
            ObjGroup = Nothing
            ObjCategory = Nothing
            ObjProd = Nothing
            ObjWU = Nothing
            ObjWU_Group = Nothing
            ObjProcess = Nothing
            ObjLocInp = Nothing
            ObjWH = Nothing
            ObjJO = Nothing
            frm_JOCari.Dispose()
        End Try
    End Sub
End Class
