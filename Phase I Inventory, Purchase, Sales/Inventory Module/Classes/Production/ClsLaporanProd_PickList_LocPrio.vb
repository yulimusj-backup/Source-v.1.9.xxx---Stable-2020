Option Explicit On
Option Strict Off
Public Class ClsLaporanProd_PickList_LocPrio

    Public Function GetPrimaryWarehouseByPirmaryLocation(ByVal vPirmaryLocation As Double) As Double
        Dim DTA As New DseLaporanProd_PickListTableAdapters.tblmaster_storage_areaTableAdapter
        Dim Retval As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimaryWByPrimaryL(vPirmaryLocation)
            If Retval.HasValue = True Then
                If Retval.Value <> -1 Then
                    Return Retval.Value
                Else
                    Return -1
                End If
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function HapusData(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseLaporanProd_PickListTableAdapters.tbltrans_custom_locationpriorityTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery(vUserCode)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanData(ByVal dgvItem1 As DataGridView, ByVal dgvItem2 As DataGridView, ByVal i As FrmLaporanProd_PickList_LocPrio, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseLaporanProd_PickListTableAdapters.tbltrans_custom_locationpriorityTableAdapter
        Dim ir As New DataGridViewRow
        Dim vLocationLvl As FrmLaporanProd_PickList_LocPrio.EnumLevel = FrmLaporanProd_PickList_LocPrio.EnumLevel._Lvl_Location
        Try
            dgvItem1.AllowUserToAddRows = False
            dgvItem2.AllowUserToAddRows = False
            If dgvItem1.RowCount > 0 Or dgvItem2.RowCount > 0 Then
                Me.HapusData(vUserCode)
                For Each ir In dgvItem1.Rows
                    With ir
                        If .Cells(i.ColLocation_Group.Index).Value <> "Warehouse" Then
                            vLocationLvl = FrmLaporanProd_PickList_LocPrio.EnumLevel._Lvl_Location
                        Else
                            vLocationLvl = FrmLaporanProd_PickList_LocPrio.EnumLevel._Lvl_Warehouse
                        End If
                        DTA.InsertQuery((.Cells(i.ColLocation_Primary.Index).Value), (.Cells(i.ColLocation_PrimaryW.Index).Value), (.Cells(i.ColLocation_Code.Index).Value), _
                         (.Cells(i.ColLocation_Desc.Index).Value), FrmLaporanProd_PickList_LocPrio.EnumTipe._PriorityLocation, vLocationLvl, (.Cells(i.ColLocation_Urutan.Index).Value), _
                         vUserCode, Now)
                    End With
                Next

                For Each ir In dgvItem2.Rows
                    With ir
                        If .Cells(i.ColLocation_Group.Index).Value <> "Warehouse" Then
                            vLocationLvl = FrmLaporanProd_PickList_LocPrio.EnumLevel._Lvl_Location
                        Else
                            vLocationLvl = FrmLaporanProd_PickList_LocPrio.EnumLevel._Lvl_Warehouse
                        End If
                        DTA.InsertQuery((.Cells(i.ColE_Primary.Index).Value), (.Cells(i.ColE_PrimaryW.Index).Value), (.Cells(i.ColE_Code.Index).Value), _
                         (.Cells(i.ColE_Desc.Index).Value), FrmLaporanProd_PickList_LocPrio.EnumTipe._ExcludeLocation, vLocationLvl, (.Cells(i.ColLocation_Urutan.Index).Value), _
                         vUserCode, Now)
                    End With
                Next
            End If

            dgvItem1.AllowUserToAddRows = True
            dgvItem2.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiDataDgv(ByVal dgvItem1 As DataGridView, ByVal dgvItem2 As DataGridView, ByVal vUserCode As String, ByVal i As FrmLaporanProd_PickList_LocPrio) As Boolean
        Dim DTA As New DseLaporanProd_PickListTableAdapters.tbltrans_custom_locationpriorityTableAdapter
        Dim TBL As New DseLaporanProd_PickList.tbltrans_custom_locationpriorityDataTable
        Dim Row As DseLaporanProd_PickList.tbltrans_custom_locationpriorityRow
        Dim rowDGV1 As New DataGridViewRow
        Dim rowDGV2 As New DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            dgvItem1.Rows.Clear()
            dgvItem1.AllowUserToAddRows = False
            dgvItem2.Rows.Clear()
            dgvItem2.AllowUserToAddRows = False
            dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvItem1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvItem2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            If DTA.FillBy(TBL, vUserCode) > 0 Then
                For ij As Integer = 0 To TBL.Rows.Count - 1
                    Row = TBL.Rows(ij)
                    With Row
                        If Row.Tipe = FrmLaporanProd_PickList_LocPrio.EnumTipe._PriorityLocation Then
                            dgvItem1.Rows.Add()
                            rowDGV1 = dgvItem1.Rows(dgvItem1.Rows.Count - 1)
                            rowDGV1.Cells(i.ColLocation_No.Index).Value = ""
                            rowDGV1.Cells(i.ColLocation_Code.Index).Value = .Code
                            rowDGV1.Cells(i.ColLocation_Desc.Index).Value = .Desc
                            rowDGV1.Cells(i.ColLocation_Primary.Index).Value = .Primary_Storage_Area
                            rowDGV1.Cells(i.ColLocation_PrimaryW.Index).Value = .Primary_Warehouse
                            If .Location_Level = FrmLaporanProd_PickList_LocPrio.EnumLevel._Lvl_Location Then
                                rowDGV1.Cells(i.ColLocation_Group.Index).Value = "Location"
                            Else
                                rowDGV1.Cells(i.ColLocation_Group.Index).Value = "Warehouse"
                            End If
                        ElseIf Row.Tipe = FrmLaporanProd_PickList_LocPrio.EnumTipe._ExcludeLocation Then
                            dgvItem2.Rows.Add()
                            rowDGV2 = dgvItem2.Rows(dgvItem2.Rows.Count - 1)
                            rowDGV2.Cells(i.ColLocation_No.Index).Value = ""
                            rowDGV2.Cells(i.ColLocation_Code.Index).Value = .Code
                            rowDGV2.Cells(i.ColLocation_Desc.Index).Value = .Desc
                            rowDGV2.Cells(i.ColLocation_Primary.Index).Value = .Primary_Storage_Area
                            rowDGV2.Cells(i.ColLocation_PrimaryW.Index).Value = .Primary_Warehouse
                            If .Location_Level = FrmLaporanProd_PickList_LocPrio.EnumLevel._Lvl_Location Then
                                rowDGV2.Cells(i.ColLocation_Group.Index).Value = "Location"
                            Else
                                rowDGV2.Cells(i.ColLocation_Group.Index).Value = "Warehouse"
                            End If
                        End If
                    End With
                Next
                ObjCD.SetNumber(dgvItem1)
                ObjCD.SetNumber(dgvItem2)
            End If
            dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvItem1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvItem2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            dgvItem1.AllowUserToAddRows = True
            dgvItem2.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Function
End Class
