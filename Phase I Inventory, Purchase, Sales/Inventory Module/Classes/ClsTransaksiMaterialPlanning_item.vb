Option Explicit On
Option Strict Off

Public Class ClsTransaksiMaterialPlanning_item
    Public Sub Simpan(ByVal vPrimaryTransaksi As Double, _
    ByVal vDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable, _
    ByVal vDT_ItemGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable, _
    ByVal vDT_ItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable, _
    ByVal vDT_ItemPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, _
    ByVal vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, _
    ByVal vDT_ItemSettingLocation As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, _
    ByVal vDT_ItemQty As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable, _
    ByVal vDT_ItemMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable, _
    ByVal vDT_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable, _
    ByVal vDT_ItemFollowUp As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupDataTable, _
    ByVal vDT_ItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable, _
    ByVal vDT_ItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable)
        Try
            Me.Simpan_Item_UsingScript(vDT, vPrimaryTransaksi)
            Me.Simpan_ItemGIT_UsingScript(vDT_ItemGIT, vPrimaryTransaksi)
            Me.Simpan_ItemDO_UsingScript(vDT_ItemDO, vPrimaryTransaksi)
            Me.Simpan_ItemPOB_UsingScript(vDT_ItemPOB, vPrimaryTransaksi)
            Me.Simpan_ItemSOB_UsingScript(vDT_ItemSOB, vPrimaryTransaksi)
            Me.Simpan_ItemPrOB_UsingScript(vDT_ItemPrOB, vPrimaryTransaksi)
            Me.Simpan_ItemQty_UsingScript(vDT_ItemQty, vPrimaryTransaksi)
            Me.Simpan_ItemExcLocation_UsingScript(vDT_ItemSettingLocation, vPrimaryTransaksi)
            Me.Simpan_ItemMOB_UsingScript(vDT_ItemMOB, vPrimaryTransaksi)
            Me.Simpan_Item_RMReq_UsingScript(vDT_ItemRMReq, vPrimaryTransaksi)
            Me.Simpan_ItemSOBDelievery_UsingScript(vDT_ItemSOBDelivery, vPrimaryTransaksi)
            Me.Simpan_ItemPrOBDelivery_UsingScript(vDT_ItemPrOBDelivery, vPrimaryTransaksi)
            Me.Simpan_Item_FollowUpDetail_UsingScrip(vDT_ItemFollowUp, vPrimaryTransaksi)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Simpan_Item_FollowUpDetail_UsingScrip(ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupDataTable, ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_Item_FollowUpDetail(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function Simpan_Item_RMReq_UsingScript(ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable, ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_Item_RMReq(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function GetInsertString_Item_FollowUpDetail(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            retval = "INSERT INTO tbltrans_material_planning_item_followup (`primary transaksi`, `primary product`, tipe, `primary po mo`, " _
            & "`no po mo`, `tanggal po mo`, `identifier jo`, `no jo`, `qty default`, `primary unit default`, `ratio default`, `qty multi`, " _
            & "`primary unit multi`, `ratio multi`, `user code`, `followup time`, `remarks`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If
                    With row
                        If .no_po_mo <> "" Then
                            If .no_po_mo.Contains("\") = True Then
                                .no_po_mo = .no_po_mo.Replace("\", "\\")
                            End If
                            If .no_po_mo.Contains(vbCrLf) = True Then
                                .no_po_mo = .no_po_mo.Replace(vbCrLf, "\r\n")
                            End If
                            If .no_po_mo.Contains("'") = True Then
                                .no_po_mo = .no_po_mo.Replace("'", "''")
                            End If
                        End If
                        If .no_jo <> "" Then
                            If .no_jo.Contains("\") = True Then
                                .no_jo = .no_jo.Replace("\", "\\")
                            End If
                            If .no_jo.Contains(vbCrLf) = True Then
                                .no_jo = .no_jo.Replace(vbCrLf, "\r\n")
                            End If
                            If .no_jo.Contains("'") = True Then
                                .no_jo = .no_jo.Replace("'", "''")
                            End If
                        End If
                        If .remarks <> "" Then
                            If .remarks.Contains("\") = True Then
                                .remarks = .remarks.Replace("\", "\\")
                            End If
                            If .remarks.Contains(vbCrLf) = True Then
                                .remarks = .remarks.Replace(vbCrLf, "\r\n")
                            End If
                            If .remarks.Contains("'") = True Then
                                .remarks = .remarks.Replace("'", "''")
                            End If
                        End If
                        retval = retval & vbCrLf & "(" & vPrimaryTransaksi.ToString & "," & .primary_product.ToString & "," _
                        & .tipe.ToString & "," & row.primary_po_mo.ToString & ",'" & .no_po_mo & "','" & Format(CDate(.tanggal_po_mo), "yyyy-MM-dd") & "','" & .identifier_jo & "','" _
                        & .no_jo & "'," & .qty_default.ToString & "," & .primary_unit_default.ToString & "," & .ratio_default.ToString & "," & .qty_multi.ToString & "," _
                        & .primary_unit_multi.ToString & "," & .ratio_multi.ToString & ",'" & .user_code & "','" & Format(.followup_time, "yyyy-MM-dd HH:mm:ss") & "','" & .remarks & "')" & Delimiter
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

    Private Function GetInsertString_Item_RMReq(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_material_planning_item_rm_req` (`primary transaksi`,`primary product`,`qty`,`primary mo rec`," + _
            "`identifier jo rec`,`identifier wo rec`,`mo date rec`,`mo no rec`,`jo no rec`,`wo no rec`,`primary proses`, `MP Spare Qty`, `required date`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .mo_no_rec <> "" Then
                            If .mo_no_rec.Contains("\") = True Then
                                .mo_no_rec = .mo_no_rec.Replace("\", "\\")
                            End If
                            If .mo_no_rec.Contains(vbCrLf) = True Then
                                .mo_no_rec = .mo_no_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .mo_no_rec.Contains("'") = True Then
                                .mo_no_rec = .mo_no_rec.Replace("'", "''")
                            End If
                        End If
                        If .jo_no_rec <> "" Then
                            If .jo_no_rec.Contains("\") = True Then
                                .jo_no_rec = .jo_no_rec.Replace("\", "\\")
                            End If
                            If .jo_no_rec.Contains(vbCrLf) = True Then
                                .jo_no_rec = .jo_no_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .jo_no_rec.Contains("'") = True Then
                                .jo_no_rec = .jo_no_rec.Replace("'", "''")
                            End If
                        End If
                        If .wo_no_rec <> "" Then
                            If .wo_no_rec.Contains("\") = True Then
                                .wo_no_rec = .wo_no_rec.Replace("\", "\\")
                            End If
                            If .wo_no_rec.Contains(vbCrLf) = True Then
                                .wo_no_rec = .wo_no_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .wo_no_rec.Contains("'") = True Then
                                .wo_no_rec = .wo_no_rec.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_product).ToString + "," _
                        + CDec(.qty).ToString + "," + CDbl(.primary_mo_rec).ToString + "," + "'" + .identifier_jo_rec + "'" + "," + "'" + .identifier_wo_rec + "'" + "," _
                        + "'" + Format(CDate(.mo_date_rec), "yyyy-MM-dd") + "'" + "," + "'" + .mo_no_rec + "'" + "," + "'" + .jo_no_rec + "'" + "," _
                        + "'" + .wo_no_rec + "'" + "," + CDbl(.primary_proses).ToString + "," + CDec(.MP_Spare_Qty).ToString + ",'" + Format(.required_date, "yyyy-MM-dd") + "'" + "" _
                        + ")" + Delimiter
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

    Private Function Simpan_ItemMOB_UsingScript(ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable, ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemMOB(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function GetInsertString_ItemMOB(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_material_planning_itemmob` (`primary transaksi`,`primary product`,`qty`,`primary mo rec`," + _
            "`identifier jo rec`,`mo date rec`,`mo no rec`,`jo no rec`,`primary proses`,`identifier wo rec`,`wo no rec`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .mo_no_rec <> "" Then
                            If .mo_no_rec.Contains("\") = True Then
                                .mo_no_rec = .mo_no_rec.Replace("\", "\\")
                            End If
                            If .mo_no_rec.Contains(vbCrLf) = True Then
                                .mo_no_rec = .mo_no_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .mo_no_rec.Contains("'") = True Then
                                .mo_no_rec = .mo_no_rec.Replace("'", "''")
                            End If
                        End If
                        If .jo_no_rec <> "" Then
                            If .jo_no_rec.Contains("\") = True Then
                                .jo_no_rec = .jo_no_rec.Replace("\", "\\")
                            End If
                            If .jo_no_rec.Contains(vbCrLf) = True Then
                                .jo_no_rec = .jo_no_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .jo_no_rec.Contains("'") = True Then
                                .jo_no_rec = .jo_no_rec.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_product).ToString + "," _
                        + CDec(.qty).ToString + "," + CDbl(.primary_mo_rec).ToString + "," + "'" + .identifier_jo_rec + "'" + "," _
                        + "'" + Format(CDate(.mo_date_rec), "yyyy-MM-dd") + "'" + "," + "'" + .mo_no_rec + "'" + "," + "'" + .jo_no_rec + "'" _
                        + "," + CDbl(.primary_proses).ToString + "," + "'" + .identifier_wo_rec + "'" + "," + "'" + .wo_no_rec + "'" + "" _
                        + ")" + Delimiter
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

    Private Function Simpan_ItemExcLocation_UsingScript(ByVal vTBL As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, _
    ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemExcLocation(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function GetInsertString_ItemExcLocation(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_material_planning_item_location_exclusion` (`primary acuan`,`level`,`remarks`," + _
            "`include_exclude`,`primary transaksi`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        .remarks = .remarks.Trim
                        If .remarks <> "" Then
                            If .remarks.Contains("\") = True Then
                                .remarks = .remarks.Replace("\", "\\")
                            End If
                            If .remarks.Contains(vbCrLf) = True Then
                                .remarks = .remarks.Replace(vbCrLf, "\r\n")
                            End If
                            If .remarks.Contains("'") = True Then
                                .remarks = .remarks.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(.primary_acuan).ToString + "," _
                        + CDec(.level).ToString + "," _
                        + "'" + .remarks + "'" + "," _
                        + CDbl(.include_exclude).ToString + "," _
                        + CDbl(vPrimaryTransaksi).ToString + ")" + Delimiter
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

    Private Function Simpan_ItemQty_UsingScript(ByVal vTBL As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable, _
    ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemQty(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function GetInsertString_ItemQty(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_material_planning_item_qty` (`Primary Product`,`Qty`,`Lot No`," + _
            "`Expiry Date`,`Primary Storage Area`,`primary transaksi`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        .Lot_No = .Lot_No.Trim
                        If .Lot_No <> "" Then
                            If .Lot_No.Contains("\") = True Then
                                .Lot_No = .Lot_No.Replace("\", "\\")
                            End If
                            If .Lot_No.Contains(vbCrLf) = True Then
                                .Lot_No = .Lot_No.Replace(vbCrLf, "\r\n")
                            End If
                            If .Lot_No.Contains("'") = True Then
                                .Lot_No = .Lot_No.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(.Primary_Product).ToString + "," _
                        + CDec(.Qty).ToString + "," _
                        + "'" + .Lot_No + "'" + "," + "'" + Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'" + "," _
                        + CDbl(.Primary_Storage_Area).ToString + "," _
                        + CDbl(vPrimaryTransaksi).ToString + ")" + Delimiter
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

    Private Function Simpan_ItemDO_UsingScript(ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable, ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemDO(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function GetInsertString_ItemDO(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_material_planning_itemdo` (`primary transaksi`,`primary product`,`qty`," + _
            "`primary do rec`,`primary customer rec`,`do date rec`,`do no rec`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .do_no_rec <> "" Then
                            If .do_no_rec.Contains("\") = True Then
                                .do_no_rec = .do_no_rec.Replace("\", "\\")
                            End If
                            If .do_no_rec.Contains(vbCrLf) = True Then
                                .do_no_rec = .do_no_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .do_no_rec.Contains("'") = True Then
                                .do_no_rec = .do_no_rec.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_product).ToString + "," _
                        + CDec(.qty).ToString + "," + CDbl(.primary_do_rec).ToString + "," _
                        + CDbl(.primary_customer_rec).ToString + "," + "'" + Format(CDate(.do_date_rec), "yyyy-MM-dd") + "'" + "," _
                        + "'" + .do_no_rec + "'" _
                        + ")" + Delimiter
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

    Private Function Simpan_ItemSOB_UsingScript(ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemSOB(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function GetInsertString_ItemSOB(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_material_planning_itemsob` (`primary transaksi`,`primary product`,`qty`,`primary so rec`," + _
            "`primary customer rec`,`so date rec`,`so no rec`,`MP Spare Qty`,`qty sob multi`,`primary unit multi`,`delivery date cont`," + _
            "`ratio reference`,`ratio default`,`identifier so`,`primary unit`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .so_no_rec <> "" Then
                            If .so_no_rec.Contains("\") = True Then
                                .so_no_rec = .so_no_rec.Replace("\", "\\")
                            End If
                            If .so_no_rec.Contains(vbCrLf) = True Then
                                .so_no_rec = .so_no_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .so_no_rec.Contains("'") = True Then
                                .so_no_rec = .so_no_rec.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_product).ToString + "," _
                        + CDec(.qty).ToString + "," + CDbl(.primary_so_rec).ToString + "," _
                        + CDbl(.primary_customer_rec).ToString + "," + "'" + Format(CDate(.so_date_rec), "yyyy-MM-dd") + "'" + "," _
                        + "'" + .so_no_rec + "'," + CDec(.MP_Spare_Qty).ToString + "," + CDec(.qty_sob_multi).ToString + "," + CDbl(.primary_unit_multi).ToString _
                        + "," + CInt(.delivery_date_cont).ToString + "," + CDec(.ratio_reference).ToString + "," + CDec(.ratio_default).ToString _
                        + ",'" + .identifier_so + "'" + "," + CDbl(.primary_unit).ToString _
                        + ")" + Delimiter
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

    Private Function Simpan_ItemSOBDelievery_UsingScript(ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable, _
    ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemSOBDelivery(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function GetInsertString_ItemSOBDelivery(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_material_planning_itemsob_delivery` (`primary transaksi`,`primary so`,`identifier`,`delivery qty`,`delivery date`," + _
            "`remarks`,`delivery priority`,`delivery qty multi`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .remarks <> "" Then
                            If .remarks.Contains("\") = True Then
                                .remarks = .remarks.Replace("\", "\\")
                            End If
                            If .remarks.Contains(vbCrLf) = True Then
                                .remarks = .remarks.Replace(vbCrLf, "\r\n")
                            End If
                            If .remarks.Contains("'") = True Then
                                .remarks = .remarks.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_so).ToString + "," _
                        + "'" + .identifier + "'," + CDec(.delivery_qty).ToString + "," _
                        + "'" + Format(CDate(.delivery_date), "yyyy-MM-dd") + "'" + "," _
                        + "'" + .remarks + "'," + CInt(.delivery_priority).ToString + "," + CDec(.delivery_qty_multi).ToString _
                        + ")" + Delimiter
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

    Private Function Simpan_ItemPrOBDelivery_UsingScript(ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable, _
    ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemPrOBDelivery(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function GetInsertString_ItemPrOBDelivery(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_material_planning_itempurch_req_delivery` (`primary transaksi`,`primary purc req`," + _
            "`identifier`,`delivery qty`,`delivery date`,`remarks`,`delivery priority`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .remarks <> "" Then
                            If .remarks.Contains("\") = True Then
                                .remarks = .remarks.Replace("\", "\\")
                            End If
                            If .remarks.Contains(vbCrLf) = True Then
                                .remarks = .remarks.Replace(vbCrLf, "\r\n")
                            End If
                            If .remarks.Contains("'") = True Then
                                .remarks = .remarks.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_purc_req).ToString + "," _
                        + "'" + .identifier + "'," + CDec(.delivery_qty).ToString + "," _
                        + "'" + Format(.delivery_date, "yyyy-MM-dd") + "'," + "'" + .remarks + "'," + CDbl(.delivery_priority).ToString + ")" + Delimiter
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

    Private Function Simpan_ItemPrOB_UsingScript(ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, _
    ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemPrOB(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function GetInsertString_ItemPrOB(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_material_planning_itempurch_req` (`primary transaksi`,`primary product`,`qty`,`primary pr rec`," + _
            "`primary ref rec`,`pr date rec`,`pr no rec`, `remarks`,`MP Spare Qty`,`delivery date cont`,`identifier pr item`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .pr_no_rec <> "" Then
                            If .pr_no_rec.Contains("\") = True Then
                                .pr_no_rec = .pr_no_rec.Replace("\", "\\")
                            End If
                            If .pr_no_rec.Contains(vbCrLf) = True Then
                                .pr_no_rec = .pr_no_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .pr_no_rec.Contains("'") = True Then
                                .pr_no_rec = .pr_no_rec.Replace("'", "''")
                            End If
                        End If
                        If .remarks <> "" Then
                            If .remarks.Contains("\") = True Then
                                .remarks = .remarks.Replace("\", "\\")
                            End If
                            If .remarks.Contains(vbCrLf) = True Then
                                .remarks = .remarks.Replace(vbCrLf, "\r\n")
                            End If
                            If .remarks.Contains("'") = True Then
                                .remarks = .remarks.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_product).ToString + "," _
                        + CDec(.qty).ToString + "," + CDbl(.primary_pr_rec).ToString + "," _
                        + CDbl(.primary_ref_rec).ToString + "," + "'" + Format(CDate(.pr_date_rec), "yyyy-MM-dd") + "'" + "," _
                        + "'" + .pr_no_rec + "'," + "'" + .remarks + "'," + CDec(.MP_Spare_Qty).ToString + "," + .delivery_date_cont.ToString + ",'" + .identifier_pr_item + "'" + ")" + Delimiter
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

    Private Function Simpan_ItemGIT_UsingScript(ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable, ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemGIT(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function GetInsertString_ItemGIT(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_material_planning_itemgit` (`primary transaksi`,`primary product`,`qty`,`primary packing list rec`," + _
            "`primary suppier rec`,`packing list date rec`,`packing list no rec`,`supplier packing list no rec`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .packing_list_no_rec <> "" Then
                            If .packing_list_no_rec.Contains("\") = True Then
                                .packing_list_no_rec = .packing_list_no_rec.Replace("\", "\\")
                            End If
                            If .packing_list_no_rec.Contains(vbCrLf) = True Then
                                .packing_list_no_rec = .packing_list_no_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .packing_list_no_rec.Contains("'") = True Then
                                .packing_list_no_rec = .packing_list_no_rec.Replace("'", "''")
                            End If
                        End If
                        If .packing_list_no_rec <> "" Then
                            If .packing_list_no_rec.Contains("\") = True Then
                                .packing_list_no_rec = .packing_list_no_rec.Replace("\", "\\")
                            End If
                            If .packing_list_no_rec.Contains(vbCrLf) = True Then
                                .packing_list_no_rec = .packing_list_no_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .packing_list_no_rec.Contains("'") = True Then
                                .packing_list_no_rec = .packing_list_no_rec.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_product).ToString + "," _
                        + CDec(.qty).ToString + "," + CDbl(.primary_packing_list_rec).ToString + "," _
                        + CDbl(.primary_suppier_rec).ToString + "," + "'" + Format(CDate(.packing_list_date_rec), "yyyy-MM-dd") + "'" + "," _
                        + "'" + .packing_list_no_rec + "'" + "," + "'" + .supplier_packing_list_no_rec + "'" _
                        + ")" + Delimiter
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

    Private Function Simpan_ItemPOB_UsingScript(ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable, ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemPOB(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function GetInsertString_ItemPOB(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_material_planning_itempob` (`primary transaksi`,`primary product`,`qty`,`primary po rec`," + _
            "`primary supplier rec`,`po date rec`,`po no rec`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .po_no_rec <> "" Then
                            If .po_no_rec.Contains("\") = True Then
                                .po_no_rec = .po_no_rec.Replace("\", "\\")
                            End If
                            If .po_no_rec.Contains(vbCrLf) = True Then
                                .po_no_rec = .po_no_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .po_no_rec.Contains("'") = True Then
                                .po_no_rec = .po_no_rec.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_product).ToString + "," _
                        + CDec(.qty).ToString + "," + CDbl(.primary_po_rec).ToString + "," + CDbl(.primary_supplier_rec).ToString + "," _
                        + "'" + Format(CDate(.po_date_rec), "yyyy-MM-dd") + "'" + "," + "'" + .po_no_rec + "'" + "" _
                        + ")" + Delimiter
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

    Private Function Simpan_Item_UsingScript(ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable, ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_Item(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
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

    Private Function GetInsertString_Item(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_material_planning_item` (`primary transaksi`,`primary code no`,`primary unit`,`qty on hand`,`qty pob`,`qty git`,`qty inquiry`,`qty pending do`," + _
            "`qty sob`,`qty safety stock`,`qty lead time`,`qty required`,`default prec unit`,`qty order maksimum`,`qty order minimum`,`qty order optimum`,`qty max`,`qty min`," + _
            "`qty optimum`,`qty order`,`primary order unit`,`def prec order unit`,`order point`,`procurement type`,`primary prefered supplier`,`primary currency`,`unit price`," + _
            "`discount`,`nett price`,`sub total`,`def prec currency`,`order strategy`,`next order point`,`qty_order_multiunit`,`lead time days`,`order point delay`,`next order point delay`," + _
            "`qty avg montly demand`,`supplier selection mode`,`ExplainAvgQty`,`Explain Safety Stock`,`Explain Lead Time`,`SafetyStockStrategy`,`Lead Time Strategy`,`modify`, `procurement strategy`," + _
            "`optimum stock strategy`, `explain optimum stock`, `qty prob`, `product type`, `primary unit multi`, `ratio default`, `ratio multi`, `def prec unit multi`, `disc amount`, `disc from perct`" + _
            ", `Tolerance Down`, `Tolerance Up`,`qty mob`,`qty rm req`, `MP Spare Qty SOB`, `MP Spare Qty PrOB`, `MP Spare Qty Rm Req`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_code_no).ToString + "," _
                        + CDbl(.primary_unit).ToString + "," + CDec(.qty_on_hand).ToString + "," _
                        + CDec(.qty_pob).ToString + "," + CDec(.qty_git).ToString + "," _
                        + CDbl(.qty_inquiry).ToString + "," + CDec(.qty_pending_do).ToString + "," _
                        + CDbl(.qty_sob).ToString + "," + CDec(.qty_safety_stock).ToString + "," _
                        + CDbl(.qty_lead_time).ToString + "," + CDec(.qty_required).ToString + "," _
                        + CInt(.default_prec_unit).ToString + "," + CDec(.qty_order_maksimum).ToString + "," _
                        + CDec(.qty_order_minimum).ToString + "," + CDec(.qty_order_optimum).ToString + "," _
                        + CDec(.qty_max).ToString + "," + CDec(.qty_min).ToString + "," _
                        + CDec(.qty_optimum).ToString + "," + CDec(.qty_order).ToString + "," _
                        + CDbl(.primary_order_unit).ToString + "," + CInt(.def_prec_order_unit).ToString + "," _
                        + "'" + Format(CDate(.order_point), "yyyy-MM-dd") + "'," + CInt(.procurement_type).ToString + "," _
                        + CDbl(.primary_prefered_supplier).ToString + "," + CDec(.primary_currency).ToString + "," _
                        + CDec(.unit_price).ToString + "," + CDec(.discount).ToString + "," _
                        + CDec(.nett_price).ToString + "," + CDec(.sub_total).ToString + "," _
                        + CDec(.def_prec_currency).ToString + "," + CInt(.order_strategy).ToString + "," _
                        + "'" + Format(CDate(.next_order_point), "yyyy-MM-dd") + "'," + CDec(.qty_order_multiunit).ToString + "," _
                        + CDec(.lead_time_days).ToString + "," + CDbl(.order_point_delay).ToString + "," + CDbl(.next_order_point_delay).ToString + "," _
                        + CDec(.qty_avg_montly_demand).ToString + "," + CInt(.supplier_selection_mode).ToString + "," _
                        + "'" + .ExplainAvgQty + "'" + "," + "'" + .Explain_Safety_Stock + "'" + "," _
                        + "'" + .Explain_Lead_Time + "'" + "," + CInt(.SafetyStockStrategy).ToString + "," _
                        + CInt(.Lead_Time_Strategy).ToString + "," + CInt(.modify).ToString + "," + CInt(.procurement_strategy).ToString + "," _
                        + CInt(.optimum_stock_strategy).ToString + "," + "'" + .explain_optimum_stock + "'" + "," _
                        + CDec(.qty_prob).ToString + "," + .product_type.ToString + "," _
                        + CDbl(.primary_unit_multi).ToString + "," + CDec(.ratio_default).ToString + "," _
                        + CDec(.ratio_multi).ToString + "," + CDbl(.def_prec_unit_multi).ToString + "," _
                        + CDec(.disc_amount).ToString + "," + CInt(.disc_from_perct).ToString + "," _
                        + CDec(.Tolerance_Down).ToString + "," + CDec(.Tolerance_Up).ToString + "," _
                        + CDec(.qty_mob).ToString + "," + CDec(.qty_rm_req).ToString + "," + _
                        CDec(.MP_Spare_Qty_SOB).ToString + "," + CDec(.MP_Spare_Qty_PrOB).ToString + "," + CDec(.MP_Spare_Qty_Rm_Req).ToString + ")" + Delimiter
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

    Public Sub Hapus(ByVal vPrimaryTransaksi As Double)
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itemTableAdapter
        Dim DTAGIT As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itemgitTableAdapter
        Dim DTADO As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itemdoTableAdapter
        Dim DTAPOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itempobTableAdapter
        Dim DTASOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itemsobTableAdapter
        Dim DTASOBDelivery As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itemsob_deliveryTableAdapter
        Dim DTAPrOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itempurch_reqTableAdapter
        Dim DTAPrOBDelivery As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itempurch_req_deliveryTableAdapter
        Dim DTALocationSetting As New DseTransaksiMaterialPlanning_WarehouseExcTableAdapters.tbltrans_material_planning_item_location_exclusionTableAdapter
        Dim DTAItemQty As New DseTransaksiMaterialPlanning_WarehouseExcTableAdapters.tbltrans_material_planning_item_qtyTableAdapter
        Dim DTAItemFollowUp As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_item_followupTableAdapter
        Dim DTAMOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itemmobTableAdapter
        Dim DTARMReq As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_item_rm_reqTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAGIT.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTADO.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAPOB.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTASOB.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTASOBDelivery.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAPrOB.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAMOB.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAPrOB.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAPrOBDelivery.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTARMReq.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTAItemFollowUp.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemFollowUp.DeleteQuery(vPrimaryTransaksi)

            DTA.Delete(vPrimaryTransaksi)
            DTAGIT.Delete(vPrimaryTransaksi)
            DTADO.Delete(vPrimaryTransaksi)
            DTAPOB.Delete(vPrimaryTransaksi)
            DTASOB.Delete(vPrimaryTransaksi)
            DTASOBDelivery.DeleteQuery(vPrimaryTransaksi)
            DTAPrOB.Delete(vPrimaryTransaksi)
            DTAPrOBDelivery.DeleteQuery(vPrimaryTransaksi)
            DTALocationSetting.Delete(vPrimaryTransaksi)
            DTAItemQty.Delete(vPrimaryTransaksi)
            DTAMOB.Delete(vPrimaryTransaksi)
            DTARMReq.Delete(vPrimaryTransaksi)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAGIT.Dispose()
            DTADO.Dispose()
            DTAPOB.Dispose()
            DTASOB.Dispose()
            DTAPrOB.Dispose()
            DTALocationSetting.Dispose()
            DTAItemQty.Dispose()
        End Try
    End Sub

    Public Sub IsiDGV(ByVal vPrimaryTransaksi As Double, _
    ByVal vDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable, _
    ByVal vDT_ItemGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable, _
    ByVal vDT_ItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable, _
    ByVal vDT_ItemPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, _
    ByVal vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, _
    ByVal vDT_ItemSettingLocation As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, _
    ByVal vDT_ItemQty As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable, _
    ByVal vDT_ItemMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable, _
    ByVal vDT_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable, _
    ByVal vDT_ItemFollowUp As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupDataTable, _
    ByVal vDT_ItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable, _
    ByVal vDT_ItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable)
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_item1TableAdapter
        Dim DTAGIT As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itemgit1TableAdapter
        Dim TBLDummyGIT As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgit1DataTable
        Dim rowDummyGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgit1Row
        Dim rowGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitRow
        Dim DTADummyDO As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itemdo1TableAdapter
        Dim TBLDummyDO As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdo1DataTable
        Dim rowDummyDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdo1Row
        Dim rowDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoRow
        Dim DTADummyPOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itempob1TableAdapter
        Dim TBLDummyPOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempob1DataTable
        Dim rowDummyPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempob1Row
        Dim rowPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobRow
        Dim DTADummySOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itemsob1TableAdapter
        Dim TBLDummySOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob1DataTable
        Dim rowDummySOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob1Row
        Dim rowSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobRow
        Dim DTADummyPrOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itempurch_req2TableAdapter
        Dim TBLDummyPrOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req2DataTable
        Dim rowDummyPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req2Row
        Dim rowPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqRow
        Dim DTAQtyOnHand As New DseTransaksiMaterialPlanning_WarehouseExcTableAdapters.tbltrans_material_planning_item_qtyTableAdapter
        Dim DTALocationSetting As New DseTransaksiMaterialPlanning_WarehouseExcTableAdapters.tbltrans_material_planning_item_location_exclusionTableAdapter
        Dim DTAItemQty As New DseTransaksiMaterialPlanning_WarehouseExcTableAdapters.tbltrans_material_planning_item_qtyTableAdapter
        Dim DTADummyMOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itemmob1TableAdapter
        Dim TBLDummyMOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmob1DataTable
        Dim rowDummyMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmob1Row
        Dim rowMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobRow
        Dim DTADummyRMReq As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_item_rm_req1TableAdapter
        Dim TBLDummyRMReq As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_req1DataTable
        Dim rowDummyRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_req1Row
        Dim rowRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow
        Dim DTAFollowUp As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_item_followupTableAdapter
        Dim DTASOBDelievery As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itemsob_deliveryTableAdapter
        Dim DTAPrOBDelivery As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itempurch_req_deliveryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTASOBDelievery.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAGIT.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAQtyOnHand.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTALocationSetting.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTADummyDO.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTADummyPOB.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTADummySOB.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAItemQty.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTALocationSetting.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTADummyMOB.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTADummyRMReq.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAPrOBDelivery.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTAFollowUp.Connection.ConnectionString = Core_Module.GetConstring
            DTAFollowUp.Fill(vDT_ItemFollowUp, vPrimaryTransaksi)

            vDT_ItemQty.Rows.Clear()
            DTAItemQty.FillBy_Primary_Transaksi(vDT_ItemQty, vPrimaryTransaksi)

            vDT_ItemSettingLocation.Rows.Clear()
            DTALocationSetting.FillBy_PrimaryTransaksi(vDT_ItemSettingLocation, vPrimaryTransaksi)

            vDT.Rows.Clear()
            DTA.Fill(vDT, vPrimaryTransaksi)

            vDT_ItemGIT.Rows.Clear()
            DTAGIT.Fill(TBLDummyGIT, vPrimaryTransaksi)
            For Each rowDummyGIT In TBLDummyGIT.Rows
                rowGIT = vDT_ItemGIT.Newtbltrans_material_planning_itemgitRow

                With rowGIT
                    .No_Packing_List = rowDummyGIT.No_Transaksi
                    .Tgl_Packing_List = rowDummyGIT.Tanggal_Transaksi
                    .Supplier_Packing_List_No = rowDummyGIT.Supplier_Packing_List_No

                    .packing_list_date_rec = rowDummyGIT.packing_list_date_rec
                    .packing_list_no_rec = rowDummyGIT.packing_list_no_rec
                    .primary = rowDummyGIT.primary
                    .primary_packing_list_rec = rowDummyGIT.primary_packing_list_rec
                    .primary_product = rowDummyGIT.primary_product
                    .primary_suppier_rec = rowDummyGIT.primary_suppier_rec
                    .primary_transaksi = vPrimaryTransaksi
                    .qty = rowDummyGIT.qty
                    .supplier_packing_list_no_rec = rowDummyGIT.supplier_packing_list_no_rec

                    If rowDummyGIT.IsPrimary_SupplierNull = False Then
                        If rowDummyGIT.primary_suppier_rec <> rowDummyGIT.Primary_Supplier Then
                            'jika tidak sama primary supplier dan yang terrekam, packing list sudah pernah diedit
                            'ambil data dari record
                            .No_Packing_List = rowDummyGIT.packing_list_no_rec
                            .Tgl_Packing_List = rowDummyGIT.packing_list_date_rec
                            .Supplier_Packing_List_No = rowDummyGIT.supplier_packing_list_no_rec
                        End If
                    Else
                        .No_Packing_List = rowDummyGIT.packing_list_no_rec
                        .Tgl_Packing_List = rowDummyGIT.packing_list_date_rec
                        .Supplier_Packing_List_No = rowDummyGIT.supplier_packing_list_no_rec
                    End If
                End With

                vDT_ItemGIT.Rows.Add(rowGIT)
            Next

            vDT_ItemDO.Rows.Clear()
            DTADummyDO.Fill(TBLDummyDO, vPrimaryTransaksi)
            For Each rowDummyDO In TBLDummyDO.Rows
                rowDO = vDT_ItemDO.Newtbltrans_material_planning_itemdoRow

                With rowDO
                    .do_date_rec = rowDummyDO.do_date_rec
                    .do_no_rec = rowDummyDO.do_no_rec
                    .primary = rowDummyDO.primary
                    .primary_customer_rec = rowDummyDO.primary_customer_rec
                    .primary_do_rec = rowDummyDO.primary_do_rec
                    .primary_product = rowDummyDO.primary_product
                    .primary_transaksi = rowDummyDO.primary_transaksi
                    .qty = rowDummyDO.qty
                    .DO_Number = rowDummyDO.DO_Number
                    .DO_Date = rowDummyDO.Do_Date

                    'jika do sudah dihapuskan atau ada perubahan customer pada do, maka ambil data record
                    If rowDummyDO.IsPrimary_CustomerNull = False Then
                        If rowDummyDO.Primary_Customer <> rowDummyDO.primary_customer_rec Then
                            .DO_Date = .do_date_rec
                            .DO_Number = .do_no_rec
                        End If
                    Else
                        .DO_Date = .do_date_rec
                        .DO_Number = .do_no_rec
                    End If
                End With

                vDT_ItemDO.Rows.Add(rowDO)
            Next

            vDT_ItemPOB.Rows.Clear()
            DTADummyPOB.Fill(TBLDummyPOB, vPrimaryTransaksi)
            For Each rowDummyPOB In TBLDummyPOB.Rows
                rowPOB = vDT_ItemPOB.Newtbltrans_material_planning_itempobRow

                With rowPOB
                    .No_Transaksi = rowDummyPOB.No_Transaksi
                    .po_date_rec = rowDummyPOB.po_date_rec
                    .po_no_rec = rowDummyPOB.po_no_rec
                    .primary = rowDummyPOB.primary
                    .primary_po_rec = rowDummyPOB.primary_po_rec
                    .primary_product = rowDummyPOB.primary_product
                    .primary_supplier_rec = rowDummyPOB.primary_supplier_rec
                    .primary_transaksi = rowDummyPOB.primary_transaksi
                    .qty = rowDummyPOB.qty
                    .Tanggal_Transaksi = rowDummyPOB.Tanggal_Transaksi

                    'jika do sudah dihapuskan atau ada perubahan customer pada do, maka ambil data record
                    If rowDummyPOB.IsPrimary_SupplierNull = False Then
                        If rowDummyPOB.Primary_Supplier <> rowDummyPOB.primary_supplier_rec Then
                            .No_Transaksi = .po_no_rec
                            .Tanggal_Transaksi = .po_date_rec
                        End If
                    Else
                        .No_Transaksi = .po_no_rec
                        .Tanggal_Transaksi = .po_date_rec
                    End If
                End With

                vDT_ItemPOB.Rows.Add(rowPOB)
            Next

            vDT_ItemSOBDelivery.Rows.Clear()
            DTASOBDelievery.Fill(vDT_ItemSOBDelivery, vPrimaryTransaksi)

            vDT_ItemSOB.Rows.Clear()
            DTADummySOB.Fill(TBLDummySOB, vPrimaryTransaksi)
            For Each rowDummySOB In TBLDummySOB.Rows
                rowSOB = vDT_ItemSOB.Newtbltrans_material_planning_itemsobRow

                With rowSOB
                    If rowDummySOB.IsNo_TransaksiNull = True Then
                        .No_Transaksi = ""
                        .Tanggal_Transaksi = #1/1/1950#
                    Else
                        .No_Transaksi = rowDummySOB.No_Transaksi
                        .Tanggal_Transaksi = rowDummySOB.Tanggal_Transaksi
                    End If
                    .primary = rowDummySOB.primary
                    .primary_customer_rec = rowDummySOB.primary_customer_rec
                    .primary_product = rowDummySOB.primary_product
                    .primary_so_rec = rowDummySOB.primary_so_rec
                    .primary_transaksi = rowDummySOB.primary_transaksi
                    .qty = rowDummySOB.qty
                    .so_date_rec = rowDummySOB.so_date_rec
                    .so_no_rec = rowDummySOB.so_no_rec
                    .primary_unit_multi = rowDummySOB.primary_unit_multi
                    .ratio_default = rowDummySOB.ratio_default
                    .ratio_reference = rowDummySOB.ratio_reference
                    .qty_sob_multi = rowDummySOB.qty_sob_multi
                    .delivery_date_cont = rowDummySOB.delivery_date_cont
                    .identifier_so = rowDummySOB.identifier_so
                    .primary_unit = rowDummySOB.primary_unit

                    'jika do sudah dihapuskan atau ada perubahan customer pada do, maka ambil data record
                    If rowDummySOB.IsPrimary_CustomerNull = False Then
                        If rowDummySOB.Primary_Customer <> rowDummySOB.primary_customer_rec Then
                            .No_Transaksi = .so_no_rec
                            .Tanggal_Transaksi = .so_date_rec
                        End If
                    Else
                        .No_Transaksi = .so_no_rec
                        .Tanggal_Transaksi = .so_date_rec
                    End If
                    .MP_Spare_Qty = rowDummySOB.MP_Spare_Qty
                End With

                vDT_ItemSOB.Rows.Add(rowSOB)
            Next

            vDT_ItemPrOBDelivery.Rows.Clear()
            DTAPrOBDelivery.Fill(vDT_ItemPrOBDelivery, vPrimaryTransaksi)

            vDT_ItemPrOB.Rows.Clear()
            DTADummyPrOB.Fill(TBLDummyPrOB, vPrimaryTransaksi)

            For Each rowDummyPrOB In TBLDummyPrOB.Rows
                rowPrOB = vDT_ItemPrOB.Newtbltrans_material_planning_itempurch_reqRow

                With rowPrOB
                    If rowDummyPrOB.no_transaksi = "" Then
                        .Tanggal_Transaksi = #1/1/1950#
                        .No_Transaksi = ""
                    Else
                        .No_Transaksi = rowDummyPrOB.no_transaksi
                        .Tanggal_Transaksi = rowDummyPrOB.tanggal
                    End If
                    .primary = rowDummyPrOB.primary
                    .primary_ref_rec = rowDummyPrOB.primary_ref_rec
                    .primary_product = rowDummyPrOB.primary_product
                    .primary_pr_rec = rowDummyPrOB.primary_pr_rec
                    .primary_transaksi = rowDummyPrOB.primary_transaksi
                    .qty = rowDummyPrOB.qty
                    .pr_date_rec = rowDummyPrOB.pr_date_rec
                    .pr_no_rec = rowDummyPrOB.pr_no_rec
                    .remarks = rowDummyPrOB.remarks
                    .MP_Spare_Qty = rowDummyPrOB.MP_Spare_Qty
                    .delivery_date_cont = rowDummyPrOB.delivery_date_cont
                    .identifier_pr_item = rowDummyPrOB.identifier_pr_item
                End With

                vDT_ItemPrOB.Rows.Add(rowPrOB)
            Next

            vDT_ItemMOB.Rows.Clear()
            DTADummyMOB.Fill(TBLDummyMOB, vPrimaryTransaksi)
            For Each rowDummyMOB In TBLDummyMOB.Rows
                rowMOB = vDT_ItemMOB.Newtbltrans_material_planning_itemmobRow
                With rowMOB
                    .identifier_jo_rec = rowDummyMOB.identifier_jo_rec
                    If rowDummyMOB.JO_Number = "" Then
                        .jo_no_rec = rowDummyMOB.jo_no_rec
                    Else
                        .jo_no_rec = rowDummyMOB.JO_Number
                    End If
                    If rowDummyMOB.MO_Number = "" Then
                        .mo_date_rec = rowDummyMOB.mo_date_rec
                        .mo_no_rec = rowDummyMOB.mo_no_rec
                    Else
                        .mo_date_rec = rowDummyMOB.MO_Date
                        .mo_no_rec = rowDummyMOB.mo_no_rec
                    End If
                    .primary = rowDummyMOB.primary
                    .primary_mo_rec = rowDummyMOB.primary_mo_rec
                    .primary_product = rowDummyMOB.primary_product
                    .primary_transaksi = rowDummyMOB.primary_transaksi
                    .qty = rowDummyMOB.qty
                    .primary_proses = rowDummyMOB.primary_proses
                    .identifier_wo_rec = rowDummyMOB.identifier_wo_rec
                    .wo_no_rec = rowDummyMOB.wo_no_rec
                End With
                vDT_ItemMOB.Rows.Add(rowMOB)
            Next

            vDT_ItemRMReq.Rows.Clear()
            DTADummyRMReq.Fill(TBLDummyRMReq, vPrimaryTransaksi)
            For Each rowDummyRMReq In TBLDummyRMReq.Rows
                rowRMReq = vDT_ItemRMReq.Newtbltrans_material_planning_item_rm_reqRow
                With rowRMReq
                    .identifier_wo_rec = rowDummyRMReq.identifier_wo_rec
                    If rowDummyRMReq.WO_Number = "" Then
                        .wo_no_rec = rowDummyRMReq.wo_no_rec
                    Else
                        .wo_no_rec = rowDummyRMReq.WO_Number
                    End If
                    .identifier_jo_rec = rowDummyRMReq.identifier_jo_rec
                    If rowDummyRMReq.JO_Number = "" Then
                        .jo_no_rec = rowDummyRMReq.jo_no_rec
                    Else
                        .jo_no_rec = rowDummyRMReq.JO_Number
                    End If
                    If rowDummyRMReq.MO_Number = "" Then
                        .mo_date_rec = rowDummyRMReq.mo_date_rec
                        .mo_no_rec = rowDummyRMReq.mo_no_rec
                    Else
                        .mo_date_rec = rowDummyRMReq.MO_Date
                        .mo_no_rec = rowDummyRMReq.mo_no_rec
                    End If
                    .primary = rowDummyRMReq.primary
                    .primary_mo_rec = rowDummyRMReq.primary_mo_rec
                    .primary_product = rowDummyRMReq.primary_product
                    .primary_transaksi = rowDummyRMReq.primary_transaksi
                    .qty = rowDummyRMReq.qty
                    .primary_proses = rowDummyRMReq.primary_proses
                    .is_wip = rowDummyRMReq.is_wip
                    .MP_Spare_Qty = rowDummyRMReq.MP_Spare_Qty
                    .required_date = rowDummyRMReq.required_date
                End With
                vDT_ItemRMReq.Rows.Add(rowRMReq)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAGIT.Dispose()
            TBLDummyGIT.Dispose()
            DTADummyDO.Dispose()
            TBLDummyDO.Dispose()
            DTADummyPOB.Dispose()
            TBLDummyPOB.Dispose()
            DTADummySOB.Dispose()
            TBLDummySOB.Dispose()
            DTADummyPrOB.Dispose()
            TBLDummyPrOB.Dispose()
            DTAItemQty.Dispose()
            DTALocationSetting.Dispose()
        End Try
    End Sub
End Class
