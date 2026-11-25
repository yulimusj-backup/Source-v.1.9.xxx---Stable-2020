Option Explicit On
Option Strict Off

Imports Core_Module
Public Class ClsMasterCuti_Item
    Dim ObjKaryawan As New Core_Module.ClsMasterKaryawan
    Dim ObjCd As New Core_Module.ClsCekData

    Public Enum EnumDeleteItem As Integer
        DeleteAll
        DeleteItem
        DeleteItemDesc
    End Enum

    Public Function IsiDGVItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryCuti As Double, _
    ByVal Frm As FrmMasterCuti, ByVal vObjCd As Core_Module.ClsCekData, ByVal vTblItemDesc As DseMasterCuti.tblmaster_cuti_item_descDataTable) As Boolean
        Dim DTA As New DseMasterCutiTableAdapters.tblmaster_cuti_item2TableAdapter
        Dim TBL As New DseMasterCuti.tblmaster_cuti_item2DataTable
        Dim row As DseMasterCuti.tblmaster_cuti_item2Row
        Dim rowDGV As DataGridViewRow

        ''@rosadinnn
        ''@20150909 1610
        Dim DtaItemDesc As New DseMasterCutiTableAdapters.tblmaster_cuti_item_descTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DtaItemDesc.Connection.ConnectionString = Core_Module.GetConstring

            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            ''isikan(tabel)
            DtaItemDesc.FillBy_LeavePrimary(vTblItemDesc, vPrimaryCuti)
            If DTA.FillBy_PrimaryCuti(TBL, vPrimaryCuti) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColNo.Index).Value = vDGVItem.RowCount
                        rowDGV.Cells(Frm.ColCode.Index).Value = .Employee_Code 'Me.ObjKaryawan.GetCodeByPrimary(.primary_karyawan)
                        rowDGV.Cells(Frm.ColName.Index).Value = .Employee_Name 'Me.ObjKaryawan.GetFullNameByPrimary(.primary_karyawan)
                        rowDGV.Cells(Frm.ColLeaveAddition.Index).Value = Me.ObjCd.DigitGrouping(.leave_addition, 1)
                        rowDGV.Cells(Frm.ColLeaveDeduction.Index).Value = Me.ObjCd.DigitGrouping(.leave_deduction, 1)
                        rowDGV.Cells(Frm.ColBalance.Index).Value = Me.ObjCd.DigitGrouping(.balance, 1)
                        rowDGV.Cells(Frm.ColRemarks.Index).Value = .remarks
                        rowDGV.Cells(Frm.ColPrimary.Index).Value = .primary_karyawan
                        rowDGV.Cells(Frm.ColIdentifier.Index).Value = .identifier
                    End With
                Next
            End If

            'set nomer

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Function

    Public Sub SimpanItem(ByVal vDgv As DataGridView, ByVal vPrimaryCuti As Double, ByVal frm As FrmMasterCuti, ByVal vTblDesc As DseMasterCuti.tblmaster_cuti_item_descDataTable)
        ''@rosadinnn
        ''20150909 1556
        Dim DtaDesc As New DseMasterCutiTableAdapters.tblmaster_cuti_item_descTableAdapter
        Dim RowDesc As DseMasterCuti.tblmaster_cuti_item_descRow

        Dim Dta As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim row As DataGridViewRow
        Dim Remarks As String = ""

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            DtaDesc.Connection.ConnectionString = Core_Module.GetConstring

            If Me.SimpanUsingScript_Item(vDgv, vPrimaryCuti, frm) = False Then
                'jika ada error saat menggunakan using scrip maka simpan dengan cara manual
                Me.HapusItem(vPrimaryCuti, EnumDeleteItem.DeleteItem)
                For Each row In vDgv.Rows
                    With row
                        If .Cells(frm.ColPrimary.Index).Value <> -1 Then
                            If String.IsNullOrEmpty(.Cells(frm.ColRemarks.Index).Value) = False Then
                                Remarks = .Cells(frm.ColRemarks.Index).Value.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
                            Else
                                Remarks = ""
                            End If
                            ''@rosadinnn
                            ''@20150908 1433
                            If row.Cells(frm.ColIdentifier.Index).Value.ToString.StartsWith("-1") Then
                                row.Cells(frm.ColIdentifier.Index).Value = vPrimaryCuti.ToString + row.Cells(frm.ColIdentifier.Index).Value.ToString.Substring(2)
                            End If
                            Dta.Insert(vPrimaryCuti, .Cells(frm.ColPrimary.Index).Value, Remarks, _
                            .Cells(frm.ColLeaveAddition.Index).Value, .Cells(frm.ColLeaveDeduction.Index).Value, .Cells(frm.ColIdentifier.Index).Value)
                        End If
                    End With
                Next
            End If

            ''@rosadinnn
            ''@20150909 1622
            If Me.SimpanUsingScript_ItemDesc(vTblDesc, vPrimaryCuti) = False Then
                'jika ada error saat menggunakan using scrip maka simpan dengan cara manual
                Me.HapusItem(vPrimaryCuti, EnumDeleteItem.DeleteItemDesc)
                For Each RowDesc In vTblDesc.Rows
                    With RowDesc
                        If .identifier.StartsWith("-1") = True Then
                            .identifier = vPrimaryCuti.ToString + .identifier.Substring(2)
                        End If
                        .desc = .desc.Replace("'", "").Replace("*", "").Replace("%", "")
                        DtaDesc.Insert(vPrimaryCuti, .identifier, .tipe, ._date, .desc, .qty, .tipe2)
                    End With
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Sub

    Private Function SimpanUsingScript_Item(ByVal vDgv As DataGridView, ByVal vPrimaryCuti As Double, ByVal vfrm As FrmMasterCuti) As Boolean
        Dim conStr As String = My.Settings.Core_DbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()
            vDgv.AllowUserToAddRows = False
            While StartRowIndex < vDgv.Rows.Count
                sql = Me.GetInsertString_Item(StartRowIndex, BatchQty, vDgv, vPrimaryCuti, vfrm)
                If sql = "Error" Then
                    Return False
                End If
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While
            vDgv.AllowUserToAddRows = True

            Conn.Close()

            Return True
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_Item(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryCuti As Double, ByVal vfrm As FrmMasterCuti) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","

        Try
            retval = "INSERT INTO `tblmaster_cuti_item` (`primary leave`,`primary karyawan`,`remarks`,`leave addition`,`leave deduction`,`balance`,`identifier`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem.Rows.Count > vStartRowIndex + i Then
                    row = vdgvItem.Rows(vStartRowIndex + i)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    If String.IsNullOrEmpty(row.Cells(vfrm.ColRemarks.Index).Value) = False Then
                        row.Cells(vfrm.ColRemarks.Index).Value = row.Cells(vfrm.ColRemarks.Index).Value.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
                    Else
                        row.Cells(vfrm.ColRemarks.Index).Value = ""
                    End If
                    If row.Cells(vfrm.ColIdentifier.Index).Value.ToString.StartsWith("-1") Then
                        row.Cells(vfrm.ColIdentifier.Index).Value = vPrimaryCuti.ToString + row.Cells(vfrm.ColIdentifier.Index).Value.ToString.Substring(2)
                    End If

                    retval = retval + vbCrLf + "(" + CDbl(vPrimaryCuti).ToString + "," _
                    + CDbl(row.Cells(vfrm.ColPrimary.Index).Value).ToString + "," _
                    + "'" + row.Cells(vfrm.ColRemarks.Index).Value.ToString + "'," _
                    + CDec(row.Cells(vfrm.ColLeaveAddition.Index).Value).ToString + "," _
                    + CDec(row.Cells(vfrm.ColLeaveDeduction.Index).Value).ToString + "," _
                    + CDec(row.Cells(vfrm.ColBalance.Index).Value).ToString + "," _
                    + "'" + row.Cells(vfrm.ColIdentifier.Index).Value.ToString + "'" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Error"
        End Try
    End Function

    Private Function SimpanUsingScript_ItemDesc(ByVal vTblDesc As DseMasterCuti.tblmaster_cuti_item_descDataTable, ByVal vPrimaryCuti As Double) As Boolean
        Dim conStr As String = My.Settings.Core_DbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTblDesc.Rows.Count
                sql = Me.GetInsertString_ItemDesc(StartRowIndex, BatchQty, vTblDesc, vPrimaryCuti)
                If sql = "Error" Then
                    Return False
                End If
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemDesc(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vTblDesc As DseMasterCuti.tblmaster_cuti_item_descDataTable, _
        ByVal vPrimaryCuti As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseMasterCuti.tblmaster_cuti_item_descRow

        Try
            retval = "INSERT INTO `tblmaster_cuti_item_desc` (`primary leave`,`identifier`,`tipe`,`date`,`desc`,`qty`,`tipe2`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTblDesc.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTblDesc.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTblDesc.Rows(vStartRowIndex + i)
                    If row.identifier.StartsWith("-1") = True Then
                        row.identifier = vPrimaryCuti.ToString + row.identifier.Substring(2)
                    End If
                    row.desc = row.desc.Replace("'", "").Replace("*", "").Replace("%", "")

                    retval = retval + vbCrLf + "(" + CDbl(vPrimaryCuti).ToString + "," _
                    + "'" + row.identifier.ToString + "'," _
                    + CInt(row.tipe).ToString + "," _
                    + "'" + Format(row._date, "yyyy-MM-dd").ToString + "'," _
                    + "'" + row.desc.ToString + "'," _
                    + CDbl(row.qty).ToString + "," _
                    + CInt(row.tipe2).ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Error"
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryCuti As Double, Optional ByVal vEnumDelete As EnumDeleteItem = EnumDeleteItem.DeleteAll) As Boolean
        Dim DtaItem As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter

        ''@rosadinnn
        ''20150909 1606
        Dim DtaItemDesc As New DseMasterCutiTableAdapters.tblmaster_cuti_item_descTableAdapter
        Dim ApakahHapusItem As Boolean = False
        Dim ApakahHapusItemDesc As Boolean = False

        Try
            'hapus data item
            DtaItem.Connection.ConnectionString = Core_Module.GetConstring
            DtaItemDesc.Connection.ConnectionString = Core_Module.GetConstring

            If vEnumDelete = EnumDeleteItem.DeleteAll Then
                ApakahHapusItem = True
                ApakahHapusItemDesc = True
            ElseIf vEnumDelete = EnumDeleteItem.DeleteItem Then
                ApakahHapusItem = True
            ElseIf vEnumDelete = EnumDeleteItem.DeleteItemDesc Then
                ApakahHapusItemDesc = True
            End If

            If ApakahHapusItem = True Then
                DtaItem.Delete(vPrimaryCuti)
            End If
            If ApakahHapusItemDesc = True Then
                DtaItemDesc.Delete(vPrimaryCuti)
            End If

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaItem.Dispose()
        End Try
    End Function

    Public Function GetAdditionByPrimary(ByVal vPrimary As Double) As Decimal
        Dim dta As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim Dummy As Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetAdditionByPrimary(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetDeductionByPrimary(ByVal vPrimary As Double) As Decimal
        Dim dta As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim Dummy As Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetDeductionByPrimary(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetRemarksByPrimary(ByVal vPrimary As Double) As Decimal
        Dim dta As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetDeductionByPrimary(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetJmlDiLeaveApp(ByVal vPrimaryCuti As Double, ByVal vPrimaryKaryawan As Double) As Long
        Dim dta As New DseMasterCutiTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetJmlByPrimaryCutiNKaryawan(vPrimaryCuti, vPrimaryKaryawan)
            If Dummy.HasValue = False Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetAddition(ByVal vPrimaryLeave As Double, ByVal vPrimaryKaryawan As Double) As Decimal
        Dim DTAITem As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim Dummy As New Nullable(Of Decimal)
        Try
            DTAITem.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTAITem.ScalarQuery_GetAddition(vPrimaryLeave, vPrimaryKaryawan)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAITem.Dispose()
        End Try
    End Function

    Public Function GetDeduction(ByVal vPrimaryLeave As Double, ByVal vPrimaryKaryawan As Double) As Decimal
        Dim DTAITem As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim Dummy As New Nullable(Of Decimal)
        Try
            DTAITem.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTAITem.ScalarQuery_GetDeduction(vPrimaryLeave, vPrimaryKaryawan)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAITem.Dispose()
        End Try
    End Function

    Public Function GetNilaiTerpakai(ByVal vPrimaryLeave As Double, ByVal vPrimaryKaryawan As Double) As Decimal
        Dim JmlLeave, JmlAddition, JmlDeduction, JmlBalance As Decimal
        Dim DTAH As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim DTAITem As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim Dummy As New Nullable(Of Decimal)
        Dim JmlItemCuti As Nullable(Of Long)

        Try
            DTAH.Connection.ConnectionString = Core_Module.GetConstring
            DTAITem.Connection.ConnectionString = DTAH.Connection.ConnectionString

            JmlItemCuti = DTAITem.ScalarQuery_GetJmlItem(vPrimaryLeave, vPrimaryKaryawan)
            If JmlItemCuti.HasValue = False Then
                'karyawan ybs tidak terdaftar pada cuti ybs
                'tidak ada nilai terpakai
                Return 0
            Else
                If JmlItemCuti.Value = 0 Then
                    'karyawan ybs tidak terdaftar pada cuti ybs
                    'tidak ada nilai terpakai
                    Return 0
                Else
                    Dummy = DTAH.ScalarQuery_GetLeaveQty(vPrimaryLeave)
                    If Dummy.HasValue Then
                        JmlLeave = Dummy.Value
                    Else
                        JmlLeave = 0
                    End If

                    Dummy = DTAITem.ScalarQuery_GetAddition(vPrimaryLeave, vPrimaryKaryawan)
                    If Dummy.HasValue Then
                        JmlAddition = Dummy.Value
                    Else
                        JmlAddition = 0
                    End If

                    Dummy = DTAITem.ScalarQuery_GetDeduction(vPrimaryLeave, vPrimaryKaryawan)
                    If Dummy.HasValue Then
                        JmlDeduction = Dummy.Value
                    Else
                        JmlDeduction = 0
                    End If

                    Dummy = DTAITem.ScalarQuery_GetNilaiBalance(vPrimaryLeave, vPrimaryKaryawan)
                    If Dummy.HasValue Then
                        JmlBalance = Dummy.Value
                    Else
                        JmlBalance = 0
                    End If

                    'jml balance = jml leave + jml addition - jml deduction - jml terpakai
                    'jml terpakai = jml balance - jml leave - jml addition + jml deductioin
                    Return JmlLeave + JmlAddition - JmlDeduction - JmlBalance
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAH.Dispose()
            DTAITem.Dispose()
        End Try
        'Return Me.GetAddition(vPrimaryLeave, vPrimaryKaryawan) - Me.GetDeduction(vPrimaryLeave, vPrimaryKaryawan)
    End Function

    Public Function GetNilaiBalance(ByVal vPrimaryLeave As Double, ByVal vPrimaryKaryawan As Double) As Decimal
        Dim DTAITem As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim Dummy As New Nullable(Of Decimal)
        Try
            DTAITem.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTAITem.ScalarQuery_GetNilaiBalance(vPrimaryLeave, vPrimaryKaryawan)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAITem.Dispose()
        End Try
    End Function
End Class
