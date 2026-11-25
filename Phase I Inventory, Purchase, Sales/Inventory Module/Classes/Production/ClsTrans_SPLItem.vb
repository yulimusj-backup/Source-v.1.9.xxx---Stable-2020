Option Explicit On
Option Strict Off
Public Class ClsTrans_SPLItem
    Public Function SimpanItemWS(ByVal vDGVItemWS As DataGridView, ByVal vPrimaryDokumen As Double, _
           ByVal vFrm As FrmTrans_SPL) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_item_overtimeTableAdapter

        Dim DefaultPrecision As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each row In vDGVItemWS.Rows

                If String.IsNullOrEmpty(row.Cells(vFrm.colRemarks.Index).Value) = True Then
                    row.Cells(vFrm.colRemarks.Index).Value = ""
                End If

                Dim EOT As Integer = 0
                If row.Cells(vFrm.colEarlyOT.Index).Value = False Then
                    EOT = 0
                Else
                    EOT = 1
                End If
                Dim OT As Integer = 0
                If row.Cells(vFrm.colOT.Index).Value = False Then
                    OT = 0
                Else
                    OT = 1
                End If

                If row.Cells(vFrm.colPrimaryWorkshift.Index).Value <> -1 Then
                    DTA.Insert(vPrimaryDokumen, row.Cells(vFrm.colPrimaryWorkshift.Index).Value, _
                     EOT, ObjCD.Convert_ToMinutes(row.Cells(vFrm.colEarlyOTDuration.Index).Value), _
                    ObjCD.Convert_ToMinutes(row.Cells(vFrm.colEarlyDelayDuration.Index).Value), _
                    ObjCD.Convert_ToMinutes(row.Cells(vFrm.ColEarlyBreak.Index).Value), _
                    ObjCD.Convert_ToMinutes(row.Cells(vFrm.ColDelayBeforeEarlyBreak.Index).Value), OT, _
                    ObjCD.Convert_ToMinutes(row.Cells(vFrm.colOTDuration.Index).Value), _
                    ObjCD.Convert_ToMinutes(row.Cells(vFrm.colDelayDuration.Index).Value), _
                    ObjCD.Convert_ToMinutes(row.Cells(vFrm.ColBreak.Index).Value), _
                    ObjCD.Convert_ToMinutes(row.Cells(vFrm.ColDelayBeforeBreak.Index).Value), _
                    row.Cells(vFrm.colRemarks.Index).Value)
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function SimpanItemWU(ByVal vDGVItemWU As DataGridView, ByVal vPrimaryDokumen As Double, _
           ByVal vFrm As FrmTrans_SPL, ByVal vTblLog As DseTrans_SPL.tbltrans_spl_item_logDataTable) As Boolean
        Dim rowLog As DseTrans_SPL.tbltrans_spl_item_logRow
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim objdokumen As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter
        Dim objWU As New Core_Module.ClsMasterWorkUnit
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_item_workunitTableAdapter
        Dim DefaultPrecision As Integer = 0

        Dim dtaLog As New DseTrans_SPLTableAdapters.tbltrans_spl_item_logTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            dtaLog.Connection.ConnectionString = DTA.Connection.ConnectionString
            For Each row In vDGVItemWU.Rows

                If row.Cells(vFrm.colWU_Code.Index).Value <> "" Then
                    DTA.Insert(vPrimaryDokumen, vFrm.dgvWorkUnit.Item(vFrm.colWU_Primary.Index, row.Index).Value)
                End If
            Next

            For Each rowLog In vTblLog
                With rowLog
                    .primary_spl = vPrimaryDokumen
                    dtaLog.Insert(.primary_spl, ._date, .primary_wokunit, .primary_workshift, .is_early_ot, .early_ot_duration, _
                    .early_delay_duration, .early_break_duration, .delay_before_early_break, .is_ot, .ot_duration, .delay_duration, .break, .delay_before_break)
                End With
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function IsiDGVItemWS(ByVal vPrimaryDokumen As Double, ByVal vDGVItemWS As DataGridView, _
                              ByVal vFrm As FrmTrans_SPL) As Boolean

        Dim ObjCD As New Core_Module.ClsCekData
        'Dim ObjWS As New Core_Module.ClsMasterWorkShift
        Dim DTA As New DseTrans_SPLTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTrans_SPL.DataTable2DataTable
        Dim rowItem As DseTrans_SPL.DataTable2Row
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring


            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGVItemWS.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItemWS.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemWS.Rows.Clear()

            If DTA.Fill(TBL, vPrimaryDokumen) > 0 Then
                For Each rowItem In TBL
                    vDGVItemWS.Rows.Add()
                    rowDGV = vDGVItemWS.Rows(vDGVItemWS.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.colNo.Index).Value = ""
                        .Cells(vFrm.colShiftCode.Index).Value = rowItem.code
                        .Cells(vFrm.colDesc.Index).Value = rowItem.description
                        If rowItem.is_early_ot = 0 Then
                            .Cells(vFrm.colEarlyOT.Index).Value = False
                        ElseIf rowItem.is_early_ot = 1 Then
                            .Cells(vFrm.colEarlyOT.Index).Value = True
                        End If
                        .Cells(vFrm.colEarlyOTDuration.Index).Value = ObjCD.Convert_ToString(rowItem.early_ot_duration)
                        .Cells(vFrm.colEarlyDelayDuration.Index).Value = ObjCD.Convert_ToString(rowItem.early_delay_duration)
                        .Cells(vFrm.ColEarlyBreak.Index).Value = ObjCD.Convert_ToString(rowItem.early_break_duration)
                        .Cells(vFrm.ColDelayBeforeEarlyBreak.Index).Value = ObjCD.Convert_ToString(rowItem.delay_before_early_break)
                        If rowItem.is_ot = 0 Then
                            .Cells(vFrm.colOT.Index).Value = False
                        ElseIf rowItem.is_ot = 1 Then
                            .Cells(vFrm.colOT.Index).Value = True
                        End If
                        .Cells(vFrm.colOTDuration.Index).Value = ObjCD.Convert_ToString(rowItem.ot_duration)
                        .Cells(vFrm.ColBreak.Index).Value = ObjCD.Convert_ToString(rowItem.break)
                        .Cells(vFrm.colDelayDuration.Index).Value = ObjCD.Convert_ToString(rowItem.delay_duration)
                        .Cells(vFrm.ColDelayBeforeBreak.Index).Value = ObjCD.Convert_ToString(rowItem.delay_before_break)
                        .Cells(vFrm.colRemarks.Index).Value = rowItem.remarks
                        .Cells(vFrm.colPrimaryWorkshift.Index).Value = rowItem.primary_workshift

                    End With
                Next
            End If
            ObjCD.SetNumber(vDGVItemWS)
            vDGVItemWS.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItemWS.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
        End Try
    End Function
    Public Function IsiDGVItemWU(ByVal vPrimaryDokumen As Double, ByVal vDGVItemWU As DataGridView, _
                              ByVal vFrm As FrmTrans_SPL) As Boolean

        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjWU As New Core_Module.ClsMasterWorkUnit
        Dim DTA As New DseTrans_SPLTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseTrans_SPL.DataTable1DataTable
        Dim rowItem As DseTrans_SPL.DataTable1Row
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring


            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGVItemWU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItemWU.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemWU.Rows.Clear()
            If DTA.Fill(TBL, vPrimaryDokumen) > 0 Then
                For Each rowItem In TBL
                    vDGVItemWU.Rows.Add()
                    rowDGV = vDGVItemWU.Rows(vDGVItemWU.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.colNo.Index).Value = ""
                        .Cells(vFrm.colWU_Code.Index).Value = rowItem.work_unit_code
                        .Cells(vFrm.colWU_Desc.Index).Value = rowItem.description
                        .Cells(vFrm.colWU_Primary.Index).Value = rowItem.primary_work_unit
                    End With
                Next
            End If
            ObjCD.SetNumber(vDGVItemWU)
            vDGVItemWU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItemWU.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Function

    Public Function HapusItemWS(ByVal vPrimarySPL As Double) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_item_overtimeTableAdapter
        Try
            DTA.Delete(vPrimarySPL)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function HapusItemWU(ByVal vPrimarySPL As Double) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_item_workunitTableAdapter
        Dim DTALOG As New DseTrans_SPLTableAdapters.tbltrans_spl_item_logTableAdapter
        Try

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTALOG.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(vPrimarySPL)
            DTALOG.DeleteQuery(vPrimarySPL)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
