''by    : rosadin
''date  : 2014-08-07

Option Explicit On
Option Strict Off

Imports System.Windows.Forms, Core_Module
Public Class ClsSettingLeave

    Dim objCd As New ClsCekData

    Enum EnumStat
        general
        special
    End Enum

    Public Sub isiDgv(ByVal vDgv As DataGridView, ByVal vFrm As FrmSettingLeave, _
    ByVal vDate As Date)
        Dim dtaSettingLeave As New DseLaporanAttendanceTableAdapters.leaveSettingTableAdapter
        Dim tblSettingLeave As New DseLaporanAttendance.leaveSettingDataTable
        Dim rowSettingLeave As DseLaporanAttendance.leaveSettingRow
        Dim rowDgv As DataGridViewRow
        Dim cboDgv As DataGridViewComboBoxCell
        Try
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDgv.AllowUserToAddRows = False
            vDgv.Rows.Clear()
            If dtaSettingLeave.FillBy_tglDari_tglSampai(tblSettingLeave, Me.objCd.GetFisrtDateOfTheYear(vDate), Me.objCd.GetLastDateOfTheYear(vDate)) Then
                For Each rowSettingLeave In tblSettingLeave
                    With rowSettingLeave
                        vDgv.Rows.Add("")
                        rowDgv = vDgv.Rows(vDgv.RowCount - 1)
                        rowDgv.Cells(vFrm.ColNo.Index).Value = ""
                        rowDgv.Cells(vFrm.ColLeaveCode.Index).Value = .code
                        cboDgv = rowDgv.Cells(vFrm.ColStatus.Index)
                        If .Status = 0 Then
                            rowDgv.Cells(vFrm.ColStatus.Index).Value = cboDgv.Items.Item(EnumStat.general)
                        ElseIf .Status = 1 Then
                            rowDgv.Cells(vFrm.ColStatus.Index).Value = cboDgv.Items.Item(EnumStat.special)
                        End If
                        rowDgv.Cells(vFrm.ColPrimaryLeave.Index).Value = .primary
                    End With
                Next
            End If
            Me.objCd.SetNumber(vDgv)
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Information")
        Finally
            dtaSettingLeave.Dispose()
            tblSettingLeave.Dispose()
        End Try
    End Sub

    Public Sub hapusSettingLeave(ByVal vDgv As DataGridView, ByVal vFrm As FrmSettingLeave)
        Dim dtaSettingLeave As New DseLaporanAttendanceTableAdapters.leaveSettingTableAdapter
        Dim rowDgv As DataGridViewRow
        Try
            If vDgv.Rows.Count > 0 Then
                For Each rowDgv In vDgv.Rows
                    dtaSettingLeave.DeleteQuery_primaryLeave(CDec(rowDgv.Cells(vFrm.ColPrimaryLeave.Index).Value.ToString))
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Information")
        Finally
            dtaSettingLeave.Dispose()
        End Try
    End Sub

    Public Sub simpanSettingLeave(ByVal vDgv As DataGridView, ByVal vFrm As FrmSettingLeave)
        Dim dtaSettingLeave As New DseLaporanAttendanceTableAdapters.leaveSettingTableAdapter
        Dim rowDgv As DataGridViewRow
        Dim status As Integer
        Dim cboDgv As DataGridViewComboBoxCell
        Try
            If vDgv.Rows.Count Then
                For Each rowDgv In vDgv.Rows
                    If rowDgv.Cells(vFrm.ColPrimaryLeave.Index).Value.ToString <> "" Then
                        cboDgv = rowDgv.Cells(vFrm.ColStatus.Index)
                        If rowDgv.Cells(vFrm.ColStatus.Index).Value = cboDgv.Items.Item(EnumStat.general) Then
                            status = EnumStat.general
                        ElseIf rowDgv.Cells(vFrm.ColStatus.Index).Value = cboDgv.Items.Item(EnumStat.special) Then
                            status = EnumStat.special
                        End If
                        dtaSettingLeave.InsertQuery_primaryLeave_status(rowDgv.Cells(vFrm.ColPrimaryLeave.Index).Value, status)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Information")
        Finally
            dtaSettingLeave.Dispose()
        End Try
    End Sub
End Class
