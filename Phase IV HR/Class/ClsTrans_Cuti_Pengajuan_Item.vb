Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Public Class ClsTrans_Cuti_Pengajuan_Item
    Public Sub SimpanItem(ByVal vDgv As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal vFrmExplain As FrmTransaksiCuti_Pengajuan_ExplainDuration, ByVal vTblList_StatusExplain As DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable, ByVal vFrm As FrmTransaksiCuti_Pengajuan)
        Dim dta As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuan_itemTableAdapter
        Dim row As DataGridViewRow
        Try
            Windows.Forms.Cursor.Current = Cursors.Default
            dta.Connection.ConnectionString = Core_Module.GetConstring
            For Each row In vDgv.Rows
                With row
                    dta.Insert(vPrimaryTransaksi, .Cells(vFrmExplain.ColPrimaryWorkShift.Index).Value, _
                    (.Cells(vFrmExplain.ColDate.Index).Value), .Cells(vFrmExplain.ColDuration.Index).Value, Me.GetInteger_ByStringStatusExplain(.Cells(vFrmExplain.ColStatus.Index).Value))
                End With
            Next
            vTblList_StatusExplain.Rows.Clear()
            Me.Add_To_List(vTblList_StatusExplain, vFrmExplain, vFrm)
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Sub

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double, ByVal vFrmExplain As FrmTransaksiCuti_Pengajuan_ExplainDuration, _
    ByVal vTblList_StatusExplain As DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable, ByVal vFrm As FrmTransaksiCuti_Pengajuan) As Boolean
        Dim dta As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuan_itemTableAdapter
        Try
            'hapus data item
            dta.Connection.ConnectionString = Core_Module.GetConstring

            If dta.Delete(vPrimaryTransaksi) > 0 Then
                vTblList_StatusExplain.Rows.Clear()
                Me.Add_To_List(vTblList_StatusExplain, vFrmExplain, vFrm)
                'nilai balikan
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function IsiDGVItem(ByRef vDGV As DataGridView, ByVal vPrimaryTransaksi As Double, ByRef vFrmExplain As FrmTransaksiCuti_Pengajuan_ExplainDuration, _
                               ByRef vTblList_StatusExplain As DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable, ByVal vFrm As FrmTransaksiCuti_Pengajuan) As Boolean
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuan_itemTableAdapter
        Dim TBL As New DseTrans_Cuti_Pengajuan.tbltrans_cuti_pengajuan_itemDataTable
        Dim row As DseTrans_Cuti_Pengajuan.tbltrans_cuti_pengajuan_itemRow
        Dim rowDGV As DataGridViewRow
        Dim ObjWorkShift As New Core_Module.ClsMasterWorkShift
        Dim ObjCd As New Core_Module.ClsCekData
        Try
            ''isikan(tabel)
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.Rows.Clear()
            If DTA.FillByPrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGV.Rows.Add("")
                        rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                        rowDGV.Cells(vFrmExplain.ColPrimaryWorkShift.Index).Value = .primary_workshift
                        rowDGV.Cells(vFrmExplain.ColWorkShift.Index).Value = ObjWorkShift.GetCodeByPrimary(.primary_workshift)
                        rowDGV.Cells(vFrmExplain.ColDate.Index).Value = Format(._date, "dd MMM yyyy")
                        If .status <> ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Leave Then
                            rowDGV.Cells(vFrmExplain.ColDuration.Index).Value = ObjCd.DigitGrouping(0, 1)
                        Else
                            rowDGV.Cells(vFrmExplain.ColDuration.Index).Value = ObjCd.DigitGrouping(.duration, 1)
                        End If
                        rowDGV.Cells(vFrmExplain.ColStatus.Index).Value = Me.GetString_ByEnumStatusExplain(.status)
                    End With
                Next
            End If

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vTblList_StatusExplain.Rows.Clear()
            Me.Add_To_List(vTblList_StatusExplain, vFrmExplain, vFrm)
            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCd = Nothing
        End Try
    End Function

    Public Function GetInteger_ByStringStatusExplain(ByVal vStrStatus As String) As ClsTrans_Cuti_Pengajuan.EnumStatusExplain
        Try
            If vStrStatus = "Leave" Then
                Return ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Leave
            ElseIf vStrStatus = "Sick" Then
                Return ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Sick
            ElseIf vStrStatus = "Permit" Then
                Return ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Permit
            ElseIf vStrStatus = "Absent" Then
                Return ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Absent
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Leave
        End Try
    End Function

    Public Function GetString_ByEnumStatusExplain(ByVal vEnumStatusExplain As ClsTrans_Cuti_Pengajuan.EnumStatusExplain) As String
        Try
            If vEnumStatusExplain = ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Leave Then
                Return "Leave"
            ElseIf vEnumStatusExplain = ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Sick Then
                Return "Sick"
            ElseIf vEnumStatusExplain = ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Permit Then
                Return "Permit"
            ElseIf vEnumStatusExplain = ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Absent Then
                Return "Absent"
            End If
            Return "Leave"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Leave"
        End Try
    End Function

    Public Function Add_To_List(ByRef vTblList As DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable, ByVal vFrmExplain As FrmTransaksiCuti_Pengajuan_ExplainDuration, _
   ByVal vFrm As FrmTransaksiCuti_Pengajuan) As Boolean
        Dim Datagrid As DataGridView = vFrmExplain.dgvExplain
        Dim Row As DseTrans_Cuti_Pengajuan.List_StatusExplainRow
        Dim StartDate, EndDate As Date
        Try
            If vTblList.Rows.Count > 0 Then
                vTblList.DefaultView.RowFilter = Nothing
                vTblList.DefaultView.Sort = "[Date] ASC"
                Row = vTblList.DefaultView(0).Row
                StartDate = Row._Date
                vTblList.DefaultView.Sort = "[Date] DESC"
                Row = vTblList.DefaultView(0).Row
                EndDate = Row._Date
            End If

            If vFrm.dtpStart.Value.Date <> StartDate.Date Or vFrm.DtpEnd.Value.Date <> EndDate.Date Then
                vTblList.Rows.Clear()
            End If

            For i As Integer = 0 To Datagrid.Rows.Count - 1
                vTblList.DefaultView.RowFilter = "[Primary Employee]='" & vFrm.TxtPrimaryEmployee.Text & "'" & " AND [Date]='" & Datagrid.Item(vFrmExplain.ColDate.Index, i).Value & "'"
                If vTblList.DefaultView.Count > 0 Then
                    Row = vTblList.DefaultView(0).Row
                    Row.Status = Datagrid.Item(vFrmExplain.ColStatus.Index, i).Value
                Else
                    Row = vTblList.NewList_StatusExplainRow
                    Row.Primary_Employee = vFrm.TxtPrimaryEmployee.Text
                    Row.Status = Datagrid.Item(vFrmExplain.ColStatus.Index, i).Value
                    Row._Date = Datagrid.Item(vFrmExplain.ColDate.Index, i).Value
                    Row.Index = i
                    vTblList.AddList_StatusExplainRow(Row)
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetStatusExplain(ByRef vTblList As DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable, _
    ByVal vFrm As FrmTransaksiCuti_Pengajuan_ExplainDuration, ByVal vIndex As Integer, ByVal vPrimaryEmployee As Double, ByVal vDate As Date) As String
        Dim vStr As String = "Leave"
        Dim Row As DseTrans_Cuti_Pengajuan.List_StatusExplainRow
        Try
            vTblList.DefaultView.RowFilter = "[Primary Employee]='" & vPrimaryEmployee.ToString & "'" & " AND [Date]='" & vDate & "'"
            If vTblList.DefaultView.Count > 0 Then
                Row = vTblList.DefaultView(0).Row
                vStr = Row.Status
            End If
            Return vStr
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return vStr
        End Try
    End Function

    Public Function GetDuration_InList(ByRef vTblList As DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable, _
    ByVal vDate As Date, ByVal vIndex As Integer, ByVal vPrimaryEmployee As Double) As Decimal
        Dim Row As DseTrans_Cuti_Pengajuan.List_StatusExplainRow
        Try
            vTblList.DefaultView.RowFilter = "[Primary Employee]='" & vPrimaryEmployee.ToString & "'" & " AND [Date]='" & vDate & "'"
            If vTblList.DefaultView.Count > 0 Then
                Row = vTblList.DefaultView(0).Row
                Return Row.Duration
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
