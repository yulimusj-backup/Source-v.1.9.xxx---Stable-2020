Option Explicit On
Option Strict Off

Public Class ClsTrans_SchDownTime_Item

    Dim objCd As New Core_Module.ClsCekData

    Public Function SimpanItem(ByVal vPrimary As Double, ByVal vDgv As DataGridView, ByVal vFrm As FrmTrans_SchDownTime) As Boolean
        Dim d As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtime_itemTableAdapter
        Dim startTime, endTime As Date
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            For Each r As DataGridViewRow In vDgv.Rows
                With r
                    If .Cells(vFrm.colWU_Primary.Index).Value <> "-1" Then
                        startTime = CDate(.Cells(vFrm.colWU_StartDate.Index).Value & " " & .Cells(vFrm.colWU_StartTime.Index).Value.ToString.Replace(".", ":") & ":00")
                        endTime = CDate(.Cells(vFrm.colWU_EndDate.Index).Value & " " & .Cells(vFrm.colWU_EndTime.Index).Value.ToString.Replace(".", ":") & ":00")
                        If String.IsNullOrEmpty(.Cells(vFrm.colWU_Remarks.Index).Value) Then
                            .Cells(vFrm.colWU_Remarks.Index).Value = ""
                        End If
                        d.Insert(vPrimary, startTime, endTime, .Cells(vFrm.colWU_Remarks.Index).Value, .Cells(vFrm.colWU_Primary.Index).Value)
                    End If
                End With
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimary As Double) As Boolean
        Dim dta As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtime_itemTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Delete(vPrimary) <= 0 Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function IsiDgvItem(ByVal vPrimary As Double, ByVal vDgv As DataGridView, ByVal i As FrmTrans_SchDownTime) As Boolean
        Dim dta As New DseTrans_SchDownTimeTableAdapters.DataTable1TableAdapter
        Dim tbl As New DseTrans_SchDownTime.DataTable1DataTable
        Dim row As DseTrans_SchDownTime.DataTable1Row
        Dim rowDgv As DataGridViewRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            vDgv.Rows.Clear()
            vDgv.AllowUserToAddRows = False
            If dta.Fill(tbl, vPrimary) > 0 Then
                vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                vDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                For Each row In tbl
                    With row
                        vDgv.Rows.Add()
                        rowDgv = vDgv.Rows(vDgv.Rows.Count - 1)
                        rowDgv.Cells(i.colWU_No.Index).Value = ""
                        rowDgv.Cells(i.colWU_Code.Index).Value = .Work_Unit_Code
                        rowDgv.Cells(i.colWU_Desc.Index).Value = .Work_Unit_Desc
                        rowDgv.Cells(i.colWU_Primary.Index).Value = .Primary_Work_Unit
                        rowDgv.Cells(i.colWU_StartDate.Index).Value = Format(.Start_Time, "dd MMM yyyy")
                        rowDgv.Cells(i.colWU_StartTime.Index).Value = Format(.Start_Time, "HH.mm")
                        rowDgv.Cells(i.colWU_EndDate.Index).Value = Format(.End_Time, "dd MMM yyyy")
                        rowDgv.Cells(i.colWU_EndTime.Index).Value = Format(.End_Time, "HH.mm")
                        If String.IsNullOrEmpty(.Remarks_Item) Then
                            .Remarks_Item = ""
                        End If
                        rowDgv.Cells(i.colWU_Remarks.Index).Value = .Remarks_Item
                    End With
                Next
                vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                vDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                Me.objCd.SetNumber(vDgv)
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
End Class
