Option Explicit On
Option Strict Off

Public Class ClsTrans_Schedule_Combi
    Public Sub GetListKombinasi(ByVal vTBLWU_Available As DseTrans_Schedule.tblmaster_product_item_workunitDataTable, _
    ByVal vJmlWorkUnitUsed As Integer, ByVal vListKombinasi As List(Of String), ByVal vPrimaryPreDefinedWU As Double, _
    ByVal vTBLProses As DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemprocessDataTable, _
    ByVal vTBLWUConfig As DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemworkunitDataTable, ByVal vPrimaryProduct As Double, ByVal vPrimaryProses As Double)
        Dim TBLIdx As New DseTrans_Schedule.TblIdxDataTable
        Dim Dummy As String = ""
        Dim rowWUConfig As DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemworkunitRow
        Dim rowProses As DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemprocessRow

        Try
            vListKombinasi.Clear()

            If vPrimaryPreDefinedWU <> -1 Then
                For Each rowProses In vTBLProses.Rows
                    If rowProses.primary_process = vPrimaryProses And rowProses.primary_product = vPrimaryProduct Then
                        For Each rowWUConfig In vTBLWUConfig.Rows
                            If rowWUConfig.identifier_process = rowProses.identifier Then
                                If Dummy = "" Then
                                    Dummy = rowWUConfig.primary_work_unit.ToString
                                Else
                                    Dummy = Dummy + "-" + rowWUConfig.primary_work_unit.ToString
                                End If
                            End If
                        Next
                        vListKombinasi.Add(Dummy)
                        Exit Sub
                    End If
                Next
            Else
                If vJmlWorkUnitUsed = 1 Then
                    Me.GetSingleCombination(vTBLWU_Available, vListKombinasi)
                    Exit Sub
                Else
                    TBLIdx.Rows.Clear()
                    vListKombinasi.Clear()

                    Me.GetFirstCombination(TBLIdx, vJmlWorkUnitUsed)
                    Me.AddCombination(TBLIdx, vListKombinasi, vTBLWU_Available)

                    While Me.GetNextCombination(TBLIdx, vTBLWU_Available.Rows.Count) = True
                        Me.AddCombination(TBLIdx, vListKombinasi, vTBLWU_Available)
                    End While
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddCombination(ByVal vTBLIdx As DseTrans_Schedule.TblIdxDataTable, ByVal vListKombinasi As List(Of String), _
    ByVal vTBLWU_Available As DseTrans_Schedule.tblmaster_product_item_workunitDataTable)
        Dim Combi As String = ""
        Dim rowWU As DseTrans_Schedule.tblmaster_product_item_workunitRow

        Try
            For Each row As DseTrans_Schedule.TblIdxRow In vTBLIdx.Rows
                rowWU = vTBLWU_Available.Rows(row.Idx - 1)
                If Combi = "" Then
                    Combi = rowWU.primary_work_unit.ToString
                Else
                    Combi = Combi + "-" + rowWU.primary_work_unit.ToString
                End If
            Next
            vListKombinasi.Add(Combi)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetSingleCombination(ByVal vTBLWU_Available As DseTrans_Schedule.tblmaster_product_item_workunitDataTable, _
    ByVal vListKombinasi As List(Of String))
        Try
            vListKombinasi.Clear()
            For Each rowwu As DseTrans_Schedule.tblmaster_product_item_workunitRow In vTBLWU_Available.Rows
                vListKombinasi.Add(rowwu.primary_work_unit)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetFirstCombination(ByVal vTBLIndex As DseTrans_Schedule.TblIdxDataTable, ByVal JmlWU As Double)
        Dim rowIdx As DseTrans_Schedule.TblIdxRow

        Try
            vTBLIndex.Rows.Clear()
            For i As Integer = 1 To JmlWU
                rowIdx = vTBLIndex.NewRow
                rowIdx.Idx = i
                vTBLIndex.Rows.Add(rowIdx)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetNextCombination(ByVal vTBLIndex As DseTrans_Schedule.TblIdxDataTable, ByVal JmlWUAvailable As Double) As Boolean
        Dim rowIdx As DseTrans_Schedule.TblIdxRow

        Try
            'tambahkan idx baris terakhir
            rowIdx = vTBLIndex.Rows(vTBLIndex.Rows.Count - 1)
            rowIdx.Idx += 1

            While Me.ScanUp(vTBLIndex, JmlWUAvailable) = False Or Me.ScanDown(vTBLIndex) = False
                rowIdx = vTBLIndex.Rows(0)
                If rowIdx.Idx >= JmlWUAvailable Then
                    Return False
                End If
            End While

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ScanUp(ByVal vTBLIndex As DseTrans_Schedule.TblIdxDataTable, ByVal JmlWUAvailable As Double) As Boolean
        Dim rowIdx As DseTrans_Schedule.TblIdxRow
        Dim rowIdxB4 As DseTrans_Schedule.TblIdxRow
        Dim AdaPerubahan As Boolean = False

        Try
            For i As Double = vTBLIndex.Rows.Count - 1 To 1 Step -1
                rowIdx = vTBLIndex.Rows(i)

                If rowIdx.Idx > JmlWUAvailable Then
                    rowIdxB4 = vTBLIndex.Rows(i - 1)
                    rowIdxB4.Idx += 1
                    rowIdx.Idx = 0
                    AdaPerubahan = True
                End If
            Next

            If AdaPerubahan = True Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ScanDown(ByVal vTBLIndex As DseTrans_Schedule.TblIdxDataTable) As Boolean
        Dim rowIdx, rowIdxAft As DseTrans_Schedule.TblIdxRow
        Dim AdaPerubahan As Boolean = False

        Try
            For i As Double = 0 To vTBLIndex.Rows.Count - 2
                rowIdx = vTBLIndex.Rows(i)
                rowIdxAft = vTBLIndex.Rows(i + 1)

                If rowIdx.Idx >= rowIdxAft.Idx Then
                    rowIdxAft.Idx = rowIdx.Idx + 1
                    AdaPerubahan = True
                End If
            Next

            If AdaPerubahan = True Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
