Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_NewSch3_Exception
    Public TBLRef As DseTrans_Schedule2.tbltrans_mo_master_predefine_workunit_itemworkunitDataTable
    Public FrmParent As New FrmTrans_ProductionSchedule_NewSch3

    Dim RowRef As DseTrans_Schedule2.tbltrans_mo_master_predefine_workunit_itemworkunitRow
    Dim BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub FrmTrans_ProductionSchedule_NewSch3_Exception_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim idx As Integer = 0

        Try
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            For Each rowDGV In Me.FrmParent.DgvWUAllocate_Process.Rows
                idx = 0
                While idx <= Me.TBLRef.Rows.Count - 1
                    RowRef = Me.TBLRef.Rows(idx)
                    If rowDGV.Cells(Me.FrmParent.ColAll_Process_PrimaryProcess.Index).Value = RowRef.primary_process And _
                    rowDGV.Cells(Me.FrmParent.ColAll_Process_PrimaryCodeNo.Index).Value = RowRef.primary_product Then
                        'tambahkan ke dgv
                        Me.DgvItem.Rows.Add()
                        rowDGV1 = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)
                        rowDGV1.Cells(Me.ColAll_IdentifierProcess.Index).Value = rowDGV.Cells(Me.FrmParent.ColAll_IdentifierProcess.Index).Value
                        rowDGV1.Cells(Me.ColAll_Process_Code.Index).Value = rowDGV.Cells(Me.FrmParent.ColAll_Process_Code.Index).Value
                        rowDGV1.Cells(Me.ColAll_Process_CodeNo.Index).Value = rowDGV.Cells(Me.FrmParent.ColAll_Process_CodeNo.Index).Value
                        rowDGV1.Cells(Me.ColAll_Process_Desc.Index).Value = rowDGV.Cells(Me.FrmParent.ColAll_Process_Desc.Index).Value
                        rowDGV1.Cells(Me.ColAll_Process_No.Index).Value = rowDGV1.Index + 1
                        rowDGV1.Cells(Me.ColAll_Process_PrimaryCodeNo.Index).Value = rowDGV.Cells(Me.FrmParent.ColAll_Process_PrimaryCodeNo.Index).Value
                        rowDGV1.Cells(Me.ColAll_Process_PrimaryProcess.Index).Value = rowDGV.Cells(Me.FrmParent.ColAll_Process_PrimaryProcess.Index).Value
                        rowDGV1.Cells(Me.ColAll_Process_PrimaryWU.Index).Value = RowRef.primary_work_unit
                        rowDGV1.Cells(Me.ColAll_Process_WU.Index).Value = RowRef.work_unit_code

                        'buang dari tbl ref
                        Me.TBLRef.Rows.Remove(Me.RowRef)
                    Else
                        idx += 1
                    End If
                End While
            Next

            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class