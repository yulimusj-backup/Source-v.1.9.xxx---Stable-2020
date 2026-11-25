Public Class FrmTest

    Private Sub FrmTest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub PopulateTabel(ByVal TBLWU As DseTrans_Schedule.tblmaster_product_item_workunitDataTable)
        Dim rowWU As DseTrans_Schedule.tblmaster_product_item_workunitRow

        Try
            TBLWU.Rows.Clear()
            For i As Integer = 1 To CInt(Me.TxtJmlWU.Text)
                rowWU = TBLWU.Newtblmaster_product_item_workunitRow
                rowWU.effective_rate = 100
                rowWU.efficiency = 100
                rowWU.primary_work_unit = i
                rowWU.rate_per_hour = 100
                rowWU.work_unit_code = ""
                rowWU._Decimal = 0

                TBLWU.Rows.Add(rowWU)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStart_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStart.Click
        'Dim TBLWU As New DseTrans_Schedule.tblmaster_product_item_workunitDataTable
        'Dim ObjCombi As New ClsTrans_Schedule_Combi
        'Dim ListPrimaryWU As New List(Of String)

        'Try
        '    Me.PopulateTabel(TBLWU)
        '    ObjCombi.GetListKombinasi(TBLWU, Me.TxtWUUsedQty.Text, ListPrimaryWU)
        '    Me.dgvDown.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        '    Me.dgvDown.Rows.Clear()
        '    For Each Combi As String In ListPrimaryWU
        '        Me.dgvDown.Rows.Add(Me.dgvDown.Rows.Count + 1, Combi)
        '    Next
        '    Me.dgvDown.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    ObjCombi = Nothing
        '    TBLWU.Dispose()
        '    ListPrimaryWU = Nothing
        'End Try
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub
End Class