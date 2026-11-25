Public Class FrmTrans_MaterialPlanningDashboard_HistoricalPurchasePrice
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vPrimaryProduct As Double = -1
    Public vRefDate As Date = Today
    Public Obj As New ClsTransaksiMaterialPlanning
    Public ObjCD As New Core_Module.ClsCekData

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_HistoricalPurchasePrice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.chkRefDate.Checked = False
            Me.chkRefDate.Checked = True
            Me.dtpRefDate.Value = Me.vRefDate
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadData()
        Dim RefDate As Date

        Try
            If Me.chkRefDate.Checked = True Then
                RefDate = Me.dtpRefDate.Value
            Else
                RefDate = #1/1/1950#
            End If

            Me.DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Obj.ExplainHistoricalPrice(Me.vPrimaryProduct, Me.DataGridView1, Me.ObjCD, RefDate, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkRefDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRefDate.CheckedChanged
        Try
            If Me.chkRefDate.Checked = True Then
                Me.dtpRefDate.Enabled = True
            Else
                Me.dtpRefDate.Enabled = False
            End If

            Me.LoadData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class