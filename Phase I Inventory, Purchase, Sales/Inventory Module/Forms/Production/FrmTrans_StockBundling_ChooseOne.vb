Option Explicit On
Option Strict Off
Public Class FrmTrans_StockBundling_ChooseOne
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public t As New DseTrans_Schedule.DataTable5DataTable
    Public PrimaryBundling As Double = -1

    Dim objCd As New Core_Module.ClsCekData

    Private Sub FrmTrans_StockBundling_ChooseOne_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim r As DseTrans_Schedule.DataTable5Row
        Dim rd As DataGridViewRow

        Try
            Me.dgvProductionLog.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvProductionLog.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each r In Me.t
                rd = Me.dgvProductionLog.Rows(Me.dgvProductionLog.Rows.Add)
                rd.Cells(Me.ColDate.Index).Value = Format(r.Tanggal_Transaksi, "dd MMM yyyy")
                rd.Cells(Me.ColDocumentNo.Index).Value = r.No_Transaksi
                rd.Cells(Me.ColNo.Index).Value = Me.dgvProductionLog.Rows.Count.ToString
                rd.Cells(Me.ColPrimary.Index).Value = r.Primary
                rd.Cells(Me.ColQty.Index).Value = Me.objCd.DigitGrouping(r.total_produce, r._Decimal)
            Next
            Me.dgvProductionLog.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub New(ByVal vt As DseTrans_Schedule.DataTable5DataTable)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.t = vt
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.PrimaryBundling = -1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.PrimaryBundling = Me.dgvProductionLog.Item(Me.ColPrimary.Index, Me.dgvProductionLog.SelectedRows(0).Index).Value
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProductionLog_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvProductionLog.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvProductionLog)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class