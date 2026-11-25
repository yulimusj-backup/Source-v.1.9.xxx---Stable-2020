Option Explicit On
Option Strict Off

Public Class FrmTransaksiDO_Quick_Result
    '     ORIGINAL CODING
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public UserCode As String = ""
    Public vRboSO As Boolean = True
    Dim tblResult As New DseTransaksiQuick_DO2.tblResultDataTable
    Dim rowResult As DseTransaksiQuick_DO2.tblResultRow

    Dim objCD As New Core_Module.ClsCekData

    Private Sub FrmTransaksiDO_Quick_Result_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.loadData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.objCD.SetNumber(Me.DgvItem)
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        End Try
    End Sub

    Private Sub loadData()
        Dim rowdgv As DataGridViewRow
        Try
            Me.DgvItem.Rows.Clear()
            For Each rowResult In Me.tblResult.Rows
                Me.DgvItem.Rows.Add()
                rowdgv = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)
                rowdgv.Cells(Me.ColDONo.Index).Value = rowResult.No_DO
                rowdgv.Cells(Me.colSONo.Index).Value = rowResult.No_Transaksi
                rowdgv.Cells(Me.ColPrimaryDO.Index).Value = rowResult.Primary_DO
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit

        Try
            If e.ColumnIndex <> Me.ColPrint.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub
End Class