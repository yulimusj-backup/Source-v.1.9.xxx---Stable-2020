Option Explicit On
Option Strict Off

Public Class FrmTransaksiPO_HistoricalPrice

    Dim objCD As New Core_Module.ClsCekData

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiPO_HistoricalPrice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow
        Dim AdaUnitMulti As Boolean = False

        Try
            Me.dgvPrice.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.ShowIcon = False

            For Each rowDGV In Me.dgvPrice.Rows
                If rowDGV.Cells(Me.ColUnitCode.Index).Value <> rowDGV.Cells(Me.ColUnitCodeMulti.Index).Value Then
                    AdaUnitMulti = True
                    Exit For
                End If
            Next

            If AdaUnitMulti = False Then
                Me.ColUnitCode.Visible = False
                Me.ColRatio.Visible = False
                Me.ColQty.Visible = False
                Me.ColPriceByMultiUnit_str.Visible = False
            Else
                Me.ColUnitCode.Visible = True
                Me.ColRatio.Visible = True
                Me.ColQty.Visible = True
                Me.ColPriceByMultiUnit_str.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPrice_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvPrice.Sorted
        Try
            objCD.SetNumber(Me.dgvPrice)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class