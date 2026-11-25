Option Explicit On
Option Strict Off

Public Class FrmTransaksiDO3_Option
    Public PrimarySO As Double = -1
    Public PrimaryProd As Double = -1
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim objCd As New Core_Module.ClsCekData
    Public DeffPrec As Integer = 0
    Private Sub btn_simpan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_simpan.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.OK
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub FrmOptionAOP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.MinimizeBox = False
            Me.MaximizeBox = False
            Me.Location = New Point((Me.Width * -1) - 5, (Me.Height * -1) - 5)
            LoadOption()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadOption()
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim Tbl As New DseTransaksiDO.tbltrans_so_itemdeliveryDataTable
        Dim row As DseTransaksiDO.tbltrans_so_itemdeliveryRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_Primary(Tbl, Me.PrimarySO, Me.ObjSO.GetIndetifier(Me.PrimaryProd, Me.PrimarySO)) > 0 Then
                If Tbl.Rows.Count > 1 Then
                    For i As Integer = 0 To Tbl.Rows.Count - 1
                        row = Tbl.Rows(i)
                        With row
                            Me.cbo_DeliveryDate.Items.Add(Format(row.delivery_date, "dd MMM yyyy"))
                            Me.cbo_DeliveryDate.SelectedIndex = 0
                        End With
                    Next
                    Me.CenterToScreen()
                Else
                    row = Tbl.Rows(0)
                    Me.cbo_DeliveryDate.Items.Add(Format(row.delivery_date, "dd MMM yyyy"))
                    Me.cbo_DeliveryDate.SelectedIndex = 0
                    Me.DialogResult = Windows.Forms.DialogResult.No
                    Me.Close()
                End If
            Else
                Me.DialogResult = Windows.Forms.DialogResult.No
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn_cancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            'Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class