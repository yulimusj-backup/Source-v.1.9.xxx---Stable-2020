Option Explicit On
Option Strict Off

Public Class FrmTrans_SalesDashboard_MA_DeliveryDetail
    Public Enum EnumTipeItem As Integer
        Stock
        POB
        SOB
        PlaceOrder
        SalesDelivery
        SalesReturn
        SalesCancelation
    End Enum

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub dgvData_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvData.DataError

    End Sub

    Public Sub SetVisible()
        Dim AdaBooked As Boolean = False
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvData.Rows
                If rowDGV.Cells(Me.ColBookingNo.Index).Value <> "" Then
                    AdaBooked = True
                    Exit For
                End If
            Next
            Me.dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If Me.CheckBox1.Checked = True Then
                Me.ColQtyPromiseToDeliver.Visible = True
                Me.ColPromiseToDeliver.Visible = True
                Me.ColDocNo.Visible = True
                If AdaBooked = True Then
                    Me.ColBookingNo.Visible = True
                    Me.ColBookQty.Visible = True
                Else
                    Me.ColBookingNo.Visible = False
                    Me.ColBookQty.Visible = False
                End If
            Else
                Me.ColQtyPromiseToDeliver.Visible = False
                Me.ColPromiseToDeliver.Visible = False
                Me.ColDocNo.Visible = False
                Me.ColBookingNo.Visible = False
                Me.ColBookQty.Visible = False
            End If
            Me.dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Me.SetVisible()
    End Sub

    Private Sub FrmTrans_SalesDashboard_MA_DeliveryDetail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.SetVisible()
        Me.PictureBox1.BackColor = Core_Module.GetSMESSBackColorAll
        Me.ShowIcon = False
    End Sub
End Class