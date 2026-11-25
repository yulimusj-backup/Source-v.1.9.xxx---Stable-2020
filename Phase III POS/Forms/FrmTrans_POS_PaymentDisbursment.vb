Public Class FrmTrans_POS_PaymentDisbursment
    'yang ditampilkan pada form ini hanya data payment dari session yang berada di server pusat, jadi kalo belum ditransfer, maka 
    'tidak bisa melakukan disbursment.
    'yang bisa dilakukan disbursment hanya tipe payment: cc, debet card, dan cash withdrawal, voucher dan cash payment tidak ada disbursment.

    Private Sub LoadCaraPembayaran()
        Try
            Me.cboPaymentType.Items.Clear()
            Me.cboPaymentType.Items.Add("Credit Card")
            Me.cboPaymentType.Items.Add("Debet Card")
            Me.cboPaymentType.Items.Add("Cash Withdrawal")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadData()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_POS_PaymentDisbursment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'load data cara pembayaran 
            Me.LoadCaraPembayaran()

            Me.cboPaymentType.SelectedIndex = 0
            Me.cboPaymentStatus.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Try
            If e.RowIndex <> Me.dgvItem.Rows.Count - 1 And e.ColumnIndex = Me.ColBtnModify.Index Then
                Dim ij As New FrmTrans_POS_PaymentDisbursment_Edit
                ij.ShowDialog()
                ij.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SplitContainer1_Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles SplitContainer1.Panel1.Paint

    End Sub
End Class