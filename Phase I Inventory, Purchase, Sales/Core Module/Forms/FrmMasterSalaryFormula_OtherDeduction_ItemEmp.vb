Option Explicit On
Option Strict Off

Public Class FrmMasterSalaryFormula_OtherDeduction_ItemEmp
    Public mPrimaryEmp As Double = -1
    Public mUserCode As String = ""
    Public BackColorAll As Color
    Public PrimaryEmp As Double = -1
    Dim ObjCd As New ClsCekData


    Private Sub FrmMasterSalaryFormula_OtherDeduction_ItemEmp_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'TODO: This line of code loads data into the 'DseTrans_OtherDeduction.DataTable2' table. You can move, or remove it, as needed.
            Me.DataTable2TableAdapter.FillBy(Me.DseTrans_OtherDeduction.DataTable2, Me.PrimaryEmp)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemDetail_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DgvItemDetail.CellMouseClick
        Dim frm As New FrmMasterSalaryFormula_OtherDeduction
        Try
            If e.ColumnIndex = Me.CodeDataGridViewTextBoxColumn.Index Then
                frm.UserCode = Me.mUserCode
                frm.ObjDed.CariDetail(Me.DgvItemDetail.Item(Me.PrimaryDataGridViewTextBoxColumn.Index, e.RowIndex).Value, _
                            frm.TblVar, frm.TblCons, frm)
                frm.SalinKeField()
                frm.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemDetail_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItemDetail.RowsAdded
        Try
            Me.ObjCd.SetNumber(Me.DgvItemDetail)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemDetail_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItemDetail.Sorted
        Try
            Me.ObjCd.SetNumber(Me.DgvItemDetail)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class