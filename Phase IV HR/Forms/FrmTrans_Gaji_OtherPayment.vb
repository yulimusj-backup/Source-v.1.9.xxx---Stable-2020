Public Class FrmTrans_Gaji_OtherPayment
    Public vTbl As DseTransaksiSalary.tbltrans_salary_item_other1DataTable
    Dim vRow As DseTransaksiSalary.tbltrans_salary_item_other1Row
    Dim ObjCD As New Core_Module.ClsCekData
    Dim DefPrec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjItem As New Core_Module.ClsMasterSalaryFormula_item

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_Gaji_OtherPayment_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As Windows.Forms.DataGridViewRow

        Try
            DefPrec = ObjCur.GetDecimal_DefaultCurrency
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll

            For Each vRow In vTbl.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColAmount.Index).Value = vRow.Amount
                rowDGV.Cells(Me.ColComponentName.Index).Value = vRow.Description
                rowDGV.Cells(Me.ColDate.Index).Value = vRow._Date
                rowDGV.Cells(Me.ColDocNo.Index).Value = vRow.Document_No
                rowDGV.Cells(Me.ColIdentifierKomponen.Index).Value = vRow.Identifier_Salary_Component
                rowDGV.Cells(Me.ColIsRutin.Index).Value = vRow.Is_Rutin
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                rowDGV.Cells(Me.ColPrimaryRefDoc.Index).Value = vRow.Primary_Ref_Doc
                rowDGV.Cells(Me.ColTaxElement.Index).Value = Me.ObjItem.GetTipeTaxComponent_Str(vRow.Tax_Component)
                rowDGV.Cells(Me.ColTipeKomponen.Index).Value = Me.ObjItem.GetTipeVariable_Desc(vRow.Type)
            Next

            Me.ColAmount.DefaultCellStyle.Format = "N" + Me.DefPrec.ToString
            Me.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick

    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError

    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class