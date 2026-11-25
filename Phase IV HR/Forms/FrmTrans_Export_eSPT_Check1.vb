Public Class FrmTrans_Export_eSPT_Check1
    Public TblRef As DseTrans_Export_eSPT.tbltrans_salary_monthly2DataTable

    Dim BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCd As New Core_Module.ClsCekData

    Private Sub FrmTrans_Export_eSPT_Check1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As Windows.Forms.DataGridViewRow
        Dim rowRef As DseTrans_Export_eSPT.tbltrans_salary_monthly2Row

        Try
            Me.dgvDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.dgvDetail.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            Me.dgvDetail.Rows.Clear()

            For Each rowRef In Me.TblRef.Rows
                Me.dgvDetail.Rows.Add()
                rowDGV = Me.dgvDetail.Rows(Me.dgvDetail.Rows.Count - 1)

                rowDGV.Cells(Me.ColGrossIncome.Index).Value = Me.ObjCd.DigitGrouping(rowRef.Penghasilan_Bruto, 0)
                rowDGV.Cells(Me.ColIncomeTax.Index).Value = Me.ObjCd.DigitGrouping(rowRef.Income_Tax - rowRef._21_PPH_Final_Pesangon, 0)
                rowDGV.Cells(Me.ColPeriod.Index).Value = Format(rowRef.Tanggal_Transaksi, "MMM yyyy")
                rowDGV.Cells(Me.ColPrevPeriodIncomeTax.Index).Value = Me.ObjCd.DigitGrouping(rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya1, 0)
            Next
            Me.dgvDetail.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class