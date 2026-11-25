Imports Core_Module

Public Class FrmTrans_Session_View
    Public mPrimaryPOS As Double
    Public BackColorAll As Color = Color.LightGoldenrodYellow

    Dim ObjCD As New ClsCekData

    Private Sub FrmTrans_Session_View_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.IsiDGV(Me.mPrimaryPOS, Me.dgvItem, Me, ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    Public Function IsiDGV(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTrans_Session_View, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItemP As New DseTrans_POSTableAdapters.DataTable2TableAdapter
        Dim TBLItemP As New DseTrans_POS.DataTable2DataTable
        Dim rowItemP As DseTrans_POS.DataTable2Row
        Try
            'update connection string
            DTAItemP.Connection.ConnectionString = Me.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'isikan data item
            DTAItemP.Fill(TBLItemP, vPrimary)
            vDGVItem.Rows.Clear()
            For Each rowItemP In TBLItemP.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColCodeNo.Index).Value = rowItemP.Product_Code
                    .Cells(i.ColDesc.Index).Value = rowItemP.Description
                    .Cells(i.ColLocation.Index).Value = rowItemP.Storage_Area_Code
                    .Cells(i.ColBatchNumber.Index).Value = rowItemP.batch_number
                    .Cells(i.ColExpiry.Index).Value = Format(rowItemP.expiry_date, "dd MMM yyyy")
                    .Cells(i.ColQty.Index).Value = vObjCD.DigitGrouping(rowItemP.qty, rowItemP._Decimal)
                    .Cells(i.ColUnit.Index).Value = rowItemP.Unit_Code
                    .Cells(i.ColUnitPrice.Index).Value = vObjCD.DigitGrouping(rowItemP.unit_price, rowItemP.Default_Precision)
                    .Cells(i.ColDisc_Amount.Index).Value = vObjCD.DigitGrouping(rowItemP.discount, rowItemP.Default_Precision)
                    .Cells(i.ColNettPrice.Index).Value = vObjCD.DigitGrouping(rowItemP.nett_price, rowItemP.Default_Precision)
                    .Cells(i.ColSubTotal.Index).Value = vObjCD.DigitGrouping(rowItemP.sub_total, rowItemP.Default_Precision)
                    .Cells(i.ColPrimaryCodeNo.Index).Value = rowItemP.primary_product
                    .Cells(i.ColPrimaryUnit.Index).Value = rowItemP.primary_unit
                    .Cells(i.ColPrimaryDisc1.Index).Value = rowItemP.primary_disc_1
                    .Cells(i.ColPrimaryDisc2.Index).Value = rowItemP.primary_disc_2
                    .Cells(i.ColPrimaryDisc3.Index).Value = rowItemP.primary_disc_3
                    .Cells(i.ColDiscAmount1.Index).Value = rowItemP.disc_amount_1
                    .Cells(i.ColDiscAmount2.Index).Value = rowItemP.disc_amount_2
                    .Cells(i.ColDiscAmount3.Index).Value = rowItemP.disc_amount_3
                    .Cells(i.ColDiscString1.Index).Value = rowItemP.disc_string_1
                    .Cells(i.ColDiscString2.Index).Value = rowItemP.disc_string_2
                    .Cells(i.ColDiscString3.Index).Value = rowItemP.disc_string_3
                    .Cells(i.ColDiscString.Index).Value = rowItemP.disc_string
                    .Cells(i.ColDisc_Amount.Index).Value = vObjCD.DigitGrouping(rowItemP.discount, rowItemP.Default_Precision)
                    .Cells(i.ColPrimaryLocation.Index).Value = rowItemP.primary_location
                    .Cells(i.ColPrimaryAutoNumber.Index).Value = rowItemP.primary
                End With
            Next

            vObjCD.SetNumber(vDGVItem)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemP.Dispose()
            TBLItemP.Dispose()
        End Try
    End Function

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class