Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_Exp_QtyRMAdd
    Public CodeNoRM, DescRM, WONumber, ProcessCode, vIdentifierJO, vIdentifireWO, UnitCode As String
    Public TblItemRM As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
    Public TblItemRMAdd As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Primary_Product_Code_RM As Double = -1
    Public RefDate As Date = Today
    Public PrimaryMO As Double = -1
    Public ReadOnly_ As Boolean = False

    Dim LangsungClose As Boolean = False
    Dim DeleteIndex As Integer = -1
    Public DefPrec As Integer = 0
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Buffer As String = ""

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        'sudah di handle oleh form closing
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_Exp_QtyRMAdd_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.ReadOnly_ = True Then
                Exit Sub
            End If

            If Me.LangsungClose = False Then
                If MsgBox("Close and discard data changes?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                Else
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim QtyBalanceLama, QtyAdditionalLama, QtyBalanceBaru, QtyAdditionalBaru As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim RowItemRMAdd As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addRow

        Try
            'pastikan dahulu data qty tidak ada yang nol
            For Each rowDGV In Me.DgvItemRMAdd.Rows
                If rowDGV.Index <> Me.DgvItemRMAdd.Rows.Count - 1 Then
                    If rowDGV.Cells(Me.ColQty.Index).Value <= 0 Then
                        MsgBox("Please specify additional qty on row " + (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                        Me.DgvItemRMAdd.ClearSelection()
                        rowDGV.Selected = True

                        Exit Sub
                    End If
                End If
            Next

            'pastikan balance qty tidak minus
            For Each rowDGV In Me.DgvItemRMAdd.Rows
                QtyAdditionalBaru += CDec(rowDGV.Cells(Me.ColQty.Index).Value)
            Next
            For Each RowItemRM In Me.TblItemRM.Rows
                If RowItemRM.Identifier_JO = Me.vIdentifierJO And RowItemRM.Identifier_WO = Me.vIdentifireWO And _
                RowItemRM.Primary_Product_Code_RM = Me.Primary_Product_Code_RM Then
                    QtyBalanceLama = RowItemRM.Balance_Qty
                    QtyAdditionalLama = RowItemRM.Additional_Qty

                    Exit For
                End If
            Next
            QtyBalanceBaru = QtyBalanceLama - QtyAdditionalLama + QtyAdditionalBaru
            If QtyBalanceBaru < 0 Then
                MsgBox("The balance qty is negative, please qty data check first.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If MsgBox("Accept data changes?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                'update tabel rm add
                Me.TblItemRMAdd.DefaultView.RowFilter = "[Identifier JO] = '" + Me.vIdentifierJO + _
                "' AND [Identifier WO] = '" + Me.vIdentifireWO + "' AND [Primary Product Code RM] = " + Me.Primary_Product_Code_RM.ToString
                While Me.TblItemRMAdd.DefaultView.Count > 0
                    Me.TblItemRMAdd.Rows.Remove(Me.TblItemRMAdd.DefaultView(0).Row)
                End While
                For Each rowDGV In Me.DgvItemRMAdd.Rows
                    If rowDGV.Index <> Me.DgvItemRMAdd.Rows.Count - 1 Then
                        RowItemRMAdd = Me.TblItemRMAdd.Newtbltrans_mo_item_rm_addRow
                        With RowItemRMAdd
                            .Add_Date = CDate(rowDGV.Cells(Me.ColDate.Index).Value)
                            .Identifier_JO = Me.vIdentifierJO
                            .Identifier_WO = Me.vIdentifireWO
                            .Primary = -1
                            .Primary_MO = Me.PrimaryMO
                            .Primary_Product_Code_RM = Me.Primary_Product_Code_RM
                            .Qty = CDec(rowDGV.Cells(Me.ColQty.Index).Value)
                            .Remarks = rowDGV.Cells(Me.ColRemarks.Index).Value
                        End With
                        Me.TblItemRMAdd.Rows.Add(RowItemRMAdd)
                    End If
                Next

                'update tabel rm
                For Each RowItemRM In Me.TblItemRM.Rows
                    If RowItemRM.Identifier_JO = Me.vIdentifierJO And RowItemRM.Identifier_WO = Me.vIdentifireWO And _
                    RowItemRM.Primary_Product_Code_RM = Me.Primary_Product_Code_RM Then
                        RowItemRM.Balance_Qty = QtyBalanceBaru
                        RowItemRM.Additional_Qty = QtyAdditionalBaru

                        Exit For
                    End If
                Next

                'tutup
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.LangsungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemBundling_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItemRMAdd.CellBeginEdit
        Try
            If e.RowIndex = Me.DgvItemRMAdd.Rows.Count - 1 Then
                'untuk baris baru, harus mulai dengan date
                If e.ColumnIndex = Me.ColDate.Index Then
                    If IsNothing(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.Buffer = Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        Me.Buffer = ""
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColDate.Index Or e.ColumnIndex = Me.ColQty.Index Or e.ColumnIndex = Me.ColRemarks.Index Then
                    If IsNothing(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.Buffer = Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        Me.Buffer = ""
                    End If
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemBundling_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItemRMAdd.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColDate.Index Then
                Me.EndEditDate(sender, e)
            ElseIf e.ColumnIndex = Me.ColQty.Index Then
                Me.EndEditQty(sender, e)
            ElseIf e.ColumnIndex = Me.ColRemarks.Index Then
                Me.EndEditRemarks(sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditDate(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim rowDGV As DataGridViewRow

        Try
            If IsNothing(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.RowIndex = Me.DgvItemRMAdd.Rows.Count - 2 Then
                    Me.DeleteIndex = e.RowIndex
                    Me.bgw1.RunWorkerAsync()
                    Exit Sub
                Else
                    Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            Else
                'cek apakah tanggal
                If IsDate(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.RowIndex = Me.DgvItemRMAdd.Rows.Count - 2 Then
                        Me.DeleteIndex = e.RowIndex
                        Me.bgw1.RunWorkerAsync()
                        Exit Sub
                    Else
                        Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If

                'format ulang tanggal
                Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")

                If IsDate(Me.Buffer) = True Then
                    If CDate(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value) = CDate(Me.Buffer) Then
                        Exit Sub
                    End If
                End If

                'minimal sama dengan tanggal transaksi
                If Me.RefDate.Date > CDate(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value).Date Then
                    MsgBox("Minimum addition date is: " + Format(Me.RefDate, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                    If e.RowIndex = Me.DgvItemRMAdd.Rows.Count - 2 Then
                        Me.DeleteIndex = e.RowIndex
                        Me.bgw1.RunWorkerAsync()
                        Exit Sub
                    Else
                        Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If

                'pastikan tanggal minimal sama dengan tanggal di atasnya
                If e.RowIndex > 0 Then
                    rowDGV = Me.DgvItemRMAdd.Rows(e.RowIndex - 1)
                    If CDate(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value).Date <= CDate(rowDGV.Cells(Me.ColDate.Index).Value) Then
                        MsgBox("Minimum date you can enter is: " + rowDGV.Cells(Me.ColDate.Index).Value.ToString, MsgBoxStyle.Information, "Information")
                        If e.RowIndex = Me.DgvItemRMAdd.Rows.Count - 2 Then
                            Me.DeleteIndex = e.RowIndex
                            Me.bgw1.RunWorkerAsync()
                            Exit Sub
                        Else
                            Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If

                    'pastikan qty baris sebelumnya > 0
                    If CDec(rowDGV.Cells(Me.ColQty.Index).Value) <= 0 Then
                        MsgBox("Please specify additional qty on previous row.", MsgBoxStyle.Information, "Information")
                        If e.RowIndex = Me.DgvItemRMAdd.Rows.Count - 2 Then
                            Me.DeleteIndex = e.RowIndex
                            Me.bgw1.RunWorkerAsync()
                            Exit Sub
                        Else
                            Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                End If

                'masukkan data remarks untuk baris baru
                If e.RowIndex = Me.DgvItemRMAdd.Rows.Count - 2 Then
                    Me.DgvItemRMAdd.Item(Me.ColRemarks.Index, e.RowIndex).Value = "Specify addition reason here"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditRemarks(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If IsNothing(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            Else
                If Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                    Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim QtyBalanceLama, QtyAdditionalLama, QtyBalanceBaru, QtyAdditionalBaru As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

        Try
            If IsNothing(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            Else
                'cek non negative
                If Me.ObjCD.CekNonNegatif(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please specify non negative number.", MsgBoxStyle.Information, "Information")
                    Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'digit grouping
                Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)

                'pastikan balance tidak jadi negative
                For Each rowDGV In Me.DgvItemRMAdd.Rows
                    QtyAdditionalBaru += CDec(rowDGV.Cells(Me.ColQty.Index).Value)
                Next
                For Each RowItemRM In Me.TblItemRM.Rows
                    If RowItemRM.Identifier_JO = Me.vIdentifierJO And RowItemRM.Identifier_WO = Me.vIdentifireWO And _
                    RowItemRM.Primary_Product_Code_RM = Me.Primary_Product_Code_RM Then
                        QtyBalanceLama = RowItemRM.Balance_Qty
                        QtyAdditionalLama = RowItemRM.Additional_Qty

                        Exit For
                    End If
                Next
                QtyBalanceBaru = QtyBalanceLama - QtyAdditionalLama + QtyAdditionalBaru
                If QtyBalanceBaru < 0 Then
                    MsgBox("Minimum qty you can enter is: " + Me.ObjCD.DigitGrouping(QtyBalanceBaru + _
                    CDec(Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value), Me.DefPrec) + " " + Me.UnitCode, MsgBoxStyle.Information, "Information")
                    Me.DgvItemRMAdd.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemRMAdd_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItemRMAdd.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.DgvItemRMAdd.Rows(Me.DgvItemRMAdd.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColUnit.Index).Value = Me.UnitCode
            rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefPrec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemBundling_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvItemRMAdd.UserDeletingRow
        Dim QtyBalanceLama, QtyAdditionalLama, QtyBalanceBaru, QtyAdditionalBaru As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

        Try
            'pastikan balance rm tidak negatif
            For Each RowItemRM In Me.TblItemRM.Rows
                If RowItemRM.Identifier_JO = Me.vIdentifierJO And RowItemRM.Identifier_WO = Me.vIdentifireWO And _
                RowItemRM.Primary_Product_Code_RM = Me.Primary_Product_Code_RM Then
                    QtyBalanceLama = RowItemRM.Balance_Qty
                    QtyAdditionalLama = RowItemRM.Additional_Qty

                    Exit For
                End If
            Next
            For Each rowDGV In Me.DgvItemRMAdd.Rows
                If rowDGV.Index <> e.Row.Index Then
                    QtyAdditionalBaru += CDec(rowDGV.Cells(Me.ColQty.Index).Value)
                End If
            Next

            rowDGV = Me.DgvItemRMAdd.Rows(e.Row.Index)
            QtyBalanceBaru = QtyBalanceLama - QtyAdditionalLama + QtyAdditionalBaru
            If QtyBalanceBaru < 0 Then
                MsgBox("Deleting this row will cause negative balance qty.", MsgBoxStyle.Information, "Information")
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_Exp_QtyRMAdd_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim RowItemRMAdd As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addRow
        Dim rowDGV As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DgvItemRMAdd.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            'load data header
            Me.TxtCodeNo.Text = Me.CodeNoRM
            Me.TxtDesc.Text = Me.DescRM
            Me.TxtWONumber.Text = Me.WONumber
            Me.TxtProcessCode.Text = Me.ProcessCode

            'load dgv bundling
            Me.DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.Primary_Product_Code_RM)
            Me.DgvItemRMAdd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItemRMAdd.AllowUserToAddRows = False
            Me.TblItemRMAdd.DefaultView.RowFilter = "[Identifier JO] = '" + Me.vIdentifierJO + "' AND [Identifier WO] = '" + Me.vIdentifireWO + "' AND [Primary Product Code RM] = " + Me.Primary_Product_Code_RM.ToString
            Me.TblItemRMAdd.DefaultView.Sort = "[Add Date], Remarks"
            For Each DV As Data.DataRowView In Me.TblItemRMAdd.DefaultView
                RowItemRMAdd = DV.Row
                Me.DgvItemRMAdd.Rows.Add()
                rowDGV = Me.DgvItemRMAdd.Rows(Me.DgvItemRMAdd.Rows.Count - 1)
                rowDGV.Cells(Me.ColDate.Index).Value = Format(RowItemRMAdd.Add_Date, "dd MMM yyyy")
                rowDGV.Cells(Me.ColNo.Index).Value = Me.DgvItemRMAdd.Rows.Count
                rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(RowItemRMAdd.Qty, Me.DefPrec)
                rowDGV.Cells(Me.ColRemarks.Index).Value = RowItemRMAdd.Remarks
                rowDGV.Cells(Me.ColUnit.Index).Value = Me.UnitCode
            Next
            Me.DgvItemRMAdd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvItemRMAdd.AllowUserToAddRows = True

            If Me.ReadOnly_ = True Then
                Me.btnOK.Enabled = False
                Me.DgvItemRMAdd.AllowUserToAddRows = False
                Me.DgvItemRMAdd.ReadOnly = True
                Me.DgvItemRMAdd.AllowUserToDeleteRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            If Me.DeleteIndex <> -1 Then
                Me.DgvItemRMAdd.Rows.RemoveAt(Me.DeleteIndex)
                Me.ObjCD.SetNumber(Me.DgvItemRMAdd)
                Me.DeleteIndex = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class