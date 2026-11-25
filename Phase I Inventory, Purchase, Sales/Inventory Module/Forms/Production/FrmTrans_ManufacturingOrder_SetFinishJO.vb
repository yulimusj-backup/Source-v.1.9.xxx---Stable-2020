Option Explicit On
Option Strict Off
Public Class FrmTrans_ManufacturingOrder_SetFinishJO
    Public UserCode As String
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim objMO As New ClsTrans_ManufacturingOrder

    Dim tbl As New DseTrans_ManufacturingOrder_SetFinishJO.DataTable1DataTable
    Dim row As DseTrans_ManufacturingOrder_SetFinishJO.DataTable1Row
    Dim objCd As New Core_Module.ClsCekData
    Dim buffer As Object

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_SetFinishJO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dtpFrom.Value = Now.AddMonths(-1).Date
            Me.dtpTo.Value = Now.Date
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim dta As New DseTrans_ManufacturingOrder_SetFinishJOTableAdapters.DataTable1TableAdapter
        Dim isClosed, isClosedII As Integer
        Dim dgr As DataGridViewRow
        Dim Lanjut As Boolean = False

        Try
            If Me.dtpTo.Value.Date < Me.dtpFrom.Value.Date Then
                MsgBox("Invalid MO Date", MsgBoxStyle.Information, "Information")
            End If
            Me.Cursor = Cursors.WaitCursor
            If Me.chkOngoingJOOnly.Checked = True Then
                isClosed = 0
                isClosedII = 0
            Else
                isClosed = 1
                isClosedII = 0
            End If
            Me.dgvItem.Rows.Clear()
            dta.Connection.ConnectionString = Core_Module.GetConstring

            If dta.Fill(Me.tbl, 1, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, isClosed, isClosedII) > 0 Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                For Each Me.row In Me.tbl
                    If Me.ChkShow100PectOnly.Checked = True Then
                        If row.Balance_Produce <= 0 Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    Else
                        Lanjut = True
                    End If
                    If Lanjut = True Then
                        Me.dgvItem.Rows.Add()
                        dgr = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        With dgr
                            .Cells(Me.ColItem_No.Index).Value = Me.dgvItem.Rows.Count
                            .Cells(Me.ColItem_MONumber.Index).Value = row.MO_Number
                            .Cells(Me.ColItem_JONumber.Index).Value = row.JO_Number
                            .Cells(Me.ColItem_CodeNo.Index).Value = row.Product_Code
                            .Cells(Me.ColItem_Desc.Index).Value = row.Description
                            .Cells(Me.ColItem_QtyTarget.Index).Value = Me.objCd.DigitGrouping(row.Target_Qty, row._Decimal)
                            .Cells(Me.ColItem_QtyProduce.Index).Value = Me.objCd.DigitGrouping(row.Produce_Qty, row._Decimal)
                            .Cells(Me.ColItem_Unit.Index).Value = row.Unit_Code
                            .Cells(Me.ColItem_Remarks.Index).Value = row.Remarks
                            If row.Is_Closed = 0 Then
                                .Cells(Me.ColItem_IsClosed.Index).Value = False
                            Else
                                .Cells(Me.ColItem_IsClosed.Index).Value = True
                            End If
                            .Cells(Me.ColItem_IdentifierJO.Index).Value = row.Identifier
                        End With
                    End If
                Next
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Else
                MsgBox("No data on this period", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColItem_IsClosed.Index Then
                e.Cancel = True
            Else
                buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim frmDate As New FrmTrans_ManufacturingOrder_SetFinishJO_Date
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> Me.buffer Then
                If CBool(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    If Me.objMO.ApakahBolehCloseJO(Me.dgvItem.Item(Me.ColItem_IdentifierJO.Index, e.RowIndex).Value, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If

                    If frmDate.ShowDialog <> Windows.Forms.DialogResult.OK Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    If UpdateStatus(Me.dgvItem.Item(Me.ColItem_IdentifierJO.Index, e.RowIndex).Value, CBool(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), _
                       frmDate.DateTimePicker1.Value.Date) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                Else
                    If UpdateStatus(Me.dgvItem.Item(Me.ColItem_IdentifierJO.Index, e.RowIndex).Value, CBool(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), _
                                            #1/1/1950#) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Public Function UpdateStatus(ByVal vIdentifierJO As String, ByVal vClosed As Boolean, ByVal vClosingDate As Date) As Boolean
        Dim dta As New DseTrans_ManufacturingOrder_SetFinishJOTableAdapters.DataTable1TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If vClosed = True Then
                dta.UpdateWOStatus(1, vClosingDate, vIdentifierJO)
                dta.UpdateBalanceRM(0, vIdentifierJO)
            Else
                dta.UpdateWOStatus(0, #1/1/1950#, vIdentifierJO)
                dta.UpdateBalanceRM_Recalculate(vIdentifierJO)
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Dim frmDate As New FrmTrans_ManufacturingOrder_SetFinishJO_Date
        Dim i As Integer = 0
        Dim listOfFail As New List(Of String)
        Try
            If Me.dgvItem.Rows.Count > 0 Then
                If MsgBox("Are you sure?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    If frmDate.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Me.Cursor = Cursors.WaitCursor
                        Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                        Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                        For Each row As DataGridViewRow In Me.dgvItem.Rows
                            If row.Cells(Me.ColItem_IsClosed.Index).Value = False Then
                                If Me.objMO.ApakahBolehCloseJO(row.Cells(Me.ColItem_IdentifierJO.Index).Value, False) = True Then
                                    If Me.UpdateStatus(row.Cells(Me.ColItem_IdentifierJO.Index).Value, True, frmDate.DateTimePicker1.Value.Date) = True Then
                                        row.Cells(Me.ColItem_IsClosed.Index).Value = True
                                        i += 1
                                    End If
                                Else
                                    listOfFail.Add(row.Cells(Me.ColItem_JONumber.Index).Value)
                                End If
                            End If
                        Next
                        Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                        Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                        If listOfFail.Count > 0 Then
                            Dim s0 As String = ""
                            For Each s1 As String In listOfFail
                                s0 += s1 & ","
                            Next
                            s0 = s0.Remove(s0.Length - 1, 1)
                            MsgBox("The following JO isn't authorize, you can only close authorized JO:" & vbNewLine & s0, MsgBoxStyle.Information, "Information")
                        Else
                            MsgBox(i.ToString & " rows is closed.", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub llbUnCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUnCheckAll.LinkClicked
        Dim i As Integer = 0
        Try
            If Me.dgvItem.Rows.Count > 0 Then
                If MsgBox("Are you sure?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    Me.Cursor = Cursors.WaitCursor
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                    For Each row As DataGridViewRow In Me.dgvItem.Rows
                        If row.Cells(Me.ColItem_IsClosed.Index).Value = True Then
                            If Me.UpdateStatus(row.Cells(Me.ColItem_IdentifierJO.Index).Value, False, #1/1/1950#) = True Then
                                row.Cells(Me.ColItem_IsClosed.Index).Value = False
                                i += 1
                            End If
                        End If
                    Next
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                    MsgBox(i.ToString & " rows is updated.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    
End Class