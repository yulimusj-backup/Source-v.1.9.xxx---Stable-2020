Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_ProcessOrder
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vPrimaryMO As Double = -1
    Public vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable
    Public vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    Public vUrutan As Integer
    Public vIdentifierJO As String
    Public vBolehEdit As Boolean = False
    Public vDateTransaksi As Date
    Public vReorderHide As Boolean = False
    Public vFrmMO As FrmTrans_ManufacturingOrder
    Public vLgsgTutup As Boolean = False

    Dim objMO As New ClsTrans_ManufacturingOrder
    Dim objMO_Item As New ClsTrans_ManufacturingOrder_Item
    Dim ObjCD As New Core_Module.ClsCekData
    Dim FromOtherButton As Boolean = False
    Dim CancelClose As Boolean = False

    Public Function ApakahBolehReOrderWO() As Boolean
        Dim RowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim ListIdentifierWO As New List(Of String)

        Try
            'cek bundling, bila wo tertentu yang akan di reorder sudah di bundling maka tidak bisa reorder
            Me.vTblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
            If Me.vTblItemWO.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In Me.vTblItemWO.DefaultView
                    RowItemWO = dv.Row
                    If ListIdentifierWO.Contains(RowItemWO.Identifier) = False Then
                        ListIdentifierWO.Add(RowItemWO.Identifier)
                    End If
                Next

                If ListIdentifierWO.Count > 0 Then
                    For Each IdentifierWO As String In ListIdentifierWO
                        If objMO.GetJmlBundling(IdentifierWO, vIdentifierJO) > 0 Then
                            Return False
                        End If
                    Next
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub FrmTrans_ManufacturingOrder_ProcessOrder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Hasil As MsgBoxResult

        Try
            If vBolehEdit = True Then
                If Me.vLgsgTutup = False Then
                    If Me.FromOtherButton = False Then
                        Hasil = MsgBox("Accept data changes before closing form?", MsgBoxStyle.YesNoCancel, "Confirm")
                        If Hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            Me.UpdateTblProsesOrder()
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                            Me.FromOtherButton = True
                            Me.Close()
                        ElseIf Hasil = MsgBoxResult.No Then
                            Me.FromOtherButton = True
                            Me.Close()
                        ElseIf Hasil = MsgBoxResult.Cancel Then
                            e.Cancel = True
                        End If
                    Else
                        If Me.CancelClose = True Then
                            Me.CancelClose = False
                            e.Cancel = True
                        End If
                        Me.FromOtherButton = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_ProcessOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.DgvProsesOrder.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            If vReorderHide = True Then
                Me.BtnReorderWO.Visible = False
            End If
            If vBolehEdit = False Then
                Me.BtnOK.Text = "Close"
                Me.BtnReorderWO.Visible = False
                Me.DgvProsesOrder.ReadOnly = True
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            Me.colPrimary.Visible = False
            Me.colPrimaryProduct.Visible = False
            Me.colUrutan.Visible = False

            Me.DgvProsesOrder.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.objMO_Item.IsiDgvProsesOrder(Me.vIdentifierJO, Me.DgvProsesOrder, vTblProsesOrder, Me)

            If Me.vLgsgTutup = True Then
                Me.UpdateTblProsesOrder()
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.FromOtherButton = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnReorderWO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnReorderWO.Click
        Dim hasil As MsgBoxResult

        Try
            If ApakahBolehReOrderWO() = True Then
                hasil = MsgBox("Reorder?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.ReOrderWO()
                    MsgBox("Reorder success", MsgBoxStyle.Information, "Information")
                End If
            Else
                MsgBox("This following work order data is reference by one or several bundling. You can't reorder this data.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Dim hasil As MsgBoxResult

        Try
            Me.FromOtherButton = False
            If vBolehEdit = True Then
                hasil = MsgBox("Close form and accept data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.UpdateTblProsesOrder()
                    Me.FromOtherButton = True
                    Me.Close()
                End If
            Else
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Public Sub ReOrderWO()
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim Urutan As Integer = 0
        Dim MaxUrutan As Integer = 0
        Dim count As Integer = 0

        Try
            Me.vTblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & Me.vIdentifierJO & "'"
            If Me.vTblItemWO.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In Me.vTblItemWO.DefaultView
                    rowItemWO = dv.Row
                    rowItemWO.WO_Number = ""
                    rowItemWO.Urutan = 0
                Next

                For Each rowD As DataGridViewRow In Me.DgvProsesOrder.Rows
                    Me.vTblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & Me.vIdentifierJO.ToString & "' AND [Primary Process] = " & CDbl(rowD.Cells(Me.colPrimary.Index).Value)
                    Me.vTblItemWO.DefaultView.Sort = "[Level] DESC"
                    If Me.vTblItemWO.DefaultView.Count > 0 Then
                        For Each dv_1 As Data.DataRowView In Me.vTblItemWO.DefaultView
                            rowItemWO = dv_1.Row
                            With rowItemWO
                                Urutan = CInt(rowD.Cells(Me.colNo.Index).Value)
                                If count = 0 Then
                                    MaxUrutan = vFrmMO.getMaxNoUrutWO(Me.vTblItemWO)
                                Else
                                    MaxUrutan = MaxUrutan + 1
                                End If
                                rowItemWO.WO_Number = Me.objMO.GetNoBaru1(vDateTransaksi, _
                                ClsTrans_ManufacturingOrder.PilihFormatNumber.WO, MaxUrutan, vFrmMO, Urutan, Me.vTblItemWO)
                                rowItemWO.Urutan = Urutan
                                count += 1
                            End With
                        Next
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdateTblProsesOrder()
        Dim rowProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderRow

        Try
            Me.vTblProsesOrder.DefaultView.RowFilter = "[identifier jo] = '" & Me.vIdentifierJO.ToString & "'"
            While Me.vTblProsesOrder.DefaultView.Count > 0
                rowProsesOrder = Me.vTblProsesOrder.DefaultView(0).Row
                Me.vTblProsesOrder.Rows.Remove(rowProsesOrder)
            End While

            For Each rowD As DataGridViewRow In Me.DgvProsesOrder.Rows
                rowProsesOrder = Me.vTblProsesOrder.Newtbltrans_mo_processorderRow
                With rowProsesOrder
                    .primary = -1
                    .primary_mo = Me.vPrimaryMO
                    .primary_process = CDbl(rowD.Cells(Me.colPrimary.Index).Value)
                    .primary_product = CDbl(rowD.Cells(Me.colPrimaryProduct.Index).Value)
                    .identifier_jo = Me.vIdentifierJO
                    .urutan = rowD.Cells(Me.colNo.Index).Value
                End With
                Me.vTblProsesOrder.Rows.Add(rowProsesOrder)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Dim hasil As MsgBoxResult

        Try
            hasil = MsgBox("Close form and discard data changes?", MsgBoxStyle.YesNo, "Confirmation")
            If hasil = MsgBoxResult.Yes Then
                Me.FromOtherButton = True
                Me.Close()
            Else
                Me.FromOtherButton = True
                Me.CancelClose = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.DgvProsesOrder.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.DgvProsesOrder.FirstDisplayedScrollingRowIndex = Me.DgvProsesOrder.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.DgvProsesOrder.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.DgvProsesOrder.Rows(row.Index).Displayed = False
                    Me.DgvProsesOrder.FirstDisplayedScrollingRowIndex = Me.DgvProsesOrder.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.DgvProsesOrder.DisplayedRowCount(True) <> Me.DgvProsesOrder.DisplayedRowCount(False) Then
                    Me.DgvProsesOrder.FirstDisplayedScrollingRowIndex = Me.DgvProsesOrder.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.DgvProsesOrder.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.DgvProsesOrder.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.DgvProsesOrder.Rows.Remove(row)
                    Me.DgvProsesOrder.Rows.InsertRange(indexbaru, row)
                    Me.DgvProsesOrder.Focus()
                    Me.DgvProsesOrder.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.DgvProsesOrder)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.DgvProsesOrder.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.DgvProsesOrder.SelectedRows(0)
                If row.Index <> Me.DgvProsesOrder.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.DgvProsesOrder.Rows.Remove(row)
                    Me.DgvProsesOrder.Rows.InsertRange(indexbaru, row)
                    Me.DgvProsesOrder.Focus()
                    Me.DgvProsesOrder.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.DgvProsesOrder)
                    Me.ShowRow(row, False)
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub DgvProsesOrder_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvProsesOrder.CellBeginEdit
    '    Try
    '        e.Cancel = True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub DgvProsesOrder_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvProsesOrder.SelectionChanged
        Try
            If Me.DgvProsesOrder.SelectedRows.Count > 0 And Me.DgvProsesOrder.Rows.Count >= 1 Then
                If Me.vBolehEdit = False Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                Else
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                End If
            Else
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvProsesOrder_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvProsesOrder.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvProsesOrder)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class