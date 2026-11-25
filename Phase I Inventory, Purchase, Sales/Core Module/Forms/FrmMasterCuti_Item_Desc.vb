''@rosadinnn
''@20150908 1504
Option Explicit On
Option Strict Off

Public Class FrmMasterCuti_Item_Desc

    Public mIdentifier As String = ""
    Public FrmRef As FrmMasterCuti
    Public mTipe As FrmMasterCuti.EnumTipeLeave = FrmMasterCuti.EnumTipeLeave.Addition
    Public mYear As Integer = 0
    Public mTotalAddition As Double = 0
    Public mTotalDeduction As Double = 0
    Public mTotalAdditionBuffer As Double = 0
    Public mTotalDeductionBuffer As Double = 0
    Public mPrimaryKaryawan As Double = -1
    Public mReadOnly As Boolean = False

    Dim vTblItemDesc As New DseMasterCuti.tblmaster_cuti_item_descDataTable
    Dim vRowItemDesc As DseMasterCuti.tblmaster_cuti_item_descRow
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCuti As New ClsMasterCuti
    Dim ObjItem As New ClsMasterCuti_Item
    Dim mDeleteIndex As Integer = -1
    Dim Buffer As String = ""
    Dim LangsungTutup As Boolean = False

    Private Sub FrmMasterCuti_Item_Desc_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow
        Try
            Me.DgvItem.AllowUserToAddRows = False
            Me.vTblItemDesc = Me.FrmRef.vTblItemDesc
            For Each Me.vRowItemDesc In Me.vTblItemDesc.Rows
                If vRowItemDesc.identifier = Me.mIdentifier Then
                    Me.DgvItem.Rows.Add()
                    rowDGV = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.DgvItem.Rows.Count
                    If vRowItemDesc.tipe = FrmMasterCuti.EnumTipeLeave.Addition Then
                        rowDGV.Cells(Me.ColTipe.Index).Value = "Addition"
                    ElseIf vRowItemDesc.tipe = FrmMasterCuti.EnumTipeLeave.Deduction Then
                        rowDGV.Cells(Me.ColTipe.Index).Value = "Deduction"
                    End If
                    rowDGV.Cells(Me.ColDate.Index).Value = Format(vRowItemDesc._date, "dd MMM yyyy")
                    rowDGV.Cells(Me.ColDesc.Index).Value = vRowItemDesc.desc
                    rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(vRowItemDesc.qty, 1)
                    If vRowItemDesc.tipe2 = FrmMasterCuti.EnumTipeLeave2.MonthlyLeaveBySystem Then
                        rowDGV.Cells(Me.ColTipe2.Index).Value = "MonthlyLeaveBySystem"
                    ElseIf vRowItemDesc.tipe2 = FrmMasterCuti.EnumTipeLeave2.LongLeaveBySystem Then
                        rowDGV.Cells(Me.ColTipe2.Index).Value = "LongLeaveBySystem"
                    ElseIf vRowItemDesc.tipe2 = FrmMasterCuti.EnumTipeLeave2.Add_Ded_ByUser Then
                        rowDGV.Cells(Me.ColTipe2.Index).Value = "Add_Ded_ByUser"
                    End If
                End If
            Next
            RefreshTotalAdditionDeduction()
            If Me.mReadOnly = False Then
                Me.DgvItem.AllowUserToAddRows = True
            Else
                Me.DgvItem.AllowUserToDeleteRows = False
                Me.DgvItem.AllowUserToAddRows = False
                Me.DgvItem.ReadOnly = True
                Me.BtnOK.Visible = False
                Me.BtnOK.TabStop = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColNo.Index Then
                e.Cancel = True
                Exit Sub
            Else
                If Me.DgvItem.Item(Me.ColTipe2.Index, e.RowIndex).Value = "LongLeaveBySystem" Or Me.DgvItem.Item(Me.ColTipe2.Index, e.RowIndex).Value = "MonthlyLeaveBySystem" Then
                    e.Cancel = True
                    Exit Sub
                End If
                If e.ColumnIndex <> Me.ColDate.Index And String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColDate.Index, e.RowIndex).Value) = True Then
                    e.Cancel = True
                    Exit Sub
                Else
                    Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Try
            If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColDate.Index Then
                    If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a valid date.", MsgBoxStyle.Information, "Information")
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        If e.RowIndex = Me.DgvItem.Rows.Count - 2 Then
                            Me.mDeleteIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                        Exit Sub
                    End If
                    If CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Year <> Me.mYear Then
                        MsgBox("Date should be in year: " & Me.mYear.ToString & ".", MsgBoxStyle.Information, "Information")
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        If e.RowIndex = Me.DgvItem.Rows.Count - 2 Then
                            Me.mDeleteIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                        Exit Sub
                    End If
                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                End If
                If e.ColumnIndex = Me.ColQty.Index Then
                    If Me.ObjCD.CekNonNegatif(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Qty must input positive number.", MsgBoxStyle.Information, "Information")
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.ObjCD.MiddleRound(0.5, Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value), 1)

                        Me.RefreshTotalAdditionDeduction()
                        If Me.DgvItem.Item(Me.ColTipe.Index, e.RowIndex).Value = "Addition" Then
                            If CDbl(Me.mTotalAddition) < CDbl(Me.mTotalAdditionBuffer) Then
                                If Me.ObjCuti.ApakahEditSah(Me.FrmRef.TxtPrimary.Text, Me.mPrimaryKaryawan, _
                                 Me.mTotalAddition, Me.mTotalDeduction, CDbl(FrmRef.TxtLeaveQty.Text), ClsMasterCuti.EnumAddDed.add) = False Then
                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                        ElseIf Me.DgvItem.Item(Me.ColTipe.Index, e.RowIndex).Value = "Deduction" Then
                            If CDbl(Me.mTotalDeduction) > CDbl(Me.mTotalDeductionBuffer) Then
                                If Me.ObjCuti.ApakahEditSah(Me.FrmRef.TxtPrimary.Text, Me.mPrimaryKaryawan, _
                                 Me.mTotalAddition, Me.mTotalDeduction, CDbl(Me.FrmRef.TxtLeaveQty.Text), ClsMasterCuti.EnumAddDed.ded) = False Then
                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                        End If

                    End If
                End If

                If e.ColumnIndex = Me.ColTipe.Index Then
                    Me.RefreshTotalAdditionDeduction()
                    If Me.DgvItem.Item(Me.ColTipe.Index, e.RowIndex).Value = "Addition" Then
                        If CDbl(Me.mTotalAddition) < CDbl(Me.mTotalAdditionBuffer) Then
                            If Me.ObjCuti.ApakahEditSah(Me.FrmRef.TxtPrimary.Text, Me.mPrimaryKaryawan, _
                             Me.mTotalAddition, Me.mTotalDeduction, CDbl(FrmRef.TxtLeaveQty.Text), ClsMasterCuti.EnumAddDed.add) = False Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If
                    ElseIf Me.DgvItem.Item(Me.ColTipe.Index, e.RowIndex).Value = "Deduction" Then
                        If CDbl(Me.mTotalDeduction) > CDbl(Me.mTotalDeductionBuffer) Then
                            If Me.ObjCuti.ApakahEditSah(Me.FrmRef.TxtPrimary.Text, Me.mPrimaryKaryawan, _
                             Me.mTotalAddition, Me.mTotalDeduction, CDbl(Me.FrmRef.TxtLeaveQty.Text), ClsMasterCuti.EnumAddDed.ded) = False Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvItem.DataError
        'escape function, ketika rows added pada saat load form, saat fungsi yang mengisikan data kolom bertipe combobox pada kolom priority desc belum diisikan
    End Sub

    Private Sub DgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            rowDGV = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.DgvItem.Rows.Count
            If Me.mTipe = FrmMasterCuti.EnumTipeLeave.Addition Then
                rowDGV.Cells(Me.ColTipe.Index).Value = "Addition"
            ElseIf Me.mTipe = FrmMasterCuti.EnumTipeLeave.Deduction Then
                rowDGV.Cells(Me.ColTipe.Index).Value = "Deduction"
            End If
            rowDGV.Cells(Me.ColDate.Index).Value = ""
            rowDGV.Cells(Me.ColDesc.Index).Value = ""
            rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping("0", 1)
            rowDGV.Cells(Me.ColTipe2.Index).Value = "Add_Ded_ByUser"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvItem.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Try
            For Each rowdgv As DataGridViewRow In Me.DgvItem.Rows
                If CDec(rowdgv.Cells(Me.ColQty.Index).Value) <= 0 And rowdgv.Index <> Me.DgvItem.RowCount - 1 Then
                    MsgBox("Please specify positive qty data for row: " + (rowdgv.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            Next
            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                RefreshTotalAdditionDeduction()
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.LangsungTutup = True
                Me.Close()
            Else
                Me.LangsungTutup = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Threading.Thread.Sleep(10)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvItem.Rows.RemoveAt(Me.mDeleteIndex)
            Me.mDeleteIndex = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCuti_Item_Desc_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.mReadOnly = False Then
                If Me.LangsungTutup = False Then
                    If MsgBox("Close dan discard data changes?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RefreshTotalAdditionDeduction()
        Try
            mTotalAddition = 0
            mTotalDeduction = 0
            For Each rowdgv As DataGridViewRow In Me.DgvItem.Rows
                If IsNothing(rowdgv.Cells(Me.ColQty.Index).Value) = False Then
                    If rowdgv.Cells(Me.ColTipe.Index).Value = "Addition" Then
                        Me.mTotalAddition += CDbl(rowdgv.Cells(Me.ColQty.Index).Value)
                    ElseIf rowdgv.Cells(Me.ColTipe.Index).Value = "Deduction" Then
                        Me.mTotalDeduction += CDbl(rowdgv.Cells(Me.ColQty.Index).Value)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvItem.UserDeletingRow
        Try
            RefreshTotalAdditionDeduction()
            If Me.DgvItem.Item(Me.ColTipe.Index, e.Row.Index).Value = "Addition" Then
                If Me.ObjCuti.ApakahEditSah(Me.FrmRef.TxtPrimary.Text, Me.mPrimaryKaryawan, _
                     Me.mTotalAddition - CDbl(Me.DgvItem.Item(Me.ColQty.Index, e.Row.Index).Value), Me.mTotalDeduction, CDbl(FrmRef.TxtLeaveQty.Text), ClsMasterCuti.EnumAddDed.add) = False Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class