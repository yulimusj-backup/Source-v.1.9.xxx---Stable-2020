Partial Public Class FrmMasterSettingPajak
    Private Sub dgvRateFinal_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvRateFinal.UserDeletingRow
        Try
            Me.dgvRateFinal.AllowUserToAddRows = False
            If Me.dgvRateFinal.Rows.Count <= 1 Then
                e.Cancel = True
            End If
            Me.dgvRateFinal.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvRateFinal_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvRateFinal.CellEndEdit
        Dim Row As Windows.Forms.DataGridViewRow = Me.dgvRateFinal.Rows(e.RowIndex)
        Dim RowIndex_Atas As Integer = e.RowIndex - 1
        Dim RowIndex_Bawah As Integer = e.RowIndex + 1

        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvRateFinal.AllowUserToAddRows = False
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Me.dgvRateFinal.AllowUserToAddRows = True
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColRate_Maximum.Index Then
                        If Me.ObjCD.CekBilanganPositif(.Cells(Me.ColRate_Maximum.Index).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            Else
                                Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                            Me.dgvRateFinal.AllowUserToAddRows = True
                            Exit Sub
                        Else
                            Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        If CDec(.Cells(Me.ColRate_Maximum.Index).Value) > Me.MAXLength_Angka Then
                            MsgBox("Max Length number 12-digits.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            Else
                                Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                            Me.dgvRateFinal.AllowUserToAddRows = True
                            Exit Sub
                        Else
                            Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        If Me.dgvRateFinal.Rows.Count > 1 Then
                            'batas atas
                            If e.RowIndex > 0 Then
                                Dim MAXRate As Decimal = Me.dgvRateFinal.Item(e.ColumnIndex, RowIndex_Atas).Value
                                If CDec(Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value) < MAXRate Then
                                    MsgBox("You cannot input maximum income lower than " + ObjCD.DigitGrouping(MAXRate, Me.DefPrec) + " .", MsgBoxStyle.Information, "Information")
                                    If Me.Buffer <> "" Then
                                        Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                                    Else
                                        Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    If Me.ApakahCellRateKosong(Row) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    End If
                                    Me.dgvRateFinal.AllowUserToAddRows = True
                                    Exit Sub
                                End If
                            End If
                            'batas bawah
                            If e.RowIndex < Me.dgvRateFinal.Rows.Count - 1 Then
                                Dim MAXRate As Decimal = Me.dgvRateFinal.Item(e.ColumnIndex, RowIndex_Bawah).Value
                                If CDec(Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value) > MAXRate Then
                                    MsgBox("You cannot input maximum income larger than " + ObjCD.DigitGrouping(MAXRate, Me.DefPrec) + " .", MsgBoxStyle.Information, "Information")
                                    If Me.Buffer <> "" Then
                                        Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                                    Else
                                        Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    If Me.ApakahCellRateKosong(Row) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    End If
                                    Me.dgvRateFinal.AllowUserToAddRows = True
                                    Exit Sub
                                End If
                            End If
                        End If
                        If Me.dgvRateFinal.Item(Me.ColRate_Tarif.Index, e.RowIndex).Value = "" Or CDec(Me.dgvRateFinal.Item(Me.ColRate_Tarif.Index, e.RowIndex).Value) = 0 Then
                            Me.dgvRateFinal.Item(Me.ColRate_Tarif.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(1, Me.DefPrec_Persen)
                        End If
                    End If
                    If e.ColumnIndex = Me.ColRate_Tarif.Index Then
                        If Me.ObjCD.CekNonNegatif(.Cells(Me.ColRate_Tarif.Index).Value) = False Then
                            MsgBox("Please enter non negative number", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            Else
                                Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                            Me.dgvRateFinal.AllowUserToAddRows = True
                            Exit Sub
                        Else
                            Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        If CDec(.Cells(Me.ColRate_Tarif.Index).Value) > Me.MAXLength_Tarif Then
                            MsgBox("Maximum value tarif = 100.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec_Persen)
                            Else
                                Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                            Me.dgvRateFinal.AllowUserToAddRows = True
                            Exit Sub
                        Else
                            Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        If Me.dgvRateFinal.Item(Me.ColRate_Maximum.Index, e.RowIndex).Value = "0" Or _
                        Me.dgvRateFinal.Item(Me.ColRate_Maximum.Index, e.RowIndex).Value = "" Then
                            MsgBox("Please input maximum income.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec_Persen)
                            Else
                                Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                            Me.dgvRateFinal.AllowUserToAddRows = True
                            Exit Sub
                        End If
                    End If
                    If e.ColumnIndex = Me.ColRate_Remarks.Index Then
                        If Me.ObjCD.CekDataKosong(.Cells(Me.ColRate_Remarks.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                            Me.dgvRateFinal.AllowUserToAddRows = True
                            Exit Sub
                        End If
                    End If
                    If Me.ApakahCellRateKosong(Row, False) = True Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker2.RunWorkerAsync()
                    End If
                    Me.dgvRateFinal.AllowUserToAddRows = True
                Else
                    If e.ColumnIndex = Me.ColRate_Remarks.Index Then
                        Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvRateFinal_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvRateFinal.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColRate_Maximum.Index Or _
            e.ColumnIndex = Me.ColRate_Remarks.Index Or _
            e.ColumnIndex = Me.ColRate_Tarif.Index Then
                Me.Buffer = Me.dgvRateFinal.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.dgvRateFinal.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
