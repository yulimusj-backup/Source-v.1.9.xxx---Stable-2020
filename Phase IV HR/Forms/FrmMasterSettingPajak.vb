Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class FrmMasterSettingPajak
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll

    Dim Buffer As String = ""
    Dim SaveDanClose As Boolean = True
    Dim RunEventClosing As Boolean = True
    Dim ObjPTKP As New ClsMasterSettingPajak
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim DefPrec As Double = Me.ObjCurr.GetDecimal_DefaultCurrency
    Dim DefPrec_Persen As Double = 2
    Dim MAXLength_Tarif As Decimal = 100
    Dim MAXLength_Angka As Decimal = 999999999999
    Dim MAXLength_Angka_Str As String = Me.ObjCD.DigitGrouping(Me.MAXLength_Angka, 0)
    Dim EraseRowIndex As Double = 0

    Private Sub FrmMasterSettingPajak_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.RunEventClosing = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ObjPTKP.ApakahDataBenar(Me.dgvPTKP, Me.dgvRate, True, Me, Me.dgvRateFinal) = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjPTKP.Simpan(Me.dgvPTKP, Me.dgvRate, Me, Me.dgvRateFinal) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Else
                            e.Cancel = True
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSettingPajak_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.RunEventClosing = True
            Me.dgvPTKP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvRate.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvRateFinal.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            'Cek Apakah Sudah Ada Data Yang Disimpan Ke database
            If Me.ObjPTKP.ApakahSudahAda_PTKP() = True Then
                'load data
                Me.ObjPTKP.IsiDGVItem(Me.dgvPTKP, Me.dgvRate, Me, Me.dgvRateFinal)
            Else
                'isi data baru
                Me.ObjPTKP.IsiDGVItem_Default(Me.dgvPTKP, Me.dgvRate, Me, Me.dgvRateFinal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.RunEventClosing = False
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPTKP_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvPTKP.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColPTKP_PTKP_Tahunan.Index Or _
            e.ColumnIndex = Me.ColPTKP_Remarks.Index Then
                Me.Buffer = Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPTKP_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPTKP.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvPTKP.Rows(e.RowIndex)
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColPTKP_PTKP_Tahunan.Index Then
                        If IsNumeric(Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If IsNumeric(Me.Buffer) = True Then
                                Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer)
                            Else
                                Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = 0
                            End If
                            Exit Sub
                        End If
                        If Me.ObjCD.CekBilanganPositif(.Cells(Me.ColPTKP_PTKP_Tahunan.Index).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                            Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            Exit Sub
                        Else
                            Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        If CDec(.Cells(Me.ColPTKP_PTKP_Tahunan.Index).Value) > Me.MAXLength_Angka Then
                            MsgBox("Max Length number 12-digits.", MsgBoxStyle.Information, "Information")
                            Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            Exit Sub
                        Else
                            Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                    End If
                    If e.ColumnIndex = Me.ColPTKP_Remarks.Index Then
                        If Me.ObjCD.CekDataKosong(.Cells(Me.ColPTKP_Remarks.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColPTKP_Remarks.Index Then
                        Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvPTKP.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvRate_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvRate.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColRate_Maximum.Index Or _
            e.ColumnIndex = Me.ColRate_Remarks.Index Or _
            e.ColumnIndex = Me.ColRate_Tarif.Index Then
                Me.Buffer = Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahCellRateKosong(ByVal vRow As DataGridViewRow, Optional ByVal CekRemarks As Boolean = True) As Boolean
        Try
            With vRow
                If .Cells(Me.ColRate_Maximum.Index).Value = "" Or CDec(.Cells(Me.ColRate_Maximum.Index).Value) = 0 Then
                    If .Cells(Me.ColRate_Tarif.Index).Value = "" Or CDec(.Cells(Me.ColRate_Tarif.Index).Value) = 0 Then
                        If CekRemarks = True Then
                            If .Cells(Me.ColRate_Remarks.Index).Value = "" Then
                                Return True
                            End If
                        Else
                            Return True
                        End If
                    End If
                End If
            End With
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvRate_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvRate.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvRate.Rows(e.RowIndex)
        Dim RowIndex_Atas As Integer = e.RowIndex - 1
        Dim RowIndex_Bawah As Integer = e.RowIndex + 1
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvRate.AllowUserToAddRows = False
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Me.dgvRate.AllowUserToAddRows = True
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColRate_Maximum.Index Then
                        If Me.ObjCD.CekBilanganPositif(.Cells(Me.ColRate_Maximum.Index).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            Else
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Me.dgvRate.AllowUserToAddRows = True
                            Exit Sub
                        Else
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        If CDec(.Cells(Me.ColRate_Maximum.Index).Value) > Me.MAXLength_Angka Then
                            MsgBox("Max Length number 12-digits.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            Else
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Me.dgvRate.AllowUserToAddRows = True
                            Exit Sub
                        Else
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        If Me.dgvRate.Rows.Count > 1 Then
                            'batas atas
                            If e.RowIndex > 0 Then
                                Dim MAXRate As Decimal = Me.dgvRate.Item(e.ColumnIndex, RowIndex_Atas).Value
                                If CDec(Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value) < MAXRate Then
                                    MsgBox("You cannot input maximum income lower than " + ObjCD.DigitGrouping(MAXRate, Me.DefPrec) + " .", MsgBoxStyle.Information, "Information")
                                    If Me.Buffer <> "" Then
                                        Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                                    Else
                                        Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    If Me.ApakahCellRateKosong(Row) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                    Me.dgvRate.AllowUserToAddRows = True
                                    Exit Sub
                                End If
                            End If
                            'batas bawah
                            If e.RowIndex < Me.dgvRate.Rows.Count - 1 Then
                                Dim MAXRate As Decimal = Me.dgvRate.Item(e.ColumnIndex, RowIndex_Bawah).Value
                                If CDec(Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value) > MAXRate Then
                                    MsgBox("You cannot input maximum income larger than " + ObjCD.DigitGrouping(MAXRate, Me.DefPrec) + " .", MsgBoxStyle.Information, "Information")
                                    If Me.Buffer <> "" Then
                                        Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                                    Else
                                        Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    If Me.ApakahCellRateKosong(Row) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                    Me.dgvRate.AllowUserToAddRows = True
                                    Exit Sub
                                End If
                            End If
                        End If
                        If Me.dgvRate.Item(Me.ColRate_Tarif.Index, e.RowIndex).Value = "" Or CDec(Me.dgvRate.Item(Me.ColRate_Tarif.Index, e.RowIndex).Value) = 0 Then
                            Me.dgvRate.Item(Me.ColRate_Tarif.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(1, Me.DefPrec_Persen)
                        End If
                    End If
                    If e.ColumnIndex = Me.ColRate_Tarif.Index Then
                        If Me.ObjCD.CekBilanganPositif(.Cells(Me.ColRate_Tarif.Index).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec)
                            Else
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Me.dgvRate.AllowUserToAddRows = True
                            Exit Sub
                        Else
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        If CDec(.Cells(Me.ColRate_Tarif.Index).Value) > Me.MAXLength_Tarif Then
                            MsgBox("Maximum value tarif = 100.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec_Persen)
                            Else
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Me.dgvRate.AllowUserToAddRows = True
                            Exit Sub
                        Else
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping( _
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec)
                        End If
                        If Me.dgvRate.Item(Me.ColRate_Maximum.Index, e.RowIndex).Value = "0" Or _
                        Me.dgvRate.Item(Me.ColRate_Maximum.Index, e.RowIndex).Value = "" Then
                            MsgBox("Please input maximum income.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.DefPrec_Persen)
                            Else
                                Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Me.dgvRate.AllowUserToAddRows = True
                            Exit Sub
                        End If
                    End If
                    If e.ColumnIndex = Me.ColRate_Remarks.Index Then
                        If Me.ObjCD.CekDataKosong(.Cells(Me.ColRate_Remarks.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.ApakahCellRateKosong(Row) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Me.dgvRate.AllowUserToAddRows = True
                            Exit Sub
                        End If
                    End If
                    If Me.ApakahCellRateKosong(Row, False) = True Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                    Me.dgvRate.AllowUserToAddRows = True
                Else
                    If e.ColumnIndex = Me.ColRate_Remarks.Index Then
                        Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvRate_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvRate.UserDeletingRow
        Try
            Me.dgvRate.AllowUserToAddRows = False
            If Me.dgvRate.Rows.Count <= 1 Then
                e.Cancel = True
            End If
            Me.dgvRate.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'konfirmasi disimpan atau tidak
            If Me.ObjPTKP.ApakahDataBenar(Me.dgvPTKP, Me.dgvRate, True, Me, Me.dgvRateFinal) = True Then
                hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If Me.ObjPTKP.Simpan(Me.dgvPTKP, Me.dgvRate, Me, Me.dgvRateFinal) = True Then
                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        If Me.SaveDanClose = True Then
                            Me.RunEventClosing = False
                            Me.Close()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvRate.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class