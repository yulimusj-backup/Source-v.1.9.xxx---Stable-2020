Option Explicit On
Option Strict Off

Public Class FrmMasterBudget
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ObjCD As New Core_Module.ClsCekData

    Dim ObjBudget As New ClsMasterBudget
    Dim mBuffer As String = ""
    Dim DataChanged As Boolean = False
    Dim AutoLoad As Boolean = False
    Dim PreviousPeriod As Date = #1/1/1950#
    Dim Def_Prec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency

    Public Enum EnumFormatYearly As Integer
        Balance_Sheet
        Profit_N_Loss
    End Enum

    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            If Me.cboFormat.SelectedIndex <> -1 Then
                If Me.AutoLoad = True Then
                    If Me.cboFormat.SelectedIndex = 0 Then
                        Me.ObjBudget.CariDetail(Me.dgvBudget, Me, Me.ObjBudget.GetPrimary(Me.dtpPeriod.Value.Date, Me.cboFilter.SelectedIndex), Me.cboFilter.SelectedIndex)
                        Me.dgvBudget.BringToFront()
                    Else
                        'simpan perubahan data jika perlu
                        If Me.DataChanged = True Then
                            Me.btnSave.PerformClick()
                        End If
                        Me.DataChanged = False
                        Me.ObjBudget.IsiDGV_Tahunan(Me.dgvYearly, Me, Me.dtpPeriod.Value.Date, Me.cboFilter.SelectedIndex)
                        Me.dgvYearly.BringToFront()
                    End If
                End If

                If Me.cboFormat.SelectedIndex = 0 Then
                    If Me.dgvBudget.SelectedCells.Count > 0 Then
                        If Me.dgvBudget.Rows(Me.dgvBudget.SelectedCells(0).RowIndex).Cells(Me.ColHeader.Index).Value = 0 Then
                            Me.btnShowDetail.Visible = True
                        Else
                            Me.btnShowDetail.Visible = False
                        End If
                    Else
                        Me.btnShowDetail.Visible = False
                    End If
                Else
                    Me.btnShowDetail.Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiBudget_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.cboFormat.SelectedIndex = 0
            Me.cboFilter.SelectedIndex = 0
            Me.dgvBudget.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvYearly.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dtpPeriod.Value = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpPeriod.Value)
            Me.ObjBudget.CariDetail(Me.dgvBudget, Me, Me.ObjBudget.GetPrimary(Me.dtpPeriod.Value.Date, Me.cboFilter.SelectedIndex), Me.cboFilter.SelectedIndex)
            Me.PreviousPeriod = Me.dtpPeriod.Value.Date
            Me.cboFormat.SelectedIndex = 1
            Me.cboFormat.SelectedIndex = 0
            Me.AutoLoad = True
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShowDetail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowDetail.Click
        Dim ij As New FrmMasterBudget_Detail

        Try
            If Me.dgvBudget.SelectedCells.Count > 0 Then
                ij.UserCode = Me.UserCode
                ij.BackColorAll = Me.BackColorAll
                ij.Ref_Form = Me
                ij.Ref_Row = Me.dgvBudget.Rows(Me.dgvBudget.SelectedCells(0).RowIndex)
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Public Sub UpdateSaldoHeader()
        Dim rowdgv As DataGridViewRow
        Me.dgvBudget.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        Me.dgvBudget.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

        Try
            For Each rowdgv In Me.dgvBudget.Rows
                If rowdgv.Cells(Me.ColHeader.Index).Value = 1 Then
                    'hitung total jan
                    rowdgv.Cells(Me.ColJan.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                    rowdgv.Cells(Me.ColLevel.Index).Value, Me.ColJan.Index), Me.Def_Prec)
                    'hitung total feb
                    rowdgv.Cells(Me.ColFeb.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                    rowdgv.Cells(Me.ColLevel.Index).Value, Me.ColFeb.Index), Me.Def_Prec)
                    'hitung total mar
                    rowdgv.Cells(Me.ColMar.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                    rowdgv.Cells(Me.ColLevel.Index).Value, Me.ColMar.Index), Me.Def_Prec)
                    'hitung total apr
                    rowdgv.Cells(Me.ColApr.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                    rowdgv.Cells(Me.ColLevel.Index).Value, Me.ColApr.Index), Me.Def_Prec)
                    'hitung total mei
                    rowdgv.Cells(Me.ColMay.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                    rowdgv.Cells(Me.ColLevel.Index).Value, Me.ColMay.Index), Me.Def_Prec)
                    'hitung total jun
                    rowdgv.Cells(Me.ColJune.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                    rowdgv.Cells(Me.ColLevel.Index).Value, Me.ColJune.Index), Me.Def_Prec)
                    'hitung total jul
                    rowdgv.Cells(Me.ColJul.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                    rowdgv.Cells(Me.ColLevel.Index).Value, Me.ColJul.Index), Me.Def_Prec)
                    'hitung total agt
                    rowdgv.Cells(Me.ColAgt.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                    rowdgv.Cells(Me.ColLevel.Index).Value, Me.ColAgt.Index), Me.Def_Prec)
                    'hitung total sep
                    rowdgv.Cells(Me.ColSept.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                    rowdgv.Cells(Me.ColLevel.Index).Value, Me.ColSept.Index), Me.Def_Prec)
                    'hitung total okt
                    rowdgv.Cells(Me.ColOct.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                    rowdgv.Cells(Me.ColLevel.Index).Value, Me.ColOct.Index), Me.Def_Prec)
                    'hitung total nop
                    rowdgv.Cells(Me.ColNov.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                    rowdgv.Cells(Me.ColLevel.Index).Value, Me.ColNov.Index), Me.Def_Prec)
                    'hitung total des
                    rowdgv.Cells(Me.ColDec.Index).Value = Me.ObjCD.DigitGrouping(Me.GetSaldoHeader(rowdgv.Index, _
                    rowdgv.Cells(Me.ColLevel.Index).Value, Me.ColDec.Index), Me.Def_Prec)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvBudget.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvBudget.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Function GetSaldoHeader(ByVal vRowIndex As Double, ByVal vLevel As Integer, _
    ByVal vColIndex As Integer) As Decimal
        Dim TargetIndex, CurrentIndex As Integer
        Dim rowdgv As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            CurrentIndex = vRowIndex + 1
            While CurrentIndex <= Me.dgvBudget.Rows.Count - 1
                rowdgv = Me.dgvBudget.Rows(CurrentIndex)
                If rowdgv.Cells(Me.ColLevel.Index).Value <= vLevel Then
                    TargetIndex = rowdgv.Index - 1
                    Exit While
                ElseIf CurrentIndex = Me.dgvBudget.RowCount - 1 Then
                    TargetIndex = CurrentIndex
                    Exit While
                End If
                CurrentIndex += 1
            End While

            For i As Integer = vRowIndex + 1 To TargetIndex
                rowdgv = Me.dgvBudget.Rows(i)
                If rowdgv.Cells(Me.ColHeader.Index).Value = 1 And rowdgv.Cells(Me.ColLevel.Index).Value = vLevel + 1 Then
                    RetVal += Me.GetSaldoHeader(i, rowdgv.Cells(Me.ColLevel.Index).Value, vColIndex)
                ElseIf rowdgv.Cells(Me.ColLevel.Index).Value = vLevel + 1 Then
                    RetVal += CDec(rowdgv.Cells(vColIndex).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvBudget_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvBudget.CellBeginEdit
        Try
            'yang bisa diedit adalah kolom:
            '1. Jan - des
            '2. remarks
            If e.ColumnIndex = Me.ColJan.Index Or _
            e.ColumnIndex = Me.ColFeb.Index Or _
            e.ColumnIndex = Me.ColMar.Index Or _
            e.ColumnIndex = Me.ColApr.Index Or _
            e.ColumnIndex = Me.ColMay.Index Or _
            e.ColumnIndex = Me.ColJune.Index Or _
            e.ColumnIndex = Me.ColJul.Index Or _
            e.ColumnIndex = Me.ColAgt.Index Or _
            e.ColumnIndex = Me.ColSept.Index Or _
            e.ColumnIndex = Me.ColOct.Index Or _
            e.ColumnIndex = Me.ColNov.Index Or _
            e.ColumnIndex = Me.ColDec.Index Then
                If Me.dgvBudget.Item(Me.ColHeader.Index, e.RowIndex).Value = 0 Then
                    'hanya non header yang bisa diedit
                    Me.mBuffer = Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.ColRemarks.Index Then
                Me.mBuffer = Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvBudget_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvBudget.CellEndEdit
        Try
            'pastikan ada isinya
            If IsNothing(Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                'cek panjang maksimum untuk string dan nilai maksimum untuk angka
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    'untuk kolom remarks cek panjang maksimum
                    If Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 200 Then
                        Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 200)
                    End If
                Else
                    'untuk kolom budget cek data angka
                    If e.ColumnIndex = Me.ColJan.Index Or _
                    e.ColumnIndex = Me.ColFeb.Index Or _
                    e.ColumnIndex = Me.ColMar.Index Or _
                    e.ColumnIndex = Me.ColApr.Index Or _
                    e.ColumnIndex = Me.ColMay.Index Or _
                    e.ColumnIndex = Me.ColJune.Index Or _
                    e.ColumnIndex = Me.ColJul.Index Or _
                    e.ColumnIndex = Me.ColAgt.Index Or _
                    e.ColumnIndex = Me.ColSept.Index Or _
                    e.ColumnIndex = Me.ColOct.Index Or _
                    e.ColumnIndex = Me.ColNov.Index Or _
                    e.ColumnIndex = Me.ColDec.Index Then
                        'cek bilangan non negatif
                        If Me.ObjCD.CekNonNegatif(Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Data must be a non negative number.", MsgBoxStyle.Information, "Information")
                            Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value = Me.mBuffer
                            Exit Sub
                        Else
                            If Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value > 999999999999999 Then
                                MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(999999999999999), _
                                MsgBoxStyle.Information, "Information")
                                Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value = Me.mBuffer
                                Exit Sub
                            End If
                        End If
                    End If

                    'sampe sini data angka ok, format
                    Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvBudget.Item(e.ColumnIndex, _
                    e.RowIndex).Value, Me.Def_Prec)
                    'update data header
                    Me.UpdateSaldoHeader()
                End If
                'set simpan perubahan data
                Me.DataChanged = True
            Else
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value = ""
                Else
                    Me.dgvBudget.Item(e.ColumnIndex, e.RowIndex).Value = Me.mBuffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            Me.ObjBudget.FinancialYear = Me.dtpPeriod.Value.Date
            Me.ObjBudget.TipeBudget = Me.cboFilter.SelectedIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjBudget
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            'simpan perubahan data jika ada
            If MsgBox("Save data change?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.SalinKeObjek()
                If Me.ObjBudget.GetPrimary(Me.dtpPeriod.Value.Date, Me.cboFilter.SelectedIndex) = -1 Then
                    'simpan baru
                    If Me.ObjBudget.SimpanBaru(Me.dgvBudget, Me, Me.UserCode) = True Then
                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Me.DataChanged = False
                    End If
                Else
                    'simpan perubahan
                    If Me.ObjBudget.SimpanPerubahan(Me.dgvBudget, Me, Me.UserCode) = True Then
                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Me.DataChanged = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpPeriod_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriod.ValueChanged
        Try
            If Me.cboFormat.SelectedIndex = 0 Then
                If Me.AutoLoad = True Then
                    'simpan perubahan data jika ada
                    If Me.DataChanged = True Then
                        If MsgBox("Save data change?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Me.ObjBudget.FinancialYear = Me.PreviousPeriod
                            If Me.ObjBudget.GetPrimary(Me.PreviousPeriod, Me.cboFilter.SelectedIndex) = -1 Then
                                'simpan baru
                                If Me.ObjBudget.SimpanBaru(Me.dgvBudget, Me, Me.UserCode) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                End If
                            Else
                                'simpan perubahan
                                If Me.ObjBudget.SimpanPerubahan(Me.dgvBudget, Me, Me.UserCode) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                End If
                            End If
                        End If
                    End If

                    'load data sesuai periode
                    Me.ObjBudget.CariDetail(Me.dgvBudget, Me, Me.ObjBudget.GetPrimary(Me.dtpPeriod.Value.Date, Me.cboFilter.SelectedIndex), Me.cboFilter.SelectedIndex)
                    If Me.ObjBudget.GetPrimary(Me.dtpPeriod.Value, Me.cboFilter.SelectedIndex) = -1 Then
                        Me.tslCreatedBy.Text = "Created By: "
                        Me.tslCreatedDate.Text = "Created Date: "
                        Me.tslModifiedBy.Text = "Modified By: "
                        Me.tslModifiedDate.Text = "Modified Date: "
                    Else
                        Me.SalinKeField()
                    End If
                    Me.DataChanged = False

                    'update data previous period
                    Me.PreviousPeriod = Me.dtpPeriod.Value.Date
                End If
            Else
                'isikan data yearly yang baru
                Me.ObjBudget.IsiDGV_Tahunan(Me.dgvYearly, Me, Me.dtpPeriod.Value.Date, Me.cboFilter.SelectedIndex)
            End If

            Me.dgvBudget.ClearSelection()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim PrimaryAkun As Double = Me.ObjBudget.GetPrimary(Me.dtpPeriod.Value.Date, Me.cboFilter.SelectedIndex)
        Try

            If PrimaryAkun <> -1 Then
                'cek apakah boleh hapus
                If Me.ObjBudget.ApakahBolehHapus(PrimaryAkun, Me.UserCode, True, True) = False Then
                    Exit Sub
                End If
            End If

            If MsgBox("Clear data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                If PrimaryAkun <> -1 Then
                    Me.ObjBudget.Hapus(PrimaryAkun, Me.UserCode, True, True)
                    Me.ObjBudget.CariDetail(Me.dgvBudget, Me, Me.ObjBudget.GetPrimary(Me.dtpPeriod.Value.Date, Me.cboFilter.SelectedIndex), Me.cboFilter.SelectedIndex)
                Else
                    Me.ObjBudget.CariDetail(Me.dgvBudget, Me, -1, Me.cboFilter.SelectedIndex)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvYearly_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvYearly.CellBeginEdit
        Try
            'tidak boleh edit semua
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvBudget_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvBudget.CellEnter

    End Sub

    Private Sub dgvBudget_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvBudget.SelectionChanged
        If Me.cboFormat.SelectedIndex = 0 Then
            If Me.dgvBudget.SelectedCells.Count > 0 Then
                If Me.dgvBudget.Rows(Me.dgvBudget.SelectedCells(0).RowIndex).Cells(Me.ColHeader.Index).Value = 0 Then
                    Me.btnShowDetail.Visible = True
                Else
                    Me.btnShowDetail.Visible = False
                End If
            Else
                Me.btnShowDetail.Visible = False
            End If
        Else
            Me.btnShowDetail.Visible = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub cboFilter_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFilter.SelectedIndexChanged
        Try
            If Me.cboFormat.SelectedIndex <> -1 Then
                If Me.AutoLoad = True Then
                    If Me.cboFormat.SelectedIndex = 0 Then
                        Me.ObjBudget.CariDetail(Me.dgvBudget, Me, Me.ObjBudget.GetPrimary(Me.dtpPeriod.Value.Date, Me.cboFilter.SelectedIndex), Me.cboFilter.SelectedIndex)
                        Me.dgvBudget.BringToFront()
                    Else
                        'simpan perubahan data jika perlu
                        If Me.DataChanged = True Then
                            Me.btnSave.PerformClick()
                        End If
                        Me.DataChanged = False
                        Me.ObjBudget.IsiDGV_Tahunan(Me.dgvYearly, Me, Me.dtpPeriod.Value.Date, Me.cboFilter.SelectedIndex)
                        Me.dgvYearly.BringToFront()
                    End If
                End If

                If Me.cboFormat.SelectedIndex = 0 Then
                    If Me.dgvBudget.SelectedCells.Count > 0 Then
                        If Me.dgvBudget.Rows(Me.dgvBudget.SelectedCells(0).RowIndex).Cells(Me.ColHeader.Index).Value = 0 Then
                            Me.btnShowDetail.Visible = True
                        Else
                            Me.btnShowDetail.Visible = False
                        End If
                    Else
                        Me.btnShowDetail.Visible = False
                    End If
                Else
                    Me.btnShowDetail.Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class