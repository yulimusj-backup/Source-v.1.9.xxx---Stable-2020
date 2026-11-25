Option Explicit On
Option Strict Off

Public Class FrmMasterDefaultValues
    Dim Berubah As Boolean = False
    Dim ObjCD As New ClsCekData
    Dim Buffer As String = ""

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.Berubah = True Then
                If MsgBox("Save changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.SimpanData()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function SimpanData() As Boolean
        Try
            'simpan data currency
            Me.SimpanDataCurrency()

            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
            Me.Berubah = False

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub LoadData()
        Try
            Me.LoadDataDefaultCurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataDefaultCurrency()
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter
        Dim TBL As New dseMasterCurrency.tblmaster_currencyDataTable
        Dim row As dseMasterCurrency.tblmaster_currencyRow
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvCurrency.Rows.Clear()
            DTA.Fill(TBL)
            For Each row In TBL.Rows
                If row.Primary <> Core_Module.PrimaryBaseCurrency Then
                    Me.dgvCurrency.Rows.Add()
                    rowDGV = Me.dgvCurrency.Rows(Me.dgvCurrency.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColCurrency.Index).Value = row.Currency_Code
                        .Cells(Me.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(row.Default_Exchange_Rate, DefPrecExchangeRate)
                        .Cells(Me.ColPrimaryCurrency.Index).Value = row.Primary
                    End With
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If Me.Berubah = True Then
                If MsgBox("Save changes before closing?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    If Me.SimpanData() = True Then
                        Me.Close()
                    End If
                Else
                    Me.Close()
                End If
            Else
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterDefaultValues_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.TabControl1.TabPages.Remove(Me.TabPage2)
            Me.LoadData()
            Me.dgvCurrency.ReadOnly = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataCurrency()
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each rowDGV In Me.dgvCurrency.Rows
                DTA.UpdateQuery_SetDefaultExchangeRate(rowDGV.Cells(Me.ColExchangeRate.Index).Value, _
                rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub dgvCurrency_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCurrency.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColExchangeRate.Index Then
                'simpan ke buffer
                Me.Buffer = Me.dgvCurrency.Item(e.ColumnIndex, e.RowIndex).Value
                'buang koma
                Me.dgvCurrency.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvCurrency.Item(e.ColumnIndex, e.RowIndex).Value)
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCurrency_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCurrency.CellEndEdit
        Try
            If IsNothing(Me.dgvCurrency.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvCurrency.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                Me.dgvCurrency.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            Else
                'harus positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvCurrency.Item(e.ColumnIndex, _
                                                  e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvCurrency.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
                'jika nilannya sama dengan buffer maka tidak ada perubahn
                If CDec(Me.Buffer) <> CDec(Me.dgvCurrency.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    Me.Berubah = True
                End If
                'digit groping
                Me.dgvCurrency.Item(e.ColumnIndex, e.RowIndex).Value = _
                ObjCD.DigitGrouping(Me.dgvCurrency.Item(e.ColumnIndex, e.RowIndex).Value, DefPrecExchangeRate)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class