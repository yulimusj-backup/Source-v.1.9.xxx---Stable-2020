Option Explicit Off
Option Strict Off
Public Class FrmOptionRate
    Public ListPrimaryCurr As New List(Of Double)
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim DgvRow As DataGridViewRow
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim Buffer As String = ""
    Dim TblCurr As New DseLaporanPembelian.tblmaster_currencyDataTable

    Private Sub FrmOptionRate_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

    End Sub

    Private Sub FrmOptionRate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim RowCurr As DseLaporanPembelian.tblmaster_currencyRow
        Try
            Me.ControlBox = False
            'Me.DgvRate.AllowUserToAddRows = True
            Me.DgvRate.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.Text = "Option Rate"
            Me.IsikanTblMaster()
            'Hapus jika tidak ada dlm list
            For i As Integer = Me.TblCurr.Rows.Count - 1 To 0 Step -1
                RowCurr = TblCurr.Rows(i)
                If ListPrimaryCurr.Contains(RowCurr.Primary) = False Or RowCurr.Primary = 1 Then
                    Me.TblCurr.Rows.RemoveAt(i)
                End If
            Next
            Me.LoadData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsikanTblMaster() As Boolean
        Dim DTA As New DseLaporanPembelianTableAdapters.tblmaster_currencyTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(Me.TblCurr)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function LoadData() As Boolean
        Dim DTA As New DseLaporanPembelianTableAdapters.tblcurrTableAdapter
        Dim TBL As New DseLaporanPembelian.tblcurrDataTable
        Dim row As DseLaporanPembelian.tblcurrRow
        Dim vPrimaryTransaksi As Double = GetPrimary()
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryTransaksi)
            For Each RowCurr In Me.TblCurr.Rows
                TBL.DefaultView.RowFilter = "[primary currency]='" & RowCurr.Primary & "'"
                If TBL.DefaultView.Count > 0 Then
                    row = TBL.DefaultView(0).Row
                    DgvRow = Me.DgvRate.Rows(Me.DgvRate.Rows.Add)
                    With DgvRow
                        .Cells(Me.ColCurrency.Index).Value = row.Currency_Code
                        .Cells(Me.ColRate.Index).Value = Me.ObjCD.DigitGrouping(row.rate, row.Default_Precision, ClsCekData.EnumTipeRounding.Middlepoint)
                        .Cells(Me.ColPrimaryCurr.Index).Value = row.primary_currency
                    End With
                Else
                    DgvRow = Me.DgvRate.Rows(Me.DgvRate.Rows.Add)
                    With DgvRow
                        .Cells(Me.ColCurrency.Index).Value = RowCurr.Currency_Code
                        .Cells(Me.ColRate.Index).Value = Me.ObjCD.DigitGrouping(1, RowCurr.Default_Precision)
                        .Cells(Me.ColPrimaryCurr.Index).Value = RowCurr.Primary
                    End With
                End If
            Next
          
            Me.ObjCD.SetNumber(Me.DgvRate)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimary() As Double
        Dim DTA As New DseLaporanPembelianTableAdapters.tblcurrTableAdapter
        Dim Retval As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimary
            If Retval.HasValue = True Then
                If Retval.Value > -1 Then
                    Return Retval.Value
                Else
                    Return -1
                End If
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub DgvRate_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvRate.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColRate.Index Then
                Me.Buffer = Me.DgvRate.Item(Me.ColRate.Index, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvRate_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvRate.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColRate.Index Then
                If String.IsNullOrEmpty(Me.DgvRate.Item(Me.ColRate.Index, e.RowIndex).Value) = False Then
                    If ObjCD.CekBilanganPositif(Me.DgvRate.Item(Me.ColRate.Index, e.RowIndex).Value) = True Then
                        Me.DgvRate.Item(Me.ColRate.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DgvRate.Item(Me.ColRate.Index, e.RowIndex).Value, ObjCur.GetDefaultPrecission(Me.DgvRate.Item(Me.ColPrimaryCurr.Index, e.RowIndex).Value))
                    Else
                        'salah
                        MsgBox("Please enter positive numnber for data.", MsgBoxStyle.Information, "Information")
                        Me.DgvRate.Item(Me.ColRate.Index, e.RowIndex).Value = Me.Buffer
                    End If
                Else
                    'salah
                    'MsgBox("Please enter positive numnber for data.", MsgBoxStyle.Information, "Information")
                    Me.DgvRate.Item(Me.ColRate.Index, e.RowIndex).Value = Me.Buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
End Class