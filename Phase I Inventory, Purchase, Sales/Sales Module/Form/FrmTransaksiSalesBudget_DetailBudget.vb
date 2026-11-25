Option Explicit On
Option Strict Off

Public Class FrmTransaksiSalesBudget_DetailBudget
    Public RefForm As FrmTransaksiSalesBudget
    Public DefPrec As Integer = 0

    Dim DgvRef As DataGridView
    Dim ColTotal As DataGridViewColumn
    Dim ColTotalPerGroup As DataGridViewColumn
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmTransaksiSalesBudget_DetailBudget_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.DgvRef = Me.RefForm.dgvItem
            Me.ComboBox1.SelectedIndex = 0
            Me.IsikanDGV()
            Me.dgvTotal.AlternatingRowsDefaultCellStyle.BackColor = Me.RefForm.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDGV()
        Try
            If Me.ComboBox1.SelectedIndex = 0 Then
                Me.IsikanDGVGrouping()
            Else
                Me.IsikanDGVNoDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTotalPerBulan(ByVal vIndex As Integer) As Decimal
        Dim RetVal As Decimal = 0
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DgvRef.Rows
                RetVal += CDec(rowDGV.Cells(vIndex).Value)
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTotalPerBulan(ByVal vIndex As Integer, ByVal vGroup As String) As Decimal
        Dim RetVal As Decimal = 0
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DgvRef.Rows
                If rowDGV.Cells(Me.RefForm.ColSalesPGroup.Index).Value = vGroup Then
                    RetVal += CDec(rowDGV.Cells(vIndex).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsikanDGVGrouping()
        Dim rowDGV As DataGridViewRow
        Dim DummyCol As DataGridViewColumn

        Try
            Me.TxtTotal.Text = Me.RefForm.TxtTotal.Text

            'buang kolom yang tidak perlu
            While Me.dgvTotal.ColumnCount > 2
                DummyCol = Me.dgvTotal.Columns(Me.dgvTotal.Columns.Count - 1)
                Me.dgvTotal.Columns.RemoveAt(Me.dgvTotal.Columns.Count - 1)
                DummyCol.Dispose()
            End While

            'tambahkan kolom total
            Me.dgvTotal.Columns.Add("1", "1")
            Me.ColTotal = Me.dgvTotal.Columns(Me.dgvTotal.Columns.Count - 1)
            Me.ColTotal.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColTotal.HeaderText = "Total"
            Me.ColTotal.Name = "Total"

            Me.dgvTotal.Rows.Clear()
            'jan
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Jan"
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalPerBulan(Me.RefForm.ColJan.Index), Me.DefPrec)
            'feb
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Feb"
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalPerBulan(Me.RefForm.ColFeb.Index), Me.DefPrec)
            'mar
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Mar"
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalPerBulan(Me.RefForm.ColMar.Index), Me.DefPrec)
            'apr
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Apr"
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalPerBulan(Me.RefForm.ColApr.Index), Me.DefPrec)
            'may
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "May"
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalPerBulan(Me.RefForm.ColMay.Index), Me.DefPrec)
            'jun
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Jun"
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalPerBulan(Me.RefForm.ColJun.Index), Me.DefPrec)
            'jul
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Jul"
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalPerBulan(Me.RefForm.ColJul.Index), Me.DefPrec)
            'agt
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Aug"
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalPerBulan(Me.RefForm.ColAug.Index), Me.DefPrec)
            'sep
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Sep"
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalPerBulan(Me.RefForm.ColSep.Index), Me.DefPrec)
            'okt
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Oct"
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalPerBulan(Me.RefForm.ColOct.Index), Me.DefPrec)
            'nov
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Nov"
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalPerBulan(Me.RefForm.ColNov.Index), Me.DefPrec)
            'dec
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Dec"
            rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalPerBulan(Me.RefForm.ColDec.Index), Me.DefPrec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDGVNoDetail()
        Dim rowDGV As DataGridViewRow
        Dim DummyCol As DataGridViewColumn
        Dim ListGroup As New List(Of String)
        Dim NamaGroup As String = ""
        Dim DummyVal As Decimal = 0

        Try
            Me.TxtTotal.Text = Me.RefForm.TxtTotal.Text

            'buang kolom yang tidak perlu
            While Me.dgvTotal.ColumnCount > 2
                DummyCol = Me.dgvTotal.Columns(Me.dgvTotal.Columns.Count - 1)
                Me.dgvTotal.Columns.RemoveAt(Me.dgvTotal.Columns.Count - 1)
                DummyCol.Dispose()
            End While

            Me.dgvTotal.Rows.Clear()
            'jan
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Jan"
            'feb
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Feb"
            'mar
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Mar"
            'apr
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Apr"
            'may
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "May"
            'jun
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Jun"
            'jul
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Jul"
            'agt
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Aug"
            'sep
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Sep"
            'okt
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Oct"
            'nov
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Nov"
            'dec
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvTotal.Rows.Count
            rowDGV.Cells(Me.ColMonth.Index).Value = "Dec"

            'tambahkan per group
            For Each rowDGV In Me.DgvRef.Rows
                If rowDGV.Cells(Me.RefForm.ColPrimarySalesPerson.Index).Value <> -1 Then
                    If ListGroup.Contains(rowDGV.Cells(Me.RefForm.ColSalesPGroup.Index).Value) = False Then
                        ListGroup.Add(rowDGV.Cells(Me.RefForm.ColSalesPGroup.Index).Value)
                    End If
                End If
            Next
            For Each NamaGroup In ListGroup
                Me.IsikanPerGroup(NamaGroup)
            Next

            'masukkan data total
            '2. total per group
            Me.dgvTotal.Rows.Add()
            rowDGV = Me.dgvTotal.Rows(Me.dgvTotal.Rows.Count - 1)
            rowDGV.Cells(Me.ColMonth.Index).Value = "Group Total"
            For i As Integer = 2 To Me.dgvTotal.Columns.Count - 1
                DummyVal = 0
                For Each rowDGV In Me.dgvTotal.Rows
                    If rowDGV.Index <> Me.dgvTotal.Rows.Count - 1 Then
                        DummyVal += CDec(rowDGV.Cells(i).Value)
                    End If
                Next
                rowDGV.Cells(i).Value = Me.ObjCD.DigitGrouping(DummyVal, Me.DefPrec)
            Next

            '1. total per bulan
            Me.dgvTotal.Columns.Add("1", "1")
            Me.ColTotal = Me.dgvTotal.Columns(Me.dgvTotal.Columns.Count - 1)
            Me.ColTotal.HeaderText = "Total Monthly"
            Me.ColTotal.Name = "ColTotalMonthly"
            For Each rowDGV In Me.dgvTotal.Rows
                DummyVal = 0
                For i As Integer = 2 To Me.dgvTotal.Columns.Count - 2
                    DummyVal += CDec(rowDGV.Cells(i).Value)
                Next
                rowDGV.Cells(Me.ColTotal.Index).Value = Me.ObjCD.DigitGrouping(DummyVal, Me.DefPrec)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListGroup = Nothing
        End Try
    End Sub

    Private Sub IsikanPerGroup(ByVal vGroupName As String)
        Dim rowDgv As DataGridViewRow

        Try
            Me.dgvTotal.Columns.Add("1", "1")
            Me.ColTotal = Me.dgvTotal.Columns(Me.dgvTotal.Columns.Count - 1)
            Me.ColTotal.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColTotal.HeaderText = "Total " + vGroupName
            Me.ColTotal.Name = "ColTotal" + vGroupName.Replace(" ", "")

            'jan
            rowDgv = Me.dgvTotal.Rows(0)
            rowDgv.Cells(Me.ColTotal.Index).Value = Me.GetTotalPerBulan(Me.RefForm.ColJan.Index, vGroupName)
            'feb
            rowDgv = Me.dgvTotal.Rows(1)
            rowDgv.Cells(Me.ColTotal.Index).Value = Me.GetTotalPerBulan(Me.RefForm.ColFeb.Index, vGroupName)
            'mar
            rowDgv = Me.dgvTotal.Rows(2)
            rowDgv.Cells(Me.ColTotal.Index).Value = Me.GetTotalPerBulan(Me.RefForm.ColMar.Index, vGroupName)
            'apr
            rowDgv = Me.dgvTotal.Rows(3)
            rowDgv.Cells(Me.ColTotal.Index).Value = Me.GetTotalPerBulan(Me.RefForm.ColApr.Index, vGroupName)
            'mei
            rowDgv = Me.dgvTotal.Rows(4)
            rowDgv.Cells(Me.ColTotal.Index).Value = Me.GetTotalPerBulan(Me.RefForm.ColMay.Index, vGroupName)
            'jun
            rowDgv = Me.dgvTotal.Rows(5)
            rowDgv.Cells(Me.ColTotal.Index).Value = Me.GetTotalPerBulan(Me.RefForm.ColJun.Index, vGroupName)
            'jul
            rowDgv = Me.dgvTotal.Rows(6)
            rowDgv.Cells(Me.ColTotal.Index).Value = Me.GetTotalPerBulan(Me.RefForm.ColJul.Index, vGroupName)
            'aug
            rowDgv = Me.dgvTotal.Rows(7)
            rowDgv.Cells(Me.ColTotal.Index).Value = Me.GetTotalPerBulan(Me.RefForm.ColAug.Index, vGroupName)
            'sep
            rowDgv = Me.dgvTotal.Rows(8)
            rowDgv.Cells(Me.ColTotal.Index).Value = Me.GetTotalPerBulan(Me.RefForm.ColSep.Index, vGroupName)
            'okt
            rowDgv = Me.dgvTotal.Rows(9)
            rowDgv.Cells(Me.ColTotal.Index).Value = Me.GetTotalPerBulan(Me.RefForm.ColOct.Index, vGroupName)
            'nov
            rowDgv = Me.dgvTotal.Rows(10)
            rowDgv.Cells(Me.ColTotal.Index).Value = Me.GetTotalPerBulan(Me.RefForm.ColNov.Index, vGroupName)
            'dec
            rowDgv = Me.dgvTotal.Rows(11)
            rowDgv.Cells(Me.ColTotal.Index).Value = Me.GetTotalPerBulan(Me.RefForm.ColDec.Index, vGroupName)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Me.IsikanDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class