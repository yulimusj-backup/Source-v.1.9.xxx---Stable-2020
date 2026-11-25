Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonthlyClosing
    Public UserCode As String = ""
    Public BAckColorAll As Color = BackColor
    Dim DefaultPrecision As Integer = 0
    Dim Buffer As String = ""
    Dim baru As Boolean = True
    Dim ObjClosing As New ClsTransaksiMonthlyClosing
    Dim ObjCLosingItem As New ClsTransaksiMonthlyClosingItem
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim objCurrency As New Core_Module.clsMasterCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim TBL As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable
    Dim TBLUnitPrice As New DseTransaksiMonthlyClosing.DataTable9DataTable
    Dim TBLItemBefore As New DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable
    Dim TBLItemAfter As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable
    Dim LoadClosing As Boolean = False
    Dim DTA_CL As New DseTransaksiMonthlyClosingTableAdapters.DataTable25TableAdapter
    Dim TBL_CL As New DseTransaksiMonthlyClosing.DataTable25DataTable
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim TBLJasa As New Core_Module.dseMasterProduct.tblmaster_jasaDataTable
    Dim DTAJasa As New Core_Module.dseMasterProductTableAdapters.tblmaster_jasaTableAdapter

    Dim TBLBeginning As New DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item1DataTable
    Dim DTABeginning As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing_item1TableAdapter

    Dim ClosingPeriodeBuffer As Date

    Delegate Sub AddNewRow(ByVal BeginQty As Decimal, ByVal BeginAmount As Decimal, _
    ByVal vPrimaryProduct As Double, ByVal vCodeNo As String, ByVal Description As String, ByVal vPrimaryUnit As Double, _
    ByVal UnitCode As String, ByVal OutQty As Decimal, ByVal EndQty As Decimal, ByVal InQty As Decimal, ByVal InAmount As Decimal)
    Delegate Sub AddInAmount(ByVal vPrimaryCodeNo As Double, ByVal vAdditionalAmount As Decimal)
    Delegate Sub SetOutAmount(ByVal rowDGV As DataGridViewRow, ByVal OutAmount As Decimal)
    Delegate Sub SetEndAmount(ByVal rowDGV As DataGridViewRow, ByVal EndAmount As Decimal)

    Public ObjCD As New Core_Module.ClsCekData

    Dim DTAClosing1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable21TableAdapter
    Dim TBLClosing1 As New DseTransaksiMonthlyClosing.DataTable21DataTable
    Dim Dum_ClosingPeriode As Date = #1/2/1950#

    Dim vTBL1 As New Core_Module.dseMasterProduct.tblmaster_product1DataTable
    Dim vDTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_product1TableAdapter

    Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.DataTable12TableAdapter
    Dim TBLBundling As New DseTransaksiMonthlyClosing.DataTable12DataTable

    Dim ListPrimaryCodeNo_Bundling As New List(Of Double)

    Private Function GetAngkaDecimal(ByVal vPrimaryProduct As Double) As Integer
        Dim row As Core_Module.dseMasterProduct.tblmaster_product1Row
        Dim filter As String = ""

        Try
            filter = "primary =" + vPrimaryProduct.ToString
            Me.vTBL1.DefaultView.RowFilter = filter
            If Me.vTBL1.DefaultView.Count > 0 Then
                row = TryCast(Me.vTBL1.DefaultView(0).Row, Core_Module.dseMasterProduct.tblmaster_product1Row)
                Return row._Decimal
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.BringToFront()

                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjClosing.SimpanBaru(Me.dgvItem, Me.dgvItemAdjust, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjClosing.SimpanPerubahan(Me.dgvItem, Me.dgvItemAdjust, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
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

    Private Sub FrmTransaksiPO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.ColBeginAmount.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColBeginQty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColInAmount.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColInQty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColOutAmount.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColOutQty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColEndAmount.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.ColEndQty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            Me.DefaultPrecision = Me.objCurrency.GetDecimal_DefaultCurrency
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        TxtRemark.BackColor = Me.BAckColorAll
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        TxtRemark.BackColor = Me.TxtPrimaryClosing.BackColor
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            'msgbox("belum")
            If Me.TxtPrimaryClosing.Text <> "" Then
                If CDbl(Me.TxtPrimaryClosing.Text) >= 0 Then
                    ObjClosing.Cetak()
                End If
            Else
                Me.btnFind.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.Enter
        Try
            Me.Buffer = Me.dtpDate.Value.Date.ToShortDateString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim AngkaDecimal As Integer = 0

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                AngkaDecimal = Me.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value)
                'digit grouping: 
                row.Cells(Me.ColBeginQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColBeginQty.Index).Value, _
                AngkaDecimal)
                row.Cells(Me.ColBeginAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColBeginAmount.Index).Value, Me.DefaultPrecision)

                row.Cells(Me.ColInQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColInQty.Index).Value, _
                AngkaDecimal)
                row.Cells(Me.ColInAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColInAmount.Index).Value, Me.DefaultPrecision)

                row.Cells(Me.ColOutQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColOutQty.Index).Value, _
                AngkaDecimal)
                row.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColOutAmount.Index).Value, Me.DefaultPrecision)

                row.Cells(Me.ColEndQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColEndQty.Index).Value, _
                AngkaDecimal)
                row.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColEndAmount.Index).Value, Me.DefaultPrecision)
            Next

            'untuk in adjustment dari awal prosesnya sudah ok.
            'For Each row In Me.dgvItemAdjust.Rows
            'row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQty.Index).Value)
            'row.Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPriceAdjust.Index).Value, Me.DefaultPrecision))
            'Next

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'Jika sudah ada transaksi closing sebelumnya maka tidak bisa diedit
            If Me.ObjClosing.GetJml > 1 Or (Me.ObjClosing.GetJml = 1 And Me.baru = True) Then
                'If Me.UserCode = "USER" Then
                'kalo pake kode "USER" maka bisa edit semuanya
                'If e.ColumnIndex = Me.ColEndQty.Index Or e.ColumnIndex = Me.ColEndAmount.Index Or _
                ' e.ColumnIndex = Me.ColRemarks.Index Or e.ColumnIndex = Me.ColInQty.Index Or _
                ' e.ColumnIndex = Me.ColInAmount.Index Or e.ColumnIndex = Me.ColOutQty.Index Or _
                'e.ColumnIndex = Me.ColOutAmount.Index Then
                'kalo sampe sini maka boleh edit
                'Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                'Exit Sub
                'Else
                'e.Cancel = True
                'Exit Sub
                'End If
                'Else
                'yang bisa ediedit adalah remarks
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    'kalo sampe sini maka boleh edit
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
                'End If
            ElseIf Me.ObjClosing.GetJml = 1 And Me.baru = False And e.ColumnIndex = Me.ColEndQty.Index Then
                e.Cancel = True
                Exit Sub
            ElseIf Me.ObjClosing.GetJml = 0 And e.ColumnIndex = Me.ColEndQty.Index Then
                e.Cancel = True
                Exit Sub
            End If

            'yang bisa diedit adalah:
            '1. End Qty
            '2. End Amount
            '3. Remarks
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColEndQty.Index Or e.ColumnIndex = Me.ColEndAmount.Index Or _
                e.ColumnIndex = Me.ColRemarks.Index Then
                    'kalo sampe sini maka boleh edit
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                'hanya bisa edit pada saat detail aktif
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Try
            If e.ColumnIndex = Me.ColDetailCalculation.Index And e.RowIndex <> -1 Then
                Dim i As New FrmTransaksiMonthlyClosing_DetailCalc
                i.i = Me
                i.PrimaryCodeNo = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value
                i.TanggalDari = Me.dtpPeriod.Value.Date
                i.TanggalSampai = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date)

                i.ShowDialog()
                i.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex = Me.ColEndQty.Index Or e.ColumnIndex = Me.ColEndAmount.Index Or _
                    e.ColumnIndex = Me.ColInQty.Index Or _
                    e.ColumnIndex = Me.ColInAmount.Index Or e.ColumnIndex = Me.ColOutQty.Index Or _
                    e.ColumnIndex = Me.ColOutAmount.Index Then
                        If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If e.ColumnIndex = Me.ColOutAmount.Index Then
                                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "-" Then
                                    Me.dgvItem.Item(Me.ColOutAmount.Index, e.RowIndex).Value = _
                                    Me.ObjCD.DigitGrouping(CDbl(Me.dgvItem.Item(Me.ColBeginAmount.Index, e.RowIndex).Value) + _
                                    CDbl(Me.dgvItem.Item(Me.ColInAmount.Index, e.RowIndex).Value) - _
                                    CDbl(Me.dgvItem.Item(Me.ColEndAmount.Index, e.RowIndex).Value))
                                Else
                                    If IsNumeric(Me.Buffer) = True Then
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = 0
                                    End If
                                End If
                            Else
                                If IsNumeric(Me.Buffer) = True Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Else
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = 0
                                End If
                            End If
                        Else
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar2() As Boolean
        Dim row As DataGridViewRow
        Dim x As Double
        Dim CodeNo As String = ""
        Dim CodeNoBegin As String = ""
        Dim CodeNoOut As String = ""
        Dim CodeNoEnd As String = ""
        Dim CodeNo_QtyBegin As String = ""
        Dim CodeNo_QtyIn As String = ""
        Dim CodeNo_QtyOut As String = ""
        Dim CodeNo_QtyEnd As String = ""
        Dim CodeUnbalanceAmount As String = ""
        Dim CodeUnbalanceQty As String = ""
        Dim WrongEndAmountCalculation As String = ""
        Dim WrongOutAmountCalculation As String = ""
        Dim BeginAmount As String = ""
        Dim InAmount As String = ""
        Dim OutAmount As String = ""
        Dim EndAmount As String = ""

        Dim Pass As Boolean = True
        Dim Dummy_Adjust As String = ""

        Try
            If Me.ObjClosing.GetJml = 0 Then
                Return Me.CekEndAmountSaja()
            ElseIf Me.baru = False And Me.ObjClosing.GetJml = 1 Then
                Return Me.CekEndAmountSaja()
            End If

            'jika bukan closing pertama
            For Each row In Me.dgvItem.Rows
                If Math.Abs((CDec(row.Cells(Me.ColBeginAmount.Index).Value) + CDec(row.Cells(Me.ColInAmount.Index).Value)) - _
                (CDec(row.Cells(Me.ColOutAmount.Index).Value) + CDec(row.Cells(Me.ColEndAmount.Index).Value))) > 2 Then
                    If CodeUnbalanceAmount = "" Then
                        CodeUnbalanceAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeUnbalanceAmount = CodeUnbalanceAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColBeginQty.Index).Value) + CDec(row.Cells(Me.ColInQty.Index).Value) <> _
                CDec(row.Cells(Me.ColOutQty.Index).Value) + CDec(row.Cells(Me.ColEndQty.Index).Value) Then
                    If CodeUnbalanceQty = "" Then
                        CodeUnbalanceQty = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeUnbalanceQty = CodeUnbalanceQty + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If (CDec(row.Cells(Me.ColBeginQty.Index).Value) + CDec(row.Cells(Me.ColInQty.Index).Value)) <> 0 Then
                    x = Math.Round((CDec(row.Cells(Me.ColBeginAmount.Index).Value) + CDec(row.Cells(Me.ColInAmount.Index).Value)) * CDec(row.Cells(Me.ColEndQty.Index).Value) / _
                    (CDec(row.Cells(Me.ColBeginQty.Index).Value) + CDec(row.Cells(Me.ColInQty.Index).Value)), Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                    If Math.Abs(x - CDec(row.Cells(Me.ColEndAmount.Index).Value)) > 2 Then
                        If WrongEndAmountCalculation = "" Then
                            WrongEndAmountCalculation = row.Cells(Me.ColCodeNo.Index).Value.ToString
                        Else
                            WrongEndAmountCalculation = WrongEndAmountCalculation + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                        End If
                    End If
                End If

                'If (CDec(row.Cells(Me.ColBeginQty.Index).Value) + CDec(row.Cells(Me.ColInQty.Index).Value)) <> 0 Then
                'x = Math.Round((CDec(row.Cells(Me.ColBeginAmount.Index).Value) + CDec(row.Cells(Me.ColInAmount.Index).Value)) * CDec(row.Cells(Me.ColOutQty.Index).Value) / _
                '(CDec(row.Cells(Me.ColBeginQty.Index).Value) + CDec(row.Cells(Me.ColInQty.Index).Value)), Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                'If Math.Abs(x - CDec(row.Cells(Me.ColOutAmount.Index).Value)) > 2 Then
                'If WrongOutAmountCalculation = "" Then
                'WrongOutAmountCalculation = row.Cells(Me.ColCodeNo.Index).Value.ToString
                'Else
                'WrongOutAmountCalculation = WrongOutAmountCalculation + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                'End If
                'End If
                'End If

                If row.Cells(Me.ColInQty.Index).Value > 0 And row.Cells(Me.ColInAmount.Index).Value = 0 Then
                    If CodeNo = "" Then
                        CodeNo = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo = CodeNo + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColBeginQty.Index).Value > 0 And row.Cells(Me.ColBeginAmount.Index).Value = 0 Then
                    If CodeNoBegin = "" Then
                        CodeNoBegin = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoBegin = CodeNoBegin + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColOutQty.Index).Value > 0 And row.Cells(Me.ColOutAmount.Index).Value = 0 Then
                    If CodeNoOut = "" Then
                        CodeNoOut = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoOut = CodeNoOut + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColEndQty.Index).Value > 0 And row.Cells(Me.ColEndAmount.Index).Value = 0 Then
                    If CodeNoEnd = "" Then
                        CodeNoEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoEnd = CodeNoEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColBeginQty.Index).Value) < 0 Then
                    If CodeNo_QtyBegin = "" Then
                        CodeNo_QtyBegin = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyBegin = CodeNo_QtyBegin + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColInQty.Index).Value) < 0 Then
                    If CodeNo_QtyIn = "" Then
                        CodeNo_QtyIn = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyIn = CodeNo_QtyIn + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColOutQty.Index).Value) < 0 Then
                    If CodeNo_QtyOut = "" Then
                        CodeNo_QtyOut = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyOut = CodeNo_QtyOut + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColEndQty.Index).Value) < 0 Then
                    If CodeNo_QtyEnd = "" Then
                        CodeNo_QtyEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyEnd = CodeNo_QtyEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColBeginAmount.Index).Value) > 0 And CDec(row.Cells(Me.ColBeginQty.Index).Value) = 0 Then
                    If BeginAmount = "" Then
                        BeginAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        BeginAmount = BeginAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColInAmount.Index).Value) > 0 And CDec(row.Cells(Me.ColInQty.Index).Value) = 0 Then
                    If InAmount = "" Then
                        InAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        InAmount = InAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColOutAmount.Index).Value) > 0 And CDec(row.Cells(Me.ColOutQty.Index).Value) = 0 Then
                    If OutAmount = "" Then
                        OutAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        OutAmount = OutAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColEndAmount.Index).Value) > 0 And CDec(row.Cells(Me.ColEndQty.Index).Value) = 0 Then
                    If EndAmount = "" Then
                        EndAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        EndAmount = EndAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If
            Next

            If CodeUnbalanceAmount <> "" Then
                MsgBox("Unbalance amount for: " + CodeUnbalanceAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeUnbalanceQty <> "" Then
                MsgBox("Unbalance qty for: " + CodeUnbalanceQty, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If WrongEndAmountCalculation <> "" Then
                MsgBox("Wrong end amount calculation for: " + WrongEndAmountCalculation, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If WrongOutAmountCalculation <> "" Then
                MsgBox("Wrong out amount calculation for: " + WrongOutAmountCalculation, MsgBoxStyle.Information, "Information")
                Pass = False
            End If


            If CodeNo_QtyBegin <> "" Then
                MsgBox("The following code no, have negative beginning qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyBegin, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If CodeNo_QtyIn <> "" Then
                MsgBox("The following code no, have negative in qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyIn, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If CodeNo_QtyOut <> "" Then
                MsgBox("The following code no, have negative out qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyOut, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If CodeNo_QtyEnd <> "" Then
                MsgBox("The following code no, have negative end qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyEnd, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If BeginAmount <> "" Then
                MsgBox("The following code no, have positive beginning amount, but zero beginning qty. Please make sure you are aware of this." + _
                vbCrLf + BeginAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If InAmount <> "" Then
                MsgBox("The following code no, have positive in amount, but zero in qty. Please make sure you are aware of this." + _
                vbCrLf + InAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If OutAmount <> "" Then
                MsgBox("The following code no, have positive out amount, but zero out qty. Please make sure you are aware of this." + _
                vbCrLf + OutAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If EndAmount <> "" Then
                MsgBox("The following code no, have positive end amount, but zero end qty. Please make sure you are aware of this." + _
                vbCrLf + EndAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeNoBegin <> "" Then
                MsgBox("The following code no, have positive beginning qty but zero beginning amount, please make sure you are aware of this." + _
                vbCrLf + CodeNoBegin, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeNo <> "" Then
                MsgBox("The following code no, have positive in qty but zero in amount, please make sure you are aware of this." + _
                vbCrLf + CodeNo, MsgBoxStyle.Information, "Information")
                Pass = False
            Else
                If CodeNoOut <> "" Then
                    MsgBox("The following code no, have positive out qty but zero out amount, please make sure you are aware of this." + _
                    vbCrLf + CodeNoOut, MsgBoxStyle.Information, "Information")
                    Pass = False
                End If

                If CodeNoEnd <> "" Then
                    MsgBox("The following code no, have positive out qty but zero end amount, please make sure you are aware of this." + _
                    vbCrLf + CodeNoEnd, MsgBoxStyle.Information, "Information")
                    Pass = False
                End If
            End If

            For Each row In Me.dgvItemAdjust.Rows
                If CDec(row.Cells(Me.ColUnitPriceAdjust.Index).Value) = 0 Then
                    If Dummy_Adjust = "" Then
                        Dummy_Adjust = "Line number: " + row.Cells(Me.ColNoAdjust.Index).Value.ToString + " - Code No: " + row.Cells(Me.ColCodeNoAdjust.Index).Value.ToString
                    Else
                        Dummy_Adjust = Dummy_Adjust + vbCrLf + "Line number: " + row.Cells(Me.ColNoAdjust.Index).Value.ToString + " - Code No: " + row.Cells(Me.ColCodeNoAdjust.Index).Value.ToString
                    End If
                End If
            Next

            If Dummy_Adjust <> "" Then
                If MsgBox("The following adjustment have zero unit price:" + vbCrLf _
                + Dummy_Adjust + vbCrLf + "Continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Pass = False
                End If
            End If

            'cek qty beginning, qty in, qty out, qty end, harus lebih besar dari nol, kalo ga salah jadinya
            If Pass = False Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Transaction number")
            KoleksiTextBox.Add(Me.TxtClosingNumber)

            If ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                If MsgBox("Saved data with no item?", MsgBoxStyle.YesNo, "COnfirmation") <> MsgBoxResult.Yes Then
                    'MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.ObjClosing.TimeStamp <> Me.ObjClosing.GetTimestamp(Me.ObjClosing.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjClosing.CariDetail(Me.ObjClosing.Primary, Me.dgvItem, Me.dgvItemAdjust)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            If Me.ApakahDataBenar2 = False Then
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeField()
        Try
            With Me.ObjClosing
                Me.TxtPrimaryClosing.Text = .Primary
                Me.TxtClosingNumber.Text = .NoTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.dtpDate.Value = .ClosingDate
                Me.dtpPeriod.Value = .ClosingPeriod

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            'Me.FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.ObjClosing
                .NoTransaksi = Me.TxtClosingNumber.Text
                .Remarks = Me.TxtRemark.Text
                .ClosingDate = Me.dtpDate.Value
                .ClosingPeriod = Me.dtpPeriod.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.ReadOnly = True
            Me.dgvItemAdjust.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.btnFindRow.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryClosing.BackColor, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, _
            True, True, True, True, True)

            ' Dim nt As New Threading.Thread(AddressOf Me.LoadClosingData)
            'nt.Start()
            Me.LoadClosingData()
            Me.dgvItem.ReadOnly = False
            Me.dgvItemAdjust.ReadOnly = False
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadClosingData()
        Dim row As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2Row
        Dim rowDGV, rowDGVAdjust As DataGridViewRow
        Dim OldText As String = Me.Text
        Dim TipeTransaksiStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
        Dim TBLInAdjust As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item6DataTable
        Dim DTAAdjustIn As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item6TableAdapter
        Dim rowAdjust As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item6Row
        Dim TotalAdjustment As Decimal = 0
        Dim counter As Double = 0

        Dim AddRowDelegate As AddNewRow = AddressOf Me.AddNewRow1
        Dim AddInAmountDelegate As AddInAmount = AddressOf Me.AddInAmount1
        Dim SetEndAmountDelegate As SetEndAmount = AddressOf Me.SetEndAmount1
        Dim SetOutAmountDelegate As SetOutAmount = AddressOf Me.SetOutAmount1

        Dim DTATaking1 As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_stock_takingTableAdapter
        Dim DTATaking2 As New DseTransaksiMonthlyClosingTableAdapters.laporan_stock_taking_view2TableAdapter
        Dim TBLTaking1 As New DseTransaksiMonthlyClosing.tbltrans_stock_takingDataTable
        Dim TBLTaking2 As New DseTransaksiMonthlyClosing.laporan_stock_taking_view2DataTable
        Dim rowTaking1 As DseTransaksiMonthlyClosing.tbltrans_stock_takingRow
        Dim rowTaking2 As DseTransaksiMonthlyClosing.laporan_stock_taking_view2Row

        Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.DataTable12TableAdapter
        Dim TBLBundling As New DseTransaksiMonthlyClosing.DataTable12DataTable
        Dim rowBundling As DseTransaksiMonthlyClosing.DataTable12Row

        Dim TBLSalesReturn As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item81DataTable
        Dim DTASalesReturn As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item81TableAdapter
        Dim rowSalesReturn As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item81Row
        Dim counter1 As Double = 0
        Dim Filter As String = ""

        Dim TBL_In_SR As New DseTransaksiMonthlyClosing.DataTable20DataTable
        Dim row_In_SR As DseTransaksiMonthlyClosing.DataTable20Row
        Dim DTA_In_SR As New DseTransaksiMonthlyClosingTableAdapters.DataTable20TableAdapter

        Dim DTA_HIstoricalPrice As New DseTransaksiMonthlyClosingTableAdapters.DataTable29TableAdapter
        Dim TBL_HistoricalPrice As New DseTransaksiMonthlyClosing.DataTable29DataTable
        Dim row_HistoricalPrice As DseTransaksiMonthlyClosing.DataTable29Row
        Dim cmd As String = ""

        Dim DTABefore As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter '--> Update ke SMESS
        Dim DTAAfter As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter
        Dim TBLBefore As New DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable
        Dim TBLAfter As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable

        Dim TBLBefore1 As New DseTransaksiMonthlyClosing.DataTable16DataTable
        Dim DTABefore1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable16TableAdapter
        Dim TBLIn As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable
        Dim DTAIn As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item2TableAdapter

        Dim TBLPurchase As New DseTransaksiMonthlyClosing.DataTable4DataTable
        Dim DTAPurchase As New DseTransaksiMonthlyClosingTableAdapters.DataTable4TableAdapter
        Dim TBLSR As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable
        Dim DTASR As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item9_1TableAdapter

        Dim Sort As Boolean = False
        Dim DTABundling1 As New Inventory_Module.dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_after1TableAdapter
        Dim DefPrec As Integer = 0
        Dim r As Double = 0

        Try
            DefPrec = Me.objCurrency.GetDecimal_DefaultCurrency

            If Me.baru = True Then
                'hapuskan data unit price sementara
                Me.TBLUnitPrice.Rows.Clear()
                Me.TBLItemAfter.Rows.Clear()
                Me.TBLItemBefore.Rows.Clear()

                Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                Me.dgvItem.Rows.Clear()
                Me.dgvItemAdjust.Rows.Clear()

                'set disable
                If Me.ObjClosing.GetJml = 0 Then
                    Me.ColInAmount.Visible = False
                    Me.ColInQty.Visible = False
                    Me.ColBeginQty.Visible = False
                    Me.ColBeginAmount.Visible = False
                    Me.ColOutQty.Visible = False
                    Me.ColOutAmount.Visible = False
                    Me.dtpPeriod.Enabled = True
                    Me.dgvItem.ReadOnly = False
                    Me.LoadClosing = False
                    Me.dtpPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
                    Me.LoadClosing = True
                    Me.ColDetailCalculation.Visible = False
                Else
                    Me.ColInAmount.Visible = True
                    Me.ColInQty.Visible = True
                    Me.ColBeginQty.Visible = True
                    Me.ColBeginAmount.Visible = True
                    Me.ColOutQty.Visible = True
                    Me.ColOutAmount.Visible = True
                    Me.dtpPeriod.Enabled = False
                    Me.LoadClosing = False
                    Me.dtpPeriod.Value = Me.ObjClosing.GetNextClosingPeriode
                    Me.LoadClosing = True
                    Me.ColDetailCalculation.Visible = True
                End If

                'load data unit price 
                DTA_HIstoricalPrice.Fill(TBL_HistoricalPrice)
                'isikan data awal adjustment 
                'dari stock in adjust
                Me.Text = OldText + " - Extracting in adjustment from database"
                TotalAdjustment = DTAAdjustIn.Fill(TBLInAdjust, TipeTransaksiStockInOut, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
                Me.dtpPeriod.Value)
                For Each rowAdjust In TBLInAdjust.Rows
                    Me.Text = OldText + "- Loading adjustment data (" + CStr(Me.dgvItemAdjust.Rows.Count) _
                    + " of " + TotalAdjustment.ToString + ")"
                    Me.dgvItemAdjust.Rows.Add("")
                    rowDGV = Me.dgvItemAdjust.Rows(Me.dgvItemAdjust.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColNoAdjust.Index).Value = Me.dgvItemAdjust.Rows.Count
                        .Cells(Me.ColJenisDocumentAdjustStr.Index).Value = "Stock In Adjust"
                        .Cells(Me.ColDocumentNo.Index).Value = rowAdjust.No_Transaksi_Stock_In_Adjust
                        .Cells(Me.ColCodeNoAdjust.Index).Value = rowAdjust.Product_Code
                        .Cells(Me.ColDescAdjust.Index).Value = rowAdjust.Description
                        .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowAdjust.Qty, rowAdjust._Decimal)
                        .Cells(Me.ColUnitAdjust.Index).Value = rowAdjust.Unit_Code

                        cmd = "[Primary Product] =" + rowAdjust.Primary_Product.ToString
                        TBL_HistoricalPrice.DefaultView.RowFilter = cmd
                        If TBL_HistoricalPrice.DefaultView.Count > 0 Then
                            row_HistoricalPrice = TBL_HistoricalPrice.DefaultView(0).Row
                            rowAdjust.Unit_Price = row_HistoricalPrice.At_Price
                            .Cells(Me.ColRemarksAdjust.Index).Value = "refer to purchase price from: " _
                            + row_HistoricalPrice.Supplier_Name + "(" + row_HistoricalPrice.Supplier_Code + _
                            ") with aop number: " + row_HistoricalPrice.No_Transaksi + " on " + Format(row_HistoricalPrice.Tanggal_Transaksi, _
                            "dd MMM yyyy")
                            .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(rowAdjust.Unit_Price, Me.DefaultPrecision)
                        Else
                            .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                            .Cells(Me.ColRemarksAdjust.Index).Value = ""
                        End If

                        .Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = rowAdjust.Primary_Product
                        .Cells(Me.ColPrimaryUnitAdjust.Index).Value = rowAdjust.Primary_Units
                        .Cells(Me.ColPrimaryStockInAdjust.Index).Value = rowAdjust.Primary_Stock_In_Adjust
                        .Cells(Me.ColJenisDokumen.Index).Value = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInAdjust
                    End With
                Next

                'dari stock taking
                Me.Text = OldText + " - Extracting in adjustment from stock taking"
                TotalAdjustment = DTATaking1.Fill(TBLTaking1, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
                Me.dtpPeriod.Value)

                'hitung total adjustment
                TotalAdjustment = 0
                For Each rowTaking1 In TBLTaking1.Rows
                    TotalAdjustment += DTATaking2.ScalarQuery_GetJmlByPrimaryTaking(rowTaking1.Primary_Stock_Taking).Value
                Next

                'masukkan ke dalam item adjustment
                counter = 0
                For Each rowTaking1 In TBLTaking1.Rows
                    DTATaking2.Fill(TBLTaking2, rowTaking1.Primary_Stock_Taking)
                    For Each rowTaking2 In TBLTaking2.Rows
                        counter += 1
                        Me.Text = OldText + "- Loading adjustment data (" + CStr(counter) _
                        + " of " + TotalAdjustment.ToString + ")"
                        Me.dgvItemAdjust.Rows.Add("")
                        rowDGV = Me.dgvItemAdjust.Rows(Me.dgvItemAdjust.Rows.Count - 1)
                        With rowDGV
                            .Cells(Me.ColNoAdjust.Index).Value = Me.dgvItemAdjust.Rows.Count
                            .Cells(Me.ColJenisDocumentAdjustStr.Index).Value = "Stock Taking In Adjustment"
                            .Cells(Me.ColDocumentNo.Index).Value = rowTaking1.Stock_Taking_Number
                            .Cells(Me.ColCodeNoAdjust.Index).Value = rowTaking2.Product_Code
                            .Cells(Me.ColDescAdjust.Index).Value = rowTaking2.Description
                            .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowTaking2.Qty, rowTaking2._Decimal)
                            .Cells(Me.ColUnitAdjust.Index).Value = rowTaking2.Unit_Code
                            'If rowTaking2.PO_Number = "" Then
                            '.Cells(Me.ColRemarksAdjust.Index).Value = ""
                            'Else
                            '.Cells(Me.ColRemarksAdjust.Index).Value = "refer to purchase price from: " _
                            '+ rowTaking2.Supplier_Name + "(" + rowTaking2.Supplier_Code + _
                            '") with po number: " + rowTaking2.PO_Number + " on " + Format(rowTaking2.PO_Date, _
                            '"dd MMM yyyy")
                            'End If

                            cmd = "[Primary Product] =" + rowTaking2.Primary_Product.ToString
                            TBL_HistoricalPrice.DefaultView.RowFilter = cmd
                            If TBL_HistoricalPrice.DefaultView.Count > 0 Then
                                row_HistoricalPrice = TBL_HistoricalPrice.DefaultView(0).Row
                                rowTaking2.Unit_Price = row_HistoricalPrice.At_Price
                                .Cells(Me.ColRemarksAdjust.Index).Value = "refer to purchase price from: " _
                                + row_HistoricalPrice.Supplier_Name + "(" + row_HistoricalPrice.Supplier_Code + _
                                ") with aop number: " + row_HistoricalPrice.No_Transaksi + " on " + Format(row_HistoricalPrice.Tanggal_Transaksi, _
                                "dd MMM yyyy")
                                .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(rowTaking2.Unit_Price, DefPrec)
                            Else
                                .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                                .Cells(Me.ColRemarksAdjust.Index).Value = ""
                            End If

                            .Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = rowTaking2.Primary_Product
                            .Cells(Me.ColPrimaryUnitAdjust.Index).Value = rowTaking2.Primary_Unit
                            .Cells(Me.ColPrimaryStockInAdjust.Index).Value = rowTaking2.Primary_Stock_Taking
                            .Cells(Me.ColJenisDokumen.Index).Value = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockTakingInAdjustment
                        End With
                    Next
                Next

                'masukkan data unit price bundling, untuk item - item yang mengakibatkan perhitungan tanpa
                'batas
                Me.Text = OldText + " - Extracting in adjustment from stock bundling"
                TotalAdjustment = DTABundling.Fill(TBLBundling, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
                Me.dtpPeriod.Value)

                'masukkan ke dalam item adjustment
                counter = 0
                For Each rowBundling In TBLBundling.Rows
                    counter += 1
                    Me.Text = OldText + "- Loading bundling data (" + CStr(counter) _
                    + " of " + TotalAdjustment.ToString + ")"
                    Me.dgvItemAdjust.Rows.Add("")
                    rowDGV = Me.dgvItemAdjust.Rows(Me.dgvItemAdjust.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColNoAdjust.Index).Value = Me.dgvItemAdjust.Rows.Count
                        .Cells(Me.ColJenisDocumentAdjustStr.Index).Value = "Stock Bundling Item Before"
                        .Cells(Me.ColDocumentNo.Index).Value = rowBundling.No_Transaksi
                        .Cells(Me.ColCodeNoAdjust.Index).Value = rowBundling.Product_Code
                        .Cells(Me.ColDescAdjust.Index).Value = rowBundling.Description
                        .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowBundling.Qty, rowBundling._Decimal)
                        .Cells(Me.ColUnitAdjust.Index).Value = rowBundling.Unit_Code
                        .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = rowBundling.Primary_Product
                        .Cells(Me.ColPrimaryUnitAdjust.Index).Value = rowBundling.Primary_Unit
                        .Cells(Me.ColPrimaryStockInAdjust.Index).Value = rowBundling.Primary_Transaksi
                        .Cells(Me.ColJenisDokumen.Index).Value = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockBundlingAdjustment
                        .Cells(Me.ColRemarksAdjust.Index).Value = ""
                    End With
                Next

                'masukkan data stock in sales return yang berasal dari bulan yang sama, untuk data2 ini, 
                'harus dituliskan secara manual, direkomendasikan untuk ambil nilainya dari unit price
                'periode sebelumnya jika ada
                TotalAdjustment = DTA_In_SR.Fill(TBL_In_SR, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
                Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value))
                counter = 0
                For Each row_In_SR In TBL_In_SR.Rows
                    counter += 1
                    Me.Text = OldText + "- Loading in sales return data (" + CStr(counter) _
                    + " of " + TotalAdjustment.ToString + ")"
                    Me.dgvItemAdjust.Rows.Add("")
                    rowDGV = Me.dgvItemAdjust.Rows(Me.dgvItemAdjust.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColNoAdjust.Index).Value = Me.dgvItemAdjust.Rows.Count
                        .Cells(Me.ColJenisDocumentAdjustStr.Index).Value = "Stock In Sales Return"
                        .Cells(Me.ColDocumentNo.Index).Value = row_In_SR.Stock_In_Sales_Return_Number
                        .Cells(Me.ColCodeNoAdjust.Index).Value = row_In_SR.Product_Code
                        .Cells(Me.ColDescAdjust.Index).Value = row_In_SR.Description
                        .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row_In_SR.Qty, row_In_SR._Decimal)
                        .Cells(Me.ColUnitAdjust.Index).Value = row_In_SR.Unit_Code
                        .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
                        .Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = row_In_SR.Primary_Product
                        .Cells(Me.ColPrimaryUnitAdjust.Index).Value = row_In_SR.Primary_Units
                        .Cells(Me.ColPrimaryStockInAdjust.Index).Value = row_In_SR.Primary
                        .Cells(Me.ColJenisDokumen.Index).Value = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama
                        'rekomendasikan closing price bulan kemarin
                        counter1 = Me.GetUnitPrice(row_In_SR.Primary_Product, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date _
                        - TimeSpan.FromDays(1)))
                        If counter1 = 0 Then
                            .Cells(Me.ColRemarksAdjust.Index).Value = ""
                        Else
                            .Cells(Me.ColUnitPriceAdjust.Index).Value = counter1
                            .Cells(Me.ColRemarksAdjust.Index).Value = "refer to unit price on " + _
                            Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date - TimeSpan.FromDays(1)), _
                            "MMM yyyy")
                        End If
                    End With
                Next

                Me.Text = OldText + " - Extracting product data from database"
                'isikan data awal
                Me.TBL.Rows.Clear()
                If DTABundling1.ScalarQuery(Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.dtpPeriod.Value.Date).HasValue = True Then '--> Update ke SMESS, berkaitan dengan urutan
                    If DTABundling1.ScalarQuery(Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.dtpPeriod.Value.Date).Value > 25 Then
                        Sort = True
                    Else
                        Sort = False
                    End If
                Else
                    Sort = False
                End If
                Me.ObjClosing.IsiTBL4Closing(Me.TBL, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
                Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value), Sort)

                Me.DTABeginning.Fill(Me.TBLBeginning)
                Me.SetBeginningBalance(Me.TBL)
                DTASalesReturn.Fill(TBLSalesReturn, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), _
                Me.dtpPeriod.Value.Date)
                counter1 = 0
                While counter1 < TBLSalesReturn.Rows.Count
                    rowSalesReturn = TBLSalesReturn.Rows(counter1)
                    If rowSalesReturn.Price_At_Closing = 0 Then
                        If rowSalesReturn.Closing_Period.Date <> Me.dtpPeriod.Value.Date Then
                            'tambahkan nilai in amount untuk produt yang qty end =0, ingat kalo qty end =0 
                            'query mengembalikan nilai in amount =0, jadi ditambahkan secara manual
                            'nilainya 0 karena stocknya keluar semua pada bulan ybs
                            Filter = "[Primary Product] = " + rowSalesReturn.Primary_Product.ToString
                            Me.TBL.DefaultView.RowFilter = Filter
                            If TBL.DefaultView.Count > 0 Then
                                row = TBL.DefaultView(0).Row
                                row.In_Amount += Me.GetUnitPrice(rowSalesReturn.Primary_Product, _
                                rowSalesReturn.Closing_Period) * rowSalesReturn.Total_Qty_Retur
                            End If
                        End If
                        counter1 += 1
                    Else
                        TBLSalesReturn.Rows.Remove(rowSalesReturn)
                    End If
                End While

                'jika ada bundling yang berputar tanpa batas maka ambil dari dgv 
                DTABundling.Fill(TBLBundling, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), _
                Me.dtpPeriod.Value.Date)

                DTABefore.FillBy_TanggalSampaiDari(TBLBefore, _
                Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value)

                'isikan data yang akan digunakan untuk perhitungan bundling pada fungsi GetUnitPrice 
                DTABefore1.FillBy_ClosingPeriod2(TBLBefore1, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date - TimeSpan.FromDays(1)))
                DTAIn.Fill(TBLIn, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), _
                Me.dtpPeriod.Value.Date)
                DTAPurchase.Fill(TBLPurchase, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.dtpPeriod.Value.Date)
                DTASR.FillBy_Period(TBLSR, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.dtpPeriod.Value.Date)
                Me.DTA_CL.FillBy_All(Me.TBL_CL)

                For Each row In Me.TBL.Rows
                    Me.Text = OldText + "- Loading product data (" + CStr(Me.dgvItem.Rows.Count) _
                    + " of " + Me.TBL.Rows.Count.ToString + ")"
                    If row.Begin_Qty <> 0 Or row.In_Qty <> 0 Or row.Out_Qty <> 0 Or row.End_Qty <> 0 Or _
                    row.Begin_Amount <> 0 Or row.In_Amount <> 0 Then
                        'tambahkan baris baru
                        AddRowDelegate.Invoke(row.Begin_Qty, row.Begin_Amount, row.Primary_Product, row.Product_Code, _
                        row.Description, row.Primary_Unit, row.Unit_Code, row.Out_Qty, row.End_Qty, row.In_Qty, _
                        row.In_Amount)
                        rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)

                        'tambahkan in amount dengan:
                        '1. Bundling
                        AddInAmountDelegate.Invoke(row.Primary_Product, _
                        Me.GetInAmountFromBundling(row.Primary_Product, DTABefore, DTAAfter, TBLBefore, TBLAfter, _
                        TBLBefore1, TBLIn, TBLPurchase, TBLSR))

                        '2. In Adjust + stock taking
                        For Each rowDGVAdjust In Me.dgvItemAdjust.Rows
                            If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = _
                            rowDGVAdjust.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value Then
                                AddInAmountDelegate.Invoke(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value, Math.Round(rowDGVAdjust.Cells(Me.ColUnitPriceAdjust.Index).Value * _
                                rowDGVAdjust.Cells(Me.ColQty.Index).Value, Me.DefaultPrecision, MidpointRounding.AwayFromZero))
                            End If
                        Next

                        'hitung unit price
                        If rowDGV.Cells(Me.ColInQty.Index).Value + rowDGV.Cells(Me.ColBeginQty.Index).Value <> 0 Then
                            'hitung end balance
                            If CDec(rowDGV.Cells(Me.ColEndQty.Index).Value) <> 0 Then
                                SetEndAmountDelegate.Invoke(rowDGV, (rowDGV.Cells(Me.ColEndQty.Index).Value * _
                                (rowDGV.Cells(Me.ColBeginAmount.Index).Value + rowDGV.Cells(Me.ColInAmount.Index).Value)) / _
                                (rowDGV.Cells(Me.ColBeginQty.Index).Value + rowDGV.Cells(Me.ColInQty.Index).Value))
                            Else
                                SetEndAmountDelegate.Invoke(rowDGV, 0)
                            End If

                            'hitung out amount
                            SetOutAmountDelegate.Invoke(rowDGV, (rowDGV.Cells(Me.ColBeginAmount.Index).Value + rowDGV.Cells(Me.ColInAmount.Index).Value) - _
                            rowDGV.Cells(Me.ColEndAmount.Index).Value)
                        Else
                            SetEndAmountDelegate.Invoke(rowDGV, 0)
                            SetOutAmountDelegate.Invoke(rowDGV, 0)
                        End If
                    End If
                        Next

                Me.FormatDGV()
            Else
                Me.dtpPeriod.Enabled = False
            End If

            'buang yang jasa
            Me.DTAJasa.Fill(Me.TBLJasa)
            For Each rowJasa As Core_Module.dseMasterProduct.tblmaster_jasaRow In Me.TBLJasa.Rows
                'buang item 
                r = 0
                While r <= Me.dgvItem.Rows.Count - 1
                    rowDGV = Me.dgvItem.Rows(r)
                    If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = rowJasa.Primary Then
                        Me.dgvItem.Rows.Remove(rowDGV)
                    Else
                        r += 1
                    End If
                End While
                'buang item adjustment
                r = 0
                While r <= Me.dgvItemAdjust.Rows.Count - 1
                    rowDGV = Me.dgvItemAdjust.Rows(r)
                    If rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = rowJasa.Primary Then
                        Me.dgvItemAdjust.Rows.Remove(rowDGV)
                    Else
                        r += 1
                    End If
                End While
            Next

            If Sort = True Then
                Me.dgvItem.Sort(Me.ColCodeNo, System.ComponentModel.ListSortDirection.Ascending) '--> update ke SMESS
            End If

            Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.Text = OldText
            Me.dgvItemAdjust.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAAdjustIn.Dispose()
            TBLInAdjust.Dispose()
            DTATaking1.Dispose()
            DTATaking2.Dispose()
            TBLTaking1.Dispose()
            TBLTaking2.Dispose()
            TBLSalesReturn.Dispose()
            DTASalesReturn.Dispose()
            TBL_In_SR.Dispose()
            DTA_In_SR.Dispose()
            DTA_HIstoricalPrice.Dispose()
            TBL_HistoricalPrice.Dispose()

            DTABefore.Dispose()
            DTAAfter.Dispose()
            TBLBefore.Dispose()
            TBLAfter.Dispose()
            TBLBefore1.Dispose()
            DTABefore1.Dispose()
            TBLAfter.Dispose()
            DTAAfter.Dispose()
            TBLPurchase.Dispose()
            DTAPurchase.Dispose()
            DTASR.Dispose()
            TBLSR.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' tidak memperhitungkan baris terakhir dalam dgv
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetInAmount_DiDGVItem(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> Me.dgvItem.Rows.Count - 1 Then
                    If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                        Return rowDGV.Cells(Me.ColInAmount.Index).Value
                    End If
                End If
            Next

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahSudahAda(ByVal vPrimaryCodeNo As Double) As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> Me.dgvItem.Rows.Count - 1 Then
                    If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                        Return True
                    End If
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetUnitPrice(ByVal vPrimaryCodeNo As Double, ByVal DTABefore As DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter, _
    ByVal DTAAfter As DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter, ByVal TBLBefore As DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable, _
    ByVal TBLAfter As DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable, ByVal ObjClosingItem As ClsTransaksiMonthlyClosingItem, _
    ByVal TBL1 As DseTransaksiMonthlyClosing.DataTable16DataTable, _
    ByVal TBL2 As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable, _
    ByVal TBLPurchase As DseTransaksiMonthlyClosing.DataTable4DataTable, _
    ByVal TBLSR As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable) As Decimal '--> update ke smess

        Dim InAmount As Decimal = 0
        Dim rowDGVAdjust As DataGridViewRow
        Dim BeginQty, BeginAmount, InQty As Decimal

        Try
            'ambil begin qty, begin amount, in qty, in amount-step1
            ObjClosingItem.Get_BeginAmount_BeginQty_InQty_IAmount(vPrimaryCodeNo, _
            BeginAmount, BeginQty, InQty, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
            Me.dtpPeriod.Value, Me.ObjCD, TBL1, TBL2)

            'hitung unit price
            If BeginQty + InQty <> 0 Then
                If Me.ApakahSudahAda(vPrimaryCodeNo) = True Then
                    InAmount = Me.GetInAmount_DiDGVItem(vPrimaryCodeNo)
                Else
                    'tambahkan in amount dengan:-step2
                    InAmount = ObjClosingItem.GetInAmount_Purchase_SalesReturnBedaPeriode(vPrimaryCodeNo, _
                    Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
                    Me.dtpPeriod.Value, Me.TBL_CL, Me.ObjCD, TBLPurchase, TBLSR, Me.DefaultPrecision)

                    '1. bundling-step3
                    InAmount = InAmount + Me.GetInAmountFromBundling(vPrimaryCodeNo, DTABefore, DTAAfter, TBLBefore, TBLAfter, TBL1, TBL2, TBLPurchase, TBLSR)

                    '2. in adjustment dari stock in dan stock taking-step4
                    For Each rowDGVAdjust In Me.dgvItemAdjust.Rows
                        If rowDGVAdjust.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = vPrimaryCodeNo Then
                            InAmount += Math.Round(rowDGVAdjust.Cells(Me.ColUnitPriceAdjust.Index).Value * _
                            rowDGVAdjust.Cells(Me.ColQty.Index).Value, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                        End If
                    Next
                End If

                'hitung unit price
                Return Math.Round((BeginAmount + InAmount) / _
                (BeginQty + InQty), Me.DefaultPrecision, MidpointRounding.AwayFromZero)
            Else
                'tidak ada unit price
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjClosingItem = Nothing
        End Try
    End Function

    Public Function GetUnitPrice(ByVal vPrimaryProduct As Double, ByVal vClosingPeriod As Date) As Decimal
        Dim Filter As String = ""
        Dim rowClosing As DseTransaksiMonthlyClosing.DataTable21Row

        Try
            If vClosingPeriod <> Me.Dum_ClosingPeriode Then
                Me.Dum_ClosingPeriode = vClosingPeriod
                Me.DTAClosing1.Fill(Me.TBLClosing1, Me.Dum_ClosingPeriode)
            End If
            Filter = "[Primary Product] = " + vPrimaryProduct.ToString
            Me.TBLClosing1.DefaultView.RowFilter = Filter
            If Me.TBLClosing1.DefaultView.Count = 0 Then
                Return 0
            Else
                rowClosing = Me.TBLClosing1.DefaultView(0).Row
                If rowClosing.Ending_Qty = 0 Then
                    If rowClosing.Out_Qty = 0 Then
                        Return 0
                    Else
                        Return Math.Round(rowClosing.Out_Amount / rowClosing.Out_Qty, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                    End If
                Else
                    Return Math.Round(rowClosing.Ending_Amount / rowClosing.Ending_Qty, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Information")
        End Try
    End Function

    Private Sub SetBeginningBalance(ByVal vTBL As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable)
        Dim row As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2Row
        Dim rowBeginning As DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item1Row
        Dim Filter As String = ""

        Try
            For Each row In vTBL.Rows
                Filter = "[Primary Product] =" + row.Primary_Product.ToString
                Me.TBLBeginning.DefaultView.RowFilter = Filter
                If Me.TBLBeginning.DefaultView.Count <> 0 Then
                    rowBeginning = Me.TBLBeginning.DefaultView(0).Row
                    row.Begin_Amount = rowBeginning.Ending_Amount
                    row.Begin_Qty = rowBeginning.Ending_Qty
                Else
                    row.Begin_Amount = 0
                    row.Begin_Qty = 0
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetOutAmount1(ByVal rowDGV As DataGridViewRow, ByVal OutAmount As Decimal)
        Try
            rowDGV.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(OutAmount, _
            Me.DefaultPrecision, MidpointRounding.AwayFromZero))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetEndAmount1(ByVal rowDGV As DataGridViewRow, ByVal EndAmount As Decimal)
        Try
            rowDGV.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(EndAmount, _
            Me.DefaultPrecision, MidpointRounding.AwayFromZero))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddInAmount1(ByVal vPrimaryCodeNo As Double, ByVal vAdditionalAmount As Decimal)
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    rowDGV.Cells(Me.ColInAmount.Index).Value = CDec(rowDGV.Cells(Me.ColInAmount.Index).Value) _
                    + vAdditionalAmount
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub AddNewRow1(ByVal BeginQty As Decimal, ByVal BeginAmount As Decimal, _
    ByVal vPrimaryProduct As Double, ByVal vCodeNo As String, ByVal Description As String, ByVal vPrimaryUnit As Double, _
    ByVal UnitCode As String, ByVal OutQty As Decimal, ByVal EndQty As Decimal, ByVal InQty As Decimal, ByVal InAmount As Decimal)
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.Rows.Add("")
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
            With rowDGV
                .Cells(Me.ColNo.Index).Value = CStr(Me.dgvItem.Rows.Count)
                .Cells(Me.ColBeginQty.Index).Value = BeginQty
                .Cells(Me.ColBeginAmount.Index).Value = BeginAmount
                .Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct
                .Cells(Me.ColCodeNo.Index).Value = vCodeNo
                .Cells(Me.ColDescription.Index).Value = Description
                .Cells(Me.ColPrimaryUnits.Index).Value = vPrimaryUnit
                .Cells(Me.ColUnits.Index).Value = UnitCode
                .Cells(Me.ColRemarks.Index).Value = ""
                .Cells(Me.ColDetailCalculation.Index).Value = "Show"
                .Cells(Me.ColOutAmount.Index).Value = 0
                .Cells(Me.ColOutQty.Index).Value = OutQty
                .Cells(Me.ColEndAmount.Index).Value = 0
                .Cells(Me.ColEndQty.Index).Value = EndQty
                .Cells(Me.ColInAmount.Index).Value = InAmount
                .Cells(Me.ColInQty.Index).Value = InQty
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                'kalo ada yang perlu dihitung ulang berkaitan dengan data bundling, maka lakukan perhitungan ulang dahulu
                If Me.ListPrimaryCodeNo_Bundling.Count <> 0 Then
                    Me.HitungUlangUnitPrice1()
                End If

                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjClosing.SimpanBaru(Me.dgvItem, Me.dgvItemAdjust, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjClosing.SimpanPerubahan(Me.dgvItem, Me.dgvItemAdjust, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.LoadClosing = False
            ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.LoadClosing = True
            Me.TxtPrimaryClosing.Text = "-1"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            Me.dgvItemAdjust.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjClosing.CariDetail(Me.TxtPrimaryClosing.Text, Me.dgvItem, Me.dgvItemAdjust)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariTransaksiMonthlyClosing

        Try
            ij.BackColorAll = Me.BAckColorAll
            ij.ShowDialog()
            If Not ij.PrimaryClosing = -1 Then
                If Me.ObjClosing.CariDetail(ij.PrimaryClosing, Me.dgvItem, Me.dgvItemAdjust) _
                = True Then
                    Me.SalinKeField()
                    'jika yang ditampilkan adalah data pertama, maka hide kolom kolom yang tidak perlu
                    If Me.ObjClosing.ApakahClosingPalingAwal(Me.TxtPrimaryClosing.Text) = True Then
                        Me.ColInAmount.Visible = False
                        Me.ColInQty.Visible = False
                        Me.ColBeginQty.Visible = False
                        Me.ColBeginAmount.Visible = False
                        Me.ColOutQty.Visible = False
                        Me.ColOutAmount.Visible = False
                        Me.ColDetailCalculation.Visible = False
                    Else
                        Me.ColInAmount.Visible = True
                        Me.ColInQty.Visible = True
                        Me.ColBeginQty.Visible = True
                        Me.ColBeginAmount.Visible = True
                        Me.ColOutQty.Visible = True
                        Me.ColOutAmount.Visible = True
                        Me.ColDetailCalculation.Visible = True
                    End If
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim Dummy As String = ""
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'isikan data tbl product untuk mendapatkan angka decimal
            Me.vDTA1.Fill(Me.vTBL1)

            'peringatan, untuk bulan yang baru, tidak boleh ada AOP dan transaksi stock taking yang
            'masih aktif pada bulan yang bersangkutan
            If Me.ObjClosing.GetJml > 0 Then
                Dim NextClosingPeriod As Date = Me.ObjClosing.GetNextClosingPeriode()
                Dim ObjGRN As New ClsTransaksiGRN
                If ObjGRN.GetJmlBelumAOP(Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod), NextClosingPeriod) > 0 Then
                    MsgBox("There GRN/Stock in data in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + " that don't have approval yet. Please create AOP for: " + _
                    ObjGRN.GetNoGRN_BelumAOP(Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod), NextClosingPeriod) + ".", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                ObjGRN = Nothing

                Dim ObjStockTaking As New Inventory_Module.clsTransaksiStockTaking
                If ObjStockTaking.GetJmlActiveStockTaking(Me.ObjCD.GetFirstDateofTheMonth(NextClosingPeriod), Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod)) > 0 Then
                    MsgBox("There is an active stock taking on progress, please be aware that stock taking processing date cannot be in the period that has been closed.", MsgBoxStyle.Information, "Confirm")
                    ObjStockTaking = Nothing
                    Exit Sub
                End If
                ObjStockTaking = Nothing

                'tidak ada GRN yang dibuat pada bulan ybs yang di AOP bukan pada bulan ybs.
                Dummy = Me.ObjClosing.GetGRN_AOP_BedaPeriode(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are GRN in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " enter in AOP on different month. Please fix following data before continuing closing transaction:" + Dummy, _
                    MsgBoxStyle.Information, "Confirmation")

                    Exit Sub
                End If

                Dummy = Me.ObjClosing.GetAOP_GRNBEdaPeriode(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are AOP in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " enter GRN from different month. Please fix following data before continuing closing transaction:" _
                    + Dummy, MsgBoxStyle.Information, "Confirmation")

                    Exit Sub
                End If

                'tidak boleh ada purchase return yang belum di stock out
                Dummy = Me.ObjClosing.Get_PR_BelumStockOut(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are purchase return in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " with no stock out purchase return. Please fix this data before continuing closing transaction: " + Dummy, _
                    MsgBoxStyle.Information, "Confirmation")
                    Exit Sub
                End If

                'tidak stock out sales yang muncul dari periode sebelumnya --> berkaitan dengan balance out amount
                Dummy = Me.ObjClosing.GetInvoice_StockOutSalesPreviousPeriode(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    If MsgBox("There are invoice in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " with stock out sales from previous month. Please fix this data before continuing closing transaction. Show detail data?", _
                    MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                        Dim ijk As New Sales_Module.FrmLaporanExplainClosingB
                        ijk.dtpFrom.Value = NextClosingPeriod
                        ijk.ShowDialog()
                        ijk = Nothing
                    End If
                    Exit Sub
                End If

                'tidak DO yang sudah stock out sales tapi belum diinvoice
                Dummy = Me.ObjClosing.Get_DO_BelumInvoice(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are delivery order having stock out sales in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " with no invoice. Please fix the following data before continue: " + Dummy, _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                'tidak boleh ada stock in sales return yang dibuatkan sales return pada beda periode
                'atau bahkan belum dibuatkan sales return 
                Dummy = Me.ObjClosing.GetStockInSalesReturn_SRBedaPeriode(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are stock in sales return in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " with sales return transaction on different month. Please fix this data before continuing closing transaction: " + Dummy, _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                'tidak boleh ada sales return yang stock in sales returnya berasal dari periode bukan 
                'bulan yang sama
                Dummy = Me.ObjClosing.GetSR_StockInSalesReturnBedaPeriode(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are sales return in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " with stock in sales return transaction on different month. Please fix this data before continuing closing transaction: " + Dummy, _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                'tidak boleh ada sales return revisi invoice yang belum dimasukkan ke dalam revisi invoice
                Dummy = Me.ObjClosing.Get_SR_RevInvoice_BelumRevisi(NextClosingPeriod, _
                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are sales return(invoice revision) in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " but no invoice. Please fix this data before continuing closing transaction: " + Dummy, _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                'tidak boleh ada purchase return revisi AOP yang belum dimasukkan pada AOP
                Dummy = Me.ObjClosing.Get_PR_Rev_BelumMasukAOP(NextClosingPeriod, _
                                Me.ObjCD.GetLasDateOfTheMonth(NextClosingPeriod))
                If Dummy <> "" Then
                    MsgBox("There are purchase return(AOP revision) in " + Me.ObjCD.GetMonthOfTheYear(NextClosingPeriod.Month) + _
                    " but no AOP. Please fix this data before continuing closing transaction: " + Dummy, _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.baru = True
            Me.KosongkanField()
            Me.ObjClosing.ResetObject(Me.dgvItem)
            Form.CheckForIllegalCrossThreadCalls = False
            Me.AktifkanDetail()

            Me.dtpDate.Focus()
            Me.TxtClosingNumber.Text = Me.ObjClosing.GetNoBaru(Me.dtpPeriod.Value)

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtClosingNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtClosingNumber.Text <> "" Then
                If Me.ObjClosing.ApakahBolehEdit(Me.ObjClosing.Primary, UserCode, True, True) = True Then
                    Me.baru = False
                    Me.ObjClosing.CariDetail(Me.ObjClosing.Primary, Me.dgvItem, Me.dgvItemAdjust)
                    Me.SalinKeField()
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtClosingNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtClosingNumber.Text <> "" Then
                If Me.ObjClosing.ApakahBolehHapus(Me.TxtPrimaryClosing.Text, True, UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjClosing.Hapus(Me.TxtPrimaryClosing.Text, UserCode, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub dtpPeriod_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriod.Enter
        Try
            Me.ClosingPeriodeBuffer = Me.dtpPeriod.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemAdjust_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemAdjust.CellBeginEdit
        Try
            'yang bisa diedit hanya unit price dan remarks saja
            If e.ColumnIndex = Me.ColUnitPriceAdjust.Index Or e.ColumnIndex = Me.ColRemarksAdjust.Index Then
                Me.Buffer = Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erro")
        End Try
    End Sub

    Private Sub dgvItemAdjust_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemAdjust.CellEndEdit
        Dim JenisDocument As ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockBundlingAdjustment

        Try
            'jika yang diedit unit price
            If e.ColumnIndex = Me.ColUnitPriceAdjust.Index Then
                'pastikan tidak nothing
                If IsNothing(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'pastikan nilainya tidak negatif
                If Me.ObjCD.CekNonNegatif(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'digit grouping
                Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefaultPrecision)

                'hitung ulang unit price untuk item yang bersangkutan
                Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value)
                If Me.dgvItemAdjust.Item(Me.ColJenisDokumen.Index, e.RowIndex).Value <> JenisDocument Then
                    Me.HitungUnitPrice(Me.dgvItemAdjust.Item(Me.ColPrimaryCodeNoAdjust.Index, _
                    e.RowIndex).Value)

                    'hitung ulang juga semua unit price yang menggunakan item ybs sebagai
                    'item bundling before
                    Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
                    Dim TBLBundling As New DseTransaksiMonthlyClosing.tbltransaksi_stockbundling_item_afterDataTable
                    Dim rowBundling As DseTransaksiMonthlyClosing.tbltransaksi_stockbundling_item_afterRow

                    DTABundling.Fill(TBLBundling, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value.Date, _
                    Me.dgvItemAdjust.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value)

                    For Each rowBundling In TBLBundling.Rows
                        'Me.HitungUnitPrice(rowBundling.Primary_Product_After)
                        If Me.ListPrimaryCodeNo_Bundling.Contains(rowBundling.Primary_Product_After) = False Then
                            Me.ListPrimaryCodeNo_Bundling.Add(rowBundling.Primary_Product_After)
                        End If
                    Next

                    DTABundling.Dispose()
                    TBLBundling.Dispose()
                Else
                    'hitung ulang semua product yang terpengaruh dengan transaksi bundling yang bersangkutan
                    Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter
                    Dim TBLBundling As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable
                    Dim rowBundling As DseTransaksiMonthlyClosing.Tbltrans_bundling_afterRow

                    DTABundling.FillBy_PrimaryBundling(TBLBundling, _
                    Me.dgvItemAdjust.Item(Me.ColPrimaryStockInAdjust.Index, e.RowIndex).Value)

                    For Each rowBundling In TBLBundling.Rows
                        'Me.HitungUnitPrice(rowBundling.Primary_Product)
                        If Me.ListPrimaryCodeNo_Bundling.Contains(rowBundling.Primary_Product) = False Then
                            Me.ListPrimaryCodeNo_Bundling.Add(rowBundling.Primary_Product)
                        End If
                    Next

                    DTABundling.Dispose()
                    TBLBundling.Dispose()
                End If
            Else
                If IsNothing(Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.dgvItemAdjust.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' menghitung ulang unit price pada dgv item yang primarynya diberikan sesuai dengan 
    ''' yang dimaksud. HANYA BOLEH DIGUNAKAN OLEH FUNGSI --> HitungUlangUnitPrice
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <remarks></remarks>
    Private Sub HitungUnitPrice(ByVal vPrimaryCodeNo As Double, ByVal DTABefore As DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter, _
    ByVal DTAAfter As DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter, ByVal TBLBefore As DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable, _
    ByVal TBLAfter As DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable, ByVal TBL1 As DseTransaksiMonthlyClosing.DataTable16DataTable, _
    ByVal TBL2 As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable, _
    ByVal TBLPurchase As DseTransaksiMonthlyClosing.DataTable4DataTable, _
    ByVal TBLSR As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable, ByVal ObjClosingItem As ClsTransaksiMonthlyClosingItem)
        Dim row, rowDGVAdjust As DataGridViewRow
        Dim InAmount As Decimal = 0
        Dim UnitPrice As Decimal = 0

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    'ambil in amount
                    InAmount = ObjClosingItem.GetInAmount_Purchase_SalesReturnBedaPeriode2(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value, Me.TBL_CL, Me.ObjCD, Me.DefaultPrecision)
                    'tambahkan in amount dengan:
                    '1. bundling
                    InAmount = InAmount + Me.GetInAmountFromBundling(vPrimaryCodeNo, DTABefore, DTAAfter, TBLBefore, _
                    TBLAfter, TBL1, TBL2, TBLPurchase, TBLSR)

                    '2. in adjustment dari stock in dan stock taking
                    For Each rowDGVAdjust In Me.dgvItemAdjust.Rows
                        If rowDGVAdjust.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = vPrimaryCodeNo Then
                            InAmount += Math.Round(rowDGVAdjust.Cells(Me.ColUnitPriceAdjust.Index).Value * _
                            rowDGVAdjust.Cells(Me.ColQty.Index).Value, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                        End If
                    Next
                    row.Cells(Me.ColInAmount.Index).Value = Me.ObjCD.DigitGrouping(InAmount, Me.DefaultPrecision)

                    'hitung unit price
                    If CDec(row.Cells(Me.ColInQty.Index).Value) + CDec(row.Cells(Me.ColBeginQty.Index).Value) <> 0 Then
                        'hitung out amount
                        row.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDbl(row.Cells(Me.ColOutQty.Index).Value) * _
                        (CDbl(row.Cells(Me.ColBeginAmount.Index).Value) + CDbl(row.Cells(Me.ColInAmount.Index).Value)) / _
                        (CDbl(row.Cells(Me.ColBeginQty.Index).Value) + CDbl(row.Cells(Me.ColInQty.Index).Value)), Me.DefaultPrecision, MidpointRounding.AwayFromZero))

                        'hitung end balance
                        row.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDbl(row.Cells(Me.ColEndQty.Index).Value) * _
                        (CDbl(row.Cells(Me.ColBeginAmount.Index).Value) + CDbl(row.Cells(Me.ColInAmount.Index).Value)) / _
                        (CDbl(row.Cells(Me.ColBeginQty.Index).Value) + CDbl(row.Cells(Me.ColInQty.Index).Value)), Me.DefaultPrecision, MidpointRounding.AwayFromZero))
                    Else
                        row.Cells(Me.ColOutAmount.Index).Value = 0
                        row.Cells(Me.ColEndAmount.Index).Value = 0
                    End If

                    'slesai
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' menghitung ulang unit price pada dgv item yang primarynya diberikan sesuai dengan 
    ''' yang dimaksud. HANYA BOLEH DIGUNAKAN OLEH FUNGSI --> dgvItemAdjust_CellEndEdit 
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <remarks></remarks>
    Private Sub HitungUnitPrice(ByVal vPrimaryCodeNo As Double)
        Dim row, rowDGVAdjust As DataGridViewRow
        Dim InAmount As Decimal = 0
        Dim UnitPrice As Decimal = 0
        Dim ObjClosingItem As New ClsTransaksiMonthlyClosingItem

        'elemen2 yang digunakan untuk pehitungan bundling
        Dim DTABefore As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter
        Dim DTAAfter As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter
        Dim TBLBefore As New DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable
        Dim TBLAfter As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable

        Dim TBLBefore1 As New DseTransaksiMonthlyClosing.DataTable16DataTable
        Dim DTABefore1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable16TableAdapter
        Dim TBLIn As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable
        Dim DTAIn As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item2TableAdapter

        Dim TBLPurchase As New DseTransaksiMonthlyClosing.DataTable4DataTable
        Dim DTAPurchase As New DseTransaksiMonthlyClosingTableAdapters.DataTable4TableAdapter
        Dim TBLSR As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable
        Dim DTASR As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item9_1TableAdapter

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    'ambil in amount
                    InAmount = ObjCLosingItem.GetInAmount_Purchase_SalesReturnBedaPeriode2(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value, Me.TBL_CL, Me.ObjCD, Me.DefaultPrecision)
                    'tambahkan in amount dengan:
                    '1. bundling
                    InAmount = InAmount + Me.GetInAmountFromBundling(vPrimaryCodeNo, DTABefore, DTAAfter, TBLBefore, _
                    TBLAfter, TBLBefore1, TBLIn, TBLPurchase, TBLSR)

                    '2. in adjustment dari stock in dan stock taking
                    For Each rowDGVAdjust In Me.dgvItemAdjust.Rows
                        If rowDGVAdjust.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = vPrimaryCodeNo Then
                            InAmount += Math.Round(rowDGVAdjust.Cells(Me.ColUnitPriceAdjust.Index).Value * _
                            rowDGVAdjust.Cells(Me.ColQty.Index).Value, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                        End If
                    Next
                    row.Cells(Me.ColInAmount.Index).Value = Me.ObjCD.DigitGrouping(InAmount, Me.DefaultPrecision)

                    'hitung unit price
                    If CDec(row.Cells(Me.ColInQty.Index).Value) + CDec(row.Cells(Me.ColBeginQty.Index).Value) <> 0 Then
                        'hitung out amount
                        row.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDbl(row.Cells(Me.ColOutQty.Index).Value) * _
                        (CDbl(row.Cells(Me.ColBeginAmount.Index).Value) + CDbl(row.Cells(Me.ColInAmount.Index).Value)) / _
                        (CDbl(row.Cells(Me.ColBeginQty.Index).Value) + CDbl(row.Cells(Me.ColInQty.Index).Value)), Me.DefaultPrecision, MidpointRounding.AwayFromZero))

                        'hitung end balance
                        row.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDbl(row.Cells(Me.ColEndQty.Index).Value) * _
                        (CDbl(row.Cells(Me.ColBeginAmount.Index).Value) + CDbl(row.Cells(Me.ColInAmount.Index).Value)) / _
                        (CDbl(row.Cells(Me.ColBeginQty.Index).Value) + CDbl(row.Cells(Me.ColInQty.Index).Value)), Me.DefaultPrecision, MidpointRounding.AwayFromZero))
                    Else
                        row.Cells(Me.ColOutAmount.Index).Value = 0
                        row.Cells(Me.ColEndAmount.Index).Value = 0
                    End If

                    'slesai
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjClosingItem = Nothing
            DTABefore.Dispose()
            DTAAfter.Dispose()
            TBLBefore.Dispose()
            TBLAfter.Dispose()

            TBLBefore1.Dispose()
            DTABefore1.Dispose()
            TBLIn.Dispose()
            DTAIn.Dispose()

            TBLPurchase.Dispose()
            DTAPurchase.Dispose()
            TBLSR.Dispose()
            DTASR.Dispose()
        End Try
    End Sub

    Private Function GetInAmountFromBundling(ByVal vPrimaryCodeNo As Double, _
    ByVal DTABefore As DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter, ByVal DTAAfter As DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter, ByVal TBLBefore As DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable, _
    ByVal TBLAfter As DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable, ByVal TBL1 As DseTransaksiMonthlyClosing.DataTable16DataTable, _
    ByVal TBL2 As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable, _
    ByVal TBLPurchase As DseTransaksiMonthlyClosing.DataTable4DataTable, _
    ByVal TBLSR As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable) As Decimal
        Dim rowBefore As DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeRow
        Dim rowAfter As DseTransaksiMonthlyClosing.Tbltrans_bundling_afterRow
        Dim Filter As String = ""
        Dim TotalInAmount, InAmountRow As Decimal

        Dim Filter2 As String = ""
        Dim counter As Double = 0

        Dim TBLBefore1 As New DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable
        Dim TBLAfter1 As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable
        DTABefore.FillBy_TanggalSampaiDari(TBLBefore1, _
        Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value)


        Try
            'isikan data dari database
            DTAAfter.FillBy_PrimaryProductAfter(TBLAfter, _
            vPrimaryCodeNo, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value)

            'jumlahkan (unit price * faktor pemberat) untuk setiap item after yang ada
            For Each rowAfter In TBLAfter.Rows
                Filter = "[Primary Bundling] = " + rowAfter.Primary_Bundling.ToString + ""
                TBLBefore.DefaultView.RowFilter = Filter
                InAmountRow = 0
                counter = 0

                While counter < TBLBefore.DefaultView.Count
                    rowBefore = TBLBefore.DefaultView(counter).Row
                    Filter2 = "[Primary Transaksi] =" + rowAfter.Primary_Bundling.ToString + _
                    " AND [Primary Product] =" + rowBefore.Primary_Product.ToString + ""
                    TBLBundling.DefaultView.RowFilter = Filter2
                    If TBLBundling.DefaultView.Count = 0 Then
                        InAmountRow += Me.GetUnitPrice(rowBefore.Primary_Product, DTABefore, DTAAfter, _
                        TBLBefore1, TBLAfter1, Me.ObjCLosingItem, TBL1, TBL2, TBLPurchase, TBLSR) * rowBefore.Qty
                    Else
                        InAmountRow += Me.GetUnitPriceBundling(rowBefore.Primary_Product, _
                        rowAfter.Primary_Bundling) * rowBefore.Qty
                    End If

                    counter += 1
                End While

                TotalInAmount += Math.Round(InAmountRow * rowAfter.Faktor_Pemberat / 100, _
                Me.DefaultPrecision, MidpointRounding.AwayFromZero)
            Next

            Return TotalInAmount
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLBefore1.Dispose()
            TBLAfter1.Dispose()
        End Try
    End Function

    Private Function GetInAmountFromBundling(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim DTABefore As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter
        Dim DTAAfter As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter
        Dim TBLBefore As New DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable
        Dim TBLAfter As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable
        Dim rowBefore As DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeRow
        Dim rowAfter As DseTransaksiMonthlyClosing.Tbltrans_bundling_afterRow
        Dim Filter As String = ""
        Dim TotalInAmount, InAmountRow As Decimal

        Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.DataTable12TableAdapter
        Dim TBLBundling As New DseTransaksiMonthlyClosing.DataTable12DataTable
        Dim Filter2 As String = ""
        Dim counter As Double = 0

        Try
            'jika ada bundling yang berputar tanpa batas maka ambil dari dgv
            DTABundling.Fill(TBLBundling, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), _
            Me.dtpPeriod.Value.Date)

            'isikan data dari database
            DTABefore.FillBy_TanggalSampaiDari(TBLBefore, _
            Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value)
            DTAAfter.FillBy_PrimaryProductAfter(TBLAfter, _
            vPrimaryCodeNo, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value)

            'jumlahkan (unit price * faktor pemberat) untuk setiap item after yang ada
            For Each rowAfter In TBLAfter.Rows
                Filter = "[Primary Bundling] = " + rowAfter.Primary_Bundling.ToString + ""
                TBLBefore.DefaultView.RowFilter = Filter
                InAmountRow = 0
                counter = 0

                While counter < TBLBefore.DefaultView.Count
                    rowBefore = TBLBefore.DefaultView(counter).Row
                    Filter2 = "[Primary Transaksi] =" + rowAfter.Primary_Bundling.ToString + _
                    " AND [Primary Product] =" + rowBefore.Primary_Product.ToString + ""
                    TBLBundling.DefaultView.RowFilter = Filter2
                    If TBLBundling.DefaultView.Count = 0 Then
                        InAmountRow += Me.GetUnitPrice(rowBefore.Primary_Product) * rowBefore.Qty
                    Else
                        InAmountRow += Me.GetUnitPriceBundling(rowBefore.Primary_Product, _
                        rowAfter.Primary_Bundling) * rowBefore.Qty
                    End If

                    counter += 1
                End While

                TotalInAmount += Math.Round((InAmountRow * rowAfter.Faktor_Pemberat) / 100, _
                Me.DefaultPrecision, MidpointRounding.AwayFromZero)
            Next

            Return TotalInAmount
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTABefore.Dispose()
            DTAAfter.Dispose()
            TBLBefore.Dispose()
            TBLAfter.Dispose()
            TBLBundling.Dispose()
            DTABundling.Dispose()
        End Try
    End Function

    Public Function GetUnitPriceBundling(ByVal vPrimaryCodeNo As Double, _
    ByVal vPrimaryTransaksiBundling As Double) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim JenisDocument As ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockBundlingAdjustment

        Try
            For Each rowDGV In Me.dgvItemAdjust.Rows
                If rowDGV.Cells(Me.ColJenisDokumen.Index).Value = JenisDocument And _
                rowDGV.Cells(Me.ColPrimaryStockInAdjust.Index).Value = vPrimaryTransaksiBundling And _
                rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = vPrimaryCodeNo Then
                    Return rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value
                End If
            Next

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dtpPeriod_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriod.ValueChanged
        Try
            If Me.ClosingPeriodeBuffer <> Me.dtpPeriod.Value And Me.TxtRemark.Enabled = True And Me.LoadClosing = True Then
                If MsgBox("Reload closing data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.LoadClosingData()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' mencari baris dimana kode product ybs ada
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnFindRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindRow.Click
        Dim i As New Core_Module.frmCariMasterProduct
        Dim PrimaryProduct As Double = -1
        Dim rowdgv As DataGridViewRow

        Try
            'cari product yang akan diset posisinya
            i.cboCriteria.SelectedIndex = 0
            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                PrimaryProduct = i.PrimaryProduct
                'cari baris yang berikan product tsb
                For Each rowdgv In Me.dgvItem.Rows
                    If rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value = PrimaryProduct Then
                        'fokus ke baris yang bersangkutan
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = rowdgv.Index
                        Me.dgvItem.ClearSelection()
                        rowdgv.Selected = True
                    End If
                Next

                If Me.TabControl1.SelectedIndex <> 0 Then
                    Me.TabControl1.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TabControl1_TabIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.TabIndexChanged
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                Me.btnFindRow.Visible = True
            Else
                Me.btnFindRow.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekEndAmountSaja() As Boolean
        Dim row As DataGridViewRow
        Dim CodeNoEnd As String = ""
        Dim CodeNo_QtyEnd As String = ""
        Dim CodeNo_AmountEnd As String = ""

        Dim Pass As Boolean = True

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColEndQty.Index).Value > 0 And row.Cells(Me.ColEndAmount.Index).Value = 0 Then
                    If CodeNoEnd = "" Then
                        CodeNoEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoEnd = CodeNoEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDbl(row.Cells(Me.ColEndQty.Index).Value) < 0 Then
                    If CodeNo_QtyEnd = "" Then
                        CodeNo_QtyEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyEnd = CodeNo_QtyEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColEndAmount.Index).Value > 0 And row.Cells(Me.ColEndQty.Index).Value = 0 Then
                    If CodeNo_AmountEnd = "" Then
                        CodeNo_AmountEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_AmountEnd = CodeNo_AmountEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If
            Next

            If CodeNo_QtyEnd <> "" Then
                MsgBox("The following code no, have negative end qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyEnd, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeNo_AmountEnd <> "" Then
                MsgBox("The following code no, have positive end amount but zero end qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_AmountEnd, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeNoEnd <> "" Then
                If MsgBox("The following code no, have positive end qty but zero end amount, please make sure you are aware of this." + _
                vbCrLf + CodeNoEnd + vbCrLf + "Continue with this data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                    Pass = False
                End If
            End If

            If Pass = True Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCekData.Click
        Dim row As DataGridViewRow
        Dim x As Double
        Dim CodeNo As String = ""
        Dim CodeNoBegin As String = ""
        Dim CodeNoOut As String = ""
        Dim CodeNoEnd As String = ""
        Dim CodeNo_QtyBegin As String = ""
        Dim CodeNo_QtyIn As String = ""
        Dim CodeNo_QtyOut As String = ""
        Dim CodeNo_QtyEnd As String = ""
        Dim CodeUnbalanceAmount As String = ""
        Dim CodeUnbalanceQty As String = ""
        Dim WrongEndAmountCalculation As String = ""
        Dim WrongOutAmountCalculation As String = ""
        Dim BeginAmount As String = ""
        Dim InAmount As String = ""
        Dim OutAmount As String = ""
        Dim EndAmount As String = ""

        Dim Pass As Boolean = True
        Dim i As Double = -1

        Try
            If Me.ObjClosing.GetJml = 0 Then
                If Me.CekEndAmountSaja() = True Then
                    MsgBox("Data checked", MsgBoxStyle.Information, "Information")
                End If
                Exit Sub
            ElseIf Me.ObjClosing.GetJml = 1 And Me.baru = False Then
                If Me.CekEndAmountSaja() = True Then
                    MsgBox("Data checked", MsgBoxStyle.Information, "Information")
                End If
                Exit Sub
            End If

            For Each row In Me.dgvItem.Rows
                If Math.Abs((CDec(row.Cells(Me.ColBeginAmount.Index).Value) + CDec(row.Cells(Me.ColInAmount.Index).Value)) - _
                (CDec(row.Cells(Me.ColOutAmount.Index).Value) + CDec(row.Cells(Me.ColEndAmount.Index).Value))) > 2 Then
                    If CodeUnbalanceAmount = "" Then
                        CodeUnbalanceAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeUnbalanceAmount = CodeUnbalanceAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColBeginQty.Index).Value) + CDec(row.Cells(Me.ColInQty.Index).Value) <> _
                CDec(row.Cells(Me.ColOutQty.Index).Value) + CDec(row.Cells(Me.ColEndQty.Index).Value) Then
                    If CodeUnbalanceQty = "" Then
                        CodeUnbalanceQty = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeUnbalanceQty = CodeUnbalanceQty + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If (CDec(row.Cells(Me.ColBeginQty.Index).Value) + CDec(row.Cells(Me.ColInQty.Index).Value)) <> 0 Then
                    x = Math.Round((CDec(row.Cells(Me.ColBeginAmount.Index).Value) + CDec(row.Cells(Me.ColInAmount.Index).Value)) / _
                    (CDec(row.Cells(Me.ColBeginQty.Index).Value) + CDec(row.Cells(Me.ColInQty.Index).Value)) * _
                    CDec(row.Cells(Me.ColEndQty.Index).Value), Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                    If Math.Abs(x - CDec(row.Cells(Me.ColEndAmount.Index).Value)) > 2 Then
                        If WrongEndAmountCalculation = "" Then
                            WrongEndAmountCalculation = row.Cells(Me.ColCodeNo.Index).Value.ToString
                        Else
                            WrongEndAmountCalculation = WrongEndAmountCalculation + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                        End If
                    End If
                End If

                If (CDec(row.Cells(Me.ColBeginQty.Index).Value) + CDec(row.Cells(Me.ColInQty.Index).Value)) <> 0 Then
                    x = Math.Round((CDec(row.Cells(Me.ColBeginAmount.Index).Value) + CDec(row.Cells(Me.ColInAmount.Index).Value)) / _
                    (CDec(row.Cells(Me.ColBeginQty.Index).Value) + CDec(row.Cells(Me.ColInQty.Index).Value)) * _
                    CDec(row.Cells(Me.ColOutQty.Index).Value), Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                    If Math.Abs(x - CDec(row.Cells(Me.ColOutAmount.Index).Value)) > 2 Then
                        If WrongOutAmountCalculation = "" Then
                            WrongOutAmountCalculation = row.Cells(Me.ColCodeNo.Index).Value.ToString
                        Else
                            WrongOutAmountCalculation = WrongOutAmountCalculation + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                        End If
                    End If
                End If

                If row.Cells(Me.ColInQty.Index).Value > 0 And row.Cells(Me.ColInAmount.Index).Value = 0 Then
                    If CodeNo = "" Then
                        CodeNo = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo = CodeNo + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColBeginQty.Index).Value > 0 And row.Cells(Me.ColBeginAmount.Index).Value = 0 Then
                    If CodeNoBegin = "" Then
                        CodeNoBegin = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoBegin = CodeNoBegin + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColOutQty.Index).Value > 0 And row.Cells(Me.ColOutAmount.Index).Value = 0 Then
                    If CodeNoOut = "" Then
                        CodeNoOut = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoOut = CodeNoOut + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If row.Cells(Me.ColEndQty.Index).Value > 0 And row.Cells(Me.ColEndAmount.Index).Value = 0 Then
                    If CodeNoEnd = "" Then
                        CodeNoEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNoEnd = CodeNoEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColBeginQty.Index).Value) < 0 Then
                    If CodeNo_QtyBegin = "" Then
                        CodeNo_QtyBegin = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyBegin = CodeNo_QtyBegin + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColInQty.Index).Value) < 0 Then
                    If CodeNo_QtyIn = "" Then
                        CodeNo_QtyIn = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyIn = CodeNo_QtyIn + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColOutQty.Index).Value) < 0 Then
                    If CodeNo_QtyOut = "" Then
                        CodeNo_QtyOut = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyOut = CodeNo_QtyOut + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColEndQty.Index).Value) < 0 Then
                    If CodeNo_QtyEnd = "" Then
                        CodeNo_QtyEnd = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        CodeNo_QtyEnd = CodeNo_QtyEnd + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If
                i = row.Cells(Me.ColPrimaryCodeNo.Index).Value

                If CDec(row.Cells(Me.ColBeginAmount.Index).Value) > 0 And CDec(row.Cells(Me.ColBeginQty.Index).Value) = 0 Then
                    If BeginAmount = "" Then
                        BeginAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        BeginAmount = BeginAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColInAmount.Index).Value) > 0 And CDec(row.Cells(Me.ColInQty.Index).Value) = 0 Then
                    If InAmount = "" Then
                        InAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        InAmount = InAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColOutAmount.Index).Value) > 0 And CDec(row.Cells(Me.ColOutQty.Index).Value) = 0 Then
                    If OutAmount = "" Then
                        OutAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        OutAmount = OutAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If

                If CDec(row.Cells(Me.ColEndAmount.Index).Value) > 0 And CDec(row.Cells(Me.ColEndQty.Index).Value) = 0 Then
                    If EndAmount = "" Then
                        EndAmount = row.Cells(Me.ColCodeNo.Index).Value.ToString
                    Else
                        EndAmount = EndAmount + ", " + row.Cells(Me.ColCodeNo.Index).Value.ToString
                    End If
                End If
            Next

            If CodeUnbalanceAmount <> "" Then
                MsgBox("Unbalance amount for: " + CodeUnbalanceAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeUnbalanceQty <> "" Then
                MsgBox("Unbalance qty for: " + CodeUnbalanceQty, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If WrongEndAmountCalculation <> "" Then
                MsgBox("Wrong end amount calculation for: " + WrongEndAmountCalculation, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If WrongOutAmountCalculation <> "" Then
                MsgBox("Wrong out amount calculation for: " + WrongOutAmountCalculation, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If BeginAmount <> "" Then
                MsgBox("The following code no, have positive beginning amount, but zero beginning qty. Please make sure you are aware of this." + _
                vbCrLf + BeginAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If InAmount <> "" Then
                MsgBox("The following code no, have positive in amount, but zero in qty. Please make sure you are aware of this." + _
                vbCrLf + InAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If OutAmount <> "" Then
                MsgBox("The following code no, have positive out amount, but zero out qty. Please make sure you are aware of this." + _
                vbCrLf + OutAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If EndAmount <> "" Then
                MsgBox("The following code no, have positive end amount, but zero end qty. Please make sure you are aware of this." + _
                vbCrLf + EndAmount, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeNo_QtyBegin <> "" Then
                MsgBox("The following code no, have negative beginning qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyBegin, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If CodeNo_QtyIn <> "" Then
                MsgBox("The following code no, have negative in qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyIn, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If CodeNo_QtyOut <> "" Then
                MsgBox("The following code no, have negative out qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyOut, MsgBoxStyle.Information, "Information")
                Pass = False
            End If
            If CodeNo_QtyEnd <> "" Then
                MsgBox("The following code no, have negative end qty, please make sure you are aware of this." + _
                vbCrLf + CodeNo_QtyEnd, MsgBoxStyle.Information, "Information")
                Pass = False
            End If


            If CodeNoBegin <> "" Then
                MsgBox("The following code no, have positive beginning qty but zero beginning amount, please make sure you are aware of this." + _
                vbCrLf + CodeNoBegin, MsgBoxStyle.Information, "Information")
                Pass = False
            End If

            If CodeNo <> "" Then
                MsgBox("The following code no, have positive in qty but zero in amount, please make sure you are aware of this." + _
                vbCrLf + CodeNo, MsgBoxStyle.Information, "Information")
                Pass = False
            Else
                If CodeNoOut <> "" Then
                    MsgBox("The following code no, have positive out qty but zero out amount, please make sure you are aware of this." + _
                    vbCrLf + CodeNoOut, MsgBoxStyle.Information, "Information")
                    Pass = False
                End If

                If CodeNoEnd <> "" Then
                    MsgBox("The following code no, have positive out qty but zero end amount, please make sure you are aware of this." + _
                    vbCrLf + CodeNoEnd, MsgBoxStyle.Information, "Information")
                    Pass = False
                End If
            End If

            'cek qty beginning, qty in, qty out, qty end, harus lebih besar dari nol, kalo ga salah jadinya

            If Pass = True Then
                MsgBox("Data checked", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' cek yang remarks = "" atau "0" atau "-" harusnya end qty dan qnd amount =0
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ListCodeNo As New List(Of String)
        Dim rowDGV As DataGridViewRow
        Dim Dummy As String = ""

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColRemarks.Index).Value = "" Or rowDGV.Cells(Me.ColRemarks.Index).Value = "0" Or _
                rowDGV.Cells(Me.ColRemarks.Index).Value = "-" Then
                    If CDbl(rowDGV.Cells(Me.ColEndQty.Index).Value) <> 0 Or _
                    CDbl(rowDGV.Cells(Me.ColEndAmount.Index).Value) <> 0 Then
                        ListCodeNo.Add(rowDGV.Cells(Me.ColCodeNo.Index).Value)
                    End If
                End If
            Next

            If ListCodeNo.Count > 0 Then
                For Each CodeNo As String In ListCodeNo
                    If Dummy = "" Then
                        Dummy = CodeNo
                    Else
                        Dummy = Dummy + "," + CodeNo
                    End If
                Next
                MsgBox("Please check: " + Dummy, MsgBoxStyle.Information, "Information")
            Else
                MsgBox("Selesai")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' fungsi ini mengembalikan unit price dari product yang diberikan primarynya
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetUnitPrice(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim ObjClosingItem As New ClsTransaksiMonthlyClosingItem
        Dim InAmount As Decimal = 0
        Dim rowDGVAdjust As DataGridViewRow
        Dim BeginQty, BeginAmount, InQty As Decimal

        Try
            'ambil begin qty, begin amount, in qty, in amount
            ObjClosingItem.Get_BeginAmount_BeginQty_InQty_IAmount(vPrimaryCodeNo, _
            BeginAmount, BeginQty, InQty, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
            Me.dtpPeriod.Value, Me.ObjCD)
            'tambahkan in amount dengan:
            InAmount = ObjClosingItem.GetInAmount_Purchase_SalesReturnBedaPeriode2(vPrimaryCodeNo, _
            Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
            Me.dtpPeriod.Value, Me.TBL_CL, Me.ObjCD, Me.DefaultPrecision)

            '1. bundling
            InAmount = InAmount + Me.GetInAmountFromBundling(vPrimaryCodeNo)
            '2. in adjustment dari stock in dan stock taking
            For Each rowDGVAdjust In Me.dgvItemAdjust.Rows
                If rowDGVAdjust.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = vPrimaryCodeNo Then
                    InAmount += Math.Round(rowDGVAdjust.Cells(Me.ColUnitPriceAdjust.Index).Value * _
                    rowDGVAdjust.Cells(Me.ColQty.Index).Value, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                End If
            Next

            'hitung unit price
            If BeginQty + InQty <> 0 Then
                'hitung unit price
                Return Math.Round((BeginAmount + InAmount) / _
                (BeginQty + InQty), Me.DefaultPrecision, MidpointRounding.AwayFromZero)
            Else
                'tidak ada unit price
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjClosingItem = Nothing
        End Try
    End Function

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim row As DataGridViewRow
        Dim RndNumber As New Random

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            For Each row In Me.dgvItem.Rows
                If CDec(row.Cells(Me.ColEndQty.Index).Value) <> 0 Then
                    row.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(RndNumber.Next(1000, 1000000), 0, MidpointRounding.AwayFromZero))
                End If
            Next

            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'button ini yang digunakan jika ada masalah dengan perhitungan inventory movement, pada "FrmTransaksiMonthlyClosing_DetailCalc",
    'ada coding yang harus diaktifkan pada fungsi form shown
    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim row As DataGridViewRow
        Dim TotalIn As Double = 0
        Dim TotalPurchase, TotalSalesReturn, TotalBundling, TotalInAdjustment As Decimal
        'Dim DTA12 As New DseTransaksiMonthlyClosingTableAdapters.tblcheckTableAdapter
        Dim DetailSR As String = ""

        Try
            For Each row In Me.dgvItem.Rows
                TotalIn += CDec(row.Cells(Me.ColInAmount.Index).Value)
            Next

            MsgBox("Total In: " + Me.ObjCD.DigitGrouping(TotalIn))

            For Each row In Me.dgvItem.Rows
                Dim i As New FrmTransaksiMonthlyClosing_DetailCalc
                i.i = Me
                i.PrimaryCodeNo = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                i.TanggalDari = Me.dtpPeriod.Value.Date
                i.TanggalSampai = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date)

                i.ShowDialog()

                TotalPurchase += i.TotalP
                TotalSalesReturn += i.TotalRtn
                TotalBundling += i.TotalBdl
                TotalInAdjustment += i.TotalIn

                If i.DetailSR <> "" Then
                    If DetailSR = "" Then
                        DetailSR = i.DetailSR
                    Else
                        DetailSR = DetailSR + vbCrLf + i.DetailSR
                    End If
                End If

                i.Dispose()

                'DTA12.Insert(i.PrimaryCodeNo, i.TotalP)
            Next

            MsgBox("Total Purchase: " + Me.ObjCD.DigitGrouping(TotalPurchase) + ", " + vbCrLf + _
            "Total Sales Return: " + Me.ObjCD.DigitGrouping(TotalSalesReturn) + ", " + vbCrLf + _
            "Total Bundling: " + Me.ObjCD.DigitGrouping(TotalBundling) + ", " + vbCrLf + _
            "Total In Adjust: " + Me.ObjCD.DigitGrouping(TotalInAdjustment))
            If DetailSR <> "" Then
                MsgBox(DetailSR)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TabControl1.SelectedIndex = 0 Then
                    'Dim th As New Threading.Thread(AddressOf HitungUlangUnitPrice)
                    'Control.CheckForIllegalCrossThreadCalls = False
                    'th.Start()
                    Me.HitungUlangUnitPrice()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungUlangUnitPrice()
        Try
            Me.ControlBox = False
            Me.gboDetail.Enabled = False
            Dim Th As New Threading.Thread(AddressOf HitungUlangUnitPrice1)
            Control.CheckForIllegalCrossThreadCalls = False
            Th.Start()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungUlangUnitPrice1()
        Dim PrimaryCodeNo As Double = -1
        Dim ij As New FrmTransaksiMonthlyClosing_HitungUnitPrice
        Dim Dummy As String = ""
        Dim ObjClosingItem As New ClsTransaksiMonthlyClosingItem

        'elemen2 yang digunakan untuk pehitungan bundling
        Dim DTABefore As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_beforeTableAdapter
        Dim DTAAfter As New DseTransaksiMonthlyClosingTableAdapters.Tbltrans_bundling_afterTableAdapter
        Dim TBLBefore As New DseTransaksiMonthlyClosing.Tbltrans_bundling_beforeDataTable
        Dim TBLAfter As New DseTransaksiMonthlyClosing.Tbltrans_bundling_afterDataTable

        Dim TBLBefore1 As New DseTransaksiMonthlyClosing.DataTable16DataTable
        Dim DTABefore1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable16TableAdapter
        Dim TBLIn As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable
        Dim DTAIn As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item2TableAdapter

        Dim TBLPurchase As New DseTransaksiMonthlyClosing.DataTable4DataTable
        Dim DTAPurchase As New DseTransaksiMonthlyClosingTableAdapters.DataTable4TableAdapter
        Dim TBLSR As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable
        Dim DTASR As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item9_1TableAdapter

        Try
            Me.ControlBox = False
            Me.gboDetail.Enabled = False

            'isikan data yang akan digunakan untuk perhitungan bundling pada fungsi GetUnitPrice 
            DTABefore.FillBy_TanggalSampaiDari(TBLBefore, _
            Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value)
            DTABefore1.FillBy_ClosingPeriod2(TBLBefore1, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date - TimeSpan.FromDays(1)))
            DTAIn.Fill(TBLIn, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), _
            Me.dtpPeriod.Value.Date)
            DTAPurchase.Fill(TBLPurchase, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.dtpPeriod.Value.Date)
            DTASR.FillBy_Period(TBLSR, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.dtpPeriod.Value.Date)

            ij.ProgressBar1.Value = 0
            ij.ProgressBar1.Maximum = Me.ListPrimaryCodeNo_Bundling.Count
            ij.Show()

            Dummy = ij.Text

            While Me.ListPrimaryCodeNo_Bundling.Count > 0
                ij.ProgressBar1.Value += 1
                ij.Text = Dummy + " (" + ij.ProgressBar1.Value.ToString + " of " + ij.ProgressBar1.Maximum.ToString + ")"
                PrimaryCodeNo = Me.ListPrimaryCodeNo_Bundling.Item(0)
                Me.HitungUnitPrice(PrimaryCodeNo, DTABefore, DTAAfter, TBLBefore, TBLAfter, TBLBefore1, TBLIn, TBLPurchase, TBLSR, ObjClosingItem)
                Me.ListPrimaryCodeNo_Bundling.RemoveAt(0)
            End While

            ij.Hide()
            Me.ControlBox = True
            Me.gboDetail.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            DTABefore.Dispose()
            DTAAfter.Dispose()
            TBLBefore.Dispose()
            TBLAfter.Dispose()

            TBLBefore1.Dispose()
            DTABefore1.Dispose()
            TBLIn.Dispose()
            DTAIn.Dispose()

            TBLPurchase.Dispose()
            DTAPurchase.Dispose()
            TBLSR.Dispose()
            DTASR.Dispose()

            ObjClosingItem = Nothing
        End Try
    End Sub

    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim rowDGV As DataGridViewRow
        Dim DTABundling As New DseTransaksiMonthlyClosingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBLBundling As New DseTransaksiMonthlyClosing.tbltransaksi_stockbundling_item_afterDataTable
        Dim rowBundling As DseTransaksiMonthlyClosing.tbltransaksi_stockbundling_item_afterRow

        Try
            Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemAdjust.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            For Each rowDGV In Me.dgvItemAdjust.Rows
                If CDec(rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value) = 0 Then
                    rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value = "5000.00"
                    Me.HitungUnitPrice(rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value)

                    'hitung ulang juga semua unit price yang menggunakan item ybs sebagai
                    'item bundling before

                    DTABundling.Fill(TBLBundling, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.dtpPeriod.Value.Date, _
                    rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value)

                    For Each rowBundling In TBLBundling.Rows
                        If Me.ListPrimaryCodeNo_Bundling.Contains(rowBundling.Primary_Product_After) = False Then
                            Me.ListPrimaryCodeNo_Bundling.Add(rowBundling.Primary_Product_After)
                        End If
                    Next
                End If
            Next

            Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            MsgBox("Start Update")
            If Me.ListPrimaryCodeNo_Bundling.Count <> 0 Then
                Me.HitungUlangUnitPrice()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTABundling.Dispose()
            TBLBundling.Dispose()
        End Try
    End Sub
End Class