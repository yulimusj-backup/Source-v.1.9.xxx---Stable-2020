Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmTrans_POS_PayoutDropCash_Input
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris
    Public Parent_Form As FrmTrans_POS

    Dim vTipe As ClsMast_PayoutDropCash.EnumTipe = ClsMast_PayoutDropCash.EnumTipe.Payout
    Dim DTA As New DseMast_PayoutDropCashTableAdapters.tblmaster_payout_dropcash2TableAdapter
    Dim TBL As New DseMast_PayoutDropCash.tblmaster_payout_dropcash2DataTable
    Dim ObjPayout As New ClsMast_PayoutDropCash
    Dim ObjSession As New ClsTrans_Session
    Dim Obj As New ClsTrans_POS_PayoutDropCash
    Dim ObjCD As New ClsCekData
    Dim Buffer As String

    Private Sub TipeCBO()
        Try
            If Me.cboPayout_DropCash.SelectedIndex = 0 Then
                vTipe = ClsMast_PayoutDropCash.EnumTipe.Payout
            Else
                vTipe = ClsMast_PayoutDropCash.EnumTipe.Drop_Cash
            End If
            ObjPayout.IsiCBOPayoutDropCash(Me.CboProdCode, Me.TBL, Me.DTA, Me.vTipe)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim Hasil As MsgBoxResult
        Try
            'If Me.dgvItem.Rows.Count > 0 Then
            Hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
            If Hasil = MsgBoxResult.Yes Then
                Obj.SimpanItem(Me.TxtPrimarySession.Text, Me.dgvItem, Me)
                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                Me.dgvItem.Rows.Clear()
                Me.cboPayout_DropCash.DataSource = Nothing
                Me.CboProdCode.DataSource = Nothing
                Me.TxtPrimaryPayOutDropCash.Text = "-1"
                Me.TxtDesc.Text = ""
                Obj.IsiDGV(Me.TxtPrimarySession.Text, Me.dgvItem, Me, Me.ObjCD)
                Me.cboPayout_DropCash.SelectedIndex = 0
                Me.cboPayout_DropCash.Focus()
                'isi cbo payout 
                ObjPayout.IsiCBOPayoutDropCash(Me.CboProdCode, Me.TBL, Me.DTA, Me.vTipe)
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_POS_PayoutDropCash_Input_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.cboPayout_DropCash.SelectedIndex = 0
            Me.cboPayout_DropCash.Focus()
            Me.TxtPrimarySession.Text = Me.Parent_Form.TxtPrimarySession.Text
            'isi cbo payout 
            ObjPayout.IsiCBOPayoutDropCash(Me.CboProdCode, Me.TBL, Me.DTA, Me.vTipe)
            'Load Data Payout drop cash berdasarkan session
            Obj.IsiDGV(Me.TxtPrimarySession.Text, Me.dgvItem, Me, ObjCD)
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboProdCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboProdCode.SelectedIndexChanged
        Dim row As DseMast_PayoutDropCash.tblmaster_payout_dropcash2Row
        Try
            If Me.CboProdCode.SelectedIndex = -1 Then
                Me.TxtPrimaryPayOutDropCash.Text = "-1"
                Me.TxtDesc.Text = ""
            Else
                row = Me.TBL.DefaultView(Me.CboProdCode.SelectedIndex).Row
                Me.TxtPrimaryPayOutDropCash.Text = row.primary
                Me.TxtDesc.Text = row.desc
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPayout_DropCash_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPayout_DropCash.SelectedIndexChanged
        Try
            Me.TipeCBO()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAmount.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAmount.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAmount.Leave
        Try
            ObjCD.Number_Leave(Me.TxtAmount, sender, e, Me.Parent_Form.DefPrec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAmount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAmount.GotFocus
        Try
            If Me.TxtAmount.Enabled = True Then
                If Me.TxtAmount.ReadOnly = False Then
                    Me.TxtAmount.BackColor = Me.BackColorAll
                Else
                    Me.TxtAmount.BackColor = Me.TxtDesc.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAmount.LostFocus
        Try
            If Me.TxtAmount.Enabled = True Then
                If Me.TxtAmount.ReadOnly = False Then
                    Me.TxtAmount.BackColor = Me.TxtPrimaryPayOutDropCash.BackColor
                Else
                    Me.TxtAmount.BackColor = Me.TxtDesc.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.BackColorAll
                Else
                    Me.TxtRemarks.BackColor = Me.TxtDesc.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.TxtPrimaryPayOutDropCash.BackColor
                Else
                    Me.TxtRemarks.BackColor = Me.TxtDesc.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vPrimaryPayoutDropCash As Double) As Double
        Dim RowDGV As DataGridViewRow
        Try
            For Each RowDGV In Me.dgvItem.Rows
                If RowDGV.Cells(Me.ColPrimaryPayout.Index).Value = vPrimaryPayoutDropCash Then
                    Return RowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function ApakahDataItemBenar() As Boolean
        Try
            If Me.TxtPrimaryPayOutDropCash.Text = "-1" Then
                MsgBox("Please choose code", MsgBoxStyle.Information, "Information")
                Me.cboPayout_DropCash.Focus()
                Return False
            End If

            If Me.TxtAmount.Text <= 0 Then
                MsgBox("Please input amount", MsgBoxStyle.Information, "Information")
                Me.TxtAmount.Focus()
                Return False
            End If

            If Me.cboPayout_DropCash.SelectedIndex = 0 Then
                If ObjSession.GetNilaiBalance(Me.TxtPrimarySession.Text, Me.ObjCD) - CDec(Me.TxtAmount.Text) < 0 Then
                    MsgBox("You cannot input payout because balance cannot be negative", MsgBoxStyle.Information, "Information")
                    Me.TxtAmount.Focus()
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInput.Click
        Dim index As Double = -1
        Dim rowDGV As DataGridViewRow
        Try
            If ApakahDataItemBenar() = True Then
                'cek apakah sudah ada item jenis pekerjaan sama yang masuk
                index = Me.GetIndexItem(Me.TxtPrimaryPayOutDropCash.Text)
                If index < 0 Then
                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColNo.Index).Value = ""
                        If Me.cboPayout_DropCash.SelectedIndex = 0 Then
                            .Cells(Me.ColTipe.Index).Value = "Payout"
                        Else
                            .Cells(Me.ColTipe.Index).Value = "Drop Cash"
                        End If
                        .Cells(Me.ColCode.Index).Value = ObjPayout.GetCode(Me.TxtPrimaryPayOutDropCash.Text)
                        .Cells(Me.ColDesc.Index).Value = ObjPayout.GetDesc(Me.TxtPrimaryPayOutDropCash.Text)
                        .Cells(Me.ColAmount.Index).Value = ObjCD.DigitGrouping(Me.TxtAmount.Text, Me.Parent_Form.DefPrec, Me.Parent_Form.RoundingTipe)
                        .Cells(Me.ColRemarks.Index).Value = Me.TxtRemarks.Text
                        .Cells(Me.ColPrimarySession.Index).Value = Me.TxtPrimarySession.Text
                        '-------
                        .Cells(Me.ColPrimaryAutoNumber.Index).Value = "-1"
                        .Cells(Me.ColPrimaryPayout.Index).Value = Me.TxtPrimaryPayOutDropCash.Text
                    End With
                    'set number
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    index = Me.dgvItem.Rows.Count - 1
                Else
                    'Tambahkan Qty dan Hitung kembali
                    rowDGV = Me.dgvItem.Rows(index)
                    With rowDGV
                        .Cells(Me.ColAmount.Index).Value = ObjCD.DigitGrouping(CDec(.Cells(Me.ColAmount.Index).Value) + CDec(Me.TxtAmount.Text), Me.Parent_Form.DefPrec, Me.Parent_Form.RoundingTipe)
                        .Cells(Me.ColRemarks.Index).Value += Me.TxtRemarks.Text
                    End With
                End If
                Me.KosongkanFieldItem()
                Me.cboPayout_DropCash.Focus()

                '-----Scrolling DGV-----
                If Me.dgvItem.FirstDisplayedScrollingRowIndex < index Then
                    While Me.dgvItem.Rows(index).Displayed = False _
                    Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                    End While
                Else
                    While Me.dgvItem.Rows(index).Displayed = False
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = _
                        Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                    End While
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanFieldItem()
        Try
            Me.TxtAmount.Text = "0"
            Me.TxtRemarks.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColAmount.Index Or e.ColumnIndex = Me.ColRemarks.Index Then
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    'Cek string keterangan
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If ObjCD.CekDataKosong(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            MsgBox("Data keterangan" _
                            + " tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    If e.ColumnIndex = Me.ColAmount.Index Then
                        If ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please enter positif number", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.Parent_Form.DefPrec, Me.Parent_Form.RoundingTipe)
                            Exit Sub
                        End If
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.Parent_Form.DefPrec, Me.Parent_Form.RoundingTipe)
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

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            If e.Row.Cells(Me.ColTipe.Index).Value = "Drop Cash" Then
                If CDec(ObjSession.GetNilaiBalance(Me.TxtPrimarySession.Text, Me.ObjCD)) - CDec(e.Row.Cells(Me.ColAmount.Index).Value) < 0 Then
                    MsgBox("You cannot delete drop cash because balance cannot be negative", MsgBoxStyle.Information, "Information")
                    Me.TxtAmount.Focus()
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class