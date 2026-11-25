Option Explicit On
Option Strict Off

Public Class FrmTransaksiModifyControl
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim Buffer As String = ""
    Dim ObjSC As New ClsTransaksiStockCard

    Public BackColorAll As Color = Color.LightGoldenrodYellow

    Private Sub btnFindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCodeNo.Click
        Try
            Me.FindCodeNo(True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FindCodeNo(ByVal vDariBtn As Boolean)
        Dim i As New Core_Module.frmCariMasterProduct

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.TxtQty.Text = ObjCD.DigitGrouping(Me.ObjProduct.GetJmlByPrimary(Me.TxtPrimaryCodeNo.Text))
                Me.LblUnit.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                Me.LblUnit2.Text = Me.LblUnit.Text
                Me.chk_Location_Ori.Checked = Me.ObjProduct.ApakahLocationControlled(Me.TxtPrimaryCodeNo.Text)
                Me.chkBatch_Ori.Checked = Me.ObjProduct.ApakahBatchNoControlled(Me.TxtPrimaryCodeNo.Text)
                Me.chkED_Ori.Checked = Me.ObjProduct.ApakahExpiryDateControlled(Me.TxtPrimaryCodeNo.Text)

                If vDariBtn = True Then
                    Me.chkBatch.Focus()
                End If
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Text = ""
                Me.TxtQty.Text = "0"
                Me.LblUnit.Text = "Unit"
                Me.LblUnit2.Text = "Unit"
                Me.chkED_Ori.Checked = False
                Me.chkBatch_Ori.Checked = False
                Me.chk_Location_Ori.Checked = False
                Me.TxtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtProductCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim i As New Core_Module.frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNo.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.TxtQty.Text = ObjCD.DigitGrouping(Me.ObjProduct.GetJmlByPrimary(Me.TxtPrimaryCodeNo.Text))
                    Me.LblUnit.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                    Me.LblUnit2.Text = Me.LblUnit.Text
                    Me.chk_Location_Ori.Checked = Me.ObjProduct.ApakahLocationControlled(Me.TxtPrimaryCodeNo.Text)
                    Me.chkBatch_Ori.Checked = Me.ObjProduct.ApakahBatchNoControlled(Me.TxtPrimaryCodeNo.Text)
                    Me.chkED_Ori.Checked = Me.ObjProduct.ApakahExpiryDateControlled(Me.TxtPrimaryCodeNo.Text)
                    Me.chkBatch.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    Me.btnFindCodeNo.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtProductCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim i As New Core_Module.frmCariMasterProduct

        Try
            If Me.TxtCodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtCodeNo.Text) = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.TxtQty.Text = ObjCD.DigitGrouping(Me.ObjProduct.GetJmlByPrimary(Me.TxtPrimaryCodeNo.Text))
                    Me.LblUnit.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                    Me.LblUnit2.Text = Me.LblUnit.Text
                    Me.chk_Location_Ori.Checked = Me.ObjProduct.ApakahLocationControlled(Me.TxtPrimaryCodeNo.Text)
                    Me.chkBatch_Ori.Checked = Me.ObjProduct.ApakahBatchNoControlled(Me.TxtPrimaryCodeNo.Text)
                    Me.chkED_Ori.Checked = Me.ObjProduct.ApakahExpiryDateControlled(Me.TxtPrimaryCodeNo.Text)
                Else
                    Me.FindCodeNo(False)
                End If
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Text = ""
                Me.TxtQty.Text = "0"
                Me.LblUnit.Text = "Unit"
                Me.chkED_Ori.Checked = False
                Me.chkBatch_Ori.Checked = False
                Me.chk_Location_Ori.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        Try
            Me.TxtCodeNo.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        Try
            Me.TxtCodeNo.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBatch.GotFocus
        Try
            Me.TxtBatch.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBatch.LostFocus
        Try
            Me.TxtBatch.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationCode.GotFocus
        Try
            Me.TxtLocationCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationCode.LostFocus
        Try
            Me.TxtLocationCode.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkBatch_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBatch.CheckedChanged
        Try
            If Me.chkBatch.Checked = True Then
                Me.chkExpiryDate.Enabled = True
                Me.ColBatchNo.Visible = True
                Me.TxtBatch.Enabled = True
            Else
                Me.TxtBatch.Enabled = False
                Me.TxtBatch.Text = ""
                Me.dtpED.Enabled = False
                Me.chkExpiryDate.Enabled = False
                Me.chkExpiryDate.Checked = False
                Me.ColBatchNo.Visible = False
                Me.ColExpiryDate.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkExpiryDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkExpiryDate.CheckedChanged
        Try
            If Me.chkExpiryDate.Checked = True Then
                Me.ColExpiryDate.Visible = True
                Me.dtpED.Enabled = True
            Else
                Me.dtpED.Enabled = False
                Me.ColExpiryDate.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkLocation_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLocation.CheckedChanged
        Try
            If Me.chkLocation.Checked = True Then
                Me.ColLocation.Visible = True
                Me.btnFindLocation.Enabled = True
                Me.TxtLocationCode.Enabled = True
            Else
                Me.ColLocation.Visible = False
                Me.btnFindLocation.Enabled = False
                Me.TxtLocationCode.Enabled = False
                Me.TxtPrimaryLocation.Text = "-1"
                Me.TxtLocationCode.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindLocation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindLocation.Click
        Dim i As New Core_Module.FrmCariMasterStorageArea

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtLocationCode.Text <> "" Then
                i.txtCari.Text = Me.TxtLocationCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.TxtPrimaryLocation.Text = i.PrimaryStorageArea
                Me.TxtLocationCode.Text = Me.ObjLocation.GetStorageAreaCode(Me.TxtPrimaryLocation.Text)
                Me.TxtQtyPerBatch.Focus()
            Else
                Me.TxtLocationCode.Clear()
                Me.TxtPrimaryLocation.Text = "-1"
                Me.TxtLocationCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtLocationCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocationCode.KeyDown
        Dim i As New Core_Module.FrmCariMasterStorageArea
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjLocation.GetJmlBy_LocationCodeLIKE(Me.TxtLocationCode.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryLocation.Text = Me.ObjLocation.GetPrimary(Me.TxtLocationCode.Text)
                    Me.TxtLocationCode.Text = Me.ObjLocation.GetStorageAreaCode(Me.TxtPrimaryLocation.Text)
                    Me.TxtQtyPerBatch.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    Me.btnFindLocation.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtLocationCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationCode.Leave
        Dim i As New Core_Module.FrmCariMasterStorageArea

        Try
            If Me.TxtLocationCode.Text <> "" Then
                If Me.ObjLocation.GetJmlBy_LocationCodeLIKE(Me.TxtLocationCode.Text) = 1 Then
                    Me.TxtPrimaryLocation.Text = Me.ObjLocation.GetPrimary(Me.TxtLocationCode.Text)
                    Me.TxtLocationCode.Text = Me.ObjLocation.GetStorageAreaCode(Me.TxtPrimaryLocation.Text)
                Else
                    Me.FindLocation(False)
                End If
            Else
                Me.TxtLocationCode.Clear()
                Me.TxtPrimaryLocation.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub FindLocation(ByVal vDariBtn As Boolean)
        Dim i As New Core_Module.FrmCariMasterStorageArea

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtLocationCode.Text <> "" Then
                i.txtCari.Text = Me.TxtLocationCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.TxtPrimaryLocation.Text = i.PrimaryStorageArea
                Me.TxtLocationCode.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryLocation.Text)

                If vDariBtn = True Then
                    Me.TxtQtyPerBatch.Focus()
                End If
            Else
                Me.TxtLocationCode.Clear()
                Me.TxtPrimaryLocation.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInput.Click
        Dim row As DataGridViewRow
        Dim Index As Integer = 0

        Try
            If Me.ApakahDataBenar = True Then
                If Me.chkExpiryDate.Checked = True Then
                    Index = Me.GetIndex(Me.TxtBatch.Text, Me.dtpED.Value, Me.TxtPrimaryLocation.Text)
                Else
                    Index = Me.GetIndex(Me.TxtBatch.Text, #1/1/1950#, Me.TxtPrimaryLocation.Text)
                End If

                If Index = -1 Then
                    'tambahkan baris baru
                    Me.dgvItem.Rows.Add()
                    row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With row
                        .Cells(Me.ColBatchNo.Index).Value = Me.TxtBatch.Text
                        If Me.chkExpiryDate.Checked = True Then
                            .Cells(Me.ColExpiryDate.Index).Value = Format(Me.dtpED.Value.Date, "dd MMM yyyy")
                        Else
                            .Cells(Me.ColExpiryDate.Index).Value = Format(#1/1/1950#, "dd MMM yyyy")
                        End If
                        .Cells(Me.ColLocation.Index).Value = Me.TxtLocationCode.Text
                        .Cells(Me.ColPrimaryLocation.Index).Value = Me.TxtPrimaryLocation.Text
                        .Cells(Me.ColQty.Index).Value = Me.TxtQtyPerBatch.Text
                        .Cells(Me.ColUnit.Index).Value = Me.LblUnit.Text
                    End With
                Else
                    'tambahkan qty saja
                    row = Me.dgvItem.Rows(Index)
                    row.Cells(Me.ColQty.Index).Value = _
                    Me.ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQty.Index).Value) + CDbl(Me.TxtQtyPerBatch.Text))
                End If

                'kosongkan gbo input
                Me.TxtBatch.Text = ""
                Me.TxtLocationCode.Text = ""
                Me.TxtPrimaryLocation.Text = "-1"
                Me.dtpED.Value = Today
                Me.TxtQtyPerBatch.Text = "0"

                'fokus untuk input data selanjutnya
                If Me.chkBatch.Checked = True Then
                    Me.TxtBatch.Focus()
                ElseIf Me.chkLocation.Focus Then
                    Me.dtpED.Focus()
                End If

                'semua chk menjadi disable, kode product juga tidak bisa diubah lagi
                Me.chkExpiryDate.Enabled = False
                Me.chkLocation.Enabled = False
                Me.chkBatch.Enabled = False
                Me.TxtCodeNo.Enabled = False
                Me.btnFindCodeNo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim ED As String = ""
        Dim Index As Integer = 0

        Try
            'data salah jika:
            '0. product code tidak diisikan
            If Me.TxtCodeNo.Text = "" Then
                MsgBox("Please specify code number first", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Focus()
                Return False
            End If
            '1. batch number controlled tapi tidak menspesifikasikan batch yang digunakan
            If Me.chkBatch.Checked = True Then
                If Me.TxtBatch.Text.Trim = "" Then
                    MsgBox("Please specify batch number", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            '2. jika location controlled dan data lokasi kosong
            If Me.chkLocation.Checked = True Then
                If Me.TxtLocationCode.Text = "" Then
                    MsgBox("Please specify location", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            '3. jika tidak ada yang kontrol, maka tidak perlu input
            If Me.chkLocation.Checked = False And Me.chkExpiryDate.Checked = False And Me.chkBatch.Checked = False Then
                MsgBox("For no control code number, you cannot specify qty data.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            '4. jika expiry data controlled dan expiry date yang dispesifikasi mundur dibandingkan dengan hari ini
            'maka berikan berikan peringatan
            If Me.chkExpiryDate.Checked = True Then
                If Me.dtpED.Value.Date < Today Then
                    If MsgBox("The expiry date value for batch number " + Me.TxtBatch.Text + _
                    " is smaller than current date, is this data correct?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
                    End If
                End If
                '5. jika expiry date controlled maka harus dicheck apakah ada batch number yang serupa expiry datenya berbeda
                Index = Me.GetIndex_BatchNumberSama_ED_Beda(Me.TxtBatch.Text, Me.dtpED.Value.Date)
                If Index <> -1 Then
                    ED = Me.dgvItem.Rows(Index).Cells(Me.ColExpiryDate.Index).Value.ToString
                    MsgBox("The batch number " + Me.TxtBatch.Text + " is already have expiry date on " + ED + _
                    ". One batch number can only be reference by one unique expiry date. Please fix data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            '5. qty tidak boleh kosong
            If Me.TxtQtyPerBatch.Text = "0" Then
                MsgBox("Please specify positive quantity", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'sampe sini data benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndex_BatchNumberSama_ED_Beda(ByVal vBatchNumber As String, ByVal vED As Date) As Double
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColBatchNo.Index).Value.ToString = vBatchNumber Then
                    If vED <> CDate(row.Cells(Me.ColExpiryDate.Index).Value.ToString) Then
                        Return row.Index
                    End If
                End If
            Next

            'kalo tidak ketemu maka kembalikan nilai -1
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndex(ByVal vBatchNumber As String, ByVal vED As Date, ByVal vPrimaryLocation As Double) As Double
        Dim row As DataGridViewRow

        Try
            If Me.chkBatch.Checked = True And Me.chkExpiryDate.Checked = False And Me.chkLocation.Checked = False Then
                'cari yang batch numbernya sama
                For Each row In Me.dgvItem.Rows
                    If row.Cells(Me.ColBatchNo.Index).Value.ToString = vBatchNumber Then
                        Return row.Index
                    End If
                Next
            ElseIf Me.chkBatch.Checked = True And Me.chkExpiryDate.Checked = True Then
                'cari yang batch number dan expiry datenya sama
                For Each row In Me.dgvItem.Rows
                    If row.Cells(Me.ColBatchNo.Index).Value.ToString = vBatchNumber And _
                    CDate(row.Cells(Me.ColExpiryDate.Index).Value) = vED Then
                        Return row.Index
                    End If
                Next
            ElseIf Me.chkBatch.Checked = True And Me.chkExpiryDate.Checked = True And Me.chkLocation.Checked = True Then
                'cari yang batch number, expiry date, dan location sama
                For Each row In Me.dgvItem.Rows
                    If row.Cells(Me.ColBatchNo.Index).Value.ToString = vBatchNumber And _
                    CDate(row.Cells(Me.ColExpiryDate.Index).Value) = vED And _
                    row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
                        Return row.Index
                    End If
                Next
            End If

            'kalo tidak ketemu maka kembalikan nilai -1
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtQtyPerBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyPerBatch.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQtyPerBatch, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyPerBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyPerBatch.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQtyPerBatch, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang boleh diubah hanya kolom qty saja
            If e.ColumnIndex = Me.ColQty.Index Then
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            Else
                'pastikan nilainya positif
                If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please specify positive number for quantity data", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
                If CDbl(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                    MsgBox("Please specify positive number for quantity data", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'digit grouping
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.chkBatch.Enabled = True
                Me.chkLocation.Enabled = True
                Me.chkBatch.Enabled = True
                Me.TxtCodeNo.Enabled = True
                Me.btnFindCodeNo.Enabled = True

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.ApakahDataBenar2 = True Then
                '1. konfirmasi perubahan, terangkan resiko yang harus di hadapi, dan sarankan untuk melakukan back up data
                If MsgBox("Your are going to change the product control. Make sure nobody else is using the software. Close all other form." + _
                vbCrLf + vbCrLf + "Please be aware of the risk:" + vbCrLf + _
                "1. You cannot edit/delete all stock related transaction created before this modification containing this product." + vbCrLf + _
                "2. All booking inventory transaction that will be inconsistent with the new control state will be deleted." + vbCrLf + _
                "3. Some delivery order that will be inconsistent with the new control state must be fixed." + vbCrLf + vbCrLf + _
                "We recommend you create data back up before proceeding. Proceed anyway?", _
                MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    '2. update perubahan control state pada product
                    If Me.ObjProduct.ApakahBatchNoControlled(Me.TxtPrimaryCodeNo.Text) <> Me.chkBatch.Checked Then
                        Me.ObjProduct.SetBatchNumberControlled(Me.TxtPrimaryCodeNo.Text, Me.chkBatch.Checked)
                    End If
                    If Me.ObjProduct.ApakahLocationControlled(Me.TxtPrimaryCodeNo.Text) <> Me.chkLocation.Checked Then
                        Me.ObjProduct.SetLocationControl(Me.TxtPrimaryCodeNo.Text, Me.chkLocation.Checked)
                    End If
                    If Me.ObjProduct.ApakahExpiryDateControlled(Me.TxtPrimaryCodeNo.Text) <> Me.chkExpiryDate.Checked Then
                        Me.ObjProduct.SetExpiryDateControlled(Me.TxtPrimaryCodeNo.Text, Me.chkExpiryDate.Checked)
                    End If

                    '3. set tanggal perubahan pada master product
                    Me.ObjProduct.SetTanggalMinimum(Me.TxtPrimaryCodeNo.Text, Now)

                    '4. Hapus semua product code yang ada didalam tblmaster product qty
                    Me.ObjProduct.DeleteProductQty(Me.TxtPrimaryCodeNo.Text)

                    '5. Masukkin data baru
                    Me.ObjProduct.InsertProductQty(Me.TxtPrimaryCodeNo.Text, Me.dgvItem, Me.ColBatchNo.Index, Me.ColExpiryDate.Index, _
                    Me.ColQty.Index, Me.ColPrimaryLocation.Index, Me.ObjSC.GetStock(Me.TxtPrimaryCodeNo.Text, #1/1/2500#))

                    '6. jika ada booking inventory yang menggunakan tipe kontrol yang diberikan, maka data harus dihapuskan, dan 
                    'kirimkan email konfirmasi

                    '5. Konfirmasi selesai
                    MsgBox("Data modified", MsgBoxStyle.Information, "Information")

                    Me.Close()
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJml() As Double
        Dim row As DataGridViewRow
        Dim retval As Double = 0

        Try
            For Each row In Me.dgvItem.Rows
                retval += CDbl(row.Cells(Me.ColQty.Index).Value)
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar2() As Boolean
        Dim JmlAcuan As Double = 0

        Try
            'data salah jika
            '0. product code tidak diisikan
            If Me.TxtCodeNo.Text = "" Then
                MsgBox("Please specify code number first", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Focus()
                Return False
            End If

            '1. Total qty sama dengan total qty yang tercatat dalam database(bukan yang tercatat pada txtbox)
            If Me.chkBatch.Checked = True Or Me.chkExpiryDate.Checked = True Or Me.chkLocation.Checked = True Then
                JmlAcuan = Me.ObjProduct.GetJmlByPrimary(Me.TxtPrimaryCodeNo.Text)
                If JmlAcuan <> Me.GetJml Then
                    MsgBox("The total qty must be " + ObjCD.DigitGrouping(JmlAcuan) + _
                    " " + Me.LblUnit.Text, MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '2. Jika tipe kontrol database sama dengan tipe kontrol yang diubah
            If Me.ObjProduct.ApakahLocationControlled(Me.TxtPrimaryCodeNo.Text) = Me.chkLocation.Checked And Me.chkBatch.Checked = _
            Me.ObjProduct.ApakahBatchNoControlled(Me.TxtPrimaryCodeNo.Text) And _
            Me.ObjProduct.ApakahExpiryDateControlled(Me.TxtPrimaryCodeNo.Text) = Me.chkExpiryDate.Checked Then
                MsgBox("The original control state and the modified control state for this product is similar, you cannot continue this process.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. Tanggal komputer harus sama dengan tanggal hari ini
            If MsgBox("Is today's date " + Format(Today, "dd MMM yyyy") + "?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                MsgBox("Please make sure the computer date is valid", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'sampe sini data sudah benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtQtyPerBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyPerBatch.GotFocus
        Try
            Me.TxtQtyPerBatch.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyPerBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyPerBatch.LostFocus
        Try
            Me.TxtQtyPerBatch.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub txtBatchNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBatch.Leave
        Try
            Me.TxtBatch.Text = Me.TxtBatch.Text.Trim
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class