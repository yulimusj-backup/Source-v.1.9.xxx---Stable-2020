Option Explicit On
Option Strict Off

Public Class FrmTransaksiHistoricalPurchase
    Public UserCode As String = ""
    Private obj As New ClsTransaksiHistoricalPurchase
    Private objItem As New ClsTransaksiHostoricalPurchaseItem
    Private Baru As Boolean
    Private ObjCD As New Core_Module.ClsCekData
    Private Buffer As String = ""
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim ObjCur As New Core_Module.clsMasterCurrency

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub txtCurrencyCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.GotFocus
        Me.txtCurrencyCode.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub txtCurrencyCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.LostFocus
        Me.txtCurrencyCode.BackColor = Me.TxtPrimaryStockIn.BackColor
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Me.TxtRemark.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Me.TxtRemark.BackColor = Me.TxtPrimaryStockIn.BackColor
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        Me.TxtCodeNo.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        Me.TxtCodeNo.BackColor = Me.TxtPrimaryStockIn.BackColor
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.GotFocus
        Me.TxtQtyperBatch.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatch.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatch, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatch.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatch, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then
                    Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatch.Text, _
                    Me.ObjProd.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.LostFocus
        Me.TxtQtyperBatch.BackColor = Me.TxtPrimaryStockIn.BackColor
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        Me.TxtRemarkItem.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        Me.TxtRemarkItem.BackColor = Me.TxtPrimaryStockIn.BackColor
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        If e.ColumnIndex = Me.ColQty.Index Or e.ColumnIndex = Me.ColRemark.Index Or e.ColumnIndex = Me.ColAmount.Index Then
            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColQty.Index Or e.ColumnIndex = Me.ColAmount.Index Then
                'pastikan diisi
                If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Try
                    Else
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "0" Then
                            If e.ColumnIndex = Me.ColQty.Index Then
                                MsgBox("Please don't insert 0 qty.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Try
                            Else
                                MsgBox("Please don't insert 0 Amount.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Try
                            End If
                        End If
                    End If
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Try
                End If
                'pastikan positif
                If ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Try
                End If

                'digit grouping
                If e.ColumnIndex = Me.ColQty.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, _
                    Me.ObjProd.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                ElseIf e.ColumnIndex = Me.ColAmount.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.ObjCur.GetDecimal_DefaultCurrency)
                End If

                Me.HitungTotal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Transaction No")
            KoleksiTextBox.Add(Me.TxtStockInNo)
            KoleksiNamaData.Add("Currency Code")
            KoleksiTextBox.Add(Me.txtCurrencyCode)
            If ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Focus()
                Return False
            End If
            Dim tglDari As Date = ObjCD.GetFirstDateofTheMonth(Me.dtpForm.Value)
            Dim Tglsampai As Date = ObjCD.GetLasDateOfTheMonth(Me.dtpThrough.Value)
            If tglDari > Tglsampai Then
                MsgBox("Date of 'From' can not bigger than date of 'Thru'.", MsgBoxStyle.Information, "Information")
                Me.dtpForm.Focus()
                Return False
            End If

            Dim ListString As New List(Of String)
            If obj.ApakahAdaTglOverSlap(ListString, tglDari, Tglsampai, Me.TxtPrimaryStockIn.Text) = False Then
                Dim i As String
                Dim NO As String = ""
                For Each i In ListString
                    NO = i + ". "
                Next
                ListString = Nothing
                MsgBox("The transaction overlap with another transaction. There are: " + NO + "Please change start date or end date.", MsgBoxStyle.Information, "Information")
                Return False
            Else
                ListString = Nothing
            End If
            If Baru = False Then
                Dim OldTimeStamp As Date = Me.obj.GetTimeStamp(CDbl(Me.TxtPrimaryStockIn.Text))
                If Me.obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.obj.CariDetail(Me.dgvItem, Me.obj.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
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
    Private Sub KosongkanItem()
        Try
            ObjCD.KosongkanControl(Me.gboItem.Controls, True, True, False)
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnits.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.txtAmount.Text = "0"
            Me.TxtQtyperBatch.Text = "0"
            Me.TxtRemarkItem.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.obj
                .PrimaryCurrency = Me.txtPrimaryCurrency.Text
                .Amount = Me.TxtTotal.Text
                If Baru = True Then
                    .CreatedBy = Me.UserCode
                    .CreatedDate = Today
                End If
                .EndPeriod = ObjCD.GetLasDateOfTheMonth(Me.dtpThrough.Value)
                .ModifiedBy = Me.UserCode
                .ModifiedDate = Today
                .NoTransaksi = Me.TxtStockInNo.Text
                .TglTransaksi = Me.dtpDate.Value
                .Remarks = Me.TxtRemark.Text
                .StartPeriod = ObjCD.GetFirstDateofTheMonth(Me.dtpForm.Value)
                '    .TimeStamp = Now
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.obj
                Me.txtPrimaryCurrency.Text = .PrimaryCurrency
                Dim objCur As New Core_Module.clsMasterCurrency
                Me.txtCurrencyCode.Text = objCur.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                objCur = Nothing
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(.Amount, Me.ObjCur.GetDecimal_DefaultCurrency)
                Me.dtpThrough.Value = Me.ObjCD.GetFirstDateofTheMonth(.EndPeriod)
                Me.TxtStockInNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TglTransaksi
                Me.dtpForm.Value = .StartPeriod
                Me.TxtRemark.Text = .Remarks

                Me.TxtPrimaryStockIn.Text = .Primary
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, False, False)
            ObjCD.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, False, False)
            ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.btnFindCodeNo2.Enabled = True

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.KosongkanItem()
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.obj.SimpanBaru(Me.dgvItem, objItem, Me.UserCode) Then
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
                        If Me.obj.SimpanPerubahan(Me.dgvItem, objItem, Me.UserCode) Then
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

    Private Sub HitungTotal()
        Try
            Dim row As DataGridViewRow
            Dim Total As Decimal = 0
            For Each row In Me.dgvItem.Rows
                If String.IsNullOrEmpty(Me.TxtTotal.Text) = False Then
                    Total = CDec(Total) + CDec(row.Cells(Me.ColAmount.Index).Value)
                Else
                    Me.TxtTotal.Text = "0"
                End If
            Next
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Total, Me.ObjCur.GetDecimal_DefaultCurrency)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField(Optional ByVal HapusDGV As Boolean = True)
        Try
            Me.KosongkanItem()
            ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimaryStockIn.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtPrimaryUnits.Text = "-1"
            Me.TxtTotal.Text = "0"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            If HapusDGV = True Then
                Me.dgvItem.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 And Me.dtpDate.Value.Date = Today And _
            Me.dtpForm.Value.Date = Today And Me.dtpThrough.Value.Date = Today And Me.txtCurrencyCode.Text = "" And _
            Me.TxtStockInNo.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.obj.CariDetail(Me.dgvItem, Me.TxtPrimaryStockIn.Text)
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

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtStockInNo.Text <> "" Then
                If Me.objItem.ApakahBolehHapus = False Then
                    MsgBox("You cannot delete this data, because this will cause negative quantity data.", _
                    MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.obj.Hapus(Me.objItem) = True Then
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
    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryStockIn.BackColor, False)

            ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, False, False)
            ObjCD.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, False, False)
            ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            Me.dgvItem.Columns(Me.ColCodeNo.Index).ReadOnly = True
            Me.dgvItem.Columns(Me.ColDescription.Index).ReadOnly = True
            Me.dgvItem.Columns(Me.ColNo.Index).ReadOnly = True
            Me.dgvItem.Columns(Me.ColUnit.Index).ReadOnly = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.Baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtStockInNo.Focus()
            Me.obj.ResetObject(Me.dgvItem)
            Me.TxtStockInNo.Text = Me.obj.GetNoBaru(Me.dtpDate.Value)
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
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtStockInNo.Text <> "" Then
                Me.Baru = False
                Me.AktifkanDetail()
                ' Me.dgvItem.Focus()
                'nilai min juga harus diatas tanggal pengkuncian data terakhir
                'msgbox("belum")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariTransaksiHistoricalPurchase
        Try
            ij.ShowDialog()
            If Not ij.PrimaryTransaksiHistory = -1 Then
                If Me.obj.CariDetail(Me.dgvItem, ij.PrimaryTransaksiHistory) = True Then
                    Me.KosongkanField(False)
                    Me.SalinKeField()
                    Me.HitungTotal()
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

    Private Sub FrmTransaksiHistoricalPurchase_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New Core_Module.FrmCariMasterCurrency
        Dim objCurrency As New Core_Module.clsMasterCurrency
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtCurrencyCode.Text <> "" Then
                i.txtCari.Text = Me.txtCurrencyCode.Text
            End If
            i.ShowDialog()
            If i.PrimaryCurrency > -1 Then
                Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                Me.txtCurrencyCode.Text = objCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.TxtRemark.Focus()
            Else
                Me.txtPrimaryCurrency.Text = "-1"
                Me.txtCurrencyCode.Clear()
                Me.txtCurrencyCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCurrency = Nothing
        End Try
    End Sub

    Private Sub txtCurrencyCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Dim objCur As New Core_Module.clsMasterCurrency
        Try
            If e.KeyCode = Keys.Enter Then
                Dim primary As Integer = objCur.GetPrimary(Me.txtCurrencyCode.Text)
                If primary = -1 Then
                    Me.btnFindCurrency.PerformClick()
                Else
                    Me.txtPrimaryCurrency.Text = primary
                    Me.TxtRemark.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCur = Nothing
        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Dim objCur As New Core_Module.clsMasterCurrency
        Try
            If Me.txtCurrencyCode.Text <> "" Then
                Dim primary As Integer = objCur.GetPrimary(Me.txtCurrencyCode.Text)
                If primary = -1 Then
                    Me.btnFindCurrency.PerformClick()
                Else
                    Me.txtPrimaryCurrency.Text = primary
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCur = Nothing
        End Try
    End Sub

    Private Sub BtnFindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim i As New Core_Module.frmCariMasterProduct
        Dim objProduct As New Core_Module.clsMasterProduct
        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > -1 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = objProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.txtDesc.Text = objProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.txtPrimaryCategory.Text = objProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                Me.lblUnits.Text = objProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                Me.TxtQtyperBatch.Focus()
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.lblUnits.Text = "Units"
                Me.txtDesc.Clear()
                Me.TxtQtyperBatch.Text = "0"
                Me.TxtRemark.Clear()
                Me.TxtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim objProduct As New Core_Module.clsMasterProduct
        Try
            If e.KeyCode = Keys.Enter Then
                Dim primary As Integer = objProduct.GetPrimary(Me.TxtCodeNo.Text)
                If primary = -1 Then
                    Me.txtDesc.Text = ""
                    Me.BtnFindCodeNo.PerformClick()
                Else
                    Me.TxtPrimaryCodeNo.Text = primary
                    Me.TxtCodeNo.Text = objProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = objProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.txtPrimaryCategory.Text = objProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.lblUnits.Text = objProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objProduct = Nothing
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim objProduct As New Core_Module.clsMasterProduct
        Try
            If Me.TxtCodeNo.Text <> "" Then
                Dim primary As Integer = objProduct.GetPrimary(Me.TxtCodeNo.Text)
                If primary = -1 Then
                    Me.txtDesc.Text = ""
                    Me.BtnFindCodeNo.PerformClick()
                Else
                    Me.TxtPrimaryCodeNo.Text = primary
                    Me.TxtCodeNo.Text = objProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = objProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.txtPrimaryCategory.Text = objProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.lblUnits.Text = objProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                End If
            Else
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objProduct = Nothing
        End Try
    End Sub
    Private Function ApakahDataItemBenar() As Boolean
        Try
            'data item salah bila:
            '1. product code yang diisikan salah
            If Me.txtDesc.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Text = ""
                Me.TxtCodeNo.Focus()
                Return False
            End If
            '2. nilai qty <= 0
            If ObjCD.CekBilanganPositif(Me.TxtQtyperBatch.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatch.Focus()
                Return False
            End If
            '3. Cek Amount
            If ObjCD.CekBilanganPositif(Me.txtAmount.Text) = False Then
                MsgBox("Please enter a positive number for Amount data.", MsgBoxStyle.Information, _
                "Information")
                Me.txtAmount.Focus()
                Return False
            End If

            Dim rowDGV As DataGridViewRow
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text And _
                rowDGV.Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnits.Text Then
                    rowDGV.Cells(Me.ColQty.Index).Value = CDbl(rowDGV.Cells(Me.ColQty.Index).Value) + CDbl(Me.TxtQtyperBatch.Text)
                    rowDGV.Cells(Me.ColAmount.Index).Value = CDbl(rowDGV.Cells(Me.ColAmount.Index).Value) + CDbl(Me.txtAmount.Text)
                    Me.KosongkanItem()
                    Me.TxtCodeNo.Focus()
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Try
            If Me.ApakahDataItemBenar = True Then
                'tambahkan baris baru
                Me.dgvItem.Rows.Add("", Me.TxtCodeNo.Text, Me.txtDesc.Text, Me.TxtQtyperBatch.Text, _
                Me.lblUnits.Text, Me.txtAmount.Text, Me.TxtRemarkItem.Text, _
                Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text)
                'set ulang nomor
                ObjCD.SetNumber(Me.dgvItem)

                'kosongkan gbo item
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
                'format dgv
                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While

                Me.HitungTotal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAmount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmount.GotFocus
        Me.txtAmount.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub txtAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtAmount, sender, e)
    End Sub

    Private Sub txtAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtAmount, sender, e)
    End Sub

    Private Sub txtAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmount.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.txtAmount.Text, Me.ObjCur.GetDecimal_DefaultCurrency)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmount.LostFocus
        Me.txtAmount.BackColor = Me.txtPrimaryUnits.BackColor
    End Sub

    Private Sub dtpDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.Leave
        Try
            If Baru = True Then
                Me.TxtStockInNo.Text = Me.obj.GetNoBaru(Me.dtpDate.Value)
            Else
                Dim NoBaru As String = Me.obj.GetNoBaru(Me.dtpDate.Value)
                If NoBaru <> obj.GetNoTransaksi(Me.TxtPrimaryStockIn.Text) And NoBaru <> Me.TxtStockInNo.Text And Me.dtpDate.Value <> Me.obj.GetTglTransaksi(Me.TxtPrimaryStockIn.Text) Then
                    If MsgBox("Month or year has been changed. Do you want to change transaction code too?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.TxtStockInNo.Text = Me.obj.GetNoBaru(Me.dtpDate.Value)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiHistoricalPurchase_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            If Me.TxtRemark.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.obj.SimpanBaru(Me.dgvItem, Me.objItem, Me.UserCode) = True Then
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
                            If Me.obj.SimpanPerubahan(Me.dgvItem, Me.objItem, Me.UserCode) = True Then
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

    Private Sub TxtTotal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtTotal.TextChanged
        TxtTotal.Text = Me.ObjCD.DigitGrouping(TxtTotal.Text)
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserAddedRow
        Try
            If Me.TxtRemark.Enabled = True Then
                ObjCD.SetNumber(Me.dgvItem)
                Me.HitungTotal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            If Me.TxtRemark.Enabled = True Then
                ObjCD.SetNumber(Me.dgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            Me.HitungTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvItem.Rows(row.Index).Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False) Then
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> Me.dgvItem.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRow(row, False)
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCodeNo2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCodeNo2.Click
        Dim i As New Core_Module.frmCariMasterProduct
        Dim PrimaryProduct As Double = -1
        Dim rowdgv As DataGridViewRow

        Try
            If Me.dgvItem.RowCount = 0 Then
                Exit Sub
            End If

            'cari product yang akan diset posisinya
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
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

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtQtyperBatch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.TextChanged

    End Sub
End Class