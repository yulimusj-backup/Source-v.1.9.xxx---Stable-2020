Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiSOC
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    '
    Public FormNo As Integer = 0
    Public vFromQuick As Boolean = False
    Public vStopCancellation As Boolean = False
    Public vDirectlySave As Boolean = False
    Public vShowDialog As Boolean = True
    '
    Dim ObjSOItem As New ClsTransaksiSalesOrderItem
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjSOCItem As New ClsTransaksiSOCItem
    Dim ObjSOC As New ClsTransaksiSOC
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim Baru As Boolean
    Dim TBL3 As New DseTransaksiSO.DataTable3DataTable
    Dim TBL4 As New DseTransaksiSOC.DataTable3DataTable
    Dim buffer As String = ""

    Private Sub txtPONO_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSONO.GotFocus
        txtSONO.BackColor = BackColorAll
    End Sub

    Private Sub txtPONO_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSONO.LostFocus
        txtSONO.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub txtCancellationReason_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCancellationReason.GotFocus
        txtCancellationReason.BackColor = BackColorAll
    End Sub

    Private Sub txtCancellationReason_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCancellationReason.LostFocus
        txtCancellationReason.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Public Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimarySOC.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimarySOC.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimarySOC.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimarySOC.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimarySOC.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimarySOC.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False

            If Me.dgvItem.RowCount > 0 Then
                Me.txtSONO.Enabled = False
                Me.btnFindSalesOrder.Enabled = False
            End If

            If Me.txtPrimarySO.Text <> "-1" And Me.txtPrimarySO.Text <> "" Then
                Me.SetGboItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, False)

            Me.KosongkanItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.txtPrimarySO.Text = "-1"
            Me.TxtPrimarySOC.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)
            Me.txtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnitsDefault.Text = "-1"
            Me.txtPrimaryUnitsMulti.Text = "-1"
            Me.TxtQtyperBatchDefault.Text = "0"
            Me.TxtQtyperBatchMulti.Text = "0"
            Me.cboCodeNo.DataSource = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Try
            'data item salah bila:
            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchMulti.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If
            '2.b. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchDefault.Text) = False Then
                MsgBox("Qty you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If

            '9. Bila setelah diinput jumlah sales order balance menjadi minus
            Dim JmlObjLama, JmlSOBBaru, JmlDGV, JmlSOB As Double
            JmlDGV = Me.GetJmlDGV_Multi(Me.txtPrimaryCodeNo.Text)
            JmlSOB = Me.ObjSOItem.GetSalesOrderBalance_Multi(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text)
            JmlObjLama = Me.ObjSOCItem.GetJml_Multi(Me.TxtPrimarySOC.Text, Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlSOBBaru = JmlSOB + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchMulti.Text)
            If JmlSOBBaru < 0 Then
                MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlSOB + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)) + " " + _
                Me.LblUnitsMulti.Text + ".", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Double) As Double
        Try
            Dim row As DataGridViewRow
            Dim Jml As Double = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    Jml = Jml + CDbl(row.Cells(Me.ColQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Multi(ByVal vPrimaryCodeNo As Double) As Double
        Try
            Dim row As DataGridViewRow
            Dim Jml As Double = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    Jml = Jml + CDbl(row.Cells(Me.ColQtyMulti.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Sales order cancellation number")
            KoleksiTextBox.Add(Me.TxtSOCNo)
            KoleksiNamaData.Add("Sales order no")
            KoleksiTextBox.Add(Me.txtSONO)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.Baru = False Then
                If Me.ObjSOC.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.ObjSOC.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtSOCNo.Text = Me.ObjSOC.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.ObjSOC.TanggalTransaksi
                        Me.TxtSOCNo.Text = Me.ObjSOC.NoTransaksi
                        Return False
                    End If
                End If
            End If

            ''cek bisi ada qty default=0
            Dim ListError As String = ""
            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If CDec(.Cells(Me.ColQty.Index).Value) <= 0 Then
                        ListError += .Cells(Me.ColCodeNo.Index).Value & ", "
                    End If
                End With
            Next
            If ListError <> "" Then
                ListError = ListError.Substring(0, ListError.Length - 3) & "."
                MsgBox("The following product has zero default qty: " & vbNewLine & ListError & vbNewLine & "Please change qty.", MsgBoxStyle.Information, "Information")
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

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.ObjSOC
                .CancellationReason = Me.txtCancellationReason.Text
                .NoTransaksi = Me.TxtSOCNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .PrimarySO = Me.txtPrimarySO.Text
                .Remarks = Me.TxtRemark.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjSOC
                Me.TxtSOCNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.TxtPrimarySOC.Text = .Primary
                Me.txtPrimarySO.Text = .PrimarySO
                Me.txtSONO.Text = Me.ObjSO.GetNoSO(Me.txtPrimarySO.Text)
                Me.txtCancellationReason.Text = .CancellationReason
                'salin data SO
                Me.ObjSO.GetCustCode_CustName_CustPONO_ContactPerson_PhoneNo(Me.txtPrimarySO.Text, _
                Me.TxtCustomerName.Text, Me.txtCustomerCode.Text, Me.txtCustPONO.Text, _
                Me.txtContactPerson.Text, Me.txtPhoneNo.Text)

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.FormatDGV()
            'digit grouping
            Dim row As DataGridViewRow
            For Each row In Me.dgvItem.Rows
                row.Cells(3).Value = Me.ObjCD.DigitGrouping(row.Cells(3).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ObjLS As New clsMasterListedSoftware
        Dim rowDGV As New DataGridViewRow
        Try
            If Me.vFromQuick = False Then
                Me.Baru = True
                Me.NonAktifkanDetail()
            End If

            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable

            If Me.vFromQuick = True Then
                Me.prgProgress.Visible = True
                Me.btnStopProcess.Visible = True
                Me.btnStopProcess.TabStop = True
                Me.lblProgress.Visible = True
                Me.btnFind.Visible = False
                Me.btnFind.TabStop = False
                Me.btnNew.Visible = False
                Me.btnNew.TabStop = False
                Me.btnEdit.Visible = False
                Me.btnEdit.TabStop = False
                Me.btnDelete.Visible = False
                Me.btnDelete.TabStop = False
                Me.btnExit.Visible = False
                Me.btnExit.TabStop = False
                Me.MinimizeBox = False
                Me.txtSONO.Enabled = False
                Me.btnFindSalesOrder.Enabled = False
                If vShowDialog = True Then
                    If IsNumeric(Me.txtPrimarySO.Text) = True Then
                        If Me.txtPrimarySO.Text > 0 Then
                            For ij As Integer = Me.dgvItem.Rows.Count - 1 To 0 Step -1
                                rowDGV = Me.dgvItem.Rows(ij)
                                If rowDGV.Cells(Me.ColQtyMulti.Index).Value = 0 Then
                                    Me.dgvItem.Rows.RemoveAt(ij)
                                End If
                            Next
                            Me.IsiCboProduct(Me.txtPrimarySO.Text)
                            Me.ObjCD.SetNumber(Me.dgvItem)
                        End If
                    End If
                End If
                If vDirectlySave = True Then
                    Me.SimpanLangsung()
                End If
            End If

            Me.cboCodeNo.Width = Me.SplitContainer2.Panel1.Width - 5 - Me.cboCodeNo.Left
            Me.txtDesc.Width = Me.cboCodeNo.Width
            Me.BtnInput.Left = Me.SplitContainer2.Panel2.Width - 5 - Me.BtnInput.Width
            Me.TxtRemarkItem.Width = Me.BtnInput.Left - 5 - Me.TxtRemarkItem.Left
            Me.LblUnitsMulti.Left = Me.BtnInput.Left
            Me.TxtQtyperBatchMulti.Width = Me.LblUnitsMulti.Left - 5 - Me.TxtQtyperBatchMulti.Left

            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjLS = Nothing
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjSOC.SimpanBaru(Me.dgvItem, Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.KosongkanItem()
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()

                            If vFromQuick = True Then
                                Me.Close()
                            End If
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjSOC.SimpanPerubahan(Me.dgvItem, Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.KosongkanItem()
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

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 And Me.dtpDate.Value.Date = Today And _
            Me.TxtRemarkItem.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
                Me.KosongkanItem()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjSOC.CariDetail(Me.TxtPrimarySOC.Text, Me.dgvItem)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                If Me.vFromQuick = True Then
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.Baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.KosongkanItem()
            Me.dtpDate.Focus()
            Me.ObjSOC.ResetObject(Me.dgvItem)
            Me.TxtSOCNo.Text = Me.ObjSOC.GetNoBaru(Me.dtpDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariTransaksiSOC
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()

            If Not hj.PrimaryTransaksi = -1 Then
                If Me.ObjSOC.CariDetail(hj.PrimaryTransaksi, Me.dgvItem) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtSOCNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtSOCNo.Text <> "" Then
                If Me.ObjSOC.ApakahBolehEdit(Me.TxtPrimarySOC.Text, True) = True Then
                    Me.Baru = False
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.TxtSOCNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtSOCNo.Text <> "" Then
                If Me.ObjSOC.ApakahBolehHapus(Me.TxtPrimarySOC.Text, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjSOC.Hapus(Me.TxtPrimarySOC.Text) = True Then
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

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjSOC.SimpanBaru(Me.dgvItem, Me.UserCode) = True Then
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
                            If Me.ObjSOC.SimpanPerubahan(Me.dgvItem, Me.UserCode) = True Then
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

    Private Sub IsiCboProduct(ByVal vPrimarySO As Double)
        Try
            Me.TBL3.Rows.Clear()
            Me.TBL4.Rows.Clear()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'isikan dari basis data-persedian terakhir
            Me.cboCodeNo.DataSource = Nothing
            Me.ObjSOItem.IsikanTBLProduct(Me.txtPrimarySO.Text, Me.TBL3, Me.cboCodeNo)
            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiSOC.DataTable3Row
            Dim row2 As DseTransaksiSO.DataTable3Row
            Dim ketemu As Boolean
            If Me.Baru = False Then
                Me.ObjSOCItem.IsiTBL(Me.TxtPrimarySOC.Text, Me.TBL4)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL4.Rows.Count > 0
                    row1 = Me.TBL4.Rows(0)
                    For Each row2 In Me.TBL3.Rows
                        If row1.Primary_Product = row2.Primary_Product Then
                            row2.Order_Balance = row2.Order_Balance + row1.Qty
                            row2.Order_Balance_Multi = row2.Order_Balance_Multi + row1.Qty_Multi
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        row2 = TBL3.NewDataTable3Row
                        row2.Primary_SO = Me.txtPrimarySO.Text
                        row2.Primary_Customer = -1
                        row2.Primary_Product = row1.Primary_Product
                        row2.Product_Code = Me.ObjProduct.GetProductCode(row1.Primary_Product)
                        row2.Order_Balance = row1.Qty
                        row2.SO_No = ""
                        row2.Order_Balance_Multi = row1.Qty_Multi
                        TBL3.Rows.Add(row2)
                        'Me.TBL3.Rows.Add(Me.txtPrimarySO.Text, -1, row1.Primary_Product, _
                        'Me.ObjProduct.GetProductCode(row1.Primary_Product), row1.Qty, "")
                    End If
                    Me.TBL4.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsiTBLProduct()
            While Me.TBL4.Rows.Count > 0
                row1 = Me.TBL4.Rows(0)
                For Each row2 In Me.TBL3.Rows
                    If row1.Primary_Product = row2.Primary_Product Then
                        row2.Order_Balance = row2.Order_Balance - row1.Qty
                        row2.Order_Balance_Multi = row2.Order_Balance_Multi - row1.Qty_Multi
                        If row2.Order_Balance_Multi = 0 Then
                            Me.TBL3.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL3.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL4.Rows.Remove(row1)
            End While

            Me.TBL3.DefaultView.Sort = "[Product Code] ASC"
            Me.cboCodeNo.DataSource = Me.TBL3.DefaultView
            Me.cboCodeNo.DisplayMember = "Product Code"
            Me.cboCodeNo.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiTBLProduct()
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiSOC.DataTable3Row
            For Each row In Me.dgvItem.Rows
                Dim ketemu As Boolean = False
                For Each row1 In Me.TBL4.Rows
                    If row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                        'kalo ketemu ubah data order balance
                        row1.Qty = row1.Qty + CDbl(row.Cells(Me.ColQty.Index).Value)
                        row1.Qty_Multi = row1.Qty_Multi + CDbl(row.Cells(Me.ColQtyMulti.Index).Value)
                        ketemu = True
                        Exit For
                    End If
                Next
                'kalo ga ketemu tambahkan baris baru
                If ketemu = False Then
                    row1 = TBL4.NewDataTable3Row '
                    row1.Primary_Transaksi = Me.TxtPrimarySOC.Text
                    row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                    row1.Product_Code = row.Cells(Me.ColCodeNo.Index).Value
                    row1.Qty = row.Cells(Me.ColQty.Index).Value
                    row1.Qty_Multi = row.Cells(Me.ColQtyMulti.Index).Value
                    TBL4.Rows.Add(row1)

                    'Me.TBL4.Rows.Add(row.Cells(3).Value, row.Cells(6).Value, Me.TxtPrimarySOC.Text, row.Cells(1).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGboItem()
        Try
            Me.cboCodeNo.DataSource = Nothing

            'sales order & product code
            Me.IsiCboProduct(Me.txtPrimarySO.Text)
            Me.ObjSO.GetCustCode_CustName_CustPONO_ContactPerson_PhoneNo(Me.txtPrimarySO.Text, _
            Me.TxtCustomerName.Text, Me.txtCustomerCode.Text, Me.txtCustPONO.Text, Me.txtContactPerson.Text, _
            Me.txtPhoneNo.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex(ByVal vPrimaryProduct As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Dim DefPrecMulti As Integer = 0
        Dim DefPrecDefault As Integer = 0
        Dim RowDgv As DataGridViewRow
        Try
            If Me.ApakahDataItemBenar = True Then
                Me.txtSONO.Enabled = False
                Me.btnFindSalesOrder.Enabled = False

                Dim Index As Integer = Me.GetIndex(Me.txtPrimaryCodeNo.Text)

                DefPrecMulti = ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
                DefPrecDefault = ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)

                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    'row.Cells(3).Value = CDbl(row.Cells(3).Value) + CDbl(Me.TxtQtyperBatchDefault.Text)

                    row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) + CDec(Me.TxtQtyperBatchMulti.Text), DefPrecMulti)
                    Me.SetQtyDefault()

                Else
                    'tambahkan baris baru
                    'Me.dgvItem.Rows.Add("", Me.cboCodeNo.Text, Me.txtDesc.Text, Me.TxtQtyperBatchDefault.Text, _
                    'Me.lblUnits.Text, Me.TxtRemarkItem.Text, Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsDefault.Text)
                    Me.dgvItem.Rows.Add()
                    RowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With RowDgv
                        .Cells(Me.ColNo.Index).Value = ""
                        .Cells(Me.ColCodeNo.Index).Value = Me.cboCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColQty.Index).Value = ObjCD.DigitGrouping(Me.TxtQtyperBatchDefault.Text, DefPrecDefault)
                        .Cells(Me.ColUnit.Index).Value = Me.lblUnits.Text
                        .Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.txtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnitsDefault.Text
                        .Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, DefPrecMulti)
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value = Me.txtPrimaryUnitsMulti.Text
                        .Cells(Me.ColUnitsMulti.Index).Value = Me.LblUnitsMulti.Text
                        .Cells(Me.ColRatioDefault.Index).Value = Me.TxtRatioDefault.Text
                        .Cells(Me.ColRatioReference.Index).Value = Me.TxtRatioReference.Text
                        If Me.txtPrimaryUnitsDefault.Text <> Me.txtPrimaryUnitsMulti.Text Then
                            .Cells(Me.ColRatio.Index).Value = Me.TxtRatioReference.Text _
                            & " " & Me.LblUnitsMulti.Text & " = " & Me.TxtRatioDefault.Text & " " & Me.lblUnits.Text
                        Else
                            .Cells(Me.ColRatio.Index).Value = "-"
                        End If
                    End With
                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
                End If

                'kendalikan focus item selanjutnya
                Me.IsiCboProduct(Me.txtPrimarySO.Text)
                Me.cboCodeNo.Focus()
                Me.TxtQtyperBatchMulti.Text = ObjCD.DigitGrouping("0", Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                Me.SetQtyDefault()
                Me.TxtRemarkItem.Clear()

                'format dgv
                Me.FormatDGV()

                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim IsMultiUnit As Boolean = False
        Try
            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If .Cells(Me.ColPrimaryUnitsMulti.Index).Value <> .Cells(Me.ColPrimaryUnits.Index).Value Then
                        IsMultiUnit = True
                        Exit For
                    End If
                End With
            Next
            If IsMultiUnit = False Then
                Me.ColQty.Visible = False
                Me.ColUnit.Visible = False
                Me.ColRatio.Visible = False
            Else
                Me.ColQty.Visible = True
                Me.ColUnit.Visible = True
                Me.ColRatio.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchDefault.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchDefault, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchDefault.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchDefault, sender, e)
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Remarks
            'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColQtyMulti.Index Or e.ColumnIndex = Me.ColRemark.Index Then
                    'kalo sampe sini maka boleh edit
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
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

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex <> Me.ColRemark.Index Then
                        Me.EndEditQty(sender, e, row)
                    End If
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Dim JMLDgv As Double = 0

        Try
            With row
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, _
                e.RowIndex).Value, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value))
                Me.SetQtyDefaultOnDgv(row.Index)

                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.SetQtyDefaultOnDgv(row.Index)
                End If

                'cek apakah sales order balance menjadi minus setelah perubahan data
                Dim JmlObjLama, JmlSOBBaru, JmlSOB, JmlSendiri As Double
                JMLDgv = Me.GetJmlDGV_Multi(.Cells(Me.ColPrimaryCodeNo.Index).Value)
                JmlSOB = Me.ObjSOItem.GetSalesOrderBalance_Multi(Me.txtPrimarySO.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                JmlObjLama = Me.ObjSOCItem.GetJml_Multi(Me.TxtPrimarySOC.Text, Me.txtPrimarySO.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                JmlSendiri = .Cells(Me.ColQtyMulti.Index).Value
                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlSOBBaru = JmlSOB - JMLDgv + JmlObjLama
                If JmlSOBBaru < 0 Then
                    MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlSOB - JMLDgv + JmlSendiri + JmlObjLama), Me.ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)) + " " + row.Cells(Me.ColUnitsMulti.Index).Value + ".", _
                    MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.SetQtyDefaultOnDgv(row.Index)
                    Exit Sub
                End If


                'jika jumlah menurun maka harus load ulang data
                Me.SetGboItem()
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Me.dgvItem.TabStop = True
        If Me.TxtRemark.Enabled = True Then
            Me.btnUP.Enabled = True
            Me.btnDown.Enabled = True
        End If
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        If Me.dgvItem.Rows.Count = 0 Then
            Me.dgvItem.TabStop = False
            If Me.TxtRemark.Enabled = True Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            If Me.TxtRemark.Enabled = True Then
                Me.txtSONO.Enabled = True
                Me.btnFindSalesOrder.Enabled = True
            End If
        End If
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.dgvItem.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem)
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

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            'set number
            Me.ObjCD.SetNumber(Me.dgvItem)

            'cek sampai dimana perlu load ulang data
            Me.SetGboItem()

            'format tampilan
            Me.FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            'tidak perlu cek akibatnya pada sales order balance
            'ketika user akan menghapus data, tidak perlu cek apakah akibatnya akan mengakibatkan ada stock yang minus, karena akan mengakibatkan
            'stock menjadi positif
            'tetapi untuk transaksi yang berkaitan dengan transaksi lainnya (DO, Invoice, GRN atau PO) harus cek validitas jumlah berkaitan dengan mereka
            'tandai apakah pada baris yang sama no SO-nya sama
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Try
            Me.TxtRemark.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Try
            Me.TxtRemark.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.GotFocus
        Try
            Me.TxtQtyperBatchDefault.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.LostFocus
        Try
            Me.TxtQtyperBatchDefault.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        Try
            Me.TxtRemarkItem.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        Try
            Me.TxtRemarkItem.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCodeNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCodeNo.SelectedValueChanged
        Dim DefPrecMulti As Integer = 0
        Dim DefPrecDefault As Integer = 0
        Dim ObjUnits As New Core_Module.ClsMasterUnit
        Try
            If Me.cboCodeNo.SelectedIndex <> -1 Then
                Dim RowView As DataRowView = Me.cboCodeNo.Items(Me.cboCodeNo.SelectedIndex)
                Dim row2 As DseTransaksiSO.DataTable3Row = RowView.Row
                Dim PrimaryCodeNoBaru As Double = row2.Primary_Product

                If Me.txtPrimaryCodeNo.Text <> PrimaryCodeNoBaru Then
                    Me.txtPrimaryCodeNo.Text = PrimaryCodeNoBaru
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                End If

                Me.txtPrimaryUnitsMulti.Text = Me.ObjSO.GetPrimaryUnitsMulti(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text)
                Me.txtPrimaryUnitsDefault.Text = Me.ObjProduct.GetPrimaryUnit(Me.txtPrimaryCodeNo.Text)
                DefPrecMulti = ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
                DefPrecDefault = ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
                Me.TxtRatioReference.Text = ObjCD.DigitGrouping(Me.ObjSO.GetSORatioReference(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text), DefPrecMulti)
                Me.TxtRatioDefault.Text = ObjCD.DigitGrouping(Me.ObjSO.GetSORatioDefault(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text), DefPrecDefault)
                If Me.TxtQtyperBatchMulti.Text = "" Then
                    Me.TxtQtyperBatchMulti.Text = "0"
                End If
                Me.TxtQtyperBatchMulti.Text = ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, DefPrecMulti)
                Me.SetQtyDefault()
                Me.lblUnits.Text = ObjUnits.GetUnitCode(Me.txtPrimaryUnitsDefault.Text)
                Me.LblUnitsMulti.Text = ObjUnits.GetUnitCode(Me.txtPrimaryUnitsMulti.Text)
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                Me.txtPrimaryUnitsMulti.Text = "-1"
                Me.txtPrimaryUnitsDefault.Text = "-1"
                Me.TxtRatioReference.Text = "1"
                Me.TxtRatioDefault.Text = "1"
                Me.TxtQtyperBatchMulti.Text = "0"
                Me.TxtQtyperBatchDefault.Text = "0"
                Me.lblUnits.Text = "Units"
                Me.LblUnitsMulti.Text = "Units"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSalesOrderNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSONO.KeyDown
        Dim i As New FrmCariTransaksiSO
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                If Me.ObjSO.ApakahSudahAda(Me.txtSONO.Text) = True Then
                    Me.txtPrimarySO.Text = Me.ObjSO.GetPrimary(Me.txtSONO.Text)
                    Me.txtCancellationReason.Focus()
                    Me.SetGboItem()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtSONO.Text <> "" Then
                        i.txtCari.Text = Me.txtSONO.Text
                    End If
                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()

                    If i.PrimaryTransaksi > 0 Then
                        Me.txtPrimarySO.Text = i.PrimaryTransaksi
                        Me.txtSONO.Text = Me.ObjSO.GetNoSO(Me.txtPrimarySO.Text)
                        Me.txtCancellationReason.Focus()
                        Me.SetGboItem()
                    Else
                        Me.txtSONO.Clear()
                        Me.txtPrimarySO.Text = "-1"
                        Me.txtSONO.Focus()
                        Me.SetGboItem()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSalesOrderNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSONO.Leave
        Dim i As New FrmCariTransaksiSO

        Try
            If Me.txtSONO.Text <> "" Then
                If Me.ObjSO.ApakahSudahAda(Me.txtSONO.Text) = True Then
                    Me.txtPrimarySO.Text = Me.ObjSO.GetPrimary(Me.txtSONO.Text)
                    Me.SetGboItem()
                Else
                    i.BackColorAll = Me.BackColorAll
                    If Me.txtSONO.Text <> "" Then
                        i.txtCari.Text = Me.txtSONO.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksi > 0 Then
                        Me.txtPrimarySO.Text = i.PrimaryTransaksi
                        Me.txtSONO.Text = Me.ObjSO.GetNoSO(Me.txtPrimarySO.Text)
                        Me.SetGboItem()
                    Else
                        Me.txtSONO.Clear()
                        Me.txtPrimarySO.Text = "-1"
                        Me.KosongkanItem()
                    End If
                End If
            Else
                Me.txtPrimarySO.Text = "-1"
                Me.SetGboItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSalesOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSalesOrder.Click
        Dim i As New FrmCariTransaksiSO

        Try
            i.BackColorAll = Me.BackColorAll

            If Me.txtSONO.Text <> "" Then
                i.txtCari.Text = Me.txtSONO.Text
            End If

            i.ShowDialog()
            If i.PrimaryTransaksi > 0 Then
                Me.txtPrimarySO.Text = i.PrimaryTransaksi
                Me.txtSONO.Text = Me.ObjSO.GetNoSO(Me.txtPrimarySO.Text)
                Me.txtCancellationReason.Focus()
                Me.SetGboItem()
            Else
                Me.txtSONO.Clear()
                Me.txtPrimarySO.Text = "-1"
                Me.txtSONO.Focus()
                Me.SetGboItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.Baru = False Then
                    If Me.ObjSOC.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.ObjSOC.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtSOCNo.Text = Me.ObjSOC.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtSOCNo.Text = Me.ObjSOC.GetNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtSOCNo.Text = Me.ObjSOC.GetNoBaru(Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtPrimaryCodeNo.Text <> "-1" And Me.txtPrimaryCodeNo.Text <> "" Then
                    Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchDefault.Text, _
                    Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Enter
        Try
            Me.buffer = Me.TxtQtyperBatchMulti.Text
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.GotFocus
        Try
            Me.TxtQtyperBatchMulti.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchMulti, sender, e)
    End Sub

    Private Sub TxtQtyperBatchMulti_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchMulti, sender, e)
    End Sub

    Private Sub TxtQtyperBatchMulti_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtPrimaryCodeNo.Text <> "-1" And Me.txtPrimaryCodeNo.Text <> "" Then
                    Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, _
                    Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    SetQtyDefault()
                    If CDec(Me.TxtQtyperBatchDefault.Text) = 0 And CDec(Me.TxtQtyperBatchMulti.Text) > 0 Then
                        Me.TxtQtyperBatchMulti.Text = Me.buffer
                        MsgBox("Qty you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.SetQtyDefault()
                        Me.TxtQtyperBatchMulti.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Try
            Me.TxtQtyperBatchDefault.Text = ObjCD.DigitGrouping(CDec(Me.TxtQtyperBatchMulti.Text) * CDec(Me.TxtRatioDefault.Text) / CDec(Me.TxtRatioReference.Text), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefaultOnDgv(ByVal vRowIdx As Integer)
        Try
            Me.dgvItem.Item(Me.ColQty.Index, vRowIdx).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, vRowIdx).Value) * _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioDefault.Index, vRowIdx).Value / _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioReference.Index, vRowIdx).Value)), _
                                                            Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, vRowIdx).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryUnitsMulti_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryUnitsMulti.GotFocus
        Try
            Me.TxtQtyperBatchMulti.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.LostFocus
        Try
            Me.TxtQtyperBatchMulti.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        'Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi

        Try
            If Me.TxtSOCNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtSOCNo.Text <> "" Then
                'If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Print") = False Then
                '    FrmO.AccessName = "FrmTransaksiSO_Print"
                '    If FrmO.ShowDialog <> Windows.Forms.DialogResult.OK Then
                '        MsgBox("You don't have permission to perform this action", MsgBoxStyle.Information, "Information")
                '        Exit Sub
                '    End If
                'End If

                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Dim objcust As New CustomSMESS.ClsTransaksiSOC
                Dim FrmCetak As New FrmTransaksiSOC_Cetak

                objcust.CetakSOC(FrmCetak.CrvSOC, Me.TxtPrimarySOC.Text, Me.UserCode)
                objcust = Nothing
                FrmCetak.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            'Finally
            '    FrmO.Dispose()
        End Try
    End Sub

    Public Sub FromCutOff_SOCancellation_Create(ByVal vPrimaryTransaksi As Double, ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core, _
    ByVal vTblPrimary As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable)

        Try
            Me.Baru = False
            Me.ObjSOC.ResetObject(Me.dgvItem)

            Me.ObjSOC.CariDetail(vPrimaryTransaksi, Me.dgvItem)
            Me.ObjSOC.PrimarySO = ObjCutOff_Core.GetPrimaryUpdate(Me.ObjSOC.PrimarySO, vTblPrimary)
            Me.UserCode = Me.ObjSOC.CreatedBy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_SOCancellation_Save()
        Try
            Me.Baru = True
            Me.ObjSOC.SimpanBaru(Me.dgvItem, Me.UserCode, False, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'tambahan utk quick soc
    Public Sub SetBaru()
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.KosongkanItem()
            Me.AktifkanDetail()
            Me.TxtRemark.BackColor = Me.TxtPrimarySOC.BackColor
            Me.txtCancellationReason.BackColor = Me.TxtPrimarySOC.BackColor
            Me.TxtRemarkItem.BackColor = Me.TxtPrimarySOC.BackColor
            Me.TxtQtyperBatchMulti.BackColor = Me.TxtPrimarySOC.BackColor

            Me.dtpDate.Focus()
            Me.ObjSOC.ResetObject(Me.dgvItem)
            Me.TxtSOCNo.Text = Me.ObjSOC.GetNoBaru(Me.dtpDate.Value)

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            If Me.vFromQuick = True Then
                Me.btnStopProcess.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanLangsung()
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If Me.ObjSOC.SimpanBaru(Me.dgvItem, Me.UserCode) Then
                        Me.KosongkanItem()
                        Me.SalinKeField()
                        Me.NonAktifkanDetail()
                        Me.btnFind.Focus()

                        If vFromQuick = True Then
                            Me.Close()
                        End If
                    End If
                Else
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Me.NonAktifkanDetail()
                    Me.vStopCancellation = True
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnStopProcess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStopProcess.Click
        Try
            If MsgBox("Stop cancellation process?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.vStopCancellation = True
                Me.NonAktifkanDetail()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class