Option Explicit On
Option Strict Off

Imports Core_Module

Partial Public Class FrmTransaksiSalesInvoiceSettingPajak
    Dim baru As Boolean
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjLicense As New ClsLicense
    Dim ObjCur As New clsMasterCurrency

    Private Sub dgvItemRate_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemRate.CellBeginEdit
        Try
            'yg boleh diedit hanya kolom kurs saja
            If e.ColumnIndex = Me.ColRateKurs.Index Then
                Me.Buffer = Me.dgvItemRate.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemRate_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemRate.CellEndEdit
        Try
            'pastikan ada isinya
            If IsNothing(Me.dgvItemRate.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.dgvItemRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If

            If e.ColumnIndex = Me.ColRateKurs.Index Then
                'pastikan nilainya non negatif
                If Me.ObjCD.CekNonNegatif(Me.dgvItemRate.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItemRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'kembalikan format angka decimal
                Me.dgvItemRate.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemRate.Item(e.ColumnIndex, _
                e.RowIndex).Value, Me.ObjTaxRate.DefaultPrec)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimary.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, False, False)

            ModCekData.KendalikanControl(Me.TabPage2.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.btnClose.Enabled = False

            Me.dgvItemRate.ReadOnly = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, False, False)

            ModCekData.KendalikanControl(Me.TabPage2.Controls, False, False, True, True, True, False, False, False)
            Me.CancelButton = Me.btnClose
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)

            Me.btnClose.Enabled = True
            Me.dgvItemRate.ReadOnly = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)

            Me.TxtPrimary.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.ObjTaxRate.IsikanDGVKosong(Me.dgvItemRate, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim rowDGV As DataGridViewRow
        Dim JmlDummy As Integer = 0

        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("KMK Number")
            KoleksiTextBox.Add(Me.TxtKMKNumber)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data item harus ada minimal 1
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please specify 1 item or more.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            For Each rowDGV In Me.dgvItemRate.Rows
                If rowDGV.Cells(Me.ColRateKurs.Index).Value <> 0 Then
                    JmlDummy += 1
                End If
            Next
            If JmlDummy = 0 Then
                MsgBox("Please specify rate for 1 item or more.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.ObjTaxRate.ApakahSudahAda(Me.TxtKMKNumber.Text) = True Then
                    MsgBox("Duplication in KMK Number, please change your KMK Number", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtKMKNumber.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.ObjTaxRate.GetKMKNumber(CDbl(Me.TxtPrimary.Text))
                If ProductGroupLama <> Me.TxtKMKNumber.Text Then
                    If Me.ObjTaxRate.ApakahSudahAda(Me.TxtKMKNumber.Text) = True Then
                        MsgBox("Duplication in KMK Number, please change your KMK Number", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtKMKNumber.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjTaxRate.GetTimeStamp(Me.TxtPrimary.Text)
                If Me.ObjTaxRate.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjTaxRate.CariDetail(Me.ObjTaxRate.Primary, Me.dgvItemRate, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'cek apakah ada yg bertentangan
            If Me.ObjTaxRate.ApakahAdaYgKonflik(Me.TxtPrimary.Text, Me.dtpStart.Value, Me.dtpEnd.Value, True) = True Then
                Return False
            End If

            'tgl sampai harus lebih besar dari tgl dari
            If Me.dtpEnd.Value.Date <= Me.dtpStart.Value.Date Then
                MsgBox("End date must be large than start date", MsgBoxStyle.Information, "Information")
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
            With Me.ObjTaxRate
                .KMKDate = Me.dtpDate.Value
                .KMKNumber = Me.TxtKMKNumber.Text.Trim
                .StartDate = Me.dtpStart.Value
                .EndDate = Me.dtpEnd.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjTaxRate
                Me.dtpDate.Value = .KMKDate
                Me.TxtKMKNumber.Text = .KMKNumber
                Me.dtpStart.Value = .StartDate
                Me.dtpEnd.Value = .EndDate

                Me.TxtPrimary.Text = .Primary
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ShowIcon = False
            Me.ObjTaxRate = New ClsTransaksiSalesInvoiceTaxRate
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItemRate.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjTaxRate.SimpanBaru(Me.UserCode, Me.dgvItemRate, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.UpdateKeDefaultExchangeRate()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjTaxRate.SimpanPerubahan(Me.UserCode, Me.dgvItemRate, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.UpdateKeDefaultExchangeRate()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateKeDefaultExchangeRate()
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtPrimary.Text <> "" And Me.TxtPrimary.Text <> "-1" Then
                If Me.ObjTaxRate.ApakahDataTerakhir(Me.TxtPrimary.Text) = True Then
                    If MsgBox("Update default exchange rate value?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.No Then
                        For Each rowDGV In Me.dgvItemRate.Rows
                            If CDec(rowDGV.Cells(Me.ColRateKurs.Index).Value) <> 0 Then
                                Me.ObjCur.Set_DefaultExchangeRate(rowDGV.Cells(Me.ColRatePrimaryCurrency.Index).Value, _
                                rowDGV.Cells(Me.ColRateKurs.Index).Value)
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtKMKNumber.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtKMKNumber.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTransaksiSalesInvoiceSettingPajak_Cari

        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                If Me.ObjTaxRate.CariDetail(hj.PrimaryTransaksi, Me.dgvItemRate, Me) = True Then
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
        If Me.TxtKMKNumber.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.TxtKMKNumber.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.TxtKMKNumber.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtKMKNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtKMKNumber.Text <> "" Then
                If Me.ObjTaxRate.ApakahBolehHapus(Me.TxtPrimary.Text) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjTaxRate.Hapus(Me.TxtPrimary.Text) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtKMKNumber.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjTaxRate.SimpanBaru(Me.UserCode, Me.dgvItemRate, Me) = True Then
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
                            If Me.ObjTaxRate.SimpanPerubahan(Me.UserCode, Me.dgvItemRate, Me) = True Then
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

    Private Sub TxtKMKNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtKMKNumber.GotFocus
        TxtKMKNumber.BackColor = BackColorAll
    End Sub

    Private Sub TxtKMKNumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtKMKNumber.LostFocus
        TxtKMKNumber.BackColor = Me.TxtPrimary.BackColor
    End Sub
End Class
