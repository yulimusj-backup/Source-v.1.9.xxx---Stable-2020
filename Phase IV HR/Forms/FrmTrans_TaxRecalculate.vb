Option Explicit On
Option Strict Off

Imports System.Drawing, System.Windows.Forms

Public Class FrmTrans_TaxRecalculate
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim baru As Boolean
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Obj As New ClsTrans_TaxRecalculate
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjKryw As New Core_Module.ClsMasterKaryawan
    Dim ObjSalary As New ClsTransaksiSalary

    Dim defprec As Integer = ObjCurr.GetDecimal_DefaultCurrency
    Dim Buffer As String = ""
    Dim vTBLItemTax As New DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable

    Private Sub FrmTransaksiLoan_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtRemark.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObject()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.dgvItem, Me, Me.vTBLItemTax) = True Then
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
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.dgvItem, Me, Me.vTBLItemTax) = True Then
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

    Private Sub FrmTransaksiLoan_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dtpPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)

            If Me.baru = False Then
                'untuk edit, period tidak boleh diganti
                Me.dtpPeriod.Enabled = False
            End If

            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.dgvItem.ReadOnly = False
            Me.CancelButton = Me.btnCancel
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.dgvItem.ReadOnly = True
            Me.CancelButton = Me.btnExit
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, True)
            Me.TxtPrimary.Text = -1
            Me.TxtDifference.Text = Me.ObjCD.DigitGrouping(0, Me.defprec)
            Me.dgvItem.Rows.Clear()
            Me.vTBLItemTax.Rows.Clear()
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dtpPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim PrimaryRef As Double = -1

        Try
            '1.tidak boleh double
            If Me.Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = True Then
                If Me.baru = True Then
                    Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value)
                Else
                    If Me.Obj.GetPrimary(Me.TxtDocumentNo.Text) <> Me.TxtPrimary.Text Then
                        Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value)
                    End If
                End If
            End If

            '2.Minimal ada 1 item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please specify at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'periode kompensasi tidak boleh desember
            If Me.dtpPeriod.Value.Month = 12 Then
                MsgBox("We can't recalculate for december.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1 periode kompensasi hanya bisa di kompensasi 1x
            PrimaryRef = Me.Obj.GetPrimaryByPeriod(Me.dtpPeriod.Value)
            If PrimaryRef <> -1 Then
                If Me.TxtPrimary.Text <> PrimaryRef Then
                    MsgBox("There's tax recalculation for period " + Format(Me.dtpPeriod.Value, "MMM yyyy") + ", process stoped.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtDocumentNo.Text = .DocNo
                Me.dtpDate.Value = .Date_
                Me.dtpPeriod.Value = .Period
                Me.TxtDifference.Text = ObjCD.DigitGrouping(.Difference, Me.defprec)
                Me.TxtRemark.Text = .Remarks

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeObject()
        Try
            With Me.Obj
                .Primary = Me.TxtPrimary.Text
                .DocNo = Me.TxtDocumentNo.Text
                .Date_ = Me.dtpDate.Value
                .Period = Me.dtpPeriod.Value
                .Difference = Me.TxtDifference.Text
                .Remarks = Me.TxtRemark.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.dtpDate.Focus()
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.dtpDate.Focus()
            Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value)

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim IsBeginingBalance As Boolean = False

        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtDocumentNo.Text <> "" Then
                If Me.Obj.ApakahBolehedit(Me.UserCode, Me.TxtPrimary.Text) = True Then
                    Me.baru = False
                    Me.Obj.CariDetail(Me.TxtPrimary.Text, Me, Me.dgvItem, Me.vTBLItemTax)
                    Me.AktifkanDetail()
                    Me.TxtRemark.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObject()
                If baru = True Then
                    hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.dgvItem, Me, Me.vTBLItemTax) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.SalinKeObject()
                            Me.NonAktifkanDetail()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.dgvItem, Me, Me.vTBLItemTax) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.SalinKeObject()
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

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocumentNo.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongkanField()
                    Else
                        Me.NonAktifkanDetail()
                        Me.Obj.CariDetail(Me.TxtPrimary.Text, Me, Me.dgvItem, Me.vTBLItemTax)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.NonAktifkanDetail()
                Me.KosongkanField()
            End If
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Dim IsBeginingBalance As Boolean = False

        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocumentNo.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.UserCode, Me.TxtPrimary.Text) = True Then
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.TxtPrimary.Text, Me.UserCode, Me) = True Then
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

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTrans_TaxRecalculate_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                If Me.Obj.CariDetail(hj.PrimaryTransaksi, Me, Me.dgvItem, Me.vTBLItemTax) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                    Me.NonAktifkanDetail()
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

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.TxtRemark.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.TxtRemark.BackColor = Me.TxtPrimary.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnReload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReload.Click
        Dim PrimaryRef As Double = -1
        Dim row1, row2 As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxRow

        Try
            'periode kompensasi tidak boleh desember
            If Me.dtpPeriod.Value.Month = 12 Then
                MsgBox("We can't recalculate for december.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '1 periode kompensasi hanya bisa di kompensasi 1x
            PrimaryRef = Me.Obj.GetPrimaryByPeriod(Me.dtpPeriod.Value)
            If PrimaryRef <> -1 Then
                If Me.TxtPrimary.Text <> PrimaryRef Then
                    MsgBox("There's tax recalculation for period " + Format(Me.dtpPeriod.Value, "MMM yyyy") + ", process stoped.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If MsgBox("Recalculate data?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                Dim ij As New FrmTrans_TaxRecalculate_Process

                ij.vPeriod = Me.dtpPeriod.Value
                If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                    'salin data ke dgv
                    Me.SalinItemDGV(ij)
                    'salin ke data ke tabel item tax
                    Me.vTBLItemTax.Rows.Clear()
                    For Each row1 In ij.vTBLItemTax
                        row2 = Me.vTBLItemTax.Newtbltrans_taxrecalculate_itemtaxRow
                        With row2
                            .amount_gabungan_disetahunkan = row1.amount_gabungan_disetahunkan
                            .amount_month = row1.amount_month
                            .amount_rutin_disetahunkan = row1.amount_rutin_disetahunkan
                            .identifier_salary_component = row1.identifier_salary_component
                            .is_rutin = row1.is_rutin
                            .primary = Me.vTBLItemTax.Rows.Count + 1
                            .primary_salary = row1.primary_salary
                            .primary_transaksi = Me.TxtPrimary.Text
                            .remarks_gabungan = row1.remarks_gabungan
                            .remarks_rutin = row1.remarks_rutin
                            .tipe_item_tax = row1.tipe_item_tax
                            .tipe_tax = row1.tipe_tax
                        End With
                        Me.vTBLItemTax.Rows.Add(row2)
                    Next
                    'update txt difference
                    Me.UpdateTotalDifference()
                End If

                ij = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateTotalDifference()
        Dim rowDGV As DataGridViewRow
        Dim Dummy As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItem.Rows
                Dummy += rowDGV.Cells(Me.ColDifference.Index).Value
            Next
            Me.TxtDifference.Text = Me.ObjCD.DigitGrouping(Dummy, 0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinItemDGV(ByVal ij As FrmTrans_TaxRecalculate_Process)
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItem.Rows.Clear()

            For Each row As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1Row In ij.vTBLItem.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.Col01Gaji.Index).Value = row._01Gaji
                rowDGV.Cells(Me.Col01GajiRutin.Index).Value = row._01_Gaji_Rutin

                rowDGV.Cells(Me.Col02NewTunjPPh.Index).Value = row._02TunjanganPPh_new
                rowDGV.Cells(Me.Col02NewTunjPPhDisetahunkan.Index).Value = row._02_Tunjangan_PPH_Disetahunkan_new
                rowDGV.Cells(Me.Col02NewTunjPPhDisetahunkanRutin.Index).Value = row._02_Tunjangan_PPH_Disetahunkan_Rutin_new
                rowDGV.Cells(Me.Col02NewTunjPPhNonRutin.Index).Value = row._02_Tunjangan_PPH_Non_Rutin_new
                rowDGV.Cells(Me.Col02NewTunjPPhRutin.Index).Value = row._02_Tunjangan_PPH_Rutin_new

                rowDGV.Cells(Me.Col02TunjPPh.Index).Value = row._02TunjanganPPh
                rowDGV.Cells(Me.Col02TunjPPhDisetahunkan.Index).Value = row._02_Tunjangan_PPH_Disetahunkan
                rowDGV.Cells(Me.Col02TunjPPhDisetahunkanRutin.Index).Value = row._02_Tunjangan_PPH_Disetahunkan_Rutin
                rowDGV.Cells(Me.Col02TunjPPHNonRutin.Index).Value = row._02_Tunjangan_PPH_Non_Rutin
                rowDGV.Cells(Me.Col02TunjPPhRutin.Index).Value = row._02_Tunjangan_PPH_Rutin

                rowDGV.Cells(Me.Col03TunjLainnya.Index).Value = row._03TunjanganLainnya
                rowDGV.Cells(Me.Col03TunjLainnyaRutin.Index).Value = row._03_Tunjangan_Lainnya_Rutin
                rowDGV.Cells(Me.Col04Honorarium.Index).Value = row._04Honorarium
                rowDGV.Cells(Me.Col04HonorariumRutin.Index).Value = row._04_Honorarium_Rutin
                rowDGV.Cells(Me.Col05PremiAsuransi.Index).Value = row._05PremiAsuransi
                rowDGV.Cells(Me.Col05PremiAsuransiRutin.Index).Value = row._05_Premi_Asurasi_Rutin
                rowDGV.Cells(Me.Col06Natura.Index).Value = row._06Natura
                rowDGV.Cells(Me.Col06NaturaRutin.Index).Value = row._06_Natura_Rutin
                rowDGV.Cells(Me.Col07Tantiem.Index).Value = row._07Tantiem
                rowDGV.Cells(Me.Col07TantiemRutin.Index).Value = row._07_Tantiem_Rutin

                rowDGV.Cells(Me.Col08NewPenghasilanBruto.Index).Value = row._08PenghasilanBruto_new
                rowDGV.Cells(Me.Col08NewPenghasilanBrutoRutin.Index).Value = row._08_Penghasilan_Bruto_Rutin_new
                rowDGV.Cells(Me.Col08PenghasilanBruto.Index).Value = row._08PenghasilanBruto
                rowDGV.Cells(Me.Col08PenghasilanBrutoRutin.Index).Value = row._08_Penghasilan_Bruto_Rutin

                rowDGV.Cells(Me.Col09BiayaJabatan.Index).Value = row._09BiayaJabatan
                rowDGV.Cells(Me.Col09BiayaJabatanRutin.Index).Value = row._09_Biaya_Jabatan_Rutin
                rowDGV.Cells(Me.Col09NewBiayaJabatan.Index).Value = row._09BiayaJabatan_new
                rowDGV.Cells(Me.Col09NewBiayaJabatanRutin.Index).Value = row._09_Biaya_Jabatan_Rutin_new

                rowDGV.Cells(Me.Col10IuranPensiun.Index).Value = row._10IuranPensiun
                rowDGV.Cells(Me.Col10IuranPensiunRutin.Index).Value = row._10_Iuran_Pensiun_Rutin

                rowDGV.Cells(Me.Col11Jumlahengurang.Index).Value = row._11JumlahPengurangan
                rowDGV.Cells(Me.Col11JumlahPengurangRutin.Index).Value = row._11_Jumlah_Pengurang_Rutin
                rowDGV.Cells(Me.Col11NewJumlahPengurang.Index).Value = row._11JumlahPengurangan_new
                rowDGV.Cells(Me.Col11NewJumlahPengurangRutin.Index).Value = row._11_Jumlah_Pengurang_Rutin_new

                rowDGV.Cells(Me.Col14NewPenghasilanNettoUntPPh.Index).Value = row._14JumlahPenghasilanNettoUtkPPh_new
                rowDGV.Cells(Me.Col14NewPenghasilanNettoUntPPhRutin.Index).Value = row._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin_new
                rowDGV.Cells(Me.Col14PenghasilanNettoPPhRutin.Index).Value = row._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin
                rowDGV.Cells(Me.Col14PenghasilanNettoUntPPH.Index).Value = row._14JumlahPenghasilanNettoUtkPPh

                rowDGV.Cells(Me.Col15NewPTKP.Index).Value = row._15PTKP_new
                rowDGV.Cells(Me.Col15PTKP.Index).Value = row._15PTKP

                rowDGV.Cells(Me.Col16NewPKPSetahun.Index).Value = row._16PKPSetahun_new
                rowDGV.Cells(Me.Col16NewPKPSetahunRutin.Index).Value = row._16_PKP_Setahun_Rutin_new
                rowDGV.Cells(Me.Col16PKPSetahun.Index).Value = row._16PKPSetahun
                rowDGV.Cells(Me.Col16PKPSetahunRutin.Index).Value = row._16_PKP_Setahun_Rutin
                rowDGV.Cells(Me.Col17NewPPh21.Index).Value = row._17PPh21_new
                rowDGV.Cells(Me.Col17NewPPh21Rutin.Index).Value = row._17_PPH_21_Rutin_new
                rowDGV.Cells(Me.Col17PPH21.Index).Value = row._17PPh21
                rowDGV.Cells(Me.Col17PPH21Rutin.Index).Value = row._17_PPH_21_Rutin
                rowDGV.Cells(Me.Col18NewPPh21YgTelahDipotongMasaPjkSblmnya.Index).Value = row._18PPh21YgTelahDipotongMasaPjkSblmnya_new
                rowDGV.Cells(Me.Col18PPh21YgTelahDipotongMasaPjkSblmnya.Index).Value = row._18PPh21YgTelahDipotongMasaPjkSblmnya

                rowDGV.Cells(Me.Col19NewPPhTerutang.Index).Value = row._19PPhTerhutang_new
                rowDGV.Cells(Me.Col19NewPPhTerutangDisetahunkan.Index).Value = row._19_PPH_Terhutang_Setahun_new
                rowDGV.Cells(Me.Col19NewPPHTerutangDisetahunkanRutin.Index).Value = row._19_PPH_Terhutang_Setahun_Rutin_new
                rowDGV.Cells(Me.Col19NewPPhTerutangSebulanNonRutin.Index).Value = row._19_PPH_Terhutang_Sebulan_Non_Rutin_new
                rowDGV.Cells(Me.Col19NewPPhTerutangSebulanRutin.Index).Value = row._19_PPH_Terutang_Sebulan_Rutin_new
                rowDGV.Cells(Me.Col19PPHTerutang.Index).Value = row._19PPhTerhutang
                rowDGV.Cells(Me.Col19PPHTerutangDisetahunakan.Index).Value = row._19_PPH_Terhutang_Setahun
                rowDGV.Cells(Me.Col19PPhTerutangDisetahunkanRutin.Index).Value = row._19_PPH_Terhutang_Setahun_Rutin
                rowDGV.Cells(Me.Col19PPHTerutangSebulanNonRutin.Index).Value = row._19_PPH_Terhutang_Sebulan_Non_Rutin
                rowDGV.Cells(Me.Col19PPhTerutangSebulanRutin.Index).Value = row._19_PPH_Terutang_Sebulan_Rutin

                rowDGV.Cells(Me.Col21Pesangon.Index).Value = row._21_Pesangon
                rowDGV.Cells(Me.Col21NewPPhPesangon.Index).Value = row._21_PPH_Final_Pesangon_new
                rowDGV.Cells(Me.Col21NewTarifPesangon.Index).Value = row.Tarif_Pesangon_new
                rowDGV.Cells(Me.Col21NewTunjPesangon.Index).Value = row._21_Tunj_PPH_FInal_Pesangon_new
                rowDGV.Cells(Me.Col21PPhPesangon.Index).Value = row._21_PPH_Final_Pesangon
                rowDGV.Cells(Me.Col21TarifPesagon.Index).Value = row.Tarif_Pesangon
                rowDGV.Cells(Me.Col21TunjPPhPesangon.Index).Value = row._21_Tunj_PPH_FInal_Pesangon

                rowDGV.Cells(Me.ColAmount.Index).Value = row.Income_Tax
                rowDGV.Cells(Me.ColRecalculatedAmount.Index).Value = row.Income_Tax_New
                rowDGV.Cells(Me.ColDifference.Index).Value = row.Income_Tax_New - row.Income_Tax

                rowDGV.Cells(Me.ColAlamat1.Index).Value = row.A_Alamat1
                rowDGV.Cells(Me.ColAlamat2.Index).Value = row.A_Alamat2
                rowDGV.Cells(Me.ColDate.Index).Value = row.Tanggal_Transaksi
                rowDGV.Cells(Me.ColDocNo.Index).Value = row.No_Transaksi
                rowDGV.Cells(Me.ColEmployee.Index).Value = row.Full_Name
                rowDGV.Cells(Me.ColIncomeTax.Index).Value = row.Income_Tax
                rowDGV.Cells(Me.ColKurangBayar.Index).Value = row.Kurang_Bayar

                rowDGV.Cells(Me.ColNewAlamat1.Index).Value = row.A_Alamat1_new
                rowDGV.Cells(Me.ColNewAlamat2.Index).Value = row.A_Alamat2_new
                rowDGV.Cells(Me.ColNewIncomeTax.Index).Value = row.Income_Tax_New
                rowDGV.Cells(Me.ColNewKurangBayar.Index).Value = row.Kurang_Bayar_New
                rowDGV.Cells(Me.ColNewNPWP.Index).Value = row.A_NPWP_new
                rowDGV.Cells(Me.ColNewPenghasilanBruto.Index).Value = row.Penghasilan_Bruto_new
                rowDGV.Cells(Me.ColNewStatusNPWP.Index).Value = row.A_Status_NPWP_new
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                rowDGV.Cells(Me.ColNPWP.Index).Value = row.A_NPWP
                rowDGV.Cells(Me.ColPenghasilBruto.Index).Value = row.Penghasilan_Bruto
                rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value = row.primary_employee
                rowDGV.Cells(Me.ColPrimarySalary.Index).Value = row.primary_salary
                rowDGV.Cells(Me.ColStatusNPWP.Index).Value = row.A_Status_NPWP
                rowDGV.Cells(Me.ColTipePajak.Index).Value = row.Tipe_Tax
            Next

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.ColAmount.DefaultCellStyle.Format = "N0"
            Me.ColRecalculatedAmount.DefaultCellStyle.Format = "N0"
            Me.ColDifference.DefaultCellStyle.Format = "N0"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColRecalculatedAmount.Index Then
                    Me.TampilkanDetailPerhitunganBaru(e.RowIndex)
                ElseIf e.ColumnIndex = Me.ColAmount.Index Then
                    Me.TampilkanDetailPerhitunganLama(e.RowIndex)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TampilkanDetailPerhitunganLama(ByVal vIndex As Integer)
        Dim k As New FrmTrans_Gaji_IncomeTax2
        Dim ResignDate As Date = #1/1/1950#
        Dim StartWorkingPeriod As Date = #1/1/1950#
        Dim rowDGV As DataGridViewRow
        Dim TBLItemTax As New DseTransaksiSalary.tbltrans_salary_itemtaxDataTable
        Dim row2 As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Dim StartDate, EndDate As Date
        Dim row1 As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxRow

        Try
            rowDGV = Me.dgvItem.Rows(vIndex)

            If Me.dgvItem.RowCount > 0 Then
                ResignDate = Me.ObjKryw.GetResignDate(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value)
                StartDate = Me.ObjSalary.GetStartDate(rowDGV.Cells(Me.ColPrimarySalary.Index).Value)
                EndDate = Me.ObjSalary.GetEndDate(rowDGV.Cells(Me.ColPrimarySalary.Index).Value)

                Me.vTBLItemTax.DefaultView.RowFilter = "[primary salary] = " + rowDGV.Cells(Me.ColPrimarySalary.Index).Value.ToString + _
                " AND [tipe item tax] = " + CInt(ClsTrans_TaxRecalculateItem.EnumTipeItemTax.B4_Recalculate).ToString
                For Each DV As DataRowView In Me.vTBLItemTax.DefaultView
                    row1 = DV.Row
                    row2 = TBLItemTax.Newtbltrans_salary_itemtaxRow
                    With row2
                        .amount_gabungan_disetahunkan = row1.amount_gabungan_disetahunkan
                        .amount_month = row1.amount_month
                        .amount_rutin_disetahunkan = row1.amount_rutin_disetahunkan
                        .identifier_salary_component = row1.identifier_salary_component
                        .is_rutin = row1.is_rutin
                        .No_Transaksi = rowDGV.Cells(Me.ColDocNo.Index).Value
                        .primary = TBLItemTax.Rows.Count + 1
                        .Primary_Employee = rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value
                        .primary_transaksi = rowDGV.Cells(Me.ColPrimarySalary.Index).Value
                        .remarks_gabungan = row1.remarks_gabungan
                        .remarks_rutin = row1.remarks_rutin
                        .Tanggal_Transaksi = rowDGV.Cells(Me.ColDate.Index).Value
                        .tipe_tax = row1.tipe_tax
                    End With
                    TBLItemTax.Rows.Add(row2)
                Next

                k.vTBLItemTax = TBLItemTax
                k.vEmployeeCode = Me.ObjKryw.GetCodeByPrimary(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value)
                k.vTransactionDate = rowDGV.Cells(Me.ColDate.Index).Value
                k.vEmployeeName = rowDGV.Cells(Me.ColEmployee.Index).Value
                k.vNIK_Kary = Me.ObjKryw.GetNIK(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value)
                k.vNPWP = rowDGV.Cells(Me.ColNPWP.Index).Value
                k.vStatusNPWP = rowDGV.Cells(Me.ColStatusNPWP.Index).Value
                k.vKodeNegaraDomisili = Me.ObjKryw.GetKodeNegara(Me.ObjKryw.GetPrimaryCountry(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value))
                If Me.ObjKryw.GetSex(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value) = Core_Module.ClsMasterKaryawan.EnumSex.Male Then
                    k.vJenisKelamin = "LAKI-LAKI"
                Else
                    k.vJenisKelamin = "PEREMPUAN"
                End If
                k.vAlamatPajak1 = rowDGV.Cells(Me.ColAlamat1.Index).Value
                k.vAlamatPajak2 = rowDGV.Cells(Me.ColAlamat2.Index).Value
                k.vPrimaryEmployee = rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value
                k.vDepartement = Me.ObjKryw.GetDeptDescription(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value)
                k.vPosition = Me.ObjKryw.GetPositionDescription(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value)

                If rowDGV.Cells(Me.ColDate.Index).Value.Month = 12 Or _
                (ResignDate.Date >= StartDate.Date And ResignDate.Date <= EndDate.Date) Then
                    k.vEndOfYear = True
                    StartWorkingPeriod = Me.ObjKryw.GetFirstTaxPeriod(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value)
                    If StartWorkingPeriod.Year < EndDate.Date.Year Then
                        StartWorkingPeriod = Me.ObjCD.GetFisrtDateOfTheYear(EndDate)
                    End If
                    k.vStartPeriod = StartWorkingPeriod
                    If ResignDate = #1/1/1950# Then
                        ResignDate = EndDate.Date
                    End If
                    k.vEndPeriod = ResignDate
                End If
                k.ShowDialog()

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLItemTax = Nothing
        End Try
    End Sub

    Private Sub TampilkanDetailPerhitunganBaru(ByVal vIndex As Integer)
        Dim k As New FrmTrans_Gaji_IncomeTax2
        Dim ResignDate As Date = #1/1/1950#
        Dim StartWorkingPeriod As Date = #1/1/1950#
        Dim rowDGV As DataGridViewRow
        Dim TBLItemTax As New DseTransaksiSalary.tbltrans_salary_itemtaxDataTable
        Dim row2 As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Dim StartDate, EndDate As Date
        Dim row1 As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxRow

        Try
            rowDGV = Me.dgvItem.Rows(vIndex)

            If Me.dgvItem.RowCount > 0 Then
                ResignDate = Me.ObjKryw.GetResignDate(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value)
                StartDate = Me.ObjSalary.GetStartDate(rowDGV.Cells(Me.ColPrimarySalary.Index).Value)
                EndDate = Me.ObjSalary.GetEndDate(rowDGV.Cells(Me.ColPrimarySalary.Index).Value)

                Me.vTBLItemTax.DefaultView.RowFilter = "[primary salary] = " + rowDGV.Cells(Me.ColPrimarySalary.Index).Value.ToString + _
                " AND [tipe item tax] = " + CInt(ClsTrans_TaxRecalculateItem.EnumTipeItemTax.Aft_Recalculate).ToString
                For Each DV As DataRowView In Me.vTBLItemTax.DefaultView
                    row1 = DV.Row
                    row2 = TBLItemTax.Newtbltrans_salary_itemtaxRow
                    With row2
                        .amount_gabungan_disetahunkan = row1.amount_gabungan_disetahunkan
                        .amount_month = row1.amount_month
                        .amount_rutin_disetahunkan = row1.amount_rutin_disetahunkan
                        .identifier_salary_component = row1.identifier_salary_component
                        .is_rutin = row1.is_rutin
                        .No_Transaksi = rowDGV.Cells(Me.ColDocNo.Index).Value
                        .primary = TBLItemTax.Rows.Count + 1
                        .Primary_Employee = rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value
                        .primary_transaksi = rowDGV.Cells(Me.ColPrimarySalary.Index).Value
                        .remarks_gabungan = row1.remarks_gabungan
                        .remarks_rutin = row1.remarks_rutin
                        .Tanggal_Transaksi = rowDGV.Cells(Me.ColDate.Index).Value
                        .tipe_tax = row1.tipe_tax
                    End With
                    TBLItemTax.Rows.Add(row2)
                Next

                k.vTBLItemTax = TBLItemTax
                k.vEmployeeCode = Me.ObjKryw.GetCodeByPrimary(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value)
                k.vTransactionDate = rowDGV.Cells(Me.ColDate.Index).Value
                k.vEmployeeName = rowDGV.Cells(Me.ColEmployee.Index).Value
                k.vNIK_Kary = Me.ObjKryw.GetNIK(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value)
                k.vNPWP = rowDGV.Cells(Me.ColNewNPWP.Index).Value
                k.vStatusNPWP = rowDGV.Cells(Me.ColNewStatusNPWP.Index).Value
                k.vKodeNegaraDomisili = Me.ObjKryw.GetKodeNegara(Me.ObjKryw.GetPrimaryCountry(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value))
                If Me.ObjKryw.GetSex(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value) = Core_Module.ClsMasterKaryawan.EnumSex.Male Then
                    k.vJenisKelamin = "LAKI-LAKI"
                Else
                    k.vJenisKelamin = "PEREMPUAN"
                End If
                k.vAlamatPajak1 = rowDGV.Cells(Me.ColNewAlamat1.Index).Value
                k.vAlamatPajak2 = rowDGV.Cells(Me.ColNewAlamat2.Index).Value
                k.vPrimaryEmployee = rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value
                k.vDepartement = Me.ObjKryw.GetDeptDescription(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value)
                k.vPosition = Me.ObjKryw.GetPositionDescription(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value)

                If rowDGV.Cells(Me.ColDate.Index).Value.Month = 12 Or _
                (ResignDate.Date >= StartDate.Date And ResignDate.Date <= EndDate.Date) Then
                    k.vEndOfYear = True
                    StartWorkingPeriod = Me.ObjKryw.GetFirstTaxPeriod(rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value)
                    If StartWorkingPeriod.Year < EndDate.Date.Year Then
                        StartWorkingPeriod = Me.ObjCD.GetFisrtDateOfTheYear(EndDate)
                    End If
                    k.vStartPeriod = StartWorkingPeriod
                    If ResignDate = #1/1/1950# Then
                        ResignDate = EndDate.Date
                    End If
                    k.vEndPeriod = ResignDate
                End If
                k.ShowDialog()

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLItemTax = Nothing
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class