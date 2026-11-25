Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing
Public Class FrmTrans_Gaji_IncomeTax
    Public ik As FrmTrans_Gaji
    'DIpanggil dari Transaksi Payment
    Public mPrimarySalary As Double = -1
    Public mDipanggilDariPayment As Boolean = False
    Public ijk As FrmTransaksiPayment
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjCD As New Core_Module.ClsCekData
    Dim Def_Prec As Double = 0

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_IncomeTax_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Filter As String = ""
        Dim RowItemSalary As DseTransPayment.tbltrans_salaryRow
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.ReadOnly = True
            If Me.mDipanggilDariPayment = False Then
                Me.TxtNomor.Text = ik.Txt_P_NoForm.Text
                Me.TxtNPWP.Text = ik.Txt_P_A_NPWP.Text
                Me.TxtNIK.Text = ik.Txt_P_NoNIK.Text
                Me.TxtNama.Text = ik.Txt_P_ANama.Text
                Me.TxtAlamat1.Text = ik.Txt_P_Alamat1.Text
                Me.TxtJenisKelamin.Text = ik.Txt_P_JenisKelamin.Text
                If ik.Txt_P_StatusNPWP.Text = "TK0" Then
                    Me.CboStatusPTKP.SelectedIndex = 0
                ElseIf ik.Txt_P_StatusNPWP.Text = "TK1" Then
                    Me.CboStatusPTKP.SelectedIndex = 1
                ElseIf ik.Txt_P_StatusNPWP.Text = "TK2" Then
                    Me.CboStatusPTKP.SelectedIndex = 2
                ElseIf ik.Txt_P_StatusNPWP.Text = "TK3" Then
                    Me.CboStatusPTKP.SelectedIndex = 3
                ElseIf ik.Txt_P_StatusNPWP.Text = "K0" Then
                    Me.CboStatusPTKP.SelectedIndex = 4
                ElseIf ik.Txt_P_StatusNPWP.Text = "K1" Then
                    Me.CboStatusPTKP.SelectedIndex = 5
                ElseIf ik.Txt_P_StatusNPWP.Text = "K2" Then
                    Me.CboStatusPTKP.SelectedIndex = 6
                ElseIf ik.Txt_P_StatusNPWP.Text = "K3" Then
                    Me.CboStatusPTKP.SelectedIndex = 7
                End If
                Me.TxtNamaJabatan.Text = ik.Txt_P_NamaJabatan.Text
                If ik.Txt_P_KaryawanAsing.Text = "YA" Then
                    Me.ChkKaryawanAsing.Checked = True
                Else
                    Me.ChkKaryawanAsing.Checked = False
                End If
                Me.TxtKodeKaryawanAsing.Text = ik.Txt_P_KodeNegaraDomisili.Text
                '3
                Me.TxtNPWPPemotong.Text = ik.Txt_P_NPWP.Text
                Me.TxtNamaPemotong.Text = ik.Txt_P_Nama.Text
                Me.DtpTanggalTransaksi.Value = ik.DtpDate.Value
                'DGV Item
                Me.BuatRowDGV("", "PENGHASILAN BRUTO", "")
                Me.BuatRowDGV("1", "GAJI/PENSIUNAN ATAU THT/JHT", Me.ObjCD.DigitGrouping(ik.Txt01Pjk_Gaji.Text, Me.Def_Prec))
                Me.BuatRowDGV("2", "TUNJANGAN PPh", Me.ObjCD.DigitGrouping(ik.Txt02PjkTunjanganPPh.Text, Me.Def_Prec))
                Me.BuatRowDGV("3", "TUNJANGAN LAINNYA, UANG LEMBUR, DAN SEBAGAINYA", Me.ObjCD.DigitGrouping(ik.Txt03Pjk_TunjLainnya.Text, Me.Def_Prec))
                Me.BuatRowDGV("4", "HONORARIUM DAN IMBALAN LAIN SEJENISNYA", Me.ObjCD.DigitGrouping(ik.Txt04PjkHonorarium.Text, Me.Def_Prec))
                Me.BuatRowDGV("5", "PREMI ASURANSI YANG DIBAYARKAN PEMBERI KERJA", Me.ObjCD.DigitGrouping(ik.Txt05PjkPremiAsuransi.Text, Me.Def_Prec))
                Me.BuatRowDGV("6", "PENERIMAAN DALAM BENTUK NATURA DAN KENIKMATAN LAINNYA YANG DIKENAKAN PEMOTONGAN PPh PASAL 21", Me.ObjCD.DigitGrouping(ik.Txt06PjkNatura.Text, Me.Def_Prec))
                Me.BuatRowDGV("7", "TANTIEM, BONUS,GRATIFIKASI, JASA PRODUKSI, DAN THR", Me.ObjCD.DigitGrouping(ik.Txt07PjkTantiem.Text, Me.Def_Prec))
                Me.BuatRowDGV("8", "JUMLAH PENGHASILAN BRUTO", Me.ObjCD.DigitGrouping(ik.Txt08PjkPenghasilanBruto.Text, Me.Def_Prec))
                Me.BuatRowDGV("", "PENGURANGAN", "")
                Me.BuatRowDGV("9", "BIAYA JABATAN / BIAYA PENSIUN", Me.ObjCD.DigitGrouping(ik.Txt09PjkBiayaJabatan.Text, Me.Def_Prec))
                Me.BuatRowDGV("10", "IURAN PENSIUN ATAU IURAN THT / JHT", Me.ObjCD.DigitGrouping(ik.Txt10PjkIuranPensiun.Text, Me.Def_Prec))
                Me.BuatRowDGV("11", "JUMLAH PENGURANGAN", Me.ObjCD.DigitGrouping(ik.Txt11PjkJmlPengurangan.Text, Me.Def_Prec))
                Me.BuatRowDGV("", "PENGHITUNGAN PPh PASAL 21", "")
                Me.BuatRowDGV("12", "JUMLAH PENGHASILAN NETTO", Me.ObjCD.DigitGrouping(ik.Txt12PjkJmlPenghasilanNetto.Text, Me.Def_Prec))
                Me.BuatRowDGV("13", "PENGHASILAN NETTO MASA PAJAK SEBELUMNYA", Me.ObjCD.DigitGrouping(ik.Txt13PjkJmlHasilMsPjkSblm.Text, Me.Def_Prec))
                Me.BuatRowDGV("14", "JUMLAH PENGHASILAN NETTO UNTUK PERHITUNGAN PPh PASAL 21", Me.ObjCD.DigitGrouping(ik.Txt14PjkJmlHasilNettoUPPh.Text, Me.Def_Prec))
                Me.BuatRowDGV("15", "PENGHASILAN TIDAK KENA PAJAK", Me.ObjCD.DigitGrouping(ik.Txt15Pjk_PTKP.Text, Me.Def_Prec))
                Me.BuatRowDGV("16", "PENGHASILAN KENA PAJAK", Me.ObjCD.DigitGrouping(ik.Txt16Pjk_PKPSetahun.Text, Me.Def_Prec))
                Me.BuatRowDGV("17", "PENGHASILAN KENA PAJAK SETAHUN / DISETAHUNKAN", Me.ObjCD.DigitGrouping(ik.Txt17Pjk_PPH21.Text, Me.Def_Prec))
                Me.BuatRowDGV("18", "PPh PASAL 21 YANG TELAH DIPOTONG MASA PAJAK SEBELUMNYA", Me.ObjCD.DigitGrouping(ik.Txt18.Text, Me.Def_Prec))
                Me.BuatRowDGV("19", "PPh PASAL 21 TERHUTANG", Me.ObjCD.DigitGrouping(ik.Txt19Pjk_PPH21Terutang.Text, Me.Def_Prec))
                Me.BuatRowDGV("20", "PPh PASAL 21 DAN PASAL 26 YANG TELAH DIPOTONG ATAU LUNAS", Me.ObjCD.DigitGrouping(ik.Txt20.Text, Me.Def_Prec))
            Else
                Filter = "[Primary] = " + Me.mPrimarySalary.ToString + ""
                ijk.TBLSalary.DefaultView.RowFilter = Filter
                If ijk.TBLSalary.DefaultView.Count > 0 Then
                    RowItemSalary = ijk.TBLSalary.DefaultView(0).Row
                    Me.TxtNomor.Text = RowItemSalary.NoForm1721A1
                    Me.TxtNPWP.Text = RowItemSalary.A_NPWP
                    Me.TxtNIK.Text = RowItemSalary.A_NIK
                    Me.TxtNama.Text = RowItemSalary.A_Nama
                    Me.TxtAlamat1.Text = RowItemSalary.A_Alamat1
                    If RowItemSalary.A_Jenis_Kelamin = Core_Module.ClsMasterKaryawan.EnumSex.Female Then
                        Me.TxtJenisKelamin.Text = "PEREMPUAN"
                    Else
                        Me.TxtJenisKelamin.Text = "LAKI-LAKI"
                    End If
                    If RowItemSalary.A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.TK0 Then
                        Me.CboStatusPTKP.SelectedIndex = 0
                    ElseIf RowItemSalary.A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.TK1 Then
                        Me.CboStatusPTKP.SelectedIndex = 1
                    ElseIf RowItemSalary.A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.TK2 Then
                        Me.CboStatusPTKP.SelectedIndex = 2
                    ElseIf RowItemSalary.A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.TK3 Then
                        Me.CboStatusPTKP.SelectedIndex = 3
                    ElseIf RowItemSalary.A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.K0 Then
                        Me.CboStatusPTKP.SelectedIndex = 4
                    ElseIf RowItemSalary.A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.K1 Then
                        Me.CboStatusPTKP.SelectedIndex = 5
                    ElseIf RowItemSalary.A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.K2 Then
                        Me.CboStatusPTKP.SelectedIndex = 6
                    ElseIf RowItemSalary.A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.K3 Then
                        Me.CboStatusPTKP.SelectedIndex = 7
                    End If
                    Me.TxtNamaJabatan.Text = RowItemSalary.A_Nama_Jabatan
                    If RowItemSalary.A_Karyawan_Asing = "YA" Then
                        Me.ChkKaryawanAsing.Checked = True
                    Else
                        Me.ChkKaryawanAsing.Checked = False
                    End If
                    Me.TxtKodeKaryawanAsing.Text = RowItemSalary.A_Kode_Negara_Domisili
                    '3
                    Me.TxtNPWPPemotong.Text = RowItemSalary.C_NPWP
                    Me.TxtNamaPemotong.Text = RowItemSalary.C_Nama
                    Me.DtpTanggalTransaksi.Value = RowItemSalary.Tanggal_Transaksi
                    'DGV Item
                    Me.BuatRowDGV("", "PENGHASILAN BRUTO", "")
                    Me.BuatRowDGV("1", "GAJI/PENSIUNAN ATAU THT/JHT", Me.ObjCD.DigitGrouping(RowItemSalary._01Gaji, Me.Def_Prec))
                    Me.BuatRowDGV("2", "TUNJANGAN PPh", Me.ObjCD.DigitGrouping(RowItemSalary._02TunjanganPPh, Me.Def_Prec))
                    Me.BuatRowDGV("3", "TUNJANGAN LAINNYA, UANG LEMBUR, DAN SEBAGAINYA", Me.ObjCD.DigitGrouping(RowItemSalary._03TunjanganLainnya, Me.Def_Prec))
                    Me.BuatRowDGV("4", "HONORARIUM DAN IMBALAN LAIN SEJENISNYA", Me.ObjCD.DigitGrouping(RowItemSalary._04Honorarium, Me.Def_Prec))
                    Me.BuatRowDGV("5", "PREMI ASURANSI YANG DIBAYARKAN PEMBERI KERJA", Me.ObjCD.DigitGrouping(RowItemSalary._05PremiAsuransi, Me.Def_Prec))
                    Me.BuatRowDGV("6", "PENERIMAAN DALAM BENTUK NATURA DAN KENRowItemSalaryMATAN LAINNYA YANG DRowItemSalaryENAKAN PEMOTONGAN PPh PASAL 21", Me.ObjCD.DigitGrouping(RowItemSalary._06Natura, Me.Def_Prec))
                    Me.BuatRowDGV("7", "TANTIEM, BONUS,GRATIFRowItemSalaryASI, JASA PRODUKSI, DAN THR", Me.ObjCD.DigitGrouping(RowItemSalary._07Tantiem, Me.Def_Prec))
                    Me.BuatRowDGV("8", "JUMLAH PENGHASILAN BRUTO", Me.ObjCD.DigitGrouping(RowItemSalary._08PenghasilanBruto, Me.Def_Prec))
                    Me.BuatRowDGV("", "PENGURANGAN", "")
                    Me.BuatRowDGV("9", "BIAYA JABATAN / BIAYA PENSIUN", Me.ObjCD.DigitGrouping(RowItemSalary._09BiayaJabatan, Me.Def_Prec))
                    Me.BuatRowDGV("10", "IURAN PENSIUN ATAU IURAN THT / JHT", Me.ObjCD.DigitGrouping(RowItemSalary._10IuranPensiun, Me.Def_Prec))
                    Me.BuatRowDGV("11", "JUMLAH PENGURANGAN", Me.ObjCD.DigitGrouping(RowItemSalary._11JumlahPengurangan, Me.Def_Prec))
                    Me.BuatRowDGV("", "PENGHITUNGAN PPh PASAL 21", "")
                    Me.BuatRowDGV("12", "JUMLAH PENGHASILAN NETTO", Me.ObjCD.DigitGrouping(RowItemSalary._12JumlahPenghasilanNetto, Me.Def_Prec))
                    Me.BuatRowDGV("13", "PENGHASILAN NETTO MASA PAJAK SEBELUMNYA", Me.ObjCD.DigitGrouping(RowItemSalary._13JumlahPenghasilanMasaPjkSblmnya, Me.Def_Prec))
                    Me.BuatRowDGV("14", "JUMLAH PENGHASILAN NETTO UNTUK PERHITUNGAN PPh PASAL 21", Me.ObjCD.DigitGrouping(RowItemSalary._14JumlahPenghasilanNettoUtkPPh, Me.Def_Prec))
                    Me.BuatRowDGV("15", "PENGHASILAN TIDAK KENA PAJAK", Me.ObjCD.DigitGrouping(RowItemSalary._15PTKP, Me.Def_Prec))
                    Me.BuatRowDGV("16", "PENGHASILAN KENA PAJAK", Me.ObjCD.DigitGrouping(RowItemSalary._16PKPSetahun, Me.Def_Prec))
                    Me.BuatRowDGV("17", "PENGHASILAN KENA PAJAK SETAHUN / DISETAHUNKAN", Me.ObjCD.DigitGrouping(RowItemSalary._17PPh21, Me.Def_Prec))
                    Me.BuatRowDGV("18", "PPh PASAL 21 YANG TELAH DIPOTONG MASA PAJAK SEBELUMNYA", Me.ObjCD.DigitGrouping(RowItemSalary._18PPh21YgTelahDipotongMasaPjkSblmnya, Me.Def_Prec))
                    Me.BuatRowDGV("19", "PPh PASAL 21 TERHUTANG", Me.ObjCD.DigitGrouping(RowItemSalary._19PPhTerhutang, Me.Def_Prec))
                    Me.BuatRowDGV("20", "PPh PASAL 21 DAN PASAL 26 YANG TELAH DIPOTONG ATAU LUNAS", Me.ObjCD.DigitGrouping(RowItemSalary._20PPhPs21DanPPhPs26YgTelahDipotongNLunas, Me.Def_Prec))
                End If
                ijk.TBLSalary.DefaultView.RowFilter = ""
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatRowDGV(ByVal vNo As String, ByVal vUraian As String, ByVal vAmount As String)
        Dim RowDGV As DataGridViewRow
        Try
            Me.dgvItem.Rows.Add()
            RowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
            RowDGV.Cells(Me.ColNo.Index).Value = vNo
            RowDGV.Cells(Me.ColUraian.Index).Value = vUraian
            RowDGV.Cells(Me.ColAmount.Index).Value = vAmount
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class