Public Class FrmTrans_TaxRecalculate_Update
    Public TipeTransaksi As EnumTipeTransaksi = EnumTipeTransaksi.SimpanBaru
    Public PrimaryTransaksi As Double = -1
    Public vFrm As FrmTrans_TaxRecalculate

    Public Enum EnumTipeTransaksi As Integer
        SimpanBaru
        SimpanPerubahan
        Hapus
    End Enum

    Dim Step1 As Integer = 0
    Dim Step2 As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmTrans_TaxRecalculate_Update_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdatePerubahanKeTblSalary(ByVal vPrimary As Double)
        Dim DTAItem As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculate_itemTableAdapter
        Dim TBLItem As New DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_itemDataTable
        Dim rowItem As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_itemRow
        Dim DTAItemTax As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_taxrecalculate_itemtaxTableAdapter
        Dim TBLItemTax As New DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable
        Dim rowItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxRow
        Dim DTAItemTax2 As New DseTransaksiSalaryTableAdapters.tbltrans_salary_itemtaxTableAdapter

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemTax.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemTax2.Connection.ConnectionString = Core_Module.GetConstring

            'ambil data baru
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, Me.PrimaryTransaksi)
            DTAItemTax.FillBy_PrimaryTrans_Tipe(TBLItemTax, Me.PrimaryTransaksi, _
            CInt(ClsTrans_TaxRecalculateItem.EnumTipeItemTax.Aft_Recalculate))

            'untuk setiap baris dalam data baru
            Me.Step1 += 1
            Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + " : Updating Tax Data"
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBLItem.Rows.Count
            For Each rowItem In TBLItem.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 500 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 25) = 0 Then
                        Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + _
                        " : Updating Tax Data (Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                    End If
                Else
                    Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + _
                    " : Updating Tax Data (Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                End If

                '1. update header
                With rowItem
                    DTAItem.UpdateQuery(.Income_Tax_New, ._02TunjanganPPh_new, ._02_Tunjangan_PPH_Non_Rutin_new, ._02_Tunjangan_PPH_Rutin_new, _
                    ._02_Tunjangan_PPH_Disetahunkan_new, ._02_Tunjangan_PPH_Disetahunkan_Rutin_new, ._08PenghasilanBruto_new, _
                    ._08_Penghasilan_Bruto_Rutin_new, ._09BiayaJabatan_new, ._09_Biaya_Jabatan_Rutin_new, ._11JumlahPengurangan_new, _
                    ._11_Jumlah_Pengurang_Rutin_new, ._14JumlahPenghasilanNettoUtkPPh_new, ._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin_new, _
                    ._15PTKP_new, ._16PKPSetahun_new, ._16_PKP_Setahun_Rutin_new, ._17PPh21_new, ._17_PPH_21_Rutin_new, ._18PPh21YgTelahDipotongMasaPjkSblmnya_new, _
                    ._19PPhTerhutang_new, ._19_PPH_Terhutang_Setahun_new, ._19_PPH_Terhutang_Setahun_Rutin_new, ._19_PPH_Terhutang_Sebulan_Non_Rutin_new, _
                    ._19_PPH_Terutang_Sebulan_Rutin_new, .A_NPWP_new, .A_Alamat1_new, .A_Alamat2_new, .A_Status_NPWP_new, .Kurang_Bayar_New, _
                    ._21_Tunj_PPH_FInal_Pesangon_new, ._21_PPH_Final_Pesangon_new, .Tarif_Pesangon_new, .Penghasilan_Bruto_new, .primary_salary)
                End With

                '2. update item tax
                '2.a. hapuskan item tax lama
                DTAItemTax2.Delete(rowItem.primary_salary)
                '2.b. masukkan item tax baru
                TBLItemTax.DefaultView.RowFilter = "[primary salary] = " + rowItem.primary_salary.ToString + _
               " AND [tipe item tax] = " + CInt(ClsTrans_TaxRecalculateItem.EnumTipeItemTax.Aft_Recalculate).ToString
                For Each DV As DataRowView In TBLItemTax.DefaultView
                    rowItemTax = DV.Row
                    With rowItemTax
                        DTAItemTax2.Insert(.primary_salary, .tipe_tax, .identifier_salary_component, .is_rutin, .amount_month, _
                        .amount_gabungan_disetahunkan, .remarks_gabungan, .amount_rutin_disetahunkan, .remarks_rutin)
                    End With
                Next
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            DTAItemTax.Dispose()
            TBLItemTax.Dispose()
            DTAItemTax2.Dispose()
        End Try
    End Sub

    Private Sub KembalikanKondisiTblSalary(ByVal vPrimary As Double, ByVal vListPrimarySalary As List(Of Double))
        Dim DTAItem As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculate_itemTableAdapter
        Dim TBLItem As New DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_itemDataTable
        Dim rowItem As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_itemRow
        Dim DTAItemTax As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_taxrecalculate_itemtaxTableAdapter
        Dim TBLItemTax As New DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable
        Dim rowItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxRow
        Dim DTAItemTax2 As New DseTransaksiSalaryTableAdapters.tbltrans_salary_itemtaxTableAdapter

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemTax.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemTax2.Connection.ConnectionString = Core_Module.GetConstring

            'ambil data baru
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, Me.PrimaryTransaksi)
            DTAItemTax.FillBy_PrimaryTrans_Tipe(TBLItemTax, Me.PrimaryTransaksi, _
            CInt(ClsTrans_TaxRecalculateItem.EnumTipeItemTax.B4_Recalculate))

            'untuk setiap baris dalam data baru
            'Me.Step2 = 1

            Me.Step1 += 1
            Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + " : Updating Tax Data"
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBLItem.Rows.Count
            For Each rowItem In TBLItem.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 500 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 25) = 0 Then
                        Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + _
                        " : Updating Tax Data (Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                    End If
                Else
                    Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + _
                    " : Updating Tax Data (Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                End If

                If vListPrimarySalary.Contains(rowItem.primary_salary) = True Then
                    '1. update header
                    With rowItem
                        DTAItem.UpdateQuery(.Income_Tax, ._02TunjanganPPh, ._02_Tunjangan_PPH_Non_Rutin, ._02_Tunjangan_PPH_Rutin, _
                        ._02_Tunjangan_PPH_Disetahunkan, ._02_Tunjangan_PPH_Disetahunkan_Rutin, ._08PenghasilanBruto, _
                        ._08_Penghasilan_Bruto_Rutin, ._09BiayaJabatan, ._09_Biaya_Jabatan_Rutin, ._11JumlahPengurangan, _
                        ._11_Jumlah_Pengurang_Rutin, ._14JumlahPenghasilanNettoUtkPPh, ._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin, _
                        ._15PTKP, ._16PKPSetahun, ._16_PKP_Setahun_Rutin, ._17PPh21, ._17_PPH_21_Rutin, ._18PPh21YgTelahDipotongMasaPjkSblmnya, _
                        ._19PPhTerhutang, ._19_PPH_Terhutang_Setahun, ._19_PPH_Terhutang_Setahun_Rutin, ._19_PPH_Terhutang_Sebulan_Non_Rutin, _
                        ._19_PPH_Terutang_Sebulan_Rutin, .A_NPWP, .A_Alamat1, .A_Alamat2, .A_Status_NPWP, .Kurang_Bayar, _
                        ._21_Tunj_PPH_FInal_Pesangon, ._21_PPH_Final_Pesangon, .Tarif_Pesangon, .Penghasilan_Bruto, .primary_salary)
                    End With

                    '2. update item tax
                    '2.a. hapuskan item tax lama
                    DTAItemTax2.Delete(rowItem.primary_salary)
                    '2.b. masukkan item tax baru
                    TBLItemTax.DefaultView.RowFilter = "[primary salary] = " + rowItem.primary_salary.ToString + _
                   " AND [tipe item tax] = " + CInt(ClsTrans_TaxRecalculateItem.EnumTipeItemTax.B4_Recalculate).ToString
                    For Each DV As DataRowView In TBLItemTax.DefaultView
                        rowItemTax = DV.Row
                        With rowItemTax
                            DTAItemTax2.Insert(.primary_salary, .tipe_tax, .identifier_salary_component, .is_rutin, .amount_month, _
                            .amount_gabungan_disetahunkan, .remarks_gabungan, .amount_rutin_disetahunkan, .remarks_rutin)
                        End With
                    Next
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            DTAItemTax.Dispose()
            TBLItemTax.Dispose()
            DTAItemTax2.Dispose()
        End Try
    End Sub

    Private Sub GetListSalary(ByVal vPrimary As Double, ByVal vListPrimary As List(Of Double))
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculate_itemTableAdapter
        Dim TBL As New DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_itemDataTable
        Dim row As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_itemRow

        Try
            Me.Step1 += 1
            Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + " : Collecting Data"
            Me.ProgressBar1.Value = 0

            vListPrimary.Clear()

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryTransaksi(TBL, vPrimary)
            For Each row In TBL.Rows
                vListPrimary.Add(row.primary_salary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub GetList_SalaryDikembalikanKeKondisiSemula(ByVal vListPrimary As List(Of Double), _
    ByVal vDGV As Windows.Forms.DataGridView, ByVal vFrm As FrmTrans_TaxRecalculate)
        Dim rowDGV As Windows.Forms.DataGridViewRow

        Try
            Me.Step1 += 1
            Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + " : Filtering Data"
            Me.ProgressBar1.Value = 0

            'buang primary salary yang ada di list, tapi ada juga di dgv
            For ij As Double = vListPrimary.Count - 1 To 0 Step -1
                For Each rowDGV In vDGV.Rows
                    If rowDGV.Cells(vFrm.ColPrimarySalary.Index).Value = vListPrimary.Item(ij) Then
                        vListPrimary.RemoveAt(ij)
                        Exit For
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim ListPrimarySalary As New List(Of Double)

        Try
            Me.ObjCD.UpdateThreading()
            If Me.TipeTransaksi = EnumTipeTransaksi.SimpanBaru Then
                Me.Step2 = 1
                Me.UpdatePerubahanKeTblSalary(Me.PrimaryTransaksi)
            ElseIf Me.TipeTransaksi = EnumTipeTransaksi.Hapus Then
                Me.Step2 = 2
                Me.GetListSalary(PrimaryTransaksi, ListPrimarySalary)
                Me.KembalikanKondisiTblSalary(PrimaryTransaksi, ListPrimarySalary)
            ElseIf Me.TipeTransaksi = EnumTipeTransaksi.SimpanPerubahan Then
                Me.Step2 = 4
                Me.GetListSalary(PrimaryTransaksi, ListPrimarySalary)
                Me.GetList_SalaryDikembalikanKeKondisiSemula(ListPrimarySalary, Me.vFrm.dgvItem, Me.vFrm)
                Me.KembalikanKondisiTblSalary(PrimaryTransaksi, ListPrimarySalary)
                Me.UpdatePerubahanKeTblSalary(Me.PrimaryTransaksi)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimarySalary = Nothing
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class