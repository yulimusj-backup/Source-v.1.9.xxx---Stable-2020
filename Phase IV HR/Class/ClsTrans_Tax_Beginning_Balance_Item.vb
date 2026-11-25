Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class ClsTrans_Tax_Beginning_Balance_Item
    Public Enum EnumCriteria As Integer
        TK0
        TK1
        TK2
        TK3
        K0
        K1
        K2
        K3
    End Enum

    Public Function IsiDGV(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal vDef_Prec As Double, ByVal i As FrmTrans_BeginingBalanceTax) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseTrans_Tax_Beginning_BalanceTableAdapters.DataTable2TableAdapter
        Dim TBLItem As New DseTrans_Tax_Beginning_Balance.DataTable2DataTable
        Dim RowItem As DseTrans_Tax_Beginning_Balance.DataTable2Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItem.AllowUserToAddRows = False

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item overhead cost
            DTAItem.Fill(TBLItem, vPrimary)
            vDGVItem.Rows.Clear()
            For Each RowItem In TBLItem.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColEmp_No.Index).Value = ""
                    .Cells(i.ColKodeKaryawan.Index).Value = RowItem.Kode_Karyawan
                    .Cells(i.ColNamaKaryawan.Index).Value = RowItem.Full_Name
                    CboDgv = .Cells(i.ColStatus.Index)
                    .Cells(i.ColStatus.Index).Value = CboDgv.Items.Item(RowItem.Category)
                    .Cells(i.ColNoNPWP.Index).Value = RowItem.No_NPWP
                    .Cells(i.ColNIK.Index).Value = RowItem.NIK
                    .Cells(i.ColDepartement.Index).Value = RowItem.Departemen
                    .Cells(i.ColPosition.Index).Value = RowItem.Position
                    .Cells(i.ColGaji.Index).Value = ObjCD.DigitGrouping(RowItem.Gaji_Pensiun, vDef_Prec)
                    .Cells(i.ColTunjanganPPH.Index).Value = ObjCD.DigitGrouping(RowItem.Tunjangan_PPh, vDef_Prec)
                    .Cells(i.ColTunjanganLainnya.Index).Value = ObjCD.DigitGrouping(RowItem.Tunjangan_Lainnya, vDef_Prec)
                    .Cells(i.ColHonor.Index).Value = ObjCD.DigitGrouping(RowItem.Honorarium, vDef_Prec)
                    .Cells(i.ColPremi.Index).Value = ObjCD.DigitGrouping(RowItem.Premi_Asuransi, vDef_Prec)
                    .Cells(i.ColNatura.Index).Value = ObjCD.DigitGrouping(RowItem.Penerimaan_Natura, vDef_Prec)
                    .Cells(i.ColTantiem.Index).Value = ObjCD.DigitGrouping(RowItem.Tantiem_Bonus, vDef_Prec)
                    .Cells(i.ColPenghasilBruto.Index).Value = ObjCD.DigitGrouping(RowItem.Jumlah_Penghasilan_Bruto, vDef_Prec)
                    .Cells(i.ColBiayaJabatan.Index).Value = ObjCD.DigitGrouping(RowItem.Biaya_Jabatan, vDef_Prec)
                    .Cells(i.ColIuranPensiun.Index).Value = ObjCD.DigitGrouping(RowItem.Iuran_Pensiun, vDef_Prec)
                    .Cells(i.ColJumlahPengurangan.Index).Value = ObjCD.DigitGrouping(RowItem.Jumlah_Pengurangan, vDef_Prec)
                    .Cells(i.ColJumlahPenghasilanNetto.Index).Value = ObjCD.DigitGrouping(RowItem.Jumlah_Penghasilan_Netto, vDef_Prec)
                    .Cells(i.ColJumlahPenghasilNettoMasaPajakSebelum.Index).Value = ObjCD.DigitGrouping(RowItem.Penghasilan_Netto_Masa_Pjk_Sblm, vDef_Prec)
                    .Cells(i.ColJumlahPenghasilanNettoUntukPPH.Index).Value = ObjCD.DigitGrouping(RowItem.Perhitungan_PPh_Pasal_21_Disetahunkan, vDef_Prec)
                    .Cells(i.ColPTKP.Index).Value = ObjCD.DigitGrouping(RowItem.PTKP, vDef_Prec)
                    .Cells(i.ColPenghasilanKenaPajak.Index).Value = ObjCD.DigitGrouping(RowItem.PKP, vDef_Prec)
                    .Cells(i.ColPPhPs21AtasPKP.Index).Value = ObjCD.DigitGrouping(RowItem.PPh_Ps_21_Atas_PKP_Disetahunkan, vDef_Prec)
                    .Cells(i.ColPPHps21TelahDipotong.Index).Value = ObjCD.DigitGrouping(RowItem.PPh_Ps_21_Dipotong_Ms_Pjk_Sblmnya, vDef_Prec)
                    .Cells(i.CollPPHTerutang.Index).Value = ObjCD.DigitGrouping(RowItem.PPh_Ps_21_Terhutang, vDef_Prec)
                    .Cells(i.ColPPhTelahDipotong.Index).Value = ObjCD.DigitGrouping(RowItem.PPh_Ps_21_Yg_Lunas, vDef_Prec)
                    .Cells(i.ColRemarks.Index).Value = RowItem.Remarks
                    .Cells(i.ColPrimaryKaryawan.Index).Value = RowItem.Primary_Karyawan
                End With
            Next

            'set number
            ObjCD.SetNumber(vDGVItem, True)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGVItem.AllowUserToAddRows = True

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimary As Double) As Boolean
        Dim DTAItem As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balance_itemTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimary As Double, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_BeginingBalanceTax) As Boolean
        Dim DTAItem As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balance_itemTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGVItem.AllowUserToAddRows = False
            For Each rowDGV In vDGVItem.Rows
                With rowDGV
                    If String.IsNullOrEmpty(.Cells(i.ColRemarks.Index).Value) = True Then
                        .Cells(i.ColRemarks.Index).Value = ""
                    End If
                    CboDgv = .Cells(i.ColStatus.Index)
                    DTAItem.Insert(vPrimary, .Cells(i.ColPrimaryKaryawan.Index).Value, _
                     CboDgv.Items.IndexOf(.Cells(i.ColStatus.Index).Value), .Cells(i.ColNoNPWP.Index).Value, .Cells(i.ColGaji.Index).Value, _
                     .Cells(i.ColTunjanganPPH.Index).Value, .Cells(i.ColTunjanganLainnya.Index).Value, _
                     .Cells(i.ColHonor.Index).Value, .Cells(i.ColPremi.Index).Value, _
                     .Cells(i.ColNatura.Index).Value, .Cells(i.ColTantiem.Index).Value, _
                     .Cells(i.ColPenghasilBruto.Index).Value, .Cells(i.ColBiayaJabatan.Index).Value, _
                     .Cells(i.ColIuranPensiun.Index).Value, .Cells(i.ColJumlahPengurangan.Index).Value, _
                     .Cells(i.ColJumlahPenghasilanNetto.Index).Value, .Cells(i.ColJumlahPenghasilNettoMasaPajakSebelum.Index).Value, _
                     .Cells(i.ColJumlahPenghasilanNettoUntukPPH.Index).Value, .Cells(i.ColPTKP.Index).Value, .Cells(i.ColPenghasilanKenaPajak.Index).Value, _
                     .Cells(i.ColPPhPs21AtasPKP.Index).Value, .Cells(i.ColPPHps21TelahDipotong.Index).Value, .Cells(i.CollPPHTerutang.Index).Value, _
                     .Cells(i.ColPPhTelahDipotong.Index).Value, .Cells(i.ColRemarks.Index).Value)
                End With
            Next
            vDGVItem.AllowUserToAddRows = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function
End Class
