Option Explicit On
Option Strict Off

Public Class FrmTrans_Export_eSPT

    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll

    Dim objCd As New Core_Module.ClsCekData
    Dim ObjTR As New ClsTrans_TaxRecalculate
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim dtaMonthly As New DseTrans_Export_eSPTTableAdapters.tbltrans_salary_monthlyTableAdapter
    Dim tblMonthly As New DseTrans_Export_eSPT.tbltrans_salary_monthlyDataTable
    Dim rowMonthly As DseTrans_Export_eSPT.tbltrans_salary_monthlyRow
    Dim dtaYearly As New DseTrans_Export_eSPTTableAdapters.tbltrans_salary_yearlyTableAdapter
    Dim tblYearly1 As New DseTrans_Export_eSPT.tbltrans_salary_yearlyDataTable
    Dim tblYearly As New DseTrans_Export_eSPT.tbltrans_salary_yearlyDataTable
    Dim rowYearly, rowYearly1 As DseTrans_Export_eSPT.tbltrans_salary_yearlyRow
    Dim dtaPA As New DseTrans_Export_eSPTTableAdapters.TblPayrollAccessTableAdapter
    Dim tblPA As New DseTrans_Export_eSPT.TblPayrollAccessDataTable
    Dim rowPA As DseTrans_Export_eSPT.TblPayrollAccessRow
    Dim DTASF As New DseTrans_Export_eSPTTableAdapters.tblmaster_salary_formulaTableAdapter
    Dim TBLSF As New DseTrans_Export_eSPT.tblmaster_salary_formulaDataTable
    Dim tblMsg As New DseTrans_Export_eSPT.TblMsgBoxDataTable
    Dim rowMsg As DseTrans_Export_eSPT.TblMsgBoxRow
    Dim csvVal As String = ""
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjSalary As New ClsTransaksiSalary
    Dim DummyDate As Date = #1/1/1950#
    Dim Dummy As String = ""
    Dim Buffer As String = ""

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        Try
            If Me.cboFormat.SelectedIndex = 0 Then
                Me.ExportMonthly()
            ElseIf Me.cboFormat.SelectedIndex = 1 Then
                Me.ExportYearly()
            ElseIf Me.cboFormat.SelectedIndex = 2 Then
                Me.ExportPPHFinalPesangon()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataDGVA1()
        Try
            Me.dtaYearly.FillBy_ResignOnly(Me.tblYearly1, Me.objCd.GetLasDateOfTheMonth(Me.objCd.GetLastDateOfTheYear(Me.DtpPeriod.Value) - TimeSpan.FromDays(32)), _
                        Me.objCd.GetFisrtDateOfTheYear(Me.DtpPeriod.Value))
            Me.dtaYearly.FillBy_TanggalLaporan(Me.tblYearly, Me.objCd.GetLastDateOfTheYear(Me.DtpPeriod.Value), Me.objCd.GetFirstDateofTheMonth(Date.ParseExact("01 12 " & Me.DtpPeriod.Value.Year & "", "dd MM yyyy", Nothing)))
            For Each rowYearly In Me.tblYearly.Rows
                rowYearly1 = tblYearly1.Newtbltrans_salary_yearlyRow
                For ij As Integer = 0 To tblYearly1.Columns.Count - 1
                    rowYearly1.Item(ij) = rowYearly.Item(ij)
                Next
                tblYearly1.Rows.Add(rowYearly1)
            Next
            tblYearly.Rows.Clear()
            For Each rowYearly1 In Me.tblYearly1.Rows
                rowYearly = Me.tblYearly.Newtbltrans_salary_yearlyRow
                For ij As Integer = 0 To Me.tblYearly1.Columns.Count - 1
                    rowYearly.Item(ij) = rowYearly1.Item(ij)
                Next
                tblYearly.Rows.Add(rowYearly)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGVA1()
        Dim lengthInit As Integer = 0
        Dim notAllowedVal As String = ""
        Dim rowDGV As Windows.Forms.DataGridViewRow
        Dim Lanjut As Boolean = False

        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            Me.tblYearly.Clear()
            Me.tblYearly1.Rows.Clear()
            Me.tblMsg.Clear()
            Me.csvVal = ""

            Me.dgvA1.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            Me.dgvA1.AutoSizeRowsMode = Windows.Forms.DataGridViewAutoSizeRowsMode.None
            Me.dgvA1.Rows.Clear()

            Me.LoadDataDGVA1()

            If Me.tblYearly.Rows.Count > 0 Then
                For Each Me.rowYearly In Me.tblYearly
                    With Me.rowYearly
                        .Pembetulan = Me.TxtPembetulan.Text
                        If Me.IsAllowed(.Primary_Salary_Formula) = True Then
                            Lanjut = False

                            If Me.chkNotExportedOnly.Checked = True Then
                                If .Export_Status = "-" Or .Export_Status = "" Then
                                    Lanjut = True
                                End If
                            Else
                                Lanjut = True
                            End If
                            If Lanjut = True Then
                                Me.dgvA1.Rows.Add()
                                rowDGV = Me.dgvA1.Rows(Me.dgvA1.Rows.Count - 1)

                                rowDGV.Cells(Me.ColA1_Primary.Index).Value = .Primary
                                rowDGV.Cells(Me.ColA1_No.Index).Value = rowDGV.Index + 1
                                rowDGV.Cells(Me.ColA1_PrimaryEmployee.Index).Value = .Primary_Employee
                                rowDGV.Cells(Me.ColA1_Export.Index).Value = Boolean.FalseString
                                rowDGV.Cells(Me.ColA1_ExportStatus.Index).Value = .Export_Status
                                rowDGV.Cells(Me.ColA1_TglSalary.Index).Value = .Tanggal_Transaksi

                                'If .Period_End.ToString.Length = 1 Then
                                '    rowDGV.Cells(Me.ColA1_MasaPajak.Index).Value = "0" & .Period_End.ToString
                                'Else
                                '    rowDGV.Cells(Me.ColA1_MasaPajak.Index).Value = .Period_End.ToString
                                'End If
                                rowDGV.Cells(Me.ColA1_MasaPajak.Index).Value = "12"
                                rowDGV.Cells(Me.ColA1_TahunPajak.Index).Value = Format(Me.DtpPeriod.Value, "yyyy")
                                rowDGV.Cells(Me.ColA1_Pembetulan.Index).Value = Me.objCd.DigitGrouping(.Pembetulan, 0)
                                rowDGV.Cells(Me.ColA1_NoBuktiPotong.Index).Value = .NoForm1721A1
                                If .Period_Start.ToString.Length = 1 Then
                                    rowDGV.Cells(Me.ColA1_MasaPerolehanAwal.Index).Value = "0" & Me.objCd.DigitGrouping(.Period_Start, 0)
                                Else
                                    rowDGV.Cells(Me.ColA1_MasaPerolehanAwal.Index).Value = Me.objCd.DigitGrouping(.Period_Start, 0)
                                End If
                                If .Period_End.ToString.Length = 1 Then
                                    rowDGV.Cells(Me.ColA1_MasaPerolehanAkhir.Index).Value = "0" & Me.objCd.DigitGrouping(.Period_End, 0)
                                Else
                                    rowDGV.Cells(Me.ColA1_MasaPerolehanAkhir.Index).Value = Me.objCd.DigitGrouping(.Period_End, 0)
                                End If
                                rowDGV.Cells(Me.ColA1_NPWP.Index).Value = .A_NPWP
                                rowDGV.Cells(Me.ColA1_NIK.Index).Value = .A_NIK
                                rowDGV.Cells(Me.ColA1_Nama.Index).Value = .A_Nama
                                rowDGV.Cells(Me.ColA1_Alamat.Index).Value = .Alamat
                                If .A_Jenis_Kelamin = 0 Then
                                    rowDGV.Cells(Me.ColA1_JenisKelamin.Index).Value = "M"
                                ElseIf .A_Jenis_Kelamin = 1 Then
                                    rowDGV.Cells(Me.ColA1_JenisKelamin.Index).Value = "F"
                                End If
                                If .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.TK0 Then
                                    rowDGV.Cells(Me.ColA1_StatusPTKP.Index).Value = "TK"
                                    rowDGV.Cells(Me.ColA1_JumlahTanggungan.Index).Value = "0"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.TK1 Then
                                    rowDGV.Cells(Me.ColA1_StatusPTKP.Index).Value = "TK"
                                    rowDGV.Cells(Me.ColA1_JumlahTanggungan.Index).Value = "1"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.TK2 Then
                                    rowDGV.Cells(Me.ColA1_StatusPTKP.Index).Value = "TK"
                                    rowDGV.Cells(Me.ColA1_JumlahTanggungan.Index).Value = "2"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.TK3 Then
                                    rowDGV.Cells(Me.ColA1_StatusPTKP.Index).Value = "TK"
                                    rowDGV.Cells(Me.ColA1_JumlahTanggungan.Index).Value = "3"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.K0 Then
                                    rowDGV.Cells(Me.ColA1_StatusPTKP.Index).Value = "K"
                                    rowDGV.Cells(Me.ColA1_JumlahTanggungan.Index).Value = "0"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.K1 Then
                                    rowDGV.Cells(Me.ColA1_StatusPTKP.Index).Value = "K"
                                    rowDGV.Cells(Me.ColA1_JumlahTanggungan.Index).Value = "1"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.K2 Then
                                    rowDGV.Cells(Me.ColA1_StatusPTKP.Index).Value = "K"
                                    rowDGV.Cells(Me.ColA1_JumlahTanggungan.Index).Value = "2"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.K3 Then
                                    rowDGV.Cells(Me.ColA1_StatusPTKP.Index).Value = "K"
                                    rowDGV.Cells(Me.ColA1_JumlahTanggungan.Index).Value = "3"
                                End If
                                rowDGV.Cells(Me.ColA1_NamaJabatan.Index).Value = .A_Nama_Jabatan
                                If .A_Kode_Negara_Domisili = "INA" Then
                                    rowDGV.Cells(Me.ColA1_WPLuarNegeri.Index).Value = "N"
                                    rowDGV.Cells(Me.ColA1_KodeNegara.Index).Value = ""
                                Else
                                    rowDGV.Cells(Me.ColA1_WPLuarNegeri.Index).Value = "Y"
                                    rowDGV.Cells(Me.ColA1_KodeNegara.Index).Value = .A_Kode_Negara_Domisili
                                End If
                                rowDGV.Cells(Me.ColA1_KodePajak.Index).Value = .Kode_Pajak

                                rowDGV.Cells(Me.ColA1_Jumlah1.Index).Value = Me.objCd.DigitGrouping(._01Gaji, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah2.Index).Value = Me.objCd.DigitGrouping(._19PPhTerhutang, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah3.Index).Value = Me.objCd.DigitGrouping(._03TunjanganLainnya, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah4.Index).Value = Me.objCd.DigitGrouping(._04Honorarium, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah5.Index).Value = Me.objCd.DigitGrouping(._05PremiAsuransi, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah6.Index).Value = Me.objCd.DigitGrouping(._06Natura, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah7.Index).Value = Me.objCd.DigitGrouping(._07Tantiem, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah8.Index).Value = Me.objCd.DigitGrouping(._08PenghasilanBruto, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah9.Index).Value = Me.objCd.DigitGrouping(._09BiayaJabatan, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah10.Index).Value = Me.objCd.DigitGrouping(._10IuranPensiun, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah11.Index).Value = Me.objCd.DigitGrouping(._11JumlahPengurangan, 0)
                                'rowDGV.Cells(Me.ColA1_Jumlah12.Index).Value = Me.objCd.DigitGrouping(._12JumlahPenghasilanNetto, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah12.Index).Value = Me.objCd.DigitGrouping(._08PenghasilanBruto - ._11JumlahPengurangan, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah13.Index).Value = Me.objCd.DigitGrouping(._13JumlahPenghasilanMasaPjkSblmnya1, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah14.Index).Value = Me.objCd.DigitGrouping(._14JumlahPenghasilanNettoUtkPPh, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah15.Index).Value = Me.objCd.DigitGrouping(._15PTKP, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah16.Index).Value = Me.objCd.DigitGrouping(._16PKPSetahun, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah17.Index).Value = Me.objCd.DigitGrouping(._19PPhTerhutang, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah18.Index).Value = Me.objCd.DigitGrouping(._18PPh21YgTelahDipotongMasaPjkSblmnya1, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah19.Index).Value = Me.objCd.DigitGrouping(._19PPhTerhutang - ._18PPh21YgTelahDipotongMasaPjkSblmnya1, 0)
                                rowDGV.Cells(Me.ColA1_Jumlah20.Index).Value = Me.objCd.DigitGrouping(._19PPhTerhutang - ._18PPh21YgTelahDipotongMasaPjkSblmnya1, 0)

                                rowDGV.Cells(Me.ColA1_StatusPindah.Index).Value = .Status_Pindah
                                rowDGV.Cells(Me.ColA1_NPWPPemotong.Index).Value = .C_NPWP
                                rowDGV.Cells(Me.ColA1_NamaPemotong.Index).Value = .C_Nama
                                rowDGV.Cells(Me.ColA1_TanggalBuktiPotong.Index).Value = Format(.Tanggal_Transaksi, "dd-MM-yyyy")
                            End If
                        Else
                            Me.rowMsg = Me.tblMsg.NewTblMsgBoxRow
                            Me.rowMsg.No_Transaksi = .No_Transaksi
                            Me.rowMsg.Employee_Name = .A_Nama
                            Me.rowMsg.Primary_Salary = .Primary
                            Me.tblMsg.AddTblMsgBoxRow(Me.rowMsg)
                        End If
                    End With
                Next
                If Me.tblMsg.Rows.Count > 0 Then
                    For Each Me.rowMsg In Me.tblMsg
                        notAllowedVal += vbNewLine & (Me.tblMsg.Rows.IndexOf(Me.rowMsg) + 1).ToString & ". Salary No: " & Me.rowMsg.No_Transaksi & ", Employee Name: " & Me.rowMsg.Employee_Name
                    Next
                    MsgBox("You don't have access for the following salary data: " & notAllowedVal, MsgBoxStyle.Information, _
                    "Information")
                End If

                If Me.dgvA1.Rows.Count > 0 Then
                    Me.dgvA1.Rows.Add()

                    Me.SetSubTotal(Me.ColA1_Jumlah1.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah2.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah3.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah4.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah5.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah6.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah7.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah8.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah9.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah10.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah11.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah12.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah13.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah14.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah15.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah16.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah17.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah18.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah19.Index)
                    Me.SetSubTotal(Me.ColA1_Jumlah20.Index)

                    Beep()
                Else
                    MsgBox("There is no data found.", MsgBoxStyle.Information, "Information")
                End If
            Else
                'MsgBox("There is no data found.", MsgBoxStyle.Information, "Information")
            End If

            Me.dgvA1.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub SetSubTotal(ByVal vIndex As Integer)
        Dim rowDGV As Windows.Forms.DataGridViewRow
        Dim Total As Decimal = 0
        Dim DefPrec As Integer

        Try
            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency

            For Each rowDGV In Me.dgvA1.Rows
                If IsNumeric(rowDGV.Cells(vIndex).Value) = True Then
                    Total += CDec(rowDGV.Cells(vIndex).Value)
                End If
            Next

            rowDGV = Me.dgvA1.Rows(Me.dgvA1.Rows.Count - 1)
            rowDGV.Cells(vIndex).Value = Me.ObjCD.DigitGrouping(Total, DefPrec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ExportPPHFinalPesangon()
        Dim lengthInit As Integer = 0
        Dim notAllowedVal As String = ""
        Dim NoUrut As String = 0

        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            Me.tblMonthly.Clear()
            Me.tblMsg.Clear()
            Me.csvVal = ""
            If Me.dtaMonthly.FillBy_TanggalLaporan(Me.tblMonthly, Me.objCd.GetLasDateOfTheMonth(Me.DtpPeriod.Value), Me.objCd.GetFirstDateofTheMonth(Me.DtpPeriod.Value)) > 0 Then
                Me.csvVal = "Masa Pajak;Tahun Pajak;Pembetulan;Nomor Bukti Potong;NPWP;NIK;Nama;Alamat;Kode Pajak;Jumlah Bruto;Tarif;Jumlah PPh;NPWP Pemotong;Nama Pemotong;Tanggal Bukti Potong"
                lengthInit = Me.csvVal.Length
                For Each Me.rowMonthly In Me.tblMonthly
                    If Me.rowMonthly._21_Pesangon <> 0 Then
                        If Me.rowMonthly.Tax <> CInt(ClsTransaksiSalary.EnumTax.Manual_Tax_Calculation) Then
                            With Me.rowMonthly
                                .Pembetulan = Me.TxtPembetulan.Text
                                If Me.IsAllowed(.Primary_Salary_Formula) = True Then
                                    Me.csvVal += vbNewLine & Format(Me.DtpPeriod.Value, "MM") & ";"
                                    Me.csvVal += Format(Me.DtpPeriod.Value, "yyyy") & ";"
                                    Me.csvVal += .Pembetulan & ";"
                                    Me.csvVal += .Prefix_Bukti_Potong_PPH_Final_Pesangon + "-" + .No_Bukti_Potong_PPH_Final & ";"
                                    Me.csvVal += .A_NPWP.Replace(".", "").Replace("-", "") & ";"
                                    Me.csvVal += .A_NIK & ";"
                                    Me.csvVal += .A_Nama & ";"
                                    Me.csvVal += .A_Alamat1 + " " + .A_Alamat2 & ";"
                                    Me.csvVal += "21-401-01" & ";"
                                    Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(._21_Pesangon + ._21_Tunj_PPH_FInal_Pesangon) & ";"
                                    Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(.Tarif_Pesangon) & ";"
                                    Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(._21_PPH_Final_Pesangon) & ";"
                                    Me.csvVal += .NPWP_Pemotong.Replace(".", "").Replace("-", "") & ";"
                                    Me.csvVal += .Nama_Pemotong & ";"
                                    Me.csvVal += Format(.Tanggal_Transaksi, "dd/MM/yyyy")
                                Else
                                    Me.rowMsg = Me.tblMsg.NewTblMsgBoxRow
                                    Me.rowMsg.No_Transaksi = .No_Transaksi
                                    Me.rowMsg.Employee_Name = .Employee_Name
                                    Me.rowMsg.Primary_Salary = .Primary
                                    Me.tblMsg.AddTblMsgBoxRow(Me.rowMsg)
                                End If
                            End With
                        End If
                    End If
                Next
                If Me.tblMsg.Rows.Count > 0 Then
                    For Each Me.rowMsg In Me.tblMsg
                        notAllowedVal += vbNewLine & (Me.tblMsg.Rows.IndexOf(Me.rowMsg) + 1).ToString & ". Salary No: " & Me.rowMsg.No_Transaksi & ", Employee Name: " & Me.rowMsg.Employee_Name
                    Next
                    If MsgBox("The following salary doesn't have access in payroll access: " & notAllowedVal & vbNewLine & "Do you want to export to csv file?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
                If Me.csvVal.Length > lengthInit Then
                    Me.SaveFileDialog1.FileName = "PPH_Final_Pesangon_e-spt_" & Format(Me.DtpPeriod.Value, "MMMM") & "_" & Format(Me.DtpPeriod.Value, "yyyy")
                    If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK And Me.SaveFileDialog1.FileName.Trim.Length > 0 Then
                        My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, csvVal, False)
                        MsgBox("Csv file created.", MsgBoxStyle.Information, "Information")
                    End If
                Else
                    MsgBox("There is no data will be exported", MsgBoxStyle.Information, "Information")
                End If
            Else
                MsgBox("There is no data in this period", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Export_eSPT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.DtpPeriod.Value = Me.objCd.GetFirstDateofTheMonth(Me.DtpPeriod.Value)
            Me.cboFormat.SelectedIndex = 0
            dtaMonthly.Connection.ConnectionString = Core_Module.GetConstring
            dtaYearly.Connection.ConnectionString = Core_Module.GetConstring
            dtaPA.Connection.ConnectionString = Core_Module.GetConstring
            DTASF.Connection.ConnectionString = Core_Module.GetConstring
            Me.FillAuthority()

            Me.dgvA1.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvPesangong.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.ColA1_Jumlah1.HeaderText = "GAJI / PENSIUN ATAU THT / JHT"
            Me.ColA1_Jumlah2.HeaderText = "TUNJANGAN PPh"
            Me.ColA1_Jumlah3.HeaderText = "TUNJANGAN LAINNYA, UANG LEMBUR, DAN SEBAGAINYA"
            Me.ColA1_Jumlah4.HeaderText = "HONORARIUM DAN IMBALAN LAIN SEJENISNYA"
            Me.ColA1_Jumlah5.HeaderText = "PREMI ASURANSI YANG DIBAYAR PEMBERI KERJA"
            Me.ColA1_Jumlah6.HeaderText = "PENERIMAAN DALAM BENTUK NATURA DAN KENIKMATAN LAINNYA YANG DIKENAKAN PEMOTONGAN PPh PASAL 21"
            Me.ColA1_Jumlah7.HeaderText = "TANTIEM, BONUS, GRATIFIKASI, JASA PRODUKSI DAN THR"
            Me.ColA1_Jumlah8.HeaderText = "JUMLAH PENGHASILAN BRUTO (1 S.D. 7)"
            Me.ColA1_Jumlah9.HeaderText = "BIAYA JABATAN/BIAYA PENSIUN"
            Me.ColA1_Jumlah10.HeaderText = "IURAN PENSIUN ATAU IURAN THT/JHT"
            Me.ColA1_Jumlah11.HeaderText = "JUMLAH PENGURANGAN ( 9 S.D. 10)"
            Me.ColA1_Jumlah12.HeaderText = "JUMLAH PENGHASILAN NETTO (8 - 11)"
            Me.ColA1_Jumlah13.HeaderText = "PENGHASILAN NETO MASA PAJAK SEBELUMNYA"
            Me.ColA1_Jumlah14.HeaderText = "JUMLAH PENGHASILAN NETO UNTUK PENGHITUNGAN PPh PASAL 21 (SETAHUN/DISETAHUNKAN)"
            Me.ColA1_Jumlah15.HeaderText = "PENGHASILAN TIDAK KENA PAJAK (PTKP)"
            Me.ColA1_Jumlah16.HeaderText = "PENGHASILAN KENA PAJAK SETAHUN/DISETAHUNKAN (14 - 15)"
            Me.ColA1_Jumlah17.HeaderText = "PPh PASAL 21 ATAS PENGHASILAN KENA PAJAK SETAHUN/DISETAHUNKAN"
            Me.ColA1_Jumlah18.HeaderText = "PPh PASAL 21 YANG TELAH DIPOTONG MASA PAJAK SEBELUMNYA"
            Me.ColA1_Jumlah19.HeaderText = "PPh PASAL 21 TERUTANG"
            Me.ColA1_Jumlah20.HeaderText = "PPh PASAL 21 DAN PPh PASAL 26 YANG TELAH DIPOTONG DAN DILUNASI"

            Me.dgvA1.RowHeadersVisible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            If Me.cboFormat.SelectedIndex = 0 Then
                Me.DtpPeriod.CustomFormat = "MMM yyyy"
                Me.btnYearlyCheck.Visible = False
                Me.btnYearlyCheck.TabStop = False
            ElseIf Me.cboFormat.SelectedIndex = 1 Then
                Me.DtpPeriod.CustomFormat = "yyyy"
                Me.TabControl1.SelectedTab = Me.TabPage2
                Me.LoadDGVA1()
                Me.btnYearlyCheck.Visible = True
                Me.btnYearlyCheck.TabStop = True
            ElseIf Me.cboFormat.SelectedIndex = 2 Then
                Me.DtpPeriod.CustomFormat = "MMM yyyy"
                Me.TabControl1.SelectedTab = Me.TabPage1
                LoadDataPesangon()
                Me.btnYearlyCheck.Visible = False
                Me.btnYearlyCheck.TabStop = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataPesangon()
        Dim rowDGV As Windows.Forms.DataGridViewRow
        Dim UrutanBuktiPotong As String = ""
        Dim PrefixBuktiPotong As String = ""

        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            Me.dgvPesangong.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            Me.dgvPesangong.Rows.Clear()
            If Me.dtaMonthly.FillBy_TanggalLaporan(Me.tblMonthly, Me.objCd.GetLasDateOfTheMonth(Me.DtpPeriod.Value), Me.objCd.GetFirstDateofTheMonth(Me.DtpPeriod.Value)) > 0 Then
                For Each Me.rowMonthly In Me.tblMonthly
                    If Me.rowMonthly._21_Pesangon <> 0 Then
                        If Me.IsAllowed(Me.rowMonthly.Primary_Salary_Formula) = True Then
                            With Me.rowMonthly
                                Me.dgvPesangong.Rows.Add()
                                rowDGV = Me.dgvPesangong.Rows(Me.dgvPesangong.Rows.Count - 1)
                                rowDGV.Cells(Me.ColDate.Index).Value = Me.rowMonthly.Tanggal_Transaksi
                                rowDGV.Cells(Me.ColDocNo.Index).Value = Me.rowMonthly.No_Transaksi
                                rowDGV.Cells(Me.ColEmployee.Index).Value = Me.rowMonthly.Employee_Name
                                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvPesangong.Rows.Count
                                If Me.rowMonthly.No_Bukti_Potong_PPH_Final = "" Then
                                    UrutanBuktiPotong = Me.ObjSalary.GetNoUrutBuktiPotongMax(Me.rowMonthly.Tanggal_Transaksi) + 1
                                    While UrutanBuktiPotong.Length < 7
                                        UrutanBuktiPotong = "0" + UrutanBuktiPotong
                                    End While
                                    PrefixBuktiPotong = "1.4-" + Format(Me.rowMonthly.Tanggal_Transaksi, "MM") + "." + Format(Me.rowMonthly.Tanggal_Transaksi, "yy")
                                    rowDGV.Cells(Me.ColNoBuktiPotong.Index).Value = PrefixBuktiPotong + "-" + UrutanBuktiPotong
                                    rowDGV.Cells(Me.ColPrefiBuktiPotonf.Index).Value = PrefixBuktiPotong
                                    rowDGV.Cells(Me.ColUrutanBuktiPotong.Index).Value = UrutanBuktiPotong
                                    Me.rowMonthly.Prefix_Bukti_Potong_PPH_Final_Pesangon = PrefixBuktiPotong
                                    Me.rowMonthly.No_Bukti_Potong_PPH_Final = UrutanBuktiPotong
                                    Me.dtaMonthly.UpdateQuery(UrutanBuktiPotong, PrefixBuktiPotong, Me.rowMonthly.Primary)
                                Else
                                    rowDGV.Cells(Me.ColNoBuktiPotong.Index).Value = Me.rowMonthly.Prefix_Bukti_Potong_PPH_Final_Pesangon + "-" + Me.rowMonthly.No_Bukti_Potong_PPH_Final
                                    rowDGV.Cells(Me.ColPrefiBuktiPotonf.Index).Value = Me.rowMonthly.Prefix_Bukti_Potong_PPH_Final_Pesangon
                                    rowDGV.Cells(Me.ColUrutanBuktiPotong.Index).Value = Me.rowMonthly.No_Bukti_Potong_PPH_Final
                                End If
                                rowDGV.Cells(Me.ColTarif.Index).Value = Me.rowMonthly.Tarif_Pesangon
                                rowDGV.Cells(Me.ColPrimarySalary.Index).Value = Me.rowMonthly.Primary
                                rowDGV.Cells(Me.ColPPHFinal.Index).Value = Me.rowMonthly._21_PPH_Final_Pesangon
                                rowDGV.Cells(Me.ColPesangon.Index).Value = Me.rowMonthly._21_Pesangon + Me.rowMonthly._21_Tunj_PPH_FInal_Pesangon
                            End With
                        End If
                    End If
                Next
            End If
            Me.ColPesangon.DefaultCellStyle.Format = "N0"
            Me.ColPPHFinal.DefaultCellStyle.Format = "N0"
            Me.ColTarif.DefaultCellStyle.Format = "N2"
            Me.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.dgvPesangong.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub ExportMonthly()
        Dim lengthInit As Integer = 0
        Dim notAllowedVal As String = ""
        Dim JmlPegawaiNonPTKP As Decimal = 0
        Dim JmlPenghasilanBrutoNonPTKP As Decimal = 0
        Dim FaktorKaliDisetahunkan As Integer = 0
        Dim StartWorkingDate As Date = Today

        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            Me.tblMonthly.Clear()
            Me.tblMsg.Clear()
            Me.csvVal = ""
            If Me.dtaMonthly.FillBy_TanggalLaporan(Me.tblMonthly, Me.objCd.GetLasDateOfTheMonth(Me.DtpPeriod.Value), Me.objCd.GetFirstDateofTheMonth(Me.DtpPeriod.Value)) > 0 Then
                Me.csvVal = "Masa Pajak;Tahun Pajak;Pembetulan;NPWP;Nama;Kode Pajak;Jumlah Bruto;Jumlah PPh;Kode Negara"
                lengthInit = Me.csvVal.Length
                For Each Me.rowMonthly In Me.tblMonthly
                    With Me.rowMonthly
                        .Pembetulan = Me.TxtPembetulan.Text
                        If Me.rowMonthly.Tax <> CInt(ClsTransaksiSalary.EnumTax.Manual_Tax_Calculation) Then
                            If Me.rowMonthly.Income_Tax <> 0 Then
                                StartWorkingDate = Me.rowMonthly.first_tax_period
                                If StartWorkingDate.Date.Year < Me.objCd.GetFirstDateofTheMonth(Me.rowMonthly.End_Date).Date.Year Then
                                    FaktorKaliDisetahunkan = 12
                                ElseIf StartWorkingDate.Date.Year = Me.objCd.GetFirstDateofTheMonth(Me.rowMonthly.End_Date).Date.Year Then
                                    If StartWorkingDate.Month = 1 Then
                                        FaktorKaliDisetahunkan = 12
                                    ElseIf StartWorkingDate.Month = 2 Then
                                        FaktorKaliDisetahunkan = 11
                                    ElseIf StartWorkingDate.Month = 3 Then
                                        FaktorKaliDisetahunkan = 10
                                    ElseIf StartWorkingDate.Month = 4 Then
                                        FaktorKaliDisetahunkan = 9
                                    ElseIf StartWorkingDate.Month = 5 Then
                                        FaktorKaliDisetahunkan = 8
                                    ElseIf StartWorkingDate.Month = 6 Then
                                        FaktorKaliDisetahunkan = 7
                                    ElseIf StartWorkingDate.Month = 7 Then
                                        FaktorKaliDisetahunkan = 6
                                    ElseIf StartWorkingDate.Month = 8 Then
                                        FaktorKaliDisetahunkan = 5
                                    ElseIf StartWorkingDate.Month = 9 Then
                                        FaktorKaliDisetahunkan = 4
                                    ElseIf StartWorkingDate.Month = 10 Then
                                        FaktorKaliDisetahunkan = 3
                                    ElseIf StartWorkingDate.Month = 11 Then
                                        FaktorKaliDisetahunkan = 2
                                    ElseIf StartWorkingDate.Month = 12 Then
                                        FaktorKaliDisetahunkan = 1
                                    End If
                                Else
                                    FaktorKaliDisetahunkan = 12
                                End If

                                If Me.IsAllowed(.Primary_Salary_Formula) = True Then
                                    Me.csvVal += vbNewLine & Format(Me.DtpPeriod.Value, "MM") & ";"
                                    Me.csvVal += Format(Me.DtpPeriod.Value, "yyyy") & ";"
                                    Me.csvVal += .Pembetulan & ";"
                                    Me.csvVal += .A_NPWP.Replace(".", "").Replace("-", "") & ";"
                                    Me.csvVal += .A_Nama & ";"
                                    Me.csvVal += .Kode_Pajak & ";"
                                    Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(.Penghasilan_Bruto) & ";"
                                    Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(.Income_Tax - ._21_PPH_Final_Pesangon) & ";"
                                    If .A_Kode_Negara_Domisili <> "INA" Then
                                        Me.csvVal += .A_Kode_Negara_Domisili
                                    End If
                                Else
                                    Me.rowMsg = Me.tblMsg.NewTblMsgBoxRow
                                    Me.rowMsg.No_Transaksi = .No_Transaksi
                                    Me.rowMsg.Employee_Name = .Employee_Name
                                    Me.rowMsg.Primary_Salary = .Primary
                                    Me.tblMsg.AddTblMsgBoxRow(Me.rowMsg)
                                End If
                            Else
                                JmlPegawaiNonPTKP += 1
                                JmlPenghasilanBrutoNonPTKP += .Penghasilan_Bruto
                            End If
                        End If
                    End With
                Next
                If Me.tblMsg.Rows.Count > 0 Then
                    For Each Me.rowMsg In Me.tblMsg
                        notAllowedVal += vbNewLine & (Me.tblMsg.Rows.IndexOf(Me.rowMsg) + 1).ToString & ". Salary No: " & Me.rowMsg.No_Transaksi & ", Employee Name: " & Me.rowMsg.Employee_Name
                    Next
                    If MsgBox("The following salary doesn't have access in payroll access: " & notAllowedVal & vbNewLine & "Do you want to export to csv file?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
                If Me.csvVal.Length > lengthInit Then
                    Me.SaveFileDialog1.FileName = "monthly_e-spt_" & Format(Me.DtpPeriod.Value, "MMMM") & "_" & Format(Me.DtpPeriod.Value, "yyyy")
                    If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK And Me.SaveFileDialog1.FileName.Trim.Length > 0 Then
                        My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, csvVal, False)
                        'tampilkan jumlah pegawai yang pph =0 dan penghasilan brutonya
                        If JmlPegawaiNonPTKP > 0 Then
                            'MsgBox("Csv file created." + vbCrLf + "Notes: Pegawai dengan penghasilan tidak melebihi PTKP: " + _
                            'JmlPegawaiNonPTKP.ToString + " org - Penghasilan Bruto: " + Me.objCd.DigitGrouping(JmlPenghasilanBrutoNonPTKP, 0), MsgBoxStyle.Information, "Information")
                            Dim ij As New FrmTrans_Export_eSPT_Msg
                            ij.JmlKary = JmlPegawaiNonPTKP
                            ij.PenghasilanBruto = Math.Round(JmlPenghasilanBrutoNonPTKP, 0, MidpointRounding.AwayFromZero)
                            ij.ShowDialog()
                            ij.Dispose()
                        Else
                            MsgBox("Csv file created.", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                Else
                    MsgBox("There is no data will be exported", MsgBoxStyle.Information, "Information")
                End If
            Else
                MsgBox("There is no data in this period", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Function ApakahExport(ByVal vPrimarySalary As Double) As Boolean
        Dim rowDgv As Windows.Forms.DataGridViewRow

        Try
            For Each rowDgv In Me.dgvA1.Rows
                If rowDgv.Cells(Me.ColA1_Primary.Index).Value = vPrimarySalary Then
                    If rowDgv.Cells(Me.ColA1_Export.Index).Value = Boolean.TrueString Then
                        Return True
                    Else
                        Return False
                    End If
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub ExportYearly()
        Dim lengthInit As Integer = 0
        Dim notAllowedVal As String = ""
        Dim ListPrimary As New List(Of Double)
        Dim JmlPegawaiNonPTKP As Decimal = 0
        Dim JmlPenghasilanBrutoNonPTKP As Decimal = 0

        Try
            '1.-ok hanya export yang di cek list saja
            '2.-ok cek datanya pastikan sama dengan yang di dgv
            '3.-ok cek import via e-SPT
            '4. adu data dengan db access e-spt citasco
            '5.-ok status export di update
            '6.-ok cek list hanya menampilkan yang belum di export saja
            '7.-ok cek list dan llb, hide ketika ganti tab control
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            Me.tblYearly.Clear()
            Me.tblMsg.Clear()
            Me.csvVal = ""

            Me.LoadDataDGVA1()

            If Me.tblYearly.Rows.Count > 0 Then
                Me.csvVal = "Masa Pajak;Tahun Pajak;Pembetulan;Nomor Bukti Potong;Masa Perolehan Awal;Masa Perolehan Akhir;NPWP;NIK;Nama;Alamat;Jenis Kelamin;Status PTKP;Jumlah Tanggungan;Nama Jabatan;WP Luar Negeri;Kode Negara;Kode Pajak;Jumlah 1;Jumlah 2;Jumlah 3;Jumlah 4;Jumlah 5;Jumlah 6;Jumlah 7;Jumlah 8;Jumlah 9;Jumlah 10;Jumlah 11;Jumlah 12;Jumlah 13;Jumlah 14;Jumlah 15;Jumlah 16;Jumlah 17;Jumlah 18;Jumlah 19;Jumlah 20;Status Pindah;NPWP Pemotong;Nama Pemotong;Tanggal Bukti Potong"
                lengthInit = Me.csvVal.Length

                For Each Me.rowYearly In Me.tblYearly
                    With Me.rowYearly
                        .Pembetulan = Me.TxtPembetulan.Text
                        If Me.IsAllowed(.Primary_Salary_Formula) = True Then
                            If Me.ApakahExport(.Primary) = True Then
                                ListPrimary.Add(.Primary)

                                'If .Period_End.ToString.Length = 1 Then
                                '    Me.csvVal += vbNewLine & "0" & .Period_End & ";"
                                'Else
                                '    Me.csvVal += vbNewLine & .Period_End & ";"
                                'End If
                                Me.csvVal += vbNewLine & "12;"
                                Me.csvVal += Format(Me.DtpPeriod.Value, "yyyy") & ";"
                                Me.csvVal += .Pembetulan & ";"
                                Me.csvVal += .NoForm1721A1 & ";"
                                If .Period_Start.ToString.Length = 1 Then
                                    Me.csvVal += "0" & .Period_Start & ";"
                                Else
                                    Me.csvVal += .Period_Start & ";"
                                End If
                                If .Period_End.ToString.Length = 1 Then
                                    Me.csvVal += "0" & .Period_End & ";"
                                Else
                                    Me.csvVal += .Period_End & ";"
                                End If
                                Me.csvVal += .A_NPWP.Replace(".", "").Replace("-", "") & ";"
                                Me.csvVal += .A_NIK & ";"
                                Me.csvVal += .A_Nama & ";"
                                Me.csvVal += .Alamat & ";"
                                If .A_Jenis_Kelamin = 0 Then
                                    Me.csvVal += "M;"
                                ElseIf .A_Jenis_Kelamin = 1 Then
                                    Me.csvVal += "F;"
                                End If
                                If .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.TK0 Then
                                    Me.csvVal += "TK;"
                                    Me.csvVal += "0;"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.TK1 Then
                                    Me.csvVal += "TK;"
                                    Me.csvVal += "1;"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.TK2 Then
                                    Me.csvVal += "TK;"
                                    Me.csvVal += "2;"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.TK3 Then
                                    Me.csvVal += "TK;"
                                    Me.csvVal += "3;"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.K0 Then
                                    Me.csvVal += "K;"
                                    Me.csvVal += "0;"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.K1 Then
                                    Me.csvVal += "K;"
                                    Me.csvVal += "1;"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.K2 Then
                                    Me.csvVal += "K;"
                                    Me.csvVal += "2;"
                                ElseIf .A_Status_NPWP = ClsMasterSettingPajak.EnumCriteria.K3 Then
                                    Me.csvVal += "K;"
                                    Me.csvVal += "3;"
                                End If
                                Me.csvVal += .A_Nama_Jabatan & ";"
                                If .A_Kode_Negara_Domisili = "INA" Then
                                    Me.csvVal += "N;"
                                    Me.csvVal += ";"
                                Else
                                    Me.csvVal += "Y;"
                                    Me.csvVal += .A_Kode_Negara_Domisili & ";"
                                End If
                                Me.csvVal += .Kode_Pajak & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._01Gaji, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._19PPhTerhutang, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._03TunjanganLainnya, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._04Honorarium, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._05PremiAsuransi, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._06Natura, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._07Tantiem, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._08PenghasilanBruto, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._09BiayaJabatan, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._10IuranPensiun, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._11JumlahPengurangan, 0, MidpointRounding.AwayFromZero)) & ";"
                                'Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(._12JumlahPenghasilanNetto) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._08PenghasilanBruto - ._11JumlahPengurangan, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._13JumlahPenghasilanMasaPjkSblmnya1, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._14JumlahPenghasilanNettoUtkPPh, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._15PTKP, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._16PKPSetahun, 0, MidpointRounding.AwayFromZero)) & ";"
                                'Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(._17PPh21) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._19PPhTerhutang, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._18PPh21YgTelahDipotongMasaPjkSblmnya1, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._19PPhTerhutang - ._18PPh21YgTelahDipotongMasaPjkSblmnya1, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += Me.objCd.BuangNolDibelakangDecimal(Math.Round(._19PPhTerhutang - ._18PPh21YgTelahDipotongMasaPjkSblmnya1, 0, MidpointRounding.AwayFromZero)) & ";"
                                Me.csvVal += .Status_Pindah & ";"
                                Me.csvVal += .C_NPWP.Replace(".", "").Replace("-", "") & ";"
                                Me.csvVal += .C_Nama & ";"
                                Me.csvVal += Format(.Tanggal_Transaksi, "dd-MM-yyyy")

                                If ._16PKPSetahun = 0 Then
                                    JmlPegawaiNonPTKP += 1
                                    JmlPenghasilanBrutoNonPTKP += ._08PenghasilanBruto
                                End If
                            End If
                        Else
                            Me.rowMsg = Me.tblMsg.NewTblMsgBoxRow
                            Me.rowMsg.No_Transaksi = .No_Transaksi
                            Me.rowMsg.Employee_Name = .A_Nama
                            Me.rowMsg.Primary_Salary = .Primary
                            Me.tblMsg.AddTblMsgBoxRow(Me.rowMsg)
                        End If
                    End With
                Next
                If Me.tblMsg.Rows.Count > 0 Then
                    For Each Me.rowMsg In Me.tblMsg
                        notAllowedVal += vbNewLine & (Me.tblMsg.Rows.IndexOf(Me.rowMsg) + 1).ToString & ". Salary No: " & Me.rowMsg.No_Transaksi & ", Employee Name: " & Me.rowMsg.Employee_Name
                    Next
                    If MsgBox("The following salary doesn't have access in payroll access: " & notAllowedVal & vbNewLine & "Do you want to export to csv file?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
                If Me.csvVal.Length > lengthInit Then
                    Me.SaveFileDialog1.FileName = "yearly_e-spt_" & Format(Me.DtpPeriod.Value, "yyyy")
                    If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK And Me.SaveFileDialog1.FileName.Trim.Length > 0 Then
                        My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, csvVal, False)

                        For Each PrimarySalary As Double In ListPrimary
                            Me.dtaMonthly.UpdateQuery_ExportStatus(Me.UserCode + " on " + Format(Now, "dd MMM yyyy - HH.mm"), PrimarySalary)
                        Next

                        'MsgBox("Csv file created.", MsgBoxStyle.Information, "Information")
                        If JmlPegawaiNonPTKP > 0 Then
                            Dim ij As New FrmTrans_Export_eSPT_Msg
                            ij.JmlKary = JmlPegawaiNonPTKP
                            ij.PenghasilanBruto = Math.Round(JmlPenghasilanBrutoNonPTKP, 0, MidpointRounding.AwayFromZero)
                            ij.ShowDialog()
                            ij.Dispose()
                        Else
                            MsgBox("Csv file created.", MsgBoxStyle.Information, "Information")
                        End If

                        Me.LoadDGVA1()
                    End If
                Else
                    MsgBox("There is no data to be be exported", MsgBoxStyle.Information, "Information")
                End If
            Else
                MsgBox("There is no data in this period", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Function IsAllowed(ByVal vPrimarySalaryFormula As Double) As Boolean
        Try
            Me.tblPA.DefaultView.RowFilter = "[Primary Salary Formula] = " & vPrimarySalaryFormula.ToString
            If Me.tblPA.DefaultView.Count <= 0 Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FillAuthority()
        Try
            Me.tblPA.Rows.Clear()
            Me.TBLSF.Rows.Clear()

            Me.dtaPA.FillBy_KodeUser(Me.tblPA, Me.UserCode)
            Me.DTASF.Fill(Me.TBLSF)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpPeriod_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpPeriod.Enter
        Try
            Me.DummyDate = Me.DtpPeriod.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpPeriod_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpPeriod.Leave
        Try
            If Format(Me.DummyDate.Date, "MMM yyyy") <> Format(Me.DtpPeriod.Value.Date, "MMM yyyy") Then
                Me.LoadDataPesangon()
            End If
            If Format(Me.DummyDate.Date, "yyyy") <> Format(Me.DtpPeriod.Value.Date, "yyyy") Then
                Me.LoadDGVA1()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpPeriod_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DtpPeriod.ValueChanged
        Try
            Me.LoadDataPesangon() 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub LoadPembetulan()
    '    Dim PrimaryRecalculate As Double = -1

    '    Try
    '        PrimaryRecalculate = Me.ObjTR.GetPrimaryByPeriod(Me.objCd.GetFirstDateofTheMonth(Me.DtpPeriod.Value))
    '        If PrimaryRecalculate = -1 Then
    '            'tidak ada pembetulan

    '        Else
    '            'ada pembetulan

    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub dgvPesangong_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPesangong.CellClick
        Dim rowDGV As Windows.Forms.DataGridViewRow

        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColNoBuktiPotong.Index Then
                    rowDGV = Me.dgvPesangong.Rows(e.RowIndex)
                    Dim ij As New FrmTrans_Export_eSPT_Edit
                    ij.vPrimarySalary = rowDGV.Cells(Me.ColPrimarySalary.Index).Value
                    ij.vPrefixBuktiPotong = rowDGV.Cells(Me.ColPrefiBuktiPotonf.Index).Value
                    ij.vUrutanBuktiPotong = rowDGV.Cells(Me.ColUrutanBuktiPotong.Index).Value
                    ij.vStartDoc = FrmTrans_Export_eSPT_Edit.EnumStartDoc.PPH21Pesangon

                    If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                        rowDGV.Cells(Me.ColPrefiBuktiPotonf.Index).Value = ij.vPrefixBuktiPotong
                        rowDGV.Cells(Me.ColUrutanBuktiPotong.Index).Value = ij.vUrutanBuktiPotong
                        rowDGV.Cells(Me.ColNoBuktiPotong.Index).Value = rowDGV.Cells(Me.ColPrefiBuktiPotonf.Index).Value + "-" + rowDGV.Cells(Me.ColUrutanBuktiPotong.Index).Value
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPembetulan_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPembetulan.Enter
        Try
            Me.objCd.Number_Enter(Me.TxtPembetulan, sender, e)
            Me.Dummy = Me.TxtPembetulan.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPembetulan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPembetulan.GotFocus
        Try
            Me.TxtPembetulan.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPembetulan_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPembetulan.KeyDown
        Try
            Me.objCd.Number_KeyDown(Me.TxtPembetulan, sender, e, 20)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPembetulan_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPembetulan.KeyUp
        Try
            Me.objCd.Number_KeyUp(Me.TxtPembetulan, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPembetulan_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPembetulan.Leave
        Try
            Me.objCd.Number_Leave(Me.TxtPembetulan, sender, e, 0, Me.objCd)
            If CInt(Me.Dummy) <> CInt(Me.TxtPembetulan.Text) Then
                Me.LoadDGVA1()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPembetulan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPembetulan.LostFocus
        Try
            Me.TxtPembetulan.BackColor = Me.TxtRef.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub TabControl1_TabIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.TabIndexChanged
    '    Try
    '        If Me.TabControl1.SelectedTab.Name = Me.TabPage1.Name Then
    '            Me.LoadDataPesangon()
    '        End If
    '        If Me.TabControl1.SelectedTab.Name = Me.TabPage2.Name Then
    '            Me.LoadDGVA1()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Dim rowDGV As Windows.Forms.DataGridViewRow

        Try
            For Each rowDGV In Me.dgvA1.Rows
                rowDGV.Cells(Me.ColA1_Export.Index).Value = Boolean.TrueString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Dim rowDGV As Windows.Forms.DataGridViewRow

        Try
            For Each rowDGV In Me.dgvA1.Rows
                rowDGV.Cells(Me.ColA1_Export.Index).Value = Boolean.FalseString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedTab.Name = Me.TabPage2.Name Then
                Me.llbCheckAll.Visible = True
                Me.llbUncheckAll.Visible = True
                Me.llbFindEmployee.Visible = True
                Me.chkNotExportedOnly.Visible = True
                Me.chkNotExportedOnly.TabStop = True
            Else
                Me.llbCheckAll.Visible = False
                Me.llbUncheckAll.Visible = False
                Me.llbFindEmployee.Visible = False
                Me.chkNotExportedOnly.Visible = False
                Me.chkNotExportedOnly.TabStop = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    Private Sub llbFindEmployee_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbFindEmployee.LinkClicked
        Dim ij As New Core_Module.FrmMasterKaryawanCari
        Dim rowDGV As Windows.Forms.DataGridViewRow
        Dim Ketemu As Boolean = False

        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()

            If ij.PrimaryEmploye <> -1 Then
                Me.dgvA1.ClearSelection()
                For Each rowDGV In Me.dgvA1.Rows
                    If rowDGV.Cells(Me.ColA1_PrimaryEmployee.Index).Value = ij.PrimaryEmploye Then
                        Me.dgvA1.ClearSelection()
                        Me.dgvA1.FirstDisplayedScrollingRowIndex = rowDGV.Index
                        rowDGV.Selected = True
                        Ketemu = True
                    End If
                Next
            End If

            If Ketemu = False Then
                MsgBox("Employee data not found.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvA1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvA1.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColA1_Export.Index Then
                Me.Buffer = Me.dgvA1.Item(Me.ColA1_Export.Index, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvA1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvA1.CellClick
        Dim rowDGV As Windows.Forms.DataGridViewRow
        Dim Dummy As String = ""

        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColA1_NoBuktiPotong.Index Then
                    rowDGV = Me.dgvA1.Rows(e.RowIndex)
                    Dim ij As New FrmTrans_Export_eSPT_Edit

                    ij.vPrimarySalary = rowDGV.Cells(Me.ColA1_Primary.Index).Value
                    Dummy = Me.ObjSalary.GetNoBaru_BuatNPWP(CDate(rowDGV.Cells(Me.ColA1_TglSalary.Index).Value), Me.objCd)

                    If rowDGV.Cells(Me.ColNoBuktiPotong.Index).Value.ToString.Length <> Dummy.Length Then
                        rowDGV.Cells(Me.ColNoBuktiPotong.Index).Value = Dummy
                    End If
                    ij.vPrefixBuktiPotong = rowDGV.Cells(Me.ColNoBuktiPotong.Index).Value.ToString.Substring(0, 9)
                    ij.vUrutanBuktiPotong = rowDGV.Cells(Me.ColNoBuktiPotong.Index).Value.ToString.Substring(10)
                    ij.vStartDoc = FrmTrans_Export_eSPT_Edit.EnumStartDoc.PPH21
                    ij.vTglSalary = rowDGV.Cells(Me.ColA1_TglSalary.Index).Value

                    If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                        rowDGV.Cells(Me.ColNoBuktiPotong.Index).Value = ij.vPrefixBuktiPotong + "-" + ij.vUrutanBuktiPotong
                    End If

                    ij.Dispose()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvA1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvA1.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColA1_Export.Index Then
                'do nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkNotExportedOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkNotExportedOnly.CheckedChanged
        Try
            Me.LoadDGVA1()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnYearlyCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnYearlyCheck.Click
        Try
            Me.FillAuthority()
            If Me.tblPA.Rows.Count <> Me.TBLSF.Rows.Count Then
                MsgBox("Restricted access. Only users with access to all salary formula can perform this check.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.chkNotExportedOnly.Checked = False

            Dim ij As New FrmTrans_Export_eSPT_Check
            ij.mParentForm1 = Me
            ij.UserCode = Me.UserCode

            ij.ShowDialog()

            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvA1_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvA1.RowsAdded
        Dim rowDGV As Windows.Forms.DataGridViewRow

        Try
            rowDGV = Me.dgvA1.Rows(Me.dgvA1.Rows.Count - 1)
            rowDGV.Cells(Me.ColA1_PrimaryEmployee.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class