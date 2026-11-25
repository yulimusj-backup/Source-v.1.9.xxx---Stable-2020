Option Explicit On
Option Strict Off
Public Class FrmTrans_ManufacturingOrder_PrintJO_PCK
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Public PrimaryKompetensi As Double = 15
    Public objCustomer As New Core_Module.clsMasterCustomer
    Public objCd As New Core_Module.ClsCekData
    Public Identifier_JO As String
    Public Type_Jo As Integer
    Public Buffer As String = ""
    Public Timestamp As Date = Now
    Dim EraseRowIndex As Integer = 0
    Public baru As Boolean = True
    Public Ubah As MsgBoxResult
    Dim tglpckFromChache As Date
    Dim tglpckToChache As Date
    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    'Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiProduksi
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Public Enum EnumTipeCetakan As Integer
        RAW_PENCIL
        PACKAGING
    End Enum

    Public Enum EnumTypeProcess As Integer
        RAW_PENCIL
        PAINTING
        FINISHING
        DIPING
        PACKING
    End Enum

    Private Sub FrmTrans_ManufacturingOrder_PrintJO_PCK_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            'Me.cboType.SelectedIndex = 0
            Me.btnCancel.DialogResult = Windows.Forms.DialogResult.None
            Me.cboRevision.SelectedIndex = 0
            Me.chkPCK.Checked = False
            Me.dtpPCKFrom.Enabled = False
            Me.dtpPCKTo.Enabled = False
            Me.Type_Jo = EnumTipeCetakan.PACKAGING
            LoadSetting()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtShrinkPack_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShrinkPack.Enter
        Try
            Me.objCd.Number_Enter(Me.txtShrinkPack, sender, e)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtOuterPack_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOuterPack.Enter
        Try
            Me.objCd.Number_Enter(Me.txtOuterPack, sender, e)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus, txtShrinkPack.GotFocus, _
    txtOuterPack.GotFocus, txtFU.GotFocus, txtPoSo.GotFocus, txtRemarks1.GotFocus

        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                If Me.objCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.txtPrimaryCustomer.Text = Me.objCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.txtCustomerDesc.Text = Me.objCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.txtFU.Focus()
                Else
                    Me.btnFindCustomer.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.Leave
        Try
            If Me.txtCustomerCode.Text <> "" Then
                If Me.objCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.txtPrimaryCustomer.Text = Me.objCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.txtCustomerDesc.Text = Me.objCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                Else

                    Me.btnFindCustomer.PerformClick()
                End If
            Else
                Me.txtPrimaryCustomer.Text = -1
                Me.txtCustomerCode.Text = ""
                Me.txtCustomerDesc.Text = Me.objCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
            End If
            Me.txtCustomerCode.BackColor = Me.txtPrimaryCustomer.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtShrinkPack_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtShrinkPack.KeyDown
        Me.objCd.Number_KeyDown(Me.txtShrinkPack, sender, e)

    End Sub

    Private Sub txtShrinkPack_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtShrinkPack.KeyUp
        Me.objCd.Number_KeyUp(Me.txtShrinkPack, sender, e)

    End Sub

    Private Sub txtOuterPack_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOuterPack.KeyDown
        Me.objCd.Number_KeyDown(Me.txtOuterPack, sender, e)
    End Sub

    Private Sub txtOuterPack_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOuterPack.KeyUp
        Me.objCd.Number_KeyUp(Me.txtOuterPack, sender, e)
    End Sub

    Private Sub txtShrinkPack_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShrinkPack.Leave
        Me.txtShrinkPack.Text = Me.objCd.DigitGrouping(Me.txtShrinkPack.Text, 4)
    End Sub

    Private Sub txtOuterPack_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOuterPack.Leave
        Me.txtOuterPack.Text = Me.objCd.DigitGrouping(Me.txtOuterPack.Text, 4)
    End Sub

    Private Sub txtCustomerCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.LostFocus, txtShrinkPack.LostFocus, _
    txtOuterPack.LostFocus, txtFU.LostFocus, txtPoSo.LostFocus, txtRemarks1.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimary.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("FU")
            KoleksiTextBox.Add(Me.txtFU)
            KoleksiNamaData.Add("PoSo")
            KoleksiTextBox.Add(Me.txtPoSo)
            KoleksiNamaData.Add("Customer")
            KoleksiTextBox.Add(Me.txtCustomerCode)
            If Core_Module.ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            If Me.txtFU.Text.Contains("#") = True Or Me.txtFU.Text.Contains("*") = True _
                      Or Me.txtFU.Text.Contains("%") Or Me.txtFU.Text.Contains("#") = True _
                      Or Me.txtFU.Text.Contains("`") = True Then
                Me.txtFU.Focus()
                MsgBox("FU data can not contain the following characters: ` , #, % , or * .", MsgBoxStyle.Information, "Error")
                Return False
            End If
            If Me.txtPoSo.Text.Contains("#") = True Or Me.txtPoSo.Text.Contains("*") = True _
                                  Or Me.txtPoSo.Text.Contains("%") Or Me.txtPoSo.Text.Contains("#") = True _
                                  Or Me.txtPoSo.Text.Contains("`") = True Then
                Me.txtPoSo.Focus()
                MsgBox("PO SO data can not contain the following characters: ` , #, % , or * .", MsgBoxStyle.Information, "Error")
                Return False
            End If
            If Me.txtRemarks1.Text.Contains("#") = True Or _
                   Me.txtRemarks1.Text.Contains("*") = True Or _
                        Me.txtRemarks1.Text.Contains("%") = True Or _
                            Me.txtRemarks1.Text.Contains("`") = True Then
                MsgBox("Remarks data can not contain the following characters: ` , #, % , or * .", MsgBoxStyle.Information, "Error")
                Me.txtRemarks1.Focus()
                Return False
            End If

            If Me.dtpPCKFrom.Value > Me.dtpPCKTo.Value Then
                MsgBox("PCK From Date must be smaller than PCK To Date", MsgBoxStyle.Information, "Error")
                Return False
            End If

            For i As Integer = 0 To Me.dgvItem.Rows.Count - 1
                If Me.ApakahSudahAdaSetting(Me.dgvItem.Item(Me.ColIdentifierJO.Index, i).Value, Me.Type_Jo, Me.txtPrimary.Text) = True Then
                    MsgBox("Jo No: " & Me.dgvItem.Item(Me.ColJONo.Index, i).Value & _
                    " is already used in other settings, please input another Jo No.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.ClearSelection()
                    Me.dgvItem.Rows(i).Selected = True
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnFindCustomer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindCustomer.Click
        Dim FrmCariCustomer As New Core_Module.FrmCariMasterCustomer

        Try
            FrmCariCustomer.BackColorAll = Me.BackColorAll
            If Me.txtPrimaryCustomer.Text = -1 Then
                FrmCariCustomer.txtCari.Text = Me.txtCustomerCode.Text
                FrmCariCustomer.ShowDialog()
            End If
            If CDbl(FrmCariCustomer.PrimaryCustomer) > 0 Then
                Me.txtPrimaryCustomer.Text = FrmCariCustomer.PrimaryCustomer
                Me.txtCustomerCode.Text = Me.objCustomer.GetCustomerCode(FrmCariCustomer.PrimaryCustomer)
                Me.txtCustomerDesc.Text = Me.objCustomer.GetCustName(FrmCariCustomer.PrimaryCustomer)
                Me.txtShrinkPack.Focus()
            Else
                Me.txtPrimaryCustomer.Text = -1
                Me.txtCustomerCode.Text = ""
                Me.txtCustomerDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariCustomer.Dispose()
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                If baru = True Then
                    hasil = MsgBox("Save configuration?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.SimpanBaru() Then
                            MsgBox("Configuration saved", MsgBoxStyle.Information, "Information")
                            Me.Close()
                        End If
                    End If
                Else
                    hasil = MsgBox("Save configuration?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.SimpanPerubahan() Then
                            MsgBox("Configuration saved", MsgBoxStyle.Information, "Information")
                            Me.Close()
                        End If
                    End If
                End If

                If hasil = MsgBoxResult.Yes Then
                    'proses cetak
                    prosesCetakan(Me.dgvItem)
                End If
            End If
            

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub prosesCetakan(ByVal vDGV As DataGridView)
        'Dim DS As New DseTrans_ManufacturingOrder_PrintJO
        ''Dim frmCetakJO As New FrmTrans_ManufacturingOrder_PrintJO_Cetak
        'Dim DTAProd As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbl_ProdDummyTableAdapter
        'Dim TBLProd As New DseTrans_ManufacturingOrder_PrintJO.tbl_ProdDummyDataTable
        'Dim DTAProdBahan As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.Tbl_ProdBahanDummyTableAdapter
        'Dim TBLProdBahan As New DseTrans_ManufacturingOrder_PrintJO.Tbl_ProdBahanDummyDataTable

        'Dim TBLCetak As DseTrans_ManufacturingOrder_PrintJO.tblCetakDataTable = DS.tblCetak
        'Dim rowCetak As DseTrans_ManufacturingOrder_PrintJO.tblCetakRow

        'Dim TBLCetakBahan As DseTrans_ManufacturingOrder_PrintJO.tblCetakBahanDataTable = DS.tblCetakBahan
        'Dim TBLTipe As New DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_jo_tipe_prosesDataTable

        'Dim ObjProd As New Core_Module.clsMasterProduct
        ''Dim Report As New CrpCetak_ManufacturingOrder_PrintJO_PCK
        'Dim RevisiNo As String = Me.cboRevision.Text
        'Dim JONo_ As String = ""
        'Dim MONo_ As String = ""
        'Dim SpecNo_ As String = ""
        'Dim Desc_ As String = ""
        'Dim ArtNo_ As String = ""
        'Dim Cust_ As String = ""
        'Dim tgl_pkg As String = "-"
        'Dim totalGross As Double = 0
        'Dim totalPack As Double = 0
        'Dim totalShrink As Double = 0
        'Dim totalOuter As Double = 0
        'Dim JumlahItem As Double = 0
        'Try
        '    DTAProd.Connection.ConnectionString = Core_Module.GetConstring
        '    Cust_ = Me.txtCustomerDesc.Text
        '    For i As Integer = 0 To vDGV.Rows.Count - 2
        '        'isi data header
        '        If JONo_ = "" Then
        '            JONo_ = vDGV.Item(Me.ColJONo.Index, i).Value
        '        Else
        '            If JONo_.Contains(vDGV.Item(Me.ColJONo.Index, i).Value) = False Then
        '                JONo_ &= ". " & vDGV.Item(Me.ColJONo.Index, i).Value
        '            End If
        '        End If
        '        If SpecNo_ = "" Then
        '            SpecNo_ = Me.GetSpecNo(ObjProd.GetPrimary_ByDesc(Me.dgvItem.Item(Me.ColDesc.Index, i).Value), Me.PrimaryKompetensi)
        '        Else
        '            Dim CustomValue As String = Me.GetSpecNo(ObjProd.GetPrimary_ByDesc(Me.dgvItem.Item(Me.ColDesc.Index, i).Value), Me.PrimaryKompetensi)
        '            If SpecNo_.Contains(CustomValue) = False Then
        '                SpecNo_ &= "; " & CustomValue
        '            End If
        '        End If

        '        If MONo_ = "" Then
        '            MONo_ = vDGV.Item(Me.ColMo.Index, i).Value
        '        Else
        '            If MONo_.Contains(vDGV.Item(Me.ColMo.Index, i).Value) = False Then
        '                MONo_ &= "; " & vDGV.Item(Me.ColMo.Index, i).Value
        '            End If
        '        End If

        '        If Desc_ = "" Then
        '            Desc_ = vDGV.Item(Me.ColDesc.Index, i).Value
        '        Else
        '            If Desc_.Contains(vDGV.Item(Me.ColDesc.Index, i).Value) = False Then
        '                Desc_ &= "; " & vDGV.Item(Me.ColDesc.Index, i).Value
        '            End If
        '        End If

        '        If ArtNo_ = "" Then
        '            ArtNo_ = vDGV.Item(Me.ColCodeNo.Index, i).Value
        '        Else
        '            If ArtNo_.Contains(vDGV.Item(Me.ColCodeNo.Index, i).Value) = False Then
        '                ArtNo_ &= "; " & vDGV.Item(Me.ColCodeNo.Index, i).Value
        '            End If
        '        End If


        '        'isi data item R/P
        '        If DTAProd.Fill(TBLProd, vDGV.Item(Me.ColIdentifierJO.Index, i).Value) > 0 Then
        '            AmbilNilai(TBLProd, TBLCetak)
        '        End If
        '        'isi data item Raw Material
        '        If DTAProdBahan.Fill(TBLProdBahan, vDGV.Item(Me.ColIdentifierJO.Index, i).Value) > 0 Then
        '            AmbilNilai2(TBLCetakBahan, TBLProdBahan)
        '        End If

        '    Next
        '    'olah tanggal
        '    If Me.chkPCK.Checked = True Then
        '        If Me.dtpPCKFrom.Value.Month = Me.dtpPCKTo.Value.Month And Me.dtpPCKFrom.Value.Year = Me.dtpPCKTo.Value.Year Then
        '            tgl_pkg = Me.dtpPCKFrom.Value.Day.ToString & " - " & Me.objCd.GetDateString_Indo_Panjang(Me.dtpPCKTo.Value.Date)
        '        Else
        '            tgl_pkg = Me.objCd.GetDateString_Indo_Panjang(Me.dtpPCKFrom.Value.Date) & " - " & _
        '            Me.objCd.GetDateString_Indo_Panjang(Me.dtpPCKTo.Value.Date)
        '        End If

        '    End If


        '    'masukan ke row cetak data headernya dan total 
        '    For Each rowCetak In TBLCetak.Rows
        '        With rowCetak
        '            If RevisiNo <> "0" Then
        '                .Revisi_No = RevisiNo
        '            End If
        '            .No_JO = JONo_
        '            .Spec_No = SpecNo_
        '            .Desc = Desc_
        '            .Customer = Cust_
        '            .Art_No = ArtNo_
        '            .Mo_No = MONo_
        '            ._PO_SO_No = txtPoSo.Text
        '            .FU = txtFU.Text
        '            ._Sche_R_P = tgl_pkg
        '            totalGross += .Gross_Qty_Pkg
        '            .Total_Qty_Gross = totalGross
        '            JumlahItem += 1
        '        End With
        '    Next
        '    totalPack = (totalGross * 144) / TBLCetak.Rows.Count
        '    If CDbl(Me.txtShrinkPack.Text) = 0 Then
        '        totalShrink = 0
        '    Else
        '        totalShrink = totalPack / CDbl(Me.txtShrinkPack.Text)
        '    End If
        '    If CDbl(Me.txtOuterPack.Text) = 0 Then
        '        totalOuter = 0
        '    Else
        '        totalOuter = totalPack / CDbl(Me.txtOuterPack.Text)
        '    End If

        '    Report.SetDataSource(DS)
        '    Report.SetParameterValue(Report.Parameter_PrintDate.ParameterFieldName, Me.objCd.GetDateString_Indo_Panjang(Me.dtpPrintDate.Value.Date))
        '    Report.SetParameterValue(Report.Parameter_Remarks.ParameterFieldName, Me.txtRemarks1.Text)
        '    Report.SetParameterValue(Report.Parameter_TotalGross.ParameterFieldName, totalGross)
        '    Report.SetParameterValue(Report.Parameter_TotalPack.ParameterFieldName, totalPack)
        '    Report.SetParameterValue(Report.Parameter_TotalShrink.ParameterFieldName, totalShrink)
        '    Report.SetParameterValue(Report.Parameter_TotalOuter.ParameterFieldName, totalOuter)
        '    Report.SetParameterValue(Report.Parameter_JumlahItem.ParameterFieldName, JumlahItem)

        '    frmCetakJO.crvJOPrint.ReportSource = Report
        '    frmCetakJO.crvJOPrint.Zoom(100)
        '    frmCetakJO.crvJOPrint.Show()
        '    frmCetakJO.ShowDialog()
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub
    Private Sub AmbilNilai(ByVal vTBLProd As DseTrans_ManufacturingOrder_PrintJO.tbl_ProdDummyDataTable, _
    ByRef vTBLCetak As DseTrans_ManufacturingOrder_PrintJO.tblCetakDataTable)
        Dim rowProdDummy As DseTrans_ManufacturingOrder_PrintJO.tbl_ProdDummyRow

        Try
            vTBLProd.DefaultView.RowFilter = "[Level] = 1"
            If vTBLProd.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTBLProd.DefaultView
                    rowProdDummy = dv.Row
                    IsiCetakan(vTBLCetak, rowProdDummy)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub IsiCetakan(ByRef vTBLCetak As DseTrans_ManufacturingOrder_PrintJO.tblCetakDataTable, _
    ByVal vrowProdDummy As DseTrans_ManufacturingOrder_PrintJO.tbl_ProdDummyRow)
        Dim RowCetak As DseTrans_ManufacturingOrder_PrintJO.tblCetakRow
        Dim RowTipe As DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_jo_tipe_prosesRow

        Try
            vTBLCetak.DefaultView.RowFilter = Nothing
            vTBLCetak.DefaultView.RowFilter = "[Primary Product] = " & vrowProdDummy.Primary_Product
            If vTBLCetak.DefaultView.Count = 0 Then
                RowCetak = vTBLCetak.NewtblCetakRow
                RowCetak.Primary_Product = vrowProdDummy.Primary_Product
                RowCetak.Kode_Pensil = vrowProdDummy.Product_Code
                RowCetak.Nama_Pensil = vrowProdDummy.Description
                RowCetak.Decimal_RP = vrowProdDummy._Decimal
                If vrowProdDummy.Primary_Unit = 9 Then ' Jika Gross, Kalkulasikan Qtynya
                    RowCetak.Gross_Qty_Pkg = vrowProdDummy.Qty_Target_Prod
                End If
                vTBLCetak.Rows.Add(RowCetak)
            Else
                RowCetak = vTBLCetak.DefaultView(0).Row
                If vrowProdDummy.Primary_Unit = 9 Then ' Jika Gross, Kalkulasikan Qtynya
                    RowCetak.Gross_Qty_Pkg += vrowProdDummy.Qty_Target_Prod
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AmbilNilai2(ByRef vTBLCetakBahan As DseTrans_ManufacturingOrder_PrintJO.tblCetakBahanDataTable, _
    ByVal vTblProdBahanDummy As DseTrans_ManufacturingOrder_PrintJO.Tbl_ProdBahanDummyDataTable)
        Dim rowProdbahanDummy As DseTrans_ManufacturingOrder_PrintJO.Tbl_ProdBahanDummyRow
        Try
            vTblProdBahanDummy.DefaultView.RowFilter = Nothing
            vTblProdBahanDummy.DefaultView.RowFilter = "[Level] = 0"
            If vTblProdBahanDummy.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblProdBahanDummy.DefaultView
                    rowProdbahanDummy = dv.Row
                    IsiCetakan2(vTBLCetakBahan, rowProdbahanDummy)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub IsiCetakan2(ByRef vTBLCetakBahan As DseTrans_ManufacturingOrder_PrintJO.tblCetakBahanDataTable, _
    ByVal vrowProdbahanDummy As DseTrans_ManufacturingOrder_PrintJO.Tbl_ProdBahanDummyRow)
        Dim RowCetakBahan As DseTrans_ManufacturingOrder_PrintJO.tblCetakBahanRow


        Try
            vTBLCetakBahan.DefaultView.RowFilter = Nothing
            vTBLCetakBahan.DefaultView.RowFilter = "[Primary Product RP] = " & vrowProdbahanDummy.Primary_Product
            If vTBLCetakBahan.DefaultView.Count = 0 Then
                RowCetakBahan = vTBLCetakBahan.NewtblCetakBahanRow
                RowCetakBahan.Primary_Product_RP = vrowProdbahanDummy.Primary_Product
                RowCetakBahan.Pck_KodeBahan = vrowProdbahanDummy.Product_Code
                RowCetakBahan.Pck_NamaBahan = vrowProdbahanDummy.Description
                RowCetakBahan.Decimal_RP = vrowProdbahanDummy._Decimal
                RowCetakBahan.Pck_Unit = vrowProdbahanDummy.Unit_Desc
                RowCetakBahan.Pck_Qty = vrowProdbahanDummy.Target_Qty
                vTBLCetakBahan.Rows.Add(RowCetakBahan)
            Else
                RowCetakBahan = vTBLCetakBahan.DefaultView(0).Row
                RowCetakBahan.Pck_Qty += vrowProdbahanDummy.Target_Qty

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetSpecNo(ByVal vPrimaryProd As Double, ByVal vPrimaryKompetensi As Double) As String
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbl_ProdDummyTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetCustomValueByPrimaryProdPrimaryKomp(vPrimaryProd, vPrimaryKompetensi)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function ApakahSudahAdaSetting(ByVal vIdentifierJo As String, ByVal vType As String, ByVal vPrimarySetting As Double) As Boolean
        Dim dta As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_joTableAdapter
        Dim dummy As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJmlPrimary(vIdentifierJo, vType, vPrimarySetting)
            If dummy.HasValue Then
                If dummy.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru() As Boolean

        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Dim CreateDate As Date = Now
        Try



            'SW.Start()
            ''ambil no antrian
            ''Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, txtPrimary.Text, _
            ''Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.FrmTrans_ManufacturingOrder_PrintJO_SimpanBaru, UserCode)
            ''Me.GetNoAntrian()

            ''coba simpan data baru
            'Me.Clock_Antri.Reset()
            'Me.Clock_Antri.Start()
            'While Start = 0
            '    If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
            '        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
            '        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
            '        Exit Function
            '    Else
            '        'hanya mencoba ke database setelah interval waktu tertentu saja
            '        If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
            '            If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
            '                'clock antri matikan
            '                Me.Clock_Antri.Stop()
            '                'sudah saatnya melakukan proses simpan baru
            '                retval = Me.SimpanBaru1(CreateDate)
            '                'hapus data antri untuk transaksi ini
            '                If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
            '                    MsgBox("Logic error(1), contact your provider when you see this message", _
            '                    MsgBoxStyle.Information, "Information")
            '                End If
            '                'nilai balikan
            '                Return retval
            '            Else
            '                'kontrol interval pengecekan
            '                TryCounter += 1
            '                CheckInterval = ((TryCounter \ 5) + 1) * 500
            '                Me.Clock_Antri.Reset()
            '                Me.Clock_Antri.Start()
            '                Threading.Thread.Sleep(1000)
            '            End If
            '        End If
            '    End If
            'End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function
    Public Overridable Function SimpanPerubahan() As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Dim ModifiedDate As Date = Now
        Try
            ''ambil no antrian
            'Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.txtPrimary.Text, _
            'Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiDO_SimpanPerubahan, UserCode)

            ''coba simpan data baru
            'Me.Clock_Antri.Reset()
            'Me.Clock_Antri.Start()
            'While Start = 0
            '    'hanya mencoba ke database setelah interval waktu tertentu saja
            '    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
            '        If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
            '            MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
            '            Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
            '            Exit Function
            '        Else
            '            If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
            '                'clock antri matikan
            '                Me.Clock_Antri.Stop()
            '                'sudah saatnya melakukan proses simpan baru
            '                retval = Me.SimpanPerubahan1(ModifiedDate)
            '                'hapus data antri untuk transaksi ini
            '                If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
            '                    MsgBox("Logic error(1), contact your provider when you see this message", _
            '                    MsgBoxStyle.Information, "Information")
            '                End If
            '                'nilai balikan
            '                Return retval
            '            Else
            '                'kontrol interval pengecekan
            '                TryCounter += 1
            '                CheckInterval = ((TryCounter \ 5) + 1) * 500
            '                Me.Clock_Antri.Reset()
            '                Me.Clock_Antri.Start()
            '            End If
            '        End If
            '    End If
            'End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Public Function SimpanBaru1(ByVal vCreateDate As Date) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_joTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.dtpPrintDate.Value, Me.txtPrimaryCustomer.Text, Me.Type_Jo, _
                "", Me.txtFU.Text, Me.txtPoSo.Text, Me.cboRevision.Text, Me.txtRemarks1.Text, "", "", _
                CDec(Me.txtShrinkPack.Text), CDec(Me.txtOuterPack.Text), Me.dtpPCKFrom.Value, Me.dtpPCKTo.Value, #1/1/1950#, #1/1/2200#, #1/1/1950#, #1/1/2200#, UserCode, _
                vCreateDate, UserCode, Now, Now) > 0 Then
                txtPrimary.Text = Me.GetPrimaryByTCDCB(Me.Type_Jo, vCreateDate, Me.UserCode)
                Me.Timestamp = DTA.ScalarQuery_GetTimestampByprimary(CDec(Me.txtPrimary.Text)).Value
                'simpan item
                Me.SimpanItemSetting(txtPrimary.Text)
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Overridable Function SimpanPerubahan1(ByVal vModifiedDate As Date) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_joTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            ' If Me.ApakahDataBenar(False) = True Then

            If DTA.Update(Me.dtpPrintDate.Value, Me.txtPrimaryCustomer.Text, Me.Type_Jo, "", _
                Me.txtFU.Text, Me.txtPoSo.Text, Me.cboRevision.Text, Me.txtRemarks1.Text, "", "", _
                CDec(Me.txtShrinkPack.Text), CDec(Me.txtOuterPack.Text), Me.dtpPCKFrom.Value, Me.dtpPCKTo.Value, #1/1/1950#, #1/1/2200#, #1/1/1950#, #1/1/2200#, _
                UserCode, vModifiedDate, Now, CDec(Me.txtPrimary.Text), Me.Timestamp) Then
                Me.Timestamp = DTA.ScalarQuery_GetTimestampByprimary(CDec(Me.txtPrimary.Text)).Value


                '//JANGAN LUPA!
                Me.HapusItemSetting(Me.txtPrimary.Text)
                Me.SimpanItemSetting(Me.txtPrimary.Text)
                Return True
            Else
                Return False
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()

        End Try
    End Function
    Public Function GetPrimaryByTCDCB(ByVal vType_Jo As Integer, ByVal vCreateDate As Date, ByVal vUserCode As String) As Double
        Dim dta As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_joTableAdapter
        Dim dummy As Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetPrimaryByTCDCB(vType_Jo, vCreateDate, vUserCode)
            If dummy.HasValue Then
                If dummy.Value > 0 Then
                    Return dummy
                Else
                    Return -1
                End If
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function
    Public Function SimpanItemSetting(ByVal vPrimarySetting As Double) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_jo_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For i As Integer = 0 To dgvItem.Rows.Count - 2
                DTA.Insert(vPrimarySetting, Me.dgvItem.Item(Me.ColIdentifierJO.Index, i).Value)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function HapusItemSetting(ByVal vPrimarySetting As Double) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_jo_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(vPrimarySetting)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try

            Me.dgvItem.Rows.RemoveAt(EraseRowIndex)
            Me.objCd.SetNumber(Me.dgvItem)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Sub LoadSetting()
        Dim DTASetting As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_jo_item2TableAdapter
        Dim TBLSetting As New DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_jo_item2DataTable
        Dim TBLSettingDummy As New DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_jo_item2DataTable
        Dim RowSetting As DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_jo_item2Row
        Dim DTASettingItem As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_itemTableAdapter
        Dim TBLSettingItem As New DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_itemDataTable
        Dim RowSettingItem As DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_itemRow
        Dim NoSetting As Boolean = False
        Dim dr As DataGridViewRow
        Dim listHapusJo As New List(Of Double)
        Dim PrimaryTransaksi_ As Double = -1
        Dim Ubah As MsgBoxResult
        Try


            DTASetting.Connection.ConnectionString = Core_Module.GetConstring
            DTASettingItem.Connection.ConnectionString = Core_Module.GetConstring

            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None


            If DTASettingItem.Fill(TBLSettingItem) > 0 Then
                If DTASetting.Fill(TBLSetting) > 0 Then
                    TBLSettingDummy = TBLSetting.Copy
                    For x As Integer = 0 To TBLSetting.Rows.Count - 1
                        RowSetting = TBLSetting.Rows(x)
                        'cek kondisi bila JO number pada MO dihapus
                        If Me.ApakahMasihAdaJO(RowSetting.identifier_jo) = False Then
                            Me.HapusItemSettingByidentifierJO(RowSetting.identifier_jo, RowSetting.primary_transaksi)
                            If listHapusJo.Contains(RowSetting.primary_transaksi) = False Then
                                listHapusJo.Add(RowSetting.primary_transaksi)
                            End If
                            GoTo kesini
                        End If
                        If RowSetting.identifier_jo = Identifier_JO And RowSetting.tipe = Me.Type_Jo Then
                            NoSetting = True
                            baru = False
                            'isi header setting
                            Me.CariDetail(RowSetting.primary_transaksi, Me.dgvItem)
                            'isi Item setting
                            TBLSettingDummy.DefaultView.RowFilter = Nothing
                            TBLSettingDummy.DefaultView.RowFilter = "[primary transaksi] = " & RowSetting.primary_transaksi
                            If TBLSettingDummy.DefaultView.Count > 0 Then
                                PrimaryTransaksi_ = RowSetting.primary_transaksi
                                For i As Integer = 0 To TBLSettingDummy.DefaultView.Count - 1
                                    Me.dgvItem.AllowUserToAddRows = False
                                    RowSetting = TBLSettingDummy.DefaultView(i).Row
                                    TBLSettingItem.DefaultView.RowFilter = Nothing
                                    TBLSettingItem.DefaultView.RowFilter = "[Identifier] = '" & RowSetting.identifier_jo & "'"
                                    If TBLSettingItem.DefaultView.Count > 0 Then
                                        RowSettingItem = TBLSettingItem.DefaultView(0).Row
                                        dgvItem.Rows.Add()
                                        dr = dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                                        dr.Cells(Me.ColJONo.Index).Value = RowSettingItem.JO_Number
                                        dr.Cells(Me.ColCodeNo.Index).Value = RowSettingItem.Product_Code
                                        dr.Cells(Me.ColDesc.Index).Value = RowSettingItem.Description
                                        dr.Cells(Me.ColMO.Index).Value = RowSettingItem.MO_Number
                                        dr.Cells(Me.ColIdentifierJO.Index).Value = RowSettingItem.Identifier
                                        Me.dgvItem.AllowUserToAddRows = True
                                    End If

                                Next
                            End If
                        Else
                            If NoSetting = False Then
                                NoSetting = False
                            End If
                        End If
kesini:
                    Next
                End If

                If NoSetting = False Then
                    baru = True
                    Me.dgvItem.AllowUserToAddRows = False
                    TBLSettingItem.DefaultView.RowFilter = Nothing
                    TBLSettingItem.DefaultView.RowFilter = "[Identifier] = '" & Identifier_JO & "'"
                    If TBLSettingItem.DefaultView.Count > 0 Then
                        RowSettingItem = TBLSettingItem.DefaultView(0).Row
                        dgvItem.Rows.Add()
                        dr = dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        dr.Cells(Me.ColJONo.Index).Value = RowSettingItem.JO_Number
                        dr.Cells(Me.ColCodeNo.Index).Value = RowSettingItem.Product_Code
                        dr.Cells(Me.ColDesc.Index).Value = RowSettingItem.Description
                        dr.Cells(Me.ColMO.Index).Value = RowSettingItem.MO_Number
                        dr.Cells(Me.ColIdentifierJO.Index).Value = Identifier_JO
                        Me.dgvItem.AllowUserToAddRows = True
                    End If
                    ProsesHitungTanggal()
                End If
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

                'cek header jika di item settingnya sudah tidak ada, maka hapus headernya
                DTASetting.Fill(TBLSetting)
                For ij As Integer = 0 To listHapusJo.Count - 1
                    TBLSetting.DefaultView.RowFilter = Nothing
                    TBLSetting.DefaultView.RowFilter = "[primary transaksi] = " & listHapusJo.Item(ij)
                    If TBLSetting.DefaultView.Count = 0 Then
                        Hapus(listHapusJo(ij))
                    End If
                    If listHapusJo(ij) = PrimaryTransaksi_ Then
                        Ubah = MsgBox("There are changes to the settings, reload the date of R/P, PA and FN?", MsgBoxStyle.YesNo, "Confirm")
                        If Ubah = MsgBoxResult.Yes Then
                            Me.dgvItem.AllowUserToAddRows = True
                            ProsesHitungTanggal()
                        End If
                    End If
                Next

            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Function ApakahMasihAdaJO(ByVal vIdentifierJo As String) As Boolean
        Dim dta As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_cari_joTableAdapter
        Dim dummy As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJmlByIdentifier(vIdentifierJo)
            If dummy.HasValue Then
                If dummy.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function
    Public Function HapusItemSettingByidentifierJO(ByVal vIdentifierJO As String, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_jo_item2TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery_ByIdentifierJO(vIdentifierJO, vPrimaryTransaksi)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Overridable Function Hapus(ByVal vPrimarySetting As Double) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_joTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Delete(CDec(vPrimarySetting)) > 0 Then
                Return True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Overridable Function CariDetail(ByVal vPrimarySetting As Double, _
       ByVal vDGVItem As DataGridView) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_joTableAdapter
        Dim TBL As New DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_joDataTable
        Dim row As DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_joRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'DTA.FillBy_PrimaryTransaksi(TBL, CDec(Primary))

            If DTA.FillBy_PrimarySetting(TBL, CDec(vPrimarySetting)) > 0 Then
                row = TryCast(TBL.Rows(0), DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_joRow)
                'row = TBL.Rows(0)

                With row

                    Me.txtPrimary.Text = vPrimarySetting.ToString
                    Me.dtpPrintDate.Value = .print_date.Date
                    Me.txtPrimaryCustomer.Text = .primary_customer.ToString
                    Me.txtCustomerCode.Text = objCustomer.GetCustomerCode(txtPrimaryCustomer.Text)
                    Me.txtCustomerDesc.Text = objCustomer.GetCustName(txtPrimaryCustomer.Text)
                    Me.txtFU.Text = .fu
                    Me.txtPoSo.Text = .po_so
                    Me.cboRevision.Text = .revision
                    Me.dtpPCKFrom.Value = .raw_pencil_from
                    Me.dtpPCKTo.Value = .raw_pencil_to
                    Me.tglpckFromChache = .raw_pencil_from
                    Me.tglpckToChache = .raw_pencil_to
                    Me.txtRemarks1.Text = .remarks_1
                    Me.txtShrinkPack.Text = Me.objCd.DigitGrouping(.shrink_pack, 4)
                    Me.txtOuterPack.Text = Me.objCd.DigitGrouping(.outer_pack, 4)
                    If Me.dtpPCKTo.Value = #1/1/1950# Then
                        Me.dtpPCKTo.Value = #1/1/2200#
                    Else
                        Me.chkPCK.Checked = True
                        Me.dtpPCKFrom.Enabled = True
                        Me.dtpPCKTo.Enabled = True
                    End If

                    Me.Timestamp = DTA.ScalarQuery_GetTimestampByprimary(Me.txtPrimary.Text)
                End With

            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            'ObjItem = Nothing
        End Try
    End Function

    Public Sub ProsesHitungTanggal()
        Dim DTAMO As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim TBLMO As New DseTrans_ManufacturingOrder_PrintJO.tbltrans_mo_item_woDataTable
        Try

            DTAMO.Connection.ConnectionString = Core_Module.GetConstring
            Me.dtpPCKFrom.Value = #1/1/2200#
            Me.dtpPCKTo.Value = #1/1/1950#
            For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                DTAMO.ClearBeforeFill = False
                DTAMO.Fillby_IdentifierJO(TBLMO, Me.dgvItem.Item(Me.ColIdentifierJO.Index, i).Value)
            Next
            'Menentukan tanggal

            TentukanDTP(Me.dgvItem, TBLMO)

            Me.tglpckFromChache = #1/1/1950#
            Me.tglpckToChache = #1/1/2200#
            If Me.dtpPCKFrom.Value = #1/1/2200# Then
                Me.dtpPCKFrom.Value = #1/1/1950#
            End If
            If Me.dtpPCKTo.Value = #1/1/1950# Then
                Me.dtpPCKTo.Value = #1/1/2200#
            Else
                Me.chkPCK.Checked = True
            End If


            TBLMO.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAMO.Dispose()
        End Try

    End Sub
    Public Sub TentukanDTP(ByVal vDGV As DataGridView, ByVal vTBLMO As DseTrans_ManufacturingOrder_PrintJO.tbltrans_mo_item_woDataTable)
        Dim DTATipe As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_jo_tipe_prosesTableAdapter
        Dim TBLTipe As New DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_jo_tipe_prosesDataTable

        Dim TBLMODummy As New DseTrans_ManufacturingOrder_PrintJO.tbltrans_mo_item_woDataTable
        Dim RowMO As DseTrans_ManufacturingOrder_PrintJO.tbltrans_mo_item_woRow

        Try
            DTATipe.Connection.ConnectionString = Core_Module.GetConstring
            DTATipe.Fill(TBLTipe)
            For j As Integer = 0 To vDGV.Rows.Count - 2
                vTBLMO.DefaultView.RowFilter = Nothing
                vTBLMO.DefaultView.RowFilter = "[Identifier JO] = '" & vDGV.Item(Me.ColIdentifierJO.Index, j).Value.ToString _
                & "' and [Identifier WO Parent] = '-'"
                If vTBLMO.DefaultView.Count > 0 Then
                    For Each DV As Data.DataRowView In vTBLMO.DefaultView
                        RowMO = DV.Row
                        If RowMO.Identifier_WO_Parent = "-" Then
                            If RowMO.Status <> 0 Then
                                Me.SetDtp(RowMO, TBLTipe)
                            End If
                        End If
                        Me.AmbilTanggal(RowMO.Identifier_JO, vTBLMO, RowMO.Identifier, TBLTipe)
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub AmbilTanggal(ByVal vIdentifierJO As String, _
    ByVal vTBLMO As DseTrans_ManufacturingOrder_PrintJO.tbltrans_mo_item_woDataTable, _
    ByVal vIdentifier As String, ByVal vTBLTipe As DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_jo_tipe_prosesDataTable)
        Dim rowMO As DseTrans_ManufacturingOrder_PrintJO.tbltrans_mo_item_woRow

        Try
            vTBLMO.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO.ToString & "' and [Identifier WO Parent] = '" & vIdentifier.ToString & "'"
            If vTBLMO.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTBLMO.DefaultView
                    rowMO = dv.Row
                    If rowMO.Status <> 0 Then
                        Me.SetDtp(rowMO, vTBLTipe)
                    End If
                    Me.AmbilTanggal(rowMO.Identifier_JO, vTBLMO, rowMO.Identifier, vTBLTipe)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub SetDtp(ByVal vRowMO As DseTrans_ManufacturingOrder_PrintJO.tbltrans_mo_item_woRow, _
   ByVal vTBLTipe As DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_jo_tipe_prosesDataTable)
        Dim RowTipe As DseTrans_ManufacturingOrder_PrintJO.tbltrans_custom_cetak_jo_tipe_prosesRow

        Try
            vTBLTipe.DefaultView.RowFilter = Nothing
            vTBLTipe.DefaultView.RowFilter = "[tipe] = " & EnumTypeProcess.PACKING
            For Each dv As Data.DataRowView In vTBLTipe.DefaultView
                RowTipe = dv.Row
                If vRowMO.Primary_Process = RowTipe.primary_process Then
                    If vRowMO.Plan_Start_Time < Me.dtpPCKFrom.Value.Date Then
                        Me.dtpPCKFrom.Value = vRowMO.Plan_Start_Time.Date
                    End If
                    If vRowMO.Plan_End_Time > Me.dtpPCKTo.Value.Date Then
                        Me.dtpPCKTo.Value = vRowMO.Plan_End_Time.Date
                    End If
                    Me.chkPCK.Checked = True
                    Me.dtpPCKFrom.Enabled = True
                    Me.dtpPCKTo.Enabled = True
                End If
            Next


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If Me.dgvItem.Item(ColJONo.Index, e.RowIndex).Value = "" Then
                If e.ColumnIndex <> ColJONo.Index Then
                    e.Cancel = True
                Else
                    If Me.dgvItem.Item(Me.ColIdentifierJO.Index, e.RowIndex).Value = Identifier_JO Then
                        e.Cancel = True
                    Else
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            Else
                If e.ColumnIndex = ColJONo.Index Then
                    If Me.dgvItem.Item(Me.ColIdentifierJO.Index, e.RowIndex).Value = Identifier_JO Then
                        e.Cancel = True
                    Else
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Dim Identifier_JO As String = ""
        Dim index As Integer = -1

        Try
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColJONo.Index Then

                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ApakahSudahAda(Me.dgvItem.Item(Me.ColJONo.Index, e.RowIndex).Value) = True Then
                            Identifier_JO = Me.GetIdentifierJo(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItem(Identifier_JO, e.RowIndex)
                            If index < 0 Then
                                If Me.ApakahSudahAdaSetting(Identifier_JO, Me.Type_Jo, Me.txtPrimary.Text) = False Then
                                    Me.dgvItem.Item(Me.ColIdentifierJO.Index, e.RowIndex).Value = Identifier_JO
                                    Me.dgvItem.Item(Me.ColJONo.Index, e.RowIndex).Value = Me.GetJoNo(Identifier_JO)
                                    Me.dgvItem.Item(Me.ColCodeNo.Index, e.RowIndex).Value = Me.GetProduct(Identifier_JO)
                                    Me.dgvItem.Item(Me.ColDesc.Index, e.RowIndex).Value = Me.GetDesc(Identifier_JO)
                                    Me.dgvItem.Item(Me.ColMO.Index, e.RowIndex).Value = Me.GetMONo(Identifier_JO)
                                Else
                                    MsgBox("Jo No is already used in other settings, please input another Jo No.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(Me.ColJONo.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.objCd.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication Jo No, please input another Jo No.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(Me.ColJONo.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.objCd.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                            End If
                            ProsesHitungTanggal()
                        Else
                            Dim ij As New FrmTrans_ManufacturingOrder_PrintJO_Cari
                            ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.IdentifierJO <> "" Then
                                index = Me.GetIndexItem(ij.IdentifierJO, e.RowIndex)
                                If index < 0 Then
                                    If Me.ApakahSudahAdaSetting(ij.IdentifierJO, Me.Type_Jo, Me.txtPrimary.Text) = False Then
                                        Me.dgvItem.Item(Me.ColIdentifierJO.Index, e.RowIndex).Value = ij.IdentifierJO
                                        Me.dgvItem.Item(Me.ColJONo.Index, e.RowIndex).Value = Me.GetJoNo(ij.IdentifierJO)
                                        Me.dgvItem.Item(Me.ColCodeNo.Index, e.RowIndex).Value = Me.GetProduct(ij.IdentifierJO)
                                        Me.dgvItem.Item(Me.ColDesc.Index, e.RowIndex).Value = Me.GetDesc(ij.IdentifierJO)
                                        Me.dgvItem.Item(Me.ColMO.Index, e.RowIndex).Value = Me.GetMONo(ij.IdentifierJO)
                                    Else
                                        MsgBox("Jo No is already used in other settings, please input another Jo No.", MsgBoxStyle.Information, "Information")
                                        Me.dgvItem.Item(Me.ColJONo.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.objCd.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    End If
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication Jo No, please input another Jo No.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.objCd.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    Else
                                        If Me.ApakahSudahAdaSetting(ij.IdentifierJO, Me.Type_Jo, Me.txtPrimary.Text) = False Then
                                            Me.dgvItem.Item(Me.ColIdentifierJO.Index, e.RowIndex).Value = ij.IdentifierJO
                                            Me.dgvItem.Item(Me.ColJONo.Index, e.RowIndex).Value = Me.GetJoNo(ij.IdentifierJO)
                                            Me.dgvItem.Item(Me.ColCodeNo.Index, e.RowIndex).Value = Me.GetProduct(ij.IdentifierJO)
                                            Me.dgvItem.Item(Me.ColDesc.Index, e.RowIndex).Value = Me.GetDesc(ij.IdentifierJO)
                                            Me.dgvItem.Item(Me.ColMO.Index, e.RowIndex).Value = Me.GetMONo(ij.IdentifierJO)
                                        Else
                                            MsgBox("Jo No is already used in other settings, please input another Jo No.", MsgBoxStyle.Information, "Information")
                                            Me.dgvItem.Item(Me.ColJONo.Index, e.RowIndex).Value = ""
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowIndex = e.RowIndex
                                                Me.BackgroundWorker1.RunWorkerAsync()
                                            Else
                                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            Me.objCd.SetNumber(Me.dgvItem)
                                            Exit Sub
                                        End If
                                    End If
                                End If
                                ProsesHitungTanggal()
                            ElseIf ij.IdentifierJO = "" Then
                                If Me.dgvItem.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.objCd.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                                If Me.dgvItem.Item(Me.ColJONo.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.objCd.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.objCd.SetNumber(Me.dgvItem)
                    End If

                Else

                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    Else
                        Me.dgvItem.Item(Me.ColJONo.Index, e.RowIndex).Value = Me.Buffer
                    End If
                    Me.objCd.SetNumber(Me.dgvItem)
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Function ApakahSudahAda(ByVal vJoNo As String) As Boolean
        Dim dta As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_cari_joTableAdapter
        Dim dummy As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJmlByCode(vJoNo)
            If dummy.HasValue Then
                If dummy.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function
    Public Function GetIdentifierJo(ByVal vJoNo As String) As String

        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_cari_joTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetIdentifierbyJONo(vJoNo)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetJoNo(ByVal vIdentifier As String) As String
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_itemTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJoNoByIdentifier(vIdentifier)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetProduct(ByVal vIdentifier As String) As String
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_itemTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetProductByIdentifier(vIdentifier)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetDesc(ByVal vIdentifier As String) As String
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_itemTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetDescByIdentifier(vIdentifier)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetMONo(ByVal vIdentifier As String) As String
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_custom_cetak_itemTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetMoNOByIdentifier(vIdentifier)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetIndexItem(ByVal vIdentifier As String, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            'pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.dgvItem.Rows.Count - 2
                RowDGV = Me.dgvItem.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColIdentifierJO.Index).Value = vIdentifier Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            Me.dgvItem.TabStop = True
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            With rowDGV
                .Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColJONo.Index).Value = ""
                .Cells(Me.ColCodeNo.Index).Value = ""
                .Cells(Me.ColDesc.Index).Value = ""
                .Cells(Me.ColMO.Index).Value = ""
                .Cells(Me.ColIdentifierJO.Index).Value = "-1-0"
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            rowDGV = Nothing
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            If Ubah = MsgBoxResult.Yes Then
                ProsesHitungTanggal()
            End If
            Me.objCd.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow

        Try
            If Me.dgvItem(Me.ColIdentifierJO.Index, Me.dgvItem.CurrentRow.Index).Value = Me.Identifier_JO Then
                e.Cancel = True
            Else
                Ubah = MsgBox("There are changes to the settings, reload the date of R/P, PA and FN?", MsgBoxStyle.YesNo, "Confirm")
                If Ubah = MsgBoxResult.No Then
                    e.Cancel = True
                End If
                Me.objCd.SetNumber(Me.dgvItem)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try

            hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
            If hasil = MsgBoxResult.Yes Then
                Me.Close()
            Else
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkPCK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPCK.Click
        Try
            If Me.chkPCK.CheckState = CheckState.Unchecked Then
                Me.dtpPCKFrom.Value = #1/1/1950#
                Me.dtpPCKTo.Value = #1/1/2200#
                Me.dtpPCKFrom.Enabled = False
                Me.dtpPCKTo.Enabled = False
            ElseIf Me.chkPCK.CheckState = CheckState.Checked Then
                Me.dtpPCKFrom.Value = Me.tglpckFromChache
                Me.dtpPCKTo.Value = Me.tglpckToChache
                Me.dtpPCKFrom.Enabled = True
                Me.dtpPCKTo.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class