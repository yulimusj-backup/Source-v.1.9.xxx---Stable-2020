Option Explicit On
Option Strict Off

Public Class FrmImport_eFaktur
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim FileExtStr As String = ""

    Dim Tbl As New DseImport_eFaktur.tbl_temporaryDataTable
    Dim row As DseImport_eFaktur.tbl_temporaryRow
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjCD As New Core_Module.ClsCekData
    Dim DefPrec As Integer = ObjCurr.GetDecimal_DefaultCurrency
    Dim ObjInv As New Sales_Module.ClsTransaksiSalesInvoice
    Dim ObjInvoiceItem As New Sales_Module.ClsTransaksiSalesInvoiceItem

    Dim dtaCust As New DseImport_eFakturTableAdapters.tblmaster_customerTableAdapter
    Dim tblCust As New DseImport_eFaktur.tblmaster_customerDataTable
    Dim rowCust As DseImport_eFaktur.tblmaster_customerRow
    Dim ObjMonitor As New Core_Module.ClsTransaksiMonitor
    Dim Identifier As Integer = -1
    Dim Timer1 As New Timer
    Dim ObjRandom As New Random
    Dim isQtyMinus As Boolean = False

    Private Sub FrmImport_eFaktur_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.Timer1.Enabled = False
            Me.ObjMonitor.EndMonitor(Core_Module.ClsTransaksiMonitor.EnumTipeTransaksi.Custom_ImportEfaktur, Me.Identifier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub FrmImport_eFaktur_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'tidak bisa import jika
        'customer pada file csv import tidak ada pada smess (cocokan berdasarkan customer name)
        'invoice no pada file csv import sudah ada pada smess
        'tax invoice no pada file csv import tidak masuk ke range tax invoice no smess /buat dulu di smess pada tax invoice setting
        'product pada file csv import tidak ada pada smess (cocokan berdasarkan product description)

        'tambahan
        'apakah invalid jika npwp & alamat pada file csv dengan yang ada di smess berbeda? ambil yang ada di smess
        Try
            Me.dgvHeader.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.Identifier = Me.ObjRandom.Next(0, 10 ^ 7)
            Me.ObjMonitor.StartMonitor(Me.UserCode, My.Computer.Name, Core_Module.ClsTransaksiMonitor.EnumTipeTransaksi.Custom_ImportEfaktur, Me.Identifier)
            Me.Timer1.Interval = Me.ObjMonitor.MonitorInterval
            Me.Timer1.Enabled = True

            Me.colPrimary.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindFilePath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindFilePath.Click

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.Openfile(FileExtStr) = False Then
                Me.txtFilePath.Text = ""
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function Openfile(ByRef vFileExtStr As String) As Boolean
        Dim FileExtStr As String = ""

        Try
            '1. cari file csv format infinity
            Me.OpenFileDialog1.Title = "Please specify file"
            Me.OpenFileDialog1.Filter = "csv files (*.csv)|*.csv"
            'Me.OpenFileDialog1.Filter = "Excel Worksheets (*.xlsx)|*.xlsx|excel files (*.xls)|*.xls|csv files (*.csv)|*.csv"
            Me.OpenFileDialog1.FilterIndex = 1
            Me.OpenFileDialog1.DefaultExt = "csv"
            Me.OpenFileDialog1.CheckFileExists = True
            Me.OpenFileDialog1.Multiselect = False
            Me.OpenFileDialog1.FileName = ""
            If Me.OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.OK Then
                Return False
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Return False
            End If

            Me.txtFilePath.Text = Me.OpenFileDialog1.FileName
            vFileExtStr = System.IO.Path.GetExtension(Me.OpenFileDialog1.FileName)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Try
            'cek saat tekan btnShow lokasi path csv nya apakah masih valid
            'filter file csv berdasarkan date from-to dan juga chkshowonlyvalid
            Me.ProgressBar1.Visible = False
            Me.ProgressBar1.Maximum = 0
            Me.llbProgressBar.Text = ""
            Me.dgvHeader.ClearSelection()
            If FileExtStr.StartsWith(".csv") Then
                Me.ShowToDgvHeader()
            End If
            Me.colInvoiceNumber.ReadOnly = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function UpdateLine(ByVal strArr As String(), ByVal IndexSalah As Integer) As String()
        Try
            strArr(IndexSalah - 1) = strArr(IndexSalah - 1).ToString + strArr(IndexSalah).ToString
            strArr(IndexSalah) = "data salah kolom"
            Dim newArr As String()
            Dim j As String = ""
            For i As Integer = 0 To strArr.Length - 1
                If strArr(i) <> "data salah kolom" Then
                    If j = "" Then
                        j = strArr(i)
                    Else
                        j &= "," & strArr(i)
                    End If
                End If
            Next
            'strArr = j
            newArr = j.Split(",")
            Return newArr
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub ShowToDgvHeader()
        Dim SR As System.IO.StreamReader = New System.IO.StreamReader(Me.txtFilePath.Text)
        Dim line As String = SR.ReadLine()
        line = line.Replace(";", ",")
        Dim iCount As String() = line.Split(","c)
        Dim ICount2 As String() = line.Split(";"c)
        Dim strArray As String()

        'aul
        Dim Primary_Temp As Double = 0
        Dim Customer_Temp As String = ""
        Dim InvoiceNo_Temp As String = ""
        Dim InvoiceDate_Temp As Date = "1/1/1950"
        Dim TaxInvoiceNo1_Temp As String = ""
        Dim TaxInvoiceNo2_Temp As String = ""
        Dim Npwp_Temp As String = ""
        Dim DPP_Temp As Decimal = 0
        Dim PPN_Temp As Decimal = 0
        Dim EmptyCount As Integer = 0
        Dim ListPrimaryTrans As New List(Of Double)
        Dim rd As DataGridViewRow
        Dim LolosFilterTgl As Boolean = False
        Dim ListInvoiceNo As New List(Of String)
        Dim t As DataTable = New DataTable
        Dim r As DataRow
        'Comment: Perubahan NPWP 16 Digit
        Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
        Dim tglBerlaku As Date = ObjSetNPWP.Tglberlaku
        Try
            Me.Tbl.Rows.Clear()
            If iCount.Length > ICount2.Length Then
                strArray = line.Split(","c)
            Else
                strArray = line.Split(";"c)
            End If
            'If Tbl.Columns.Count <> 3 Then
            '    MsgBox("Make sure csv file consist of name and amount.", MsgBoxStyle.Information, "Information")
            '    Exit Sub
            'End If

            ''''''cari invoice pengganti
            t = Nothing
            t = New DataTable
            For i As Integer = 0 To 2
                If i = 0 Then
                    t.Columns.Add("primary", GetType(String), "")
                ElseIf i = 1 Then
                    t.Columns.Add("invoice no", GetType(String), "")
                Else
                    t.Columns.Add("tipe pengganti", GetType(String), "")
                End If
            Next

            Do
                line = SR.ReadLine

                If Not line = String.Empty Then
                    line = line.Replace(";", ",")
                    Dim SplitValue As String = ","
                    Dim StrPerLine As String = ""
                    StrPerLine = line.Replace(ControlChars.Quote, "").Replace(Chr(34), "").Replace(";", ",")
                    strArray = StrPerLine.Split(","c)

                    If strArray(0).Replace(Chr(34), "") = "FK" Then
                        If strArray(0).Replace(ControlChars.Quote, "").Replace(Chr(34), "") = "FK" Then
                            While strArray.Length < 19
                                line = line + SR.ReadLine
                                StrPerLine = line.Replace(ControlChars.Quote, "").Replace(Chr(34), "")
                                strArray = StrPerLine.Split(","c)
                            End While

                            If strArray.Length > iCount.Length Then
                                While strArray.Length > iCount.Length
                                    strArray = Me.UpdateLine(strArray, 10)
                                End While
                            End If
                        End If
                    End If

                    If strArray(0).Replace(Chr(34), "") = "FK" And strArray.Length >= iCount.Length Then
                        If String.IsNullOrEmpty(strArray(18).Replace(Chr(34), "").Trim) = False Then
                            r = t.NewRow
                            r.Item("primary") = t.Rows.Count + 1
                            r.Item("invoice no") = strArray(18).Replace(Chr(34), "").Trim
                            r.Item("tipe pengganti") = strArray(2).Replace(Chr(34), "").Trim
                            t.Rows.Add(r)
                        End If
                    End If
                Else
                    'Exit Do
                    EmptyCount += 1
                    If EmptyCount > 50 Then
                        Exit Do
                    End If
                End If
            Loop

            ''''''''''''''''''''''''''''''''''''''''''''''''''
            EmptyCount = 0
            Me.Tbl.Rows.Clear()
           
            SR = New System.IO.StreamReader(Me.txtFilePath.Text)
            line = SR.ReadLine()
            Do
                line = SR.ReadLine

                If Not line = String.Empty Then
                    line = line.Replace(";", ",")
                    Dim SplitValue As String = ","
                    Dim StrPerLine As String = ""
                    StrPerLine = line.Replace(ControlChars.Quote, "").Replace(Chr(34), "").Replace(";", ",")
                    strArray = StrPerLine.Split(","c)

                    'If iCount.Length > ICount2.Length Then
                    '    strArray = line.Split(New String() {Chr(34) & "," & Chr(34)}, StringSplitOptions.None)
                    'Else
                    '    strArray = line.Split(New String() {Chr(34) & ";" & Chr(34)}, StringSplitOptions.None)
                    'End If

                    If strArray(0).Replace(Chr(34), "") = "FK" Then
                        'If strArray(3).Replace(Chr(34), "").Contains("32031650437") = True Then
                        '    MsgBox("")
                        'End If
                        'Dim StrPerLine As String = ""
                        'StrPerLine = line.Replace(ControlChars.Quote, "").Replace(Chr(34), "")
                        'strArray = StrPerLine.Split(","c)
                        If strArray(0).Replace(ControlChars.Quote, "").Replace(Chr(34), "") = "FK" Then
                            While strArray.Length < 19 'And strArr(IdxCheck).Contains(ControlChars.Quote)
                                line = line + SR.ReadLine
                                StrPerLine = line.Replace(ControlChars.Quote, "").Replace(Chr(34), "")
                                strArray = StrPerLine.Split(","c)
                            End While

                            If strArray.Length > iCount.Length Then
                                While strArray.Length > iCount.Length
                                    strArray = Me.UpdateLine(strArray, 10)
                                End While
                            End If
                        End If
                    End If

                    

                    If strArray(0).Replace(Chr(34), "") = "FK" And strArray.Length >= iCount.Length Then
                        'cek invoice pengganti
                    
                        t.DefaultView.Sort = "[tipe pengganti] DESC"
                        t.DefaultView.RowFilter = "[invoice no]='" & strArray(18).Replace(Chr(34), "").Trim & "'"
                        If t.DefaultView.Count > 1 Then
                            For i As Integer = 0 To t.DefaultView.Count - 1
                                r = t.DefaultView(i).Row
                                If CDec(r.Item("tipe pengganti")) > CDec(strArray(2).Replace(Chr(34), "").Trim) Then
                                    GoTo Line1
                                End If
                            Next
                        End If

                        Primary_Temp += 1
                        Customer_Temp = strArray(8).Replace(Chr(34), "").Trim
                        InvoiceNo_Temp = strArray(18).Replace(Chr(34), "").Trim
                        InvoiceDate_Temp = Me.GetConvertionDateFromString(strArray(6).Replace(Chr(34), ""))
                        TaxInvoiceNo1_Temp = strArray(1).Replace(Chr(34), "") & strArray(2).Replace(Chr(34), "")
                        TaxInvoiceNo2_Temp = Me.SetFormatTaxNo(strArray(3).Replace(Chr(34), ""))
                        'Comment: Perubahan NPWP 16 Digit
                        If InvoiceDate_Temp < tglBerlaku Then
                            Npwp_Temp = Me.SetFormatNPWP(strArray(7).Replace(Chr(34), ""))
                        Else
                            Npwp_Temp = strArray(7).Replace(Chr(34), "")
                        End If
                        DPP_Temp = strArray(10).Replace(Chr(34), "")
                        PPN_Temp = strArray(11).Replace(Chr(34), "")

                        'filter by date from-to
                        LolosFilterTgl = True
                        If (InvoiceDate_Temp < Me.dtpFrom.Value.Date) Or (InvoiceDate_Temp > Me.dtpTo.Value.Date) Then
                            'Else
                            LolosFilterTgl = False
                        End If
                    End If

                    'If InvoiceNo_Temp.Contains("FJ.0120.00031") = True Then
                    '    MsgBox("")
                    'End If
                    If Primary_Temp > 0 And strArray(0).Replace(Chr(34), "") = "OF" And LolosFilterTgl = True Then
                        row = Tbl.NewRow
                      
                        row.primary = Primary_Temp
                        row.customer = Customer_Temp
                        row.invoice_no = InvoiceNo_Temp
                        row.invoice_date = InvoiceDate_Temp
                        row.tax_invoice_no_1 = TaxInvoiceNo1_Temp
                        row.tax_invoice_no_2 = TaxInvoiceNo2_Temp
                        row.npwp = Npwp_Temp
                        If row.npwp = "000000000000000" Then
                            row.npwp = "00.000.000.0-000.000"
                        End If
                        row.dpp = DPP_Temp
                        row.ppn = PPN_Temp
                        row.prod_code = strArray(1).Replace(Chr(34), "")
                        row.prod_desc = strArray(2).Replace(Chr(34), "")
                        row.prod_unit_price = strArray(3).Replace(Chr(34), "")
                        row.prod_qty = strArray(4).Replace(Chr(34), "")
                        row.prod_amount = strArray(5).Replace(Chr(34), "")
                        row.prod_disc_amount = strArray(6).Replace(Chr(34), "")
                        row.prod_unit = "" 'nanti di isi dengan unit default product
                        row.prod_dpp = strArray(7).Replace(Chr(34), "")
                        row.prod_ppn = strArray(8).Replace(Chr(34), "")

                        row.invoice_valid = False
                        row.invalid_reason = ""
                        row.prod_invalid_reason = ""
                        Tbl.Rows.Add(row)
                    End If
Line1:
                Else
                    'Exit Do
                    EmptyCount += 1
                    If EmptyCount > 50 Then
                        Exit Do
                    End If
                End If
            Loop

            Me.CekFileImport(Me.Tbl)

            'filter invalid
            For i As Integer = Tbl.Rows.Count - 1 To 0 Step -1
                Me.row = Me.Tbl.Rows(i)
                Dim diHapus As Boolean = False
                If Me.ChkShowInvalidOnly.Checked = True Then
                    If String.IsNullOrEmpty(Me.row.invalid_reason) = True Then
                        Me.Tbl.Rows.Remove(Me.row)
                        diHapus = True
                    End If
                End If
                If diHapus = False Then
                    If String.IsNullOrEmpty(Me.row.invalid_reason) = False Then
                        If Me.row.invalid_reason.Contains("Invoice already registered in SMESS") = True Then
                            If Me.ChkShowInvoiceReg.Checked = False Then
                                Me.Tbl.Rows.Remove(Me.row)
                            End If
                        End If
                    End If
                End If
            Next

            Tbl.DefaultView.RowFilter = ""
            Tbl.DefaultView.Sort = "[invoice date], [invoice no], [prod invalid reason] DESC"
            Me.dgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvHeader.Rows.Clear()
            Me.dgvItem.Rows.Clear()
            For i As Integer = 0 To Tbl.DefaultView.Count - 1
                row = Tbl.DefaultView(i).Row
                With row
                    If ListPrimaryTrans.Contains(.primary) = False Then
                        ListPrimaryTrans.Add(.primary)
                        rd = Me.dgvHeader.Rows(Me.dgvHeader.Rows.Add)
                        rd.Cells(Me.colNo.Index).Value = rd.Index + 1
                        rd.Cells(Me.colCreate.Index).Value = False
                        rd.Cells(Me.colCustomer.Index).Value = .customer
                        rd.Cells(Me.ColCustCode.Index).Value = .Customer_Code
                        rd.Cells(Me.ColCustName.Index).Value = .Customer_Name
                        rd.Cells(Me.colInvoiceNumber.Index).Value = .invoice_no
                        rd.Cells(Me.colTaxInvoiceNumber.Index).Value = .tax_invoice_no_1 & .tax_invoice_no_2
                        rd.Cells(Me.colDate.Index).Value = .invoice_date
                        rd.Cells(Me.colNPWP.Index).Value = .npwp
                        rd.Cells(Me.colDPP.Index).Value = Me.ObjCD.DigitGrouping(.dpp, DefPrec)
                        rd.Cells(Me.colPPN.Index).Value = Me.ObjCD.DigitGrouping(.ppn, DefPrec)
                        rd.Cells(Me.colInvalidReason.Index).Value = .invalid_reason
                        rd.Cells(Me.colPrimary.Index).Value = .primary
                        If String.IsNullOrEmpty(.invalid_reason) = False Then
                            Me.dgvHeader.Rows(rd.Index).DefaultCellStyle.BackColor = Color.MistyRose
                        End If
                        If .invalid_reason.ToString.Contains("Invoice already registered in SMESS") = True Then
                            Me.dgvHeader.Rows(rd.Index).DefaultCellStyle.BackColor = Color.LightGreen
                        End If
                    End If
                End With
            Next
            Me.dgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            If Me.dgvHeader.RowCount > 0 Then
                Me.ShowToDgvItem(Me.dgvHeader(Me.colPrimary.Index, 0).Value)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CekFileImport(ByVal vTbl As DseImport_eFaktur.tbl_temporaryDataTable)
        Dim dtaInv As New DseImport_eFakturTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim tblInv As New DseImport_eFaktur.tbltrans_salesinvoiceDataTable
        Dim rowInv As DseImport_eFaktur.tbltrans_salesinvoiceRow
       
        Dim dtaTax As New DseImport_eFakturTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter
        Dim tblTax As New DseImport_eFaktur.tbltrans_salesinvoice_tax_invoice_settingDataTable
        Dim rowTax As DseImport_eFaktur.tbltrans_salesinvoice_tax_invoice_settingRow
        Dim dtaProd As New DseImport_eFakturTableAdapters.tblproduct_connectorTableAdapter
        Dim tblProd As New DseImport_eFaktur.tblproduct_connectorDataTable
        Dim rowProd As DseImport_eFaktur.tblproduct_connectorRow

        Dim tblCustModif As New DseImport_eFaktur.tblmaster_customerDataTable
        Dim rowCustModif As DseImport_eFaktur.tblmaster_customerRow
        Dim dtaqty As New DseImport_eFakturTableAdapters.tblprod_qtyTableAdapter
        Dim tblqty As New DseImport_eFaktur.tblprod_qtyDataTable
        Dim rowqty As DseImport_eFaktur.tblprod_qtyRow
        Try
            dtaInv.Connection.ConnectionString = Core_Module.GetConstring
            dtaCust.Connection.ConnectionString = Core_Module.GetConstring
            dtaTax.Connection.ConnectionString = Core_Module.GetConstring
            dtaProd.Connection.ConnectionString = Core_Module.GetConstring
            dtaqty.Connection.ConnectionString = Core_Module.GetConstring
            dtaCust.Fill(tblCust)
            dtaInv.Fill(tblInv)
            dtaTax.Fill(tblTax)
            dtaProd.Fill(tblProd)
            dtaqty.Fill(tblqty)
            tblCustModif = tblCust.Copy
            'invoice no pada file csv import sudah ada pada smess
            tblInv.DefaultView.RowFilter = ""
            tblInv.DefaultView.Sort = "[Invoice No]"
            vTbl.DefaultView.RowFilter = ""
            vTbl.DefaultView.Sort = "[invoice no]"
            For i As Integer = 0 To vTbl.DefaultView.Count - 1
                row = vTbl.DefaultView(i).Row
                With row
                    While tblInv.DefaultView.Count > 0
                        rowInv = tblInv.DefaultView(0).Row
                        'If rowInv.Invoice_No.Contains("FJ.0120.00022") = True Then
                        '    MsgBox("")
                        'End If
                        If String.IsNullOrEmpty(.invoice_no) = True Then
                            .invalid_reason = "Invoice number is empty"
                            Exit While
                        ElseIf .invoice_no < rowInv.Invoice_No Then
                            Exit While
                        ElseIf .invoice_no = rowInv.Invoice_No Then
                            'If Me.ChkShowInvoiceReg.Checked = True Then
                            '    .invalid_reason = "Invoice already registered in SMESS"
                            '    Exit While
                            'Else
                            '    tblInv.Rows.Remove(rowInv)
                            '    Exit While
                            'End If
                            .invalid_reason = "Invoice already registered in SMESS"
                            'tblInv.Rows.Remove(rowInv)
                            Exit While
                        Else
                            tblInv.Rows.Remove(rowInv)
                        End If
                    End While
                End With
            Next

            'customer pada file csv import tidak ada pada smess (cocokan berdasarkan customer name)
            tblCustModif.DefaultView.RowFilter = ""
            tblCustModif.DefaultView.Sort = "[Customer Name]"
            vTbl.DefaultView.RowFilter = ""
            vTbl.DefaultView.Sort = "[customer]"
            Dim CustKetemu As Boolean = False
            For i As Integer = 0 To vTbl.DefaultView.Count - 1
                row = vTbl.DefaultView(i).Row
                With row
                    'If .invoice_no.Contains("FJ.0120.00022") = True Then
                    '    MsgBox("")
                    'End If
                    'CustKetemu = False
                    'While tblCustModif.DefaultView.Count > 0
                    '    rowCustModif = tblCustModif.DefaultView(0).Row
                    '    If .customer < rowCustModif.Customer_Name Then
                    '        Exit While
                    '    ElseIf .customer = rowCustModif.Customer_Name Then
                    '        CustKetemu = True
                    '        Exit While
                    '    Else
                    '        tblCustModif.Rows.Remove(rowCustModif)
                    '    End If
                    'End While


                    If .invalid_reason.Contains("Invoice already registered in SMESS") = False Then
                        CustKetemu = False
                        tblCustModif.DefaultView.RowFilter = "[Customer Name]='" & row.customer & "'"
                        If tblCustModif.DefaultView.Count > 0 Then
                            CustKetemu = True
                            rowCustModif = tblCustModif.DefaultView(0).Row
                            .Customer_Code = rowCustModif.Customer_Code
                            .Customer_Name = rowCustModif.Name
                        End If
                        If CustKetemu = False Then
                            If .invalid_reason <> "" Then
                                .invalid_reason &= ", "
                            End If
                            .invalid_reason &= "Customer NPWP Name not registered in SMESS"
                        End If

                    Else
                        tblCustModif.DefaultView.RowFilter = "[Customer Name]='" & row.customer & "'"
                        If tblCustModif.DefaultView.Count > 0 Then
                            rowCustModif = tblCustModif.DefaultView(0).Row
                            .Customer_Code = rowCustModif.Customer_Code
                            .Customer_Name = rowCustModif.Name
                        End If
                    End If
                End With
            Next

            'tax invoice no pada file csv import tidak masuk ke range tax invoice no smess /buat dulu di smess pada tax invoice setting
            tblTax.DefaultView.RowFilter = ""
            tblTax.DefaultView.Sort = "[prefix]"
            Dim TaxKetemu As Boolean = False
            For i As Integer = 0 To vTbl.DefaultView.Count - 1
                row = vTbl.DefaultView(i).Row
                With row
                    If .invalid_reason.Contains("Invoice already registered in SMESS") = False Then
                        TaxKetemu = False
                        tblTax.DefaultView.RowFilter = "[prefix]='" & row.tax_invoice_no_2.Substring(0, 8) & "'"
                        If tblTax.DefaultView.Count > 0 Then
                            For ij As Integer = 0 To tblTax.DefaultView.Count - 1
                                rowTax = tblTax.DefaultView(ij).Row
                                Dim TaxInvoiceNo As Integer = row.tax_invoice_no_2.ToString.Substring(8, row.tax_invoice_no_2.Length - 8)
                                If TaxInvoiceNo >= rowTax.starting_number And TaxInvoiceNo <= rowTax.ending_number Then
                                    TaxKetemu = True
                                End If
                            Next
                        End If


                        If TaxKetemu = False Then
                            If .invalid_reason <> "" Then
                                .invalid_reason &= ", "
                            End If
                            .invalid_reason &= "Tax Invoice not registered in SMESS"
                        End If
                    End If
                End With
            Next

            'product pada file csv import tidak ada pada smess (cocokan berdasarkan product description)
            tblProd.DefaultView.RowFilter = ""
            tblProd.DefaultView.Sort = "[Prod Code Custom]"
            vTbl.DefaultView.RowFilter = ""
            vTbl.DefaultView.Sort = "[prod code]"
            Dim ListPrimary As New List(Of Double)
            Dim ProdKetemu As Boolean = False
            For i As Integer = 0 To vTbl.DefaultView.Count - 1
                row = vTbl.DefaultView(i).Row
                With row
                    'If ListPrimary.Contains(.primary) = False Then
                    '    ListPrimary.Add(.primary)

                    'If row.invoice_no.ToString.Contains("FJ.0520.00005") = True Then
                    '    MsgBox("")
                    'End If
                    tblProd.DefaultView.RowFilter = "[Prod Code Custom]='" & row.prod_code & "'"
                    If tblProd.DefaultView.Count > 0 Then
                        rowProd = tblProd.DefaultView(0).Row
                        ProdKetemu = True
                        .primary_code = rowProd.Primary
                        .prod_primary_unit = rowProd.primary_unit
                        .prod_unit = rowProd.unit_code
                        .prod_dec = rowProd.Prod_Decimal
                        .prod_code_smess = rowProd.Prod_Code_Smess
                        .prod_desc_smess = rowProd.Prod_Desc_Smess
                        .primary_code_smess = rowProd.Primary_Product_Smess
                        .qty_on_hand = rowProd.Qty
                    Else
                        ProdKetemu = False
                    End If


                    'ProdKetemu = False
                    'Dim Idx As Integer = 0
                    'While tblProd.DefaultView.Count > 0
                    '    rowProd = tblProd.DefaultView(Idx).Row
                    '    If .prod_code < rowProd.Prod_Code_Custom Then
                    '        'Exit While
                    '        Idx += 1
                    '    ElseIf .prod_code = rowProd.Prod_Code_Custom Then
                    '        If String.IsNullOrEmpty(rowProd.Prod_Code_Smess) = False Then
                    '            ProdKetemu = True
                    '            .primary_code = rowProd.Primary
                    '            .prod_primary_unit = rowProd.primary_unit
                    '            .prod_unit = rowProd.unit_code
                    '            .prod_dec = rowProd.Prod_Decimal
                    '            .prod_code_smess = rowProd.Prod_Code_Smess
                    '            .prod_desc_smess = rowProd.Prod_Desc_Smess
                    '            .primary_code_smess = rowProd.Primary_Product_Smess
                    '            .qty_on_hand = rowProd.Qty
                    '            Exit While
                    '        Else
                    '            Exit While
                    '        End If
                    '    Else
                    '        If tblProd.Rows.Count > 0 Then
                    '            tblProd.Rows.Remove(rowProd)
                    '        End If
                    '    End If
                    'End While

                    If .invalid_reason.Contains("Invoice already registered in SMESS") = False Then
                        If ProdKetemu = False Then
                            If .invalid_reason <> "" Then
                                .invalid_reason &= ", "
                            End If
                            .invalid_reason &= "Product not registered in SMESS"
                            .prod_invalid_reason = "Product not registered in SMESS"
                        End If
                    End If
                End With
            Next

            'Comment: Perubahan NPWP 16 Digit
            'customer pada file csv import tgl transaksi > tgl berlaku npwp 16 digit
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            vTbl.DefaultView.RowFilter = ""
            vTbl.DefaultView.Sort = "[customer]"
            Dim tglBerlaku As Date = ObjSetNPWP.Tglberlaku
            For i As Integer = 0 To vTbl.DefaultView.Count - 1
                row = vTbl.DefaultView(i).Row
                With row
                    If .invoice_date < tglBerlaku Then
                        If ObjSetNPWP.ApakahSudah15Digit(.npwp, False) = False Then
                            If .invalid_reason <> "" Then
                                .invalid_reason &= ", "
                            End If
                            .invalid_reason &= "Used 15 digit for NPWP No."
                        End If
                    Else
                        If ObjSetNPWP.ApakahSudah16Digit(.npwp, False) = False Then
                            If .invalid_reason <> "" Then
                                .invalid_reason &= ", "
                            End If
                            .invalid_reason &= "Used 16 digit for NPWP No."
                        End If
                    End If
                End With
            Next

            tblqty.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowToDgvItem(ByVal vPrimary As Double)
        Dim rd As DataGridViewRow

        Try

            Me.dgvItem.Rows.Clear()
            Tbl.DefaultView.RowFilter = "[primary]=" & vPrimary
            Tbl.DefaultView.Sort = "[prod code]"
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For i As Integer = 0 To Tbl.DefaultView.Count - 1
                row = Tbl.DefaultView(i).Row
                With row
                    rd = Me.dgvItem.Rows(Me.dgvItem.Rows.Add)
                    rd.Cells(Me.colItem_No.Index).Value = rd.Index + 1
                    rd.Cells(Me.colItem_ProductCode.Index).Value = .prod_code
                    rd.Cells(Me.colItem_ProductDesc.Index).Value = .prod_desc
                    rd.Cells(Me.colItem_Qty.Index).Value = Me.ObjCD.DigitGrouping(.prod_qty, .prod_dec)
                    rd.Cells(Me.colItem_Unit.Index).Value = .prod_unit
                    rd.Cells(Me.colItem_UnitPrice.Index).Value = Me.ObjCD.DigitGrouping(.prod_unit_price, DefPrec)
                    rd.Cells(Me.colItem_DiscAmount.Index).Value = Me.ObjCD.DigitGrouping(.prod_disc_amount, DefPrec)
                    rd.Cells(Me.colItem_Amount.Index).Value = Me.ObjCD.DigitGrouping(.prod_amount, DefPrec)
                    rd.Cells(Me.colItem_DPP.Index).Value = Me.ObjCD.DigitGrouping(.prod_dpp, DefPrec)
                    rd.Cells(Me.colItem_PPN.Index).Value = Me.ObjCD.DigitGrouping(.prod_ppn, DefPrec)
                    rd.Cells(Me.colItem_InvalidReason.Index).Value = .prod_invalid_reason
                    rd.Cells(Me.ColQtyOnHand.Index).Value = Me.ObjCD.DigitGrouping(.qty_on_hand, .prod_dec)
                End With
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowInvalidOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkShowInvalidOnly.CheckedChanged
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll.LinkClicked
        Try
            Me.SetCreateColumn(True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll.LinkClicked
        Try
            Me.SetCreateColumn(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCreateColumn(ByVal vSetColumn As Boolean)
        Try
            Me.dgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For i As Integer = 0 To Me.dgvHeader.RowCount - 1
                If String.IsNullOrEmpty(Me.dgvHeader.Item(Me.colInvalidReason.Index, i).Value) = True Then
                    Me.dgvHeader(Me.colCreate.Index, i).Value = vSetColumn
                End If
            Next
            Me.dgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim Buffer As String = ""
    Private Sub dgvHeader_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvHeader.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colCreate.Index Or e.ColumnIndex = Me.colInvoiceNumber.Index Then
                If e.ColumnIndex = Me.colCreate.Index Then
                    If String.IsNullOrEmpty(Me.dgvHeader.Item(Me.colInvalidReason.Index, e.RowIndex).Value) = False Then
                        e.Cancel = True
                    End If
                End If

                ' batasan invoice no
                If e.ColumnIndex = Me.colInvoiceNumber.Index Then
                    If Me.dgvHeader.Item(Me.colCreate.Index, e.RowIndex).Value = Boolean.TrueString Then
                        If String.IsNullOrEmpty(Me.dgvHeader.Item(Me.colInvoiceNumber.Index, e.RowIndex).Value) = True And String.IsNullOrEmpty(Me.dgvHeader.Item(Me.colInvalidReason.Index, e.RowIndex).Value) = False Then
                            Me.Buffer = Me.dgvHeader.Item(Me.colInvoiceNumber.Index, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    Else
                        e.Cancel = True
                    End If
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim Primary_Buffer As Double = -1

    Private Sub dgvHeader_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHeader.CellClick
        Try
            'If Primary_Buffer = -1 Or Me.dgvHeader(Me.colPrimary.Index, Me.dgvHeader.SelectedCells(0).RowIndex).Value <> Primary_Buffer Then
            '    Primary_Buffer = Me.dgvHeader(Me.colPrimary.Index, Me.dgvHeader.SelectedCells(0).RowIndex).Value
            '    Me.ShowToDgvItem(Me.dgvHeader(Me.colPrimary.Index, Me.dgvHeader.SelectedCells(0).RowIndex).Value)
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHeader_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHeader.CellEndEdit
        Dim row As New DataGridViewRow
        Try
            If e.ColumnIndex = Me.colInvoiceNumber.Index Then
                row = Me.dgvHeader.Rows(e.RowIndex)
                If IsNothing(row.Cells(Me.colInvoiceNumber.Index).Value) = False Then
                    With row
                        If String.IsNullOrEmpty(.Cells(Me.colInvoiceNumber.Index).Value) = False Then
                            'cek apakah sudah ada diinvoice
                            If Me.ObjInv.ApakahSudahAda(.Cells(Me.colInvoiceNumber.Index).Value) = True Then
                                MsgBox("Invoice already registered in SMESS", MsgBoxStyle.Information, "Information")
                                .Cells(Me.colInvoiceNumber.Index).Value = Me.Buffer
                                Exit Sub
                            Else
                                .Cells(Me.colInvalidReason.Index).Value = .Cells(Me.colInvalidReason.Index).Value.ToString.Replace("Invoice number is empty, ", "").Replace("Invoice number is empty", "")
                                .Cells(Me.colInvoiceNumber.Index).Value = .Cells(Me.colInvoiceNumber.Index).Value.ToString.ToUpper
                                If String.IsNullOrEmpty(.Cells(Me.colInvalidReason.Index).Value) = True Then
                                    If e.RowIndex Mod 2 = 0 Then
                                        Me.dgvHeader.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
                                    Else
                                        Me.dgvHeader.Rows(e.RowIndex).DefaultCellStyle.BackColor = Me.BackColorAll
                                    End If
                                End If
                            End If
                        End If
                    End With
                Else
                    row.Cells(Me.colInvoiceNumber.Index).Value = Me.Buffer
                End If


            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHeader_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvHeader.DataError

    End Sub

    Private Sub dgvHeader_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvHeader.SelectionChanged
        Try
            If Me.dgvHeader.Rows.Count > 0 Then
                If String.IsNullOrEmpty(Me.dgvHeader.Item(Me.colPrimary.Index, Me.dgvHeader.CurrentRow.Index).Value) = False Then
                    Me.ShowToDgvItem(Me.dgvHeader.Item(Me.colPrimary.Index, Me.dgvHeader.CurrentRow.Index).Value)
                End If
            End If

            'If Primary_Buffer = -1 Or Me.dgvHeader(Me.colPrimary.Index, Me.dgvHeader.SelectedCells(0).RowIndex).Value <> Primary_Buffer Then
            '    Primary_Buffer = Me.dgvHeader(Me.colPrimary.Index, Me.dgvHeader.SelectedCells(0).RowIndex).Value
            '    Me.ShowToDgvItem(Primary_Buffer)
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCreateInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateInvoice.Click
        Try
            If Me.GetJmlImport() = 0 Then
                MsgBox("Please chooses min. 1 row of the item data from that " & Chr(34) & "create" & Chr(34) & " coloumn.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.ObjMonitor.ApakahAdaMonitor2(Core_Module.ClsTransaksiMonitor.EnumTipeTransaksi.Custom_ImportEfaktur, Me.UserCode, Me.Identifier) = True Then
                MsgBox("Other user is on progress creating import e-faktur, " + _
                "you cannot create new transaction or saved data on this form.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.BackgroundWorker1.RunWorkerAsync()
            'Me.Import_eFaktur()

           
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.Import_eFaktur()
            Me.ProgressBar1.Visible = False
            Me.ProgressBar1.Maximum = 0
            Me.llbProgressBar.Text = ""
            MsgBox("Finish", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetConvertionDateFromString(ByVal vString As String) As Date
        Dim i As String = ""
        Dim Day As Integer = 0
        Dim Month As Integer = 0
        Dim Year As Integer = 0
        Dim count As Integer = 0
        Dim strArr() As String
        Dim d As New Date
        Try
            strArr = vString.Split("/")
            For count = 0 To strArr.Length - 1
                If count = 0 Then
                    Day = strArr(count)
                ElseIf count = 1 Then
                    Month = strArr(count)
                Else
                    Year = strArr(count)
                End If
            Next
            d = "#" & Month & "/" & Day & "/" & Year & "#"
            i = Format(d, "dd MMM yyyy")
            Return CDate(i)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick

    End Sub

    Private Function GetJmlImport() As Integer
        Dim Count As Integer = 0
        Try
            For Each rowD As DataGridViewRow In Me.dgvHeader.Rows
                If rowD.Cells(Me.colCreate.Index).Value = Boolean.TrueString Then
                    Count += 1
                End If
            Next
            Return Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SetFormatNPWP(ByVal NPWPLama As String) As String
        Dim NewNPWP As String = ""
        Try
            If NPWPLama.Contains("-") = False Or NPWPLama.Contains(".") = False Then
                If String.IsNullOrEmpty(NPWPLama) = False Then
                    For i As Integer = 0 To 14
                        NewNPWP &= NPWPLama.Chars(i)
                        If i = 1 Or i = 4 Or i = 7 Or i = 11 Then
                            NewNPWP &= "."
                        ElseIf i = 8 Then
                            NewNPWP &= "-"
                        End If
                    Next
                    Return NewNPWP
                End If
            End If
            Return NPWPLama
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SetFormatTaxNo(ByVal TaxNoLama As String) As String
        Dim NewTax As String = "."
        Try
            If TaxNoLama.Contains("-") = False Or TaxNoLama.Contains(".") = False Then
                If String.IsNullOrEmpty(TaxNoLama) = False Then
                    For i As Integer = 0 To TaxNoLama.Length - 1
                        NewTax &= TaxNoLama.Chars(i)
                        If i = 4 Then
                            NewTax &= "."
                        ElseIf i = 2 Then
                            NewTax &= "-"
                        End If
                    Next
                    Return NewTax
                End If
            End If
            Return TaxNoLama
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Import_eFaktur()
        Dim ListPrimary As New List(Of Double)
        Dim count As Integer = 0
        Dim rowN As DseImport_eFaktur.tbl_temporaryRow
        Dim ObjCust As New Core_Module.clsMasterCustomer
        Dim ObjSalesPerson As New Core_Module.ClsMasterSalesPerson
        Dim RowDGV As DataGridViewRow
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjUnit As New Core_Module.ClsMasterUnit
        Dim AngkaDecimalCurrency As Decimal = DefPrec
        Dim ObjSO As New Sales_Module.ClsTransaksiSalesOrder
        Dim ListPrimaryCode As New List(Of Double)
        Dim ListPrimaryInv As New List(Of Double)
        Dim ListInvGagalSimpan As New List(Of String)
        Dim UseTaxSetting, TaxPercentSetting As String
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Try
            UseTaxSetting = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If UseTaxSetting = "0" Or UseTaxSetting = "" Then
                TaxPercentSetting = Me.ObjCD.DigitGrouping("0", 0)
            Else
                TaxPercentSetting = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
            End If

            Me.isQtyMinus = False
            'saat create invoice pakai backgroundworker & progressbar 
            If MsgBox("Start process create sales invoice transactions?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.ProgressBar1.Visible = True
                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = Me.GetJmlImport()
                For Each rowD As DataGridViewRow In Me.dgvHeader.Rows
                    If rowD.Cells(Me.colCreate.Index).Value = Boolean.TrueString Then
                        If ListPrimary.Contains(rowD.Cells(Me.colPrimary.Index).Value) = False Then
                            ListPrimary.Add(rowD.Cells(Me.colPrimary.Index).Value)
                            If Me.ObjInv.ApakahSudahAda(rowD.Cells(Me.colInvoiceNumber.Index).Value) = True Then
                                If ListInvGagalSimpan.Contains(rowD.Cells(Me.colInvoiceNumber.Index).Value) = False Then
                                    ListInvGagalSimpan.Add(rowD.Cells(Me.colInvoiceNumber.Index).Value)
                                End If
                                Me.ProgressBar1.Value += 1
                                GoTo line_Exit
                            Else
                                Me.ProgressBar1.Value += 1
                            End If

                            Me.llbProgressBar.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - Invoice No: " + rowD.Cells(Me.colInvoiceNumber.Index).Value
                            Dim ij As New FrmTransaksiSalesInvoiceQuick
                            ij.dgvItem.Rows.Clear()
                            ij.UserCode = Me.UserCode
                            ij.Show()
                            ij.vFromQuick = True
                            ij.SetBaru()
                            ij.Location = New Point((ij.Width * -1) - 5, (ij.Height * -1) - 5)
                            'isikan Header
                            With rowD
                                'yang ada
                                ij.txtInvoiceNo.Text = .Cells(Me.colInvoiceNumber.Index).Value
                                ij.dtpDate.Value = CDate(.Cells(Me.colDate.Index).Value)
                                ij.txtCustomerCode.Text = Me.GetCustomerCode(.Cells(Me.colCustomer.Index).Value)
                                ij.GetCustomer_Leave()
                                If UseTaxSetting = "0" Or UseTaxSetting = "" Then
                                    ij.chkTax.Checked = False
                                Else
                                    ij.chkTax.Checked = True
                                End If

                                If IsNumeric(TaxPercentSetting) = True Then
                                    TaxPercentSetting = Me.ObjCD.DigitGrouping(TaxPercentSetting, 0)
                                Else
                                    TaxPercentSetting = "0"
                                End If

                                ij.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(1)
                                ij.CurrencyCode_Leave()
                                ij.SetCalculateBy()
                                ij.txtNoNPWP.Text = .Cells(Me.colNPWP.Index).Value
                                ij.txtNoFakturPajakSederhana.Text = ""
                                ij.txtNoSeriFakturPajak.Text = .Cells(Me.colTaxInvoiceNumber.Index).Value.ToString.Substring(0, 3)
                                ij.txtNoFakturPajak.Text = .Cells(Me.colTaxInvoiceNumber.Index).Value.ToString.Substring(3, .Cells(Me.colTaxInvoiceNumber.Index).Value.ToString.Length - 3)
                                ij.txtTax.Text = TaxPercentSetting
                                If TaxPercentSetting = "0" Then
                                    ij.txtTaxAmount.Text = "0"
                                Else
                                    ij.txtTaxAmount.Text = .Cells(Me.colPPN.Index).Value
                                End If



                                ij.txtSubTotalAfterDP.Text = .Cells(Me.colDPP.Index).Value
                                ij.txtSubTotalBeforeDP.Text = .Cells(Me.colDPP.Index).Value
                                ij.TxtTotal.Text = CDec(.Cells(Me.colDPP.Index).Value) + CDec(.Cells(Me.colPPN.Index).Value)

                                'yang tidak ada
                                'ij.cboCalculatesBy.SelectedIndex = 0
                                'ij.txtDownPayment.Text = 0
                                'ij.txtDutyStamp.Text = 0
                                'ij.txtExhangeRateToIDR.Text = 0
                                'ij.txtExchangeRateToUSD.Text = 0
                                'ij.txtRevisiNo.Text = ""
                                'ij.TxtShipment.Text = ""
                                'ij.txtTaxExchangeRate.Text = ""
                                'ij.txtCustomerAddIdentifier.Text = ""
                                'ij.txtPaymentTerm.Text = ""
                                'ij.txtPrimaryCurrency.Text = 1


                                'ij.chkTax.Checked = 0
                                'ij.TxtRemarks.Text = ""
                                'ij.CboFakturPajak.SelectedIndex = 0
                                ''ij.txtNoFakturPajakSederhana.Text = ""
                                ''ij.TxtNamaSigned.Text = ""
                                ''ij.cboShipmentTax.SelectedIndex = 0
                                ''ij.TxtJabatanSigned.Text = ""

                                'ij.cboPPNDipungut.SelectedIndex = 0
                                'ij.TxtDiscString.Text = 0
                                'ij.TxtDiscAmount.Text = 0

                                'ij.Txt_E_Faktur_Blok.Text = ""
                                'ij.Txt_E_Faktur_Jalan.Text = ""
                                'ij.Txt_E_Faktur_Kelurahan.Text = ""
                                'ij.Txt_E_Faktur_KodePos.Text = ""
                                'ij.Txt_E_Faktur_No.Text = ""
                                'ij.Txt_E_Faktur_NoTelepon.Text = ""
                                'ij.Txt_E_Faktur_Propinsi.Text = ""
                                'ij.Txt_E_Faktur_RT.Text = ""
                                'ij.Txt_E_Faktur_RW.Text = ""
                                'ij.Txt_E_Faktur_Kabupaten.Text = ""
                                'ij.Txt_E_Faktur_Kecamatan.Text = ""
                                'ij.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = ""
                                'ij.txtPrimaryCurrTax.Text = ""
                                'ij.txtPrimaryAccount.Text = ""
                                'ij.txtOtherCharge.Text = ""


                            End With

                            'isikan item
                            Dim FrmCustom As New CustomSMESS.FrmTransaksiSalesInvoiceQuick_AddItembyCustom
                            Dim MaxQty As Decimal = 0
                            Dim ApakahDataBenar As Boolean = True
                            Me.Tbl.DefaultView.RowFilter = "[primary]=" & rowD.Cells(Me.colPrimary.Index).Value
                            Me.Tbl.DefaultView.Sort = "[primary] ASC"
                            For i As Integer = 0 To Me.Tbl.DefaultView.Count - 1
                                Me.row = Me.Tbl.DefaultView(i).Row
                                With Me.row
                                    If FrmCustom.ApakahStockBenar(FrmCustom.dgvItem, Me.row.primary_code, Me.row.primary_code_smess, Me.row.prod_dec, _
                                    Me.row.prod_qty, Me.row.prod_primary_unit, -1, "", Me.row.qty_on_hand) = True Then
                                        RowDGV = FrmCustom.dgvItem.Rows(FrmCustom.dgvItem.Rows.Add)
                                        With RowDGV
                                            'ada
                                            .Cells(FrmCustom.ColNo.Index).Value = Me.dgvItem.Rows.Count
                                            .Cells(FrmCustom.ColPrimaryCodeNo.Index).Value = Me.row.primary_code_smess
                                            .Cells(FrmCustom.ColCodeNo.Index).Value = Me.row.prod_code_smess
                                            .Cells(FrmCustom.ColDescription.Index).Value = Me.row.prod_desc_smess
                                            .Cells(FrmCustom.ColPrimaryCustomProduct.Index).Value = Me.row.primary_code
                                            .Cells(FrmCustom.ColCustomCodeNo.Index).Value = Me.row.prod_code
                                            .Cells(FrmCustom.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.row.prod_qty), Me.row.prod_dec)
                                            .Cells(FrmCustom.ColUnit.Index).Value = Me.row.prod_unit
                                            .Cells(FrmCustom.ColQty.Index).Value = .Cells(FrmCustom.ColTotalQty.Index).Value
                                            .Cells(FrmCustom.ColUnitPrice.Index).Value = Me.row.prod_unit_price
                                            .Cells(FrmCustom.ColPrimaryUnit.Index).Value = ObjUnit.GetPrimary(Me.row.prod_unit)
                                            .Cells(FrmCustom.ColTotalQtyMulti.Index).Value = .Cells(FrmCustom.ColTotalQty.Index).Value ' cek lg aapakh disamakana dgn default
                                            .Cells(FrmCustom.ColUnitsMulti.Index).Value = Me.row.prod_unit
                                            .Cells(FrmCustom.ColPrimaryUnitsMulti.Index).Value = .Cells(FrmCustom.ColPrimaryUnit.Index).Value
                                            .Cells(FrmCustom.ColRatioDefault.Index).Value = Me.ObjCD.DigitGrouping(ObjSO.GetRatioDefault(.Cells(FrmCustom.ColPrimaryCodeNo.Index).Value, _
                                            .Cells(FrmCustom.ColPrimaryUnitsMulti.Index).Value), Me.row.prod_dec)
                                            .Cells(FrmCustom.ColRatioReference.Index).Value = Me.ObjCD.DigitGrouping(ObjSO.GetRatioReference(.Cells(FrmCustom.ColPrimaryCodeNo.Index).Value, _
                                            .Cells(FrmCustom.ColPrimaryUnitsMulti.Index).Value), Me.row.prod_dec)

                                            .Cells(FrmCustom.ColPrimaryCurrency.Index).Value = ij.txtPrimaryCurrency.Text
                                            .Cells(FrmCustom.ColCurrency.Index).Value = ij.txtCurrencyCode.Text
                                            'Disc amount = dpp / qty(qty multi)
                                            'disc amount = harga satuan - disc amount
                                            .Cells(FrmCustom.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.row.prod_disc_amount / CDec(.Cells(FrmCustom.ColTotalQtyMulti.Index).Value), AngkaDecimalCurrency)
                                            .Cells(FrmCustom.ColDiscount.Index).Value = CDec(.Cells(FrmCustom.ColDiscAmount.Index).Value) * 100 / Me.row.prod_unit_price
                                            .Cells(FrmCustom.ColNettPrice.Index).Value = Math.Round(CDec(Me.row.prod_unit_price) - CDec(.Cells(FrmCustom.ColDiscAmount.Index).Value), Me.row.prod_dec, MidpointRounding.AwayFromZero)
                                            .Cells(FrmCustom.ColAmount.Index).Value = Me.row.prod_amount - Me.row.prod_disc_amount
                                            .Cells(FrmCustom.ColNettPrice.Index).Value = CDec(.Cells(FrmCustom.ColAmount.Index).Value) / CDec(.Cells(FrmCustom.ColTotalQtyMulti.Index).Value)

                                            'tidak ada
                                            .Cells(FrmCustom.ColCustCodeNo.Index).Value = ""
                                            .Cells(FrmCustom.ColCustDesc.Index).Value = ""
                                            .Cells(FrmCustom.ColSONO.Index).Value = ""
                                            .Cells(FrmCustom.ColCustPONo.Index).Value = ""
                                            .Cells(FrmCustom.ColDONO.Index).Value = ""
                                            .Cells(FrmCustom.ColStockOutNo.Index).Value = ""
                                            .Cells(FrmCustom.ColLotNumber.Index).Value = ""
                                            'If FrmCustom.DtpExpired.Visible = False Then
                                            '    .Cells(FrmCustom.ColExpiry.Index).Value = #1/1/1950#
                                            'Else
                                            '    .Cells(FrmCustom.ColExpiry.Index).Value = FrmCustom.DtpExpired.Value
                                            'End If
                                            .Cells(FrmCustom.ColLocationCode.Index).Value = ""
                                            .Cells(FrmCustom.ColQtyPerMaster.Index).Value = 0
                                            .Cells(FrmCustom.ColQtyMaster.Index).Value = 0
                                            .Cells(FrmCustom.ColMasterUnits.Index).Value = ""
                                            .Cells(FrmCustom.ColSecQty.Index).Value = -1
                                            .Cells(FrmCustom.ColSecUnit.Index).Value = ""

                                            .Cells(FrmCustom.ColRemarksItem.Index).Value = ""
                                            .Cells(FrmCustom.ColPrimaryLocation.Index).Value = -1
                                            .Cells(FrmCustom.ColPrimaryDO.Index).Value = -1
                                            .Cells(FrmCustom.ColPrimarySO.Index).Value = -1
                                            .Cells(FrmCustom.ColPrimaryMasterUnits.Index).Value = -1

                                            .Cells(FrmCustom.ColPrimarySecUnit.Index).Value = -1

                                            .Cells(FrmCustom.ColAliasDesc.Index).Value = Me.ObjInvoiceItem.GetAliasDesc(.Cells(FrmCustom.ColPrimaryCodeNo.Index).Value, ij.txtPrimaryCustomer.Text)
                                            .Cells(FrmCustom.ColAliasCode.Index).Value = Me.ObjInvoiceItem.GetAliasCode(.Cells(FrmCustom.ColPrimaryCodeNo.Index).Value, ij.txtPrimaryCustomer.Text)
                                            .Cells(FrmCustom.ColDiscFromPrect.Index).Value = 0
                                            .Cells(FrmCustom.ColVAT.Index).Value = 0
                                            'If FrmCustom.txtPriceByMultiUnit.Text = "1" Then
                                            '    .Cells(FrmCustom.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(FrmCustom.ColVAT.Index).Value * CDec(.Cells(FrmCustom.ColTotalQtyMulti.Index).Value), AngkaDecimalCurrency)
                                            '    .Cells(FrmCustom.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(FrmCustom.ColUnitsMulti.Index).Value
                                            'Else
                                            '    .Cells(FrmCustom.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(FrmCustom.ColVAT.Index).Value * CDec(.Cells(FrmCustom.ColTotalQty.Index).Value), AngkaDecimalCurrency)
                                            '    .Cells(FrmCustom.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(FrmCustom.ColUnit.Index).Value
                                            'End If
                                            .Cells(FrmCustom.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(FrmCustom.ColVAT.Index).Value * CDec(.Cells(FrmCustom.ColTotalQty.Index).Value), AngkaDecimalCurrency)
                                            .Cells(FrmCustom.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(FrmCustom.ColVAT2.Index).Value) + CDec(.Cells(FrmCustom.ColAmount.Index).Value), AngkaDecimalCurrency)
                                            FrmCustom.SetRatioOnDgv(RowDGV)
                                        End With
                                    Else
                                        row.is_qty_minus = True
                                        ApakahDataBenar = False
                                        isQtyMinus = True
                                        If ListPrimaryCode.Contains(row.primary_code_smess) = False Then
                                            ListPrimaryCode.Add(row.primary_code_smess)
                                        End If
                                        If ListPrimaryInv.Contains(row.primary) = False Then
                                            ListPrimaryInv.Add(row.primary)
                                        End If
                                    End If


                                End With

                            Next


                            Dim ObjC As New CustomSMESS.ClsTransaksiSalesInvoiceQuick_AddByCustomProduct
                            Dim tblC As New CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1DataTable
                            ObjC.IsiTblDummyFromDgv(FrmCustom.dgvItem, FrmCustom, rowD.Cells(Me.colPrimary.Index).Value, tblC)
                            ij.TblCustomDummy = tblC.Copy
                            ij.IsikanDgvFromTblDummy()
                            ij.vApakahDatabenar = ApakahDataBenar
                            ij.Tax_Leave()
                            ij.Hide()
                            ij.ShowDialog()
line_Exit:
                        End If
                    End If
                Next
            End If
            If isQtyMinus = True Then
                Dim ij As New FrmQtyOnHand
                ij.Tbl = Me.Tbl.Copy
                ij.PrimaryInvListMinus = ListPrimaryInv
                ij.PrimaryCodeListMinus = ListPrimaryCode
                ij.BackColorAll = Me.BackColorAll
                ij.ShowDialog()
            End If

            If ListInvGagalSimpan.Count > 0 Then
                Dim Msg As String = ""
                For Each i As String In ListInvGagalSimpan
                    If Msg = "" Then
                        Msg = i
                    Else
                        Msg &= ", " & Msg
                    End If
                Next
                If Msg <> "" Then
                    MsgBox("Duplication no invoice List:" & vbNewLine & Msg, MsgBoxStyle.Information, "Information")
                End If
            End If
            Me.btnShow.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHeader_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvHeader.Sorted
        Me.ObjCD.SetNumber(Me.dgvHeader)
    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError

    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem)
    End Sub

    Private Function GetCustomerCode(ByVal vCustomerName As String) As String
        Try
            tblCust.DefaultView.RowFilter = ""
            tblCust.DefaultView.RowFilter = "[Customer Name]='" & vCustomerName & "'"
            tblCust.DefaultView.Sort = "[Customer Name]"
            If tblCust.DefaultView.Count > 0 Then
                rowCust = tblCust.DefaultView(0).Row
                Return rowCust.Customer_Code
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class