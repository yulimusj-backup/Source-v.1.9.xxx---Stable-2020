Option Explicit On
Option Strict Off
Public Class FrmImport_CheckData
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim TipeSR As FrmImport_CheckData.EnumTipeSR = EnumTipeSR.Nota_Retur_Pajak
    Dim TipePR As FrmImport_CheckData.EnumTipe_PurchaseReturn = EnumTipe_PurchaseReturn.Nota_Retur
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim DefCurr As Decimal = ObjCurr.GetDecimal_DefaultCurrency
    Dim Tbl As DataTable = New DataTable
    Dim row As DataRow
    Dim CreateTable As Boolean = False
    Dim FilePath As String
    Dim IndexLama As Integer = 0
    Dim ListFakturPajakPengganti As New List(Of String)

    Public Enum EnumTipe_PurchaseReturn As Integer
        Revisi_AOP
        Nota_Retur
    End Enum
    Public Enum EnumTipeSR As Integer
        Ganti_Barang
        Revisi_Invoice
        Nota_Retur_Pajak
    End Enum

    Private Sub FrmImport_eFaktur_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'selain data hasil import csv perlu ambil data smess nya juga berdasarkan tanggal from to
        'FP = pajak masukan (AOP)
        'FK = pajak keluaran (Sales Ivoice)
        'FP Return = purchase return
        'FK Return = sales return

        'jadi kasus nya (contoh kasuk sales invoice), 
        'invoice smess ada & invoice import tidak ada (selisih invoice = true)
        'invoice smess tidak ada & invoice import ada (selisih invoice = true)
        'invoice smess ada & invoice import ada (selisih invoice = false)


        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.cboTipe.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboTipe_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTipe.Leave
        Me.IndexLama = Me.cboTipe.SelectedIndex
        Me.CreateTable = False
    End Sub

    Private Sub cboTipe_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTipe.SelectedIndexChanged
        Try
            Me.colNoAcuan.HeaderText = Me.cboTipe.Text
            If Me.IndexLama <> Me.cboTipe.SelectedIndex Then
                Me.FilePath = ""
                Me.txtFilePath.Text = ""

                Me.dgvItem.Rows.Clear()
                Me.Tbl.Rows.Clear()
                Me.TblRev.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        'dikosongin file path ketika ubah cbo karna setiap import yang dipilih sesuai yang dibutuhkan saja contohnya fk saja efeknya biar datatable tidak banyak dan tidak lemot/lama
    End Sub

    Private Sub btnFindFilePath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindFilePath.Click

        OpenFileDialog1.Filter = "Csv files (*.Csv)|*.Csv"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Title = "Open file csv"
        OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.FileName <> "" Then
            FilePath = OpenFileDialog1.FileName
            Try
                Me.txtFilePath.Text = Me.FilePath
                Me.ImportFile()
            Catch ex As Exception
                'display error messages if they appear         
                MessageBox.Show(ex.Message)
            End Try
        End If
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

    Public Function SetFormatNPWP(ByVal NPWPLama As String) As String
        Dim NewNPWP As String = ""
        Try
            If NPWPLama.Contains("-") = False Or NPWPLama.Contains(".") = False Then
                For i As Integer = 0 To NPWPLama.Length - 1
                    NewNPWP &= NPWPLama.Chars(i)
                    If i = 1 Or i = 4 Or i = 7 Or i = 11 Then
                        NewNPWP &= "."
                    ElseIf i = 8 Then
                        NewNPWP &= "-"
                    End If
                Next
                Return NewNPWP
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

    Private Sub ImportFile()
        Dim SR As System.IO.StreamReader = New System.IO.StreamReader(Me.FilePath)
        Dim line As String = SR.ReadLine()
        line = line.ToString.Replace(";", ",")
        Dim iCOunt As String() = line.Split(","c)
        Dim Icount2 As String() = line.Split(";"c)
        Dim strArray As String()
        Dim ij As Integer = -1
        Dim TipeHeader As String = ""
        Dim TotalKolom As Integer = 19

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.cboTipe.SelectedIndex = 0 Then
                'purchase
                TipeHeader = "FM"
                TotalKolom = 14
            ElseIf Me.cboTipe.SelectedIndex = 1 Then
                'sales
                TipeHeader = "FK"
                TotalKolom = 19
            ElseIf Me.cboTipe.SelectedIndex = 2 Then
                'purchase return
                TipeHeader = "RM"
                TotalKolom = 15
            ElseIf Me.cboTipe.SelectedIndex = 3 Then
                'sales return
                TipeHeader = "RK"
                TotalKolom = 14
            End If
            '-------
            Me.Tbl.Rows.Clear()
            Me.ObjCD.UpdateThreading()

            If iCOunt.Length > Icount2.Length Then
                strArray = line.Split(","c)
            Else
                strArray = line.Split(";"c)
            End If
            'pindahkan data csv ke dalam data table

            If CreateTable = False Then
                Me.Tbl = Nothing
                Me.Tbl = New DataTable
                For Each s As String In strArray
                    Tbl.Columns.Add(s.Replace(Chr(34), ""), GetType(String), "")
                Next
                Tbl.Columns.Add(New DataColumn())
                CreateTable = True
            End If

            Me.ObjCD.UpdateThreading()
            Dim strArr() As String
            Dim str As String = ""
            Dim EmptyCount As Integer = 0
            Do
                line = SR.ReadLine

                If Not line = String.Empty Then
                    line = line.Replace(";", ",")
                    'If line.Contains("00489245") Then
                    '    MsgBox("")
                    'End If

                    Dim StrPerLine As String = ""
                    StrPerLine = line.Replace(ControlChars.Quote, "").Replace(Chr(34), "").Replace(";", ",")
                    strArr = StrPerLine.Split(","c)

                    'If line.Contains(Chr(34) & "," & Chr(34)) = False Then
                    '    StrPerLine = line.Replace(ControlChars.Quote, "").Replace(Chr(34), "")
                    '    strArr = StrPerLine.Split(","c)
                    'Else
                    '    StrPerLine = line
                    '    strArr = line.Split(New String() {Chr(34) & "," & Chr(34)}, StringSplitOptions.None)
                    'End If


                    'modifikasi jika ada entter pada kolom alamat
                    If strArr(0).Replace(ControlChars.Quote, "").Replace(Chr(34), "") = TipeHeader Then
                        While strArr.Length < TotalKolom 'And strArr(IdxCheck).Contains(ControlChars.Quote)
                            line = line + SR.ReadLine
                            StrPerLine = line.Replace(ControlChars.Quote, "").Replace(Chr(34), "")
                            If iCOunt.Length > Icount2.Length Then
                                strArr = StrPerLine.Split(","c)
                            Else
                                strArr = StrPerLine.Split(";"c)
                            End If

                        End While
                    End If


                    If strArr(0).Replace(ControlChars.Quote, "").Replace(Chr(34), "") = TipeHeader Then
                        If iCOunt.Length > Icount2.Length Then
                            strArr = StrPerLine.Split(",")
                        Else
                            strArr = StrPerLine.Split(";")
                        End If

                        If strArr.Length > iCOunt.Length Then
                            While strArr.Length > iCOunt.Length
                                strArr = Me.UpdateLine(strArr, 10)
                            End While
                        End If

                        If strArr.Length = iCOunt.Length Then
                            row = Tbl.NewRow
                            row.ItemArray = strArr
                            row.Item(0) = row.Item(0).Replace(ControlChars.Quote, "").ToString.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper
                            Tbl.Rows.Add(row)

                            'If iCOunt.Length > Icount2.Length Then
                            '    row = Tbl.NewRow
                            '    row.ItemArray = strArr 'line.Replace(ControlChars.Quote, "").Split(","c)
                            '    row.Item(2) = row.Item(0).Replace(ControlChars.Quote, "").ToString.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper
                            '    Tbl.Rows.Add(row)
                            'Else
                            '    row = Tbl.NewRow
                            '    row.ItemArray = strArr 'line.Split(";"c)
                            '    row.Item(2) = row.Item(0).Replace(ControlChars.Quote, "").ToString.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper
                            '    Tbl.Rows.Add(row)
                            'End If
                        End If
                    End If

                Else
                    EmptyCount += 1
                    If EmptyCount > 50 Then
                        Exit Do
                    End If
                End If
            Loop

            'Tbl.DefaultView.RowFilter = "[NOMOR_FAKTUR] LIKE '%00489245%'"
            'If Tbl.DefaultView.Count = 0 Then
            '    MsgBox("0")
            'Else
            '    MsgBox("Tidak 0")
            'End If
            'MsgBox("File Import Success", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error" & ", Error row=" & ij)
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

    Private Sub btnShow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Try

            If Me.dtpTo.Value < Me.dtpFrom.Value Then
                MsgBox("Please check your date.", MsgBoxStyle.Information, "Information")
                Me.dtpTo.Focus()
                Exit Sub
            End If
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetSubTotal(ByVal vColDPP_SMESS As Integer, ByVal vColPPN_SMES As Integer, ByVal vColDPP As Integer, ByVal vColPPN As Integer) As Decimal
        Dim Jml As Decimal = 0
        Try
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                'Jml += CDec(Me.dgvItem.Item(vCol, row.Index).Value)
            Next
            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Dim TblRev As New DseImport_CheckDate4.DataTable1DataTable
    Dim RowRev As DseImport_CheckDate4.DataTable1Row
    Private Function IsiDGV()
        Dim rowD As DataGridViewRow
        Dim TblDummy As DataTable = New DataTable
        Dim rowDummy As DataRow
        Dim SubTotalDPP_SMESS As Decimal = 0
        Dim SubTotalPPN_SMESS As Decimal = 0
        Dim SubTotalDPP As Decimal = 0
        Dim SubTotalPPN As Decimal = 0
        ListFakturPajakPengganti.Clear()
        Me.TblInitial.Rows.Clear()
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            TblDummy = Me.Tbl.Copy

            For i As Integer = TblDummy.Rows.Count - 1 To 0 Step -1
                rowDummy = TblDummy.Rows(i)

                If Me.GetConvertionDateFromString(rowDummy.Item("TANGGAL_FAKTUR")) < Format(Me.dtpFrom.Value.Date, "dd MMM yyyy") Or (Me.GetConvertionDateFromString(rowDummy.Item("TANGGAL_FAKTUR")) > Format(Me.dtpTo.Value.Date, "dd MMM yyyy")) Then
                    TblDummy.Rows.RemoveAt(i)
                End If
            Next

            If Me.cboTipe.SelectedIndex = 0 Then
                Me.TblRev.Rows.Clear()
                For Each rowDummy In TblDummy.Rows
                    Me.RowRev = Me.TblRev.NewDataTable1Row
                    With Me.RowRev
                        '.Invoice_Number = rowDummy.Item("REFERENSI")
                        ._Date = Format(Me.GetConvertionDateFromString(rowDummy.Item("TANGGAL_FAKTUR")), "dd MMM yyyy")
                        .Tax_Invoice_Number = rowDummy.Item("KD_JENIS_TRANSAKSI") & rowDummy.Item("FG_PENGGANTI") & Me.SetFormatTaxNo(rowDummy.Item("NOMOR_FAKTUR").ToString.Trim)

                        If rowDummy.Item("NPWP") = "000000000000000" Or rowDummy.Item("NPWP") = 0 Then
                            .NPWP = "00.000.000.0-000.000"
                        Else
                            .NPWP = Me.SetFormatNPWP(rowDummy.Item("NPWP"))
                        End If
                        .Customer = rowDummy.Item("NAMA")
                        .DPP = Me.ObjCD.DigitGrouping(rowDummy.Item("JUMLAH_DPP"), DefCurr)
                        .PPN = Me.ObjCD.DigitGrouping(rowDummy.Item("JUMLAH_PPN"), DefCurr)
                    End With
                    Me.TblRev.AddDataTable1Row(Me.RowRev)
                Next
                Me.colInvoiceNumberSMESS.Visible = False
                Me.colInvoiceNumber.Visible = False
                Me.colInvoiceNumberSMESS.HeaderText = "AOP No (SMESS)"
                Me.colCustomerSMESS.HeaderText = "Supplier (SMESS)"
                Me.ColCustomerCodeSMESS.Visible = False
                Me.ColCustomerNameSMESS.Visible = False
                Me.colInvoiceNumber.HeaderText = "AOP No"
                Me.colCustomer.HeaderText = "Supplier"
                Me.GetDataPurchase()
            ElseIf Me.cboTipe.SelectedIndex = 1 Then
                Me.TblRev.Rows.Clear()

                Me.Verication_FakturPajakPengganti(TblDummy) 'update: 2024/11/16

                For Each rowDummy In TblDummy.Rows

                    'If rowDummy.Item("NOMOR_FAKTUR").ToString.Contains("0082411073939") Then
                    '    MsgBox("")
                    'End If
                    Me.RowRev = Me.TblRev.NewDataTable1Row
                    With Me.RowRev
                        .Invoice_Number = rowDummy.Item(18)
                        ._Date = Format(Me.GetConvertionDateFromString(rowDummy.Item("TANGGAL_FAKTUR")), "dd MMM yyyy")
                        .Tax_Invoice_Number = rowDummy.Item("KD_JENIS_TRANSAKSI") & rowDummy.Item("FG_PENGGANTI") & Me.SetFormatTaxNo(rowDummy.Item("NOMOR_FAKTUR").ToString.Trim)

                        If rowDummy.Item("NPWP") = "000000000000000" Or rowDummy.Item("NPWP") = 0 Then
                            .NPWP = "00.000.000.0-000.000"
                        Else
                            .NPWP = Me.SetFormatNPWP(rowDummy.Item("NPWP"))
                        End If
                        .Customer = rowDummy.Item("NAMA")
                        .DPP = Me.ObjCD.DigitGrouping(rowDummy.Item("JUMLAH_DPP"), DefCurr)
                        .PPN = Me.ObjCD.DigitGrouping(rowDummy.Item("JUMLAH_PPN"), DefCurr)

                        If rowDummy.Item("FG_PENGGANTI") = 1 Then 'update: 2024/11/16
                            .Initial_Tax_Invoice = Me.GetInitalTaxInvoiceNo(rowDummy.Item("NOMOR_FAKTUR"))
                        End If
                    End With
                    Me.TblRev.AddDataTable1Row(Me.RowRev)
                Next
                Me.colInvoiceNumberSMESS.Visible = True
                Me.colInvoiceNumber.Visible = True
                Me.ColCustomerCodeSMESS.Visible = True
                Me.ColCustomerNameSMESS.Visible = True
                Me.colInvoiceNumberSMESS.HeaderText = "Invoice No (SMESS)"
                Me.colCustomerSMESS.HeaderText = "Customer NPWP Name (SMESS)"
                Me.colInvoiceNumber.HeaderText = "Invoice No"
                Me.colCustomer.HeaderText = "Customer"

                Me.TblRev.DefaultView.RowFilter = "[Tax Invoice Number] LIKE '%00489245%'"
                'If TblRev.DefaultView.Count = 0 Then
                '    MsgBox("0")
                'Else
                '    MsgBox("Tidak 0")
                'End If

                Me.GetDataSales()

                'Me.TblRev.DefaultView.RowFilter = "[Tax Invoice Number] LIKE '%00489245%'"
                'If TblRev.DefaultView.Count = 0 Then
                '    MsgBox("0")
                'Else
                '    MsgBox("Tidak 0")
                'End If
            ElseIf Me.cboTipe.SelectedIndex = 2 Then
                Me.TblRev.Rows.Clear()
                For Each rowDummy In TblDummy.Rows

                    'If rowDummy.Item("NOMOR_FAKTUR").ToString.Contains("00489245") Then
                    '    MsgBox("")
                    'End If
                    Me.RowRev = Me.TblRev.NewDataTable1Row
                    With Me.RowRev

                        .Invoice_Number = rowDummy.Item("NOMOR_DOKUMEN_RETUR")
                        ._Date = Format(Me.GetConvertionDateFromString(rowDummy.Item("TANGGAL_RETUR")), "dd MMM yyyy")
                        .Tax_Invoice_Number = Me.SetFormatTaxNo(rowDummy.Item("NOMOR_FAKTUR").ToString.Trim)

                        If rowDummy.Item("NPWP") = "000000000000000" Or rowDummy.Item("NPWP") = 0 Then
                            .NPWP = "00.000.000.0-000.000"
                        Else
                            .NPWP = Me.SetFormatNPWP(rowDummy.Item("NPWP"))
                        End If
                        .Customer = rowDummy.Item("NAMA")
                        .DPP = Me.ObjCD.DigitGrouping(rowDummy.Item("NILAI_RETUR_DPP"), DefCurr)
                        .PPN = Me.ObjCD.DigitGrouping(rowDummy.Item("NILAI_RETUR_PPN"), DefCurr)
                    End With
                    Me.TblRev.AddDataTable1Row(Me.RowRev)
                Next

                Me.colInvoiceNumberSMESS.HeaderText = "P. Return No (SMESS)"
                Me.colCustomerSMESS.HeaderText = "Supplier (SMESS)"
                Me.colInvoiceNumber.HeaderText = "P. Return No"
                Me.colCustomer.HeaderText = "Supplier"
                Me.ColCustomerCodeSMESS.Visible = False
                Me.ColCustomerNameSMESS.Visible = False
                Me.GetDataPurchaseReturn()
            ElseIf Me.cboTipe.SelectedIndex = 3 Then
                Me.TblRev.Rows.Clear()
                For Each rowDummy In TblDummy.Rows

                    'If rowDummy.Item("NOMOR_FAKTUR").ToString.Contains("00489245") Then
                    '    MsgBox("")
                    'End If
                    Me.RowRev = Me.TblRev.NewDataTable1Row
                    With Me.RowRev

                        .Invoice_Number = rowDummy.Item("NOMOR_DOKUMEN_RETUR")
                        ._Date = Format(Me.GetConvertionDateFromString(rowDummy.Item("TANGGAL_RETUR")), "dd MMM yyyy")
                        .Tax_Invoice_Number = Me.SetFormatTaxNo(rowDummy.Item("NOMOR_FAKTUR").ToString.Trim)

                        If rowDummy.Item("NPWP") = "000000000000000" Or rowDummy.Item("NPWP") = 0 Then
                            .NPWP = "00.000.000.0-000.000"
                        Else
                            .NPWP = Me.SetFormatNPWP(rowDummy.Item("NPWP"))
                        End If
                        .Customer = rowDummy.Item("NAMA")
                        .DPP = Me.ObjCD.DigitGrouping(rowDummy.Item("NILAI_RETUR_DPP"), DefCurr)
                        .PPN = Me.ObjCD.DigitGrouping(rowDummy.Item("NILAI_RETUR_PPN"), DefCurr)
                    End With
                    Me.TblRev.AddDataTable1Row(Me.RowRev)
                Next

                Me.colInvoiceNumberSMESS.HeaderText = "S. Return No (SMESS)"
                Me.colCustomerSMESS.HeaderText = "Customer (SMESS)"
                Me.colInvoiceNumber.HeaderText = "S. Return No"
                Me.colCustomer.HeaderText = "Customer"
                Me.ColCustomerCodeSMESS.Visible = False
                Me.ColCustomerNameSMESS.Visible = False
                Me.GetDataSalesReturn()
            End If

            'isi selisih
            For Each Me.RowRev In Me.TblRev.Rows
                With Me.RowRev
                    If .Invoice_Number_SMESS = .Invoice_Number Then
                        .Selisih_Invoice = False
                    Else
                        .Selisih_Invoice = True
                    End If

                    If .Customer_SMESS = .Customer Then
                        .Selisih_Customer = False
                    Else
                        .Selisih_Customer = True
                    End If

                    If .NPWP_SMESS = .NPWP Then
                        .Selisih_NPWP = False
                    Else
                        .Selisih_NPWP = True
                    End If

                    'saat pengecekan dpp smess wajib di round down.
                    If IsNumeric(.DPP_SMESS) = True And IsNumeric(.DPP) = True Then
                        If Math.Floor(CDec(.DPP_SMESS)) = .DPP Then
                            .Selisih_DPP = False
                        Else
                            .Selisih_DPP = True
                        End If
                    Else
                        If .DPP_SMESS = .DPP Then
                            .Selisih_DPP = False
                        Else
                            .Selisih_DPP = True
                        End If
                    End If
                  
                    If IsNumeric(.PPN_SMESS) = True And IsNumeric(.PPN) = True Then
                        'saat pengecekan ppn smess wajib di round down.
                        If Math.Floor(CDec(.PPN_SMESS)) = .PPN Then
                            .Selisih_PPN = False
                        Else
                            .Selisih_PPN = True
                        End If
                    Else
                        If .PPN_SMESS = .PPN Then
                            .Selisih_PPN = False
                        Else
                            .Selisih_PPN = True
                        End If
                    End If
                    
                End With
            Next

            'filter selisih
            If Me.ChkShowInvalidOnly.Checked = True Then
                For i As Integer = Me.TblRev.Rows.Count - 1 To 0 Step -1
                    Me.RowRev = Me.TblRev.Rows(i)
                    With Me.RowRev
                        'If Me.RowRev.Tax_Invoice_Number.Contains("00489245") Then
                        '    MsgBox("")
                        'End If

                        If Me.cboTipe.SelectedIndex = 0 Then
                            If Me.ChkShowInvalidOnly.Checked = False Then
                                If .Selisih_NPWP = True Or _
                                .Selisih_DPP = True Or _
                                .Selisih_PPN = True Then
                                    Me.TblRev.Rows.RemoveAt(i)
                                End If
                            ElseIf Me.ChkShowInvalidOnly.Checked = True Then
                                If .Selisih_NPWP = False And _
                                    .Selisih_DPP = False And _
                                    .Selisih_PPN = False Then
                                    Me.TblRev.Rows.RemoveAt(i)
                                End If
                            End If
                        Else
                            If Me.ChkShowInvalidOnly.Checked = False Then
                                If .Selisih_Invoice = True Or _
                                                           .Selisih_NPWP = True Or _
                                .Selisih_DPP = True Or _
                                .Selisih_PPN = True Then
                                    Me.TblRev.Rows.RemoveAt(i)
                                End If
                            ElseIf Me.ChkShowInvalidOnly.Checked = True Then
                                If .Selisih_Invoice = False And _
                                   .Selisih_NPWP = False And _
                                    .Selisih_DPP = False And _
                                    .Selisih_PPN = False Then
                                    Me.TblRev.Rows.RemoveAt(i)
                                End If
                            End If
                        End If
                    End With
                Next
            End If

            For Each RowRev In Me.TblRev.Rows
                'add subtotal
                SubTotalDPP_SMESS += CDec(Me.ObjCD.DigitGrouping(RowRev.DPP_SMESS, DefCurr))
                SubTotalPPN_SMESS += CDec(Me.ObjCD.DigitGrouping(RowRev.PPN_SMESS, DefCurr))
                SubTotalDPP += CDec(Me.ObjCD.DigitGrouping(RowRev.DPP, DefCurr))
                SubTotalPPN += CDec(Me.ObjCD.DigitGrouping(RowRev.PPN, DefCurr))

                If RowRev.Date_SMESS = #1/1/1950# Then
                    'MsgBox("Date smess null")
                    RowRev.Date_SMESS = RowRev._Date
                End If
            Next
            Me.dgvItem.Rows.Clear()
            Me.TblRev.DefaultView.RowFilter = ""
            Me.TblRev.DefaultView.Sort = "[Date SMESS]"
            For i As Integer = 0 To Me.TblRev.DefaultView.Count - 1
                Me.RowRev = Me.TblRev.DefaultView(i).Row
                With Me.RowRev
                    rowD = Me.dgvItem.Rows(Me.dgvItem.Rows.Add)
                    rowD.Cells(Me.colInvoiceNumberSMESS.Index).Value = .Invoice_Number_SMESS
                    If .Invoice_Number_SMESS <> "" Then
                        rowD.Cells(Me.colInvoiceDateSMESS.Index).Value = .Date_SMESS
                    Else
                        rowD.Cells(Me.colInvoiceDateSMESS.Index).Value = ""
                    End If
                    rowD.Cells(Me.colNoAcuan.Index).Value = .No_Acuan
                    rowD.Cells(Me.colNPWPSMESS.Index).Value = .NPWP_SMESS
                    rowD.Cells(Me.ColCustomerCodeSMESS.Index).Value = .Customer_Code_SMESS
                    rowD.Cells(Me.ColCustomerNameSMESS.Index).Value = .Customer_Name_SMESS
                    rowD.Cells(Me.colCustomerSMESS.Index).Value = .Customer_SMESS
                    rowD.Cells(Me.colDPPSMESS.Index).Value = Me.ObjCD.DigitGrouping(.DPP_SMESS, DefCurr)
                    rowD.Cells(Me.colPPNSMESS.Index).Value = Me.ObjCD.DigitGrouping(.PPN_SMESS, DefCurr)

                    rowD.Cells(Me.colInvoiceNumber.Index).Value = .Invoice_Number
                    If .Tax_Invoice_Number <> "" Then
                        rowD.Cells(Me.colInvoiceDate.Index).Value = ._Date
                    Else
                        rowD.Cells(Me.colInvoiceDate.Index).Value = ""
                    End If
                    rowD.Cells(Me.colTaxInvoiceNo.Index).Value = .Tax_Invoice_Number
                    rowD.Cells(Me.colNPWP.Index).Value = .NPWP
                    rowD.Cells(Me.colCustomer.Index).Value = .Customer
                    rowD.Cells(Me.colDPP.Index).Value = Me.ObjCD.DigitGrouping(.DPP, DefCurr)
                    rowD.Cells(Me.colPPN.Index).Value = Me.ObjCD.DigitGrouping(.PPN, DefCurr)
                    rowD.Cells(Me.colSelisihInvoice.Index).Value = .Selisih_Invoice
                    rowD.Cells(Me.colSelisihCustomer.Index).Value = .Selisih_Customer
                    rowD.Cells(Me.ColSelisihNPWP.Index).Value = .Selisih_NPWP
                    rowD.Cells(Me.colSelisihDPP.Index).Value = .Selisih_DPP
                    rowD.Cells(Me.colSelisihPPN.Index).Value = .Selisih_PPN

                    'If .No_Acuan.Contains("00489245") Or .Tax_Invoice_Number.Contains("00489245") Then
                    '    MsgBox("Ketemu dirow: " + (rowD.Index + 1).ToString)
                    'End If

                End With
            Next

            'add sub total
            rowD = Me.dgvItem.Rows(Me.dgvItem.Rows.Add)
            rowD.Cells(Me.colInvoiceDateSMESS.Index).Value = ""
            rowD.Cells(Me.colInvoiceDate.Index).Value = ""
            rowD.Cells(Me.colDPPSMESS.Index).Value = Me.ObjCD.DigitGrouping(SubTotalDPP_SMESS, DefCurr)
            rowD.Cells(Me.colPPNSMESS.Index).Value = Me.ObjCD.DigitGrouping(SubTotalPPN_SMESS, DefCurr)
            rowD.Cells(Me.colDPP.Index).Value = Me.ObjCD.DigitGrouping(SubTotalDPP, DefCurr)
            rowD.Cells(Me.colPPN.Index).Value = Me.ObjCD.DigitGrouping(SubTotalPPN, DefCurr)

            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Function

    Private Sub GetDataSales()
        Dim DTA As New DseImport_CheckDate4TableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim Tbl As New DseImport_CheckDate4.tbltrans_salesinvoiceDataTable
        Dim row As DseImport_CheckDate4.tbltrans_salesinvoiceRow
        Dim rowD As DataGridViewRow
        Dim NoFakturPajak As String = ""
        Dim Index As Integer = 0
        Dim rowI As DseImport_CheckDate4.DataTable2Row
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(Tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date) > 0 Then
                For Each row In Tbl.Rows
                    NoFakturPajak = row.No_Seri_Faktur_Pajak & row.No_Faktur_Pajak

                    'If NoFakturPajak.Contains("00489245") Then
                    '    MsgBox("")
                    'End If

                    Me.TblInitial.DefaultView.RowFilter = "[Initial Tax Invoice]='" & NoFakturPajak & "'" 'update: 2024/11/16
                    If Me.TblInitial.DefaultView.Count > 0 Then
                        Me.TblRev.DefaultView.RowFilter = "[Initial Tax Invoice]='" & NoFakturPajak & "'"
                    Else
                        Me.TblRev.DefaultView.RowFilter = "[Tax Invoice Number]='" & NoFakturPajak & "'"
                    End If
                    
                    If Me.TblRev.DefaultView.Count > 0 Then
                        Me.RowRev = Me.TblRev.DefaultView(0).Row
                        'If Me.RowRev.Tax_Invoice_Number.Contains("0041982256651") Then
                        '    MsgBox("")
                        'End If

                        With Me.RowRev
                            .Invoice_Number_SMESS = row.No_Transaksi
                            .Date_SMESS = Format(row.Tanggal_Transaksi, "dd MMM yyyy")
                            .No_Acuan = NoFakturPajak

                            If row.No_NPWP = "000000000000000" Then
                                .NPWP_SMESS = "00.000.000.0-000.000"
                            Else
                                .NPWP_SMESS = row.No_NPWP
                            End If
                            .Customer_SMESS = row.NPWP_Name
                            .Customer_Code_SMESS = row.Customer_Code
                            .Customer_Name_SMESS = row.Customer_Name
                            .DPP_SMESS = Math.Floor(row.Sub_Total) 'Me.ObjCD.DigitGrouping(row.Sub_Total, DefCurr)
                            .PPN_SMESS = Math.Floor(row.Sub_Total * row.Tax / 100)  'Me.ObjCD.DigitGrouping(row.Tax_Amount, DefCurr)
                            .PPN_SMESS = Me.ObjCD.DigitGrouping(.PPN_SMESS, 2)
                            .DPP_SMESS = Me.ObjCD.DigitGrouping(.DPP_SMESS)
                        End With
                    Else
                        Me.RowRev = Me.TblRev.NewDataTable1Row
                        With Me.RowRev
                            .Invoice_Number_SMESS = row.No_Transaksi
                            .Date_SMESS = Format(row.Tanggal_Transaksi, "dd MMM yyyy")
                            .No_Acuan = NoFakturPajak

                            If row.No_NPWP = "000000000000000" Then
                                .NPWP_SMESS = "00.000.000.0-000.000"
                            Else
                                .NPWP_SMESS = row.No_NPWP
                            End If
                            .Customer_SMESS = row.NPWP_Name
                            .Customer_Code_SMESS = row.Customer_Code
                            .Customer_Name_SMESS = row.Customer_Name
                            .DPP_SMESS = Math.Floor(row.Sub_Total) 'Me.ObjCD.DigitGrouping(row.Sub_Total, DefCurr)
                            .PPN_SMESS = Math.Floor(row.Sub_Total * row.Tax / 100)  'Me.ObjCD.DigitGrouping(row.Tax_Amount, DefCurr)
                            .PPN_SMESS = Me.ObjCD.DigitGrouping(.PPN_SMESS, 2)
                            .DPP_SMESS = Me.ObjCD.DigitGrouping(.DPP_SMESS)
                        End With
                        Me.TblRev.AddDataTable1Row(Me.RowRev)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex(ByVal vTaxInvoiceNo As String) As Integer
        Try
            For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                If vTaxInvoiceNo = rowdgv.Cells(Me.colTaxInvoiceNo.Index).Value Then
                    Return rowdgv.Index
                    Exit Function
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub GetDataPurchase()
        Dim DTA As New DseImport_CheckDate4TableAdapters.tbltrans_aopTableAdapter
        Dim Tbl As New DseImport_CheckDate4.tbltrans_aopDataTable
        Dim row As DseImport_CheckDate4.tbltrans_aopRow
        Dim rowD As DataGridViewRow
        Dim Index As Integer = 0
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(Tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date) > 0 Then
                For Each row In Tbl.Rows
                    '010.002-19.36217336
                    '010.002-19.36217336
                    Me.TblRev.DefaultView.RowFilter = "[Tax Invoice Number]='" & row.Tax_Invoice_Number.ToString & "'"
                    If Me.TblRev.DefaultView.Count > 0 Then
                        Me.RowRev = Me.TblRev.DefaultView(0).Row
                        With Me.RowRev
                            .Invoice_Number_SMESS = row.No_Transaksi
                            .Date_SMESS = Format(row.Tanggal_Transaksi, "dd MMM yyyy")
                            .No_Acuan = row.Tax_Invoice_Number
                            .NPWP_SMESS = row.No_NPWP
                            If row.No_NPWP = "000000000000000" Then
                                .NPWP_SMESS = "00.000.000.0-000.000"
                            End If
                            .Customer_SMESS = row.Supplier_Name
                            .DPP_SMESS = Me.ObjCD.DigitGrouping(row.Sub_Total, DefCurr)
                            .PPN_SMESS = Me.ObjCD.DigitGrouping(row.Tax_Amount, DefCurr)
                        End With
                    Else
                        Me.RowRev = Me.TblRev.NewDataTable1Row
                        With Me.RowRev
                            .Invoice_Number_SMESS = row.No_Transaksi
                            .Date_SMESS = Format(row.Tanggal_Transaksi, "dd MMM yyyy")
                            .No_Acuan = row.Tax_Invoice_Number
                            .NPWP_SMESS = row.No_NPWP
                            If row.No_NPWP = "000000000000000" Then
                                .NPWP_SMESS = "00.000.000.0-000.000"
                            End If
                            .Customer_SMESS = row.Supplier_Name
                            .DPP_SMESS = Me.ObjCD.DigitGrouping(row.Sub_Total, DefCurr)
                            .PPN_SMESS = Me.ObjCD.DigitGrouping(row.Tax_Amount, DefCurr)
                        End With
                        Me.TblRev.AddDataTable1Row(Me.RowRev)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetDataPurchaseReturn()
        Dim DTA As New DseImport_CheckDate4TableAdapters.tbltrans_purchasereturnTableAdapter
        Dim Tbl As New DseImport_CheckDate4.tbltrans_purchasereturnDataTable
        Dim row As DseImport_CheckDate4.tbltrans_purchasereturnRow
        Dim DTAOut_PR_Ref As New DseImport_CheckDate4TableAdapters.Purchase_Return_ReferenceTableAdapter
        Dim TBLOut_PR_Ref As New DseImport_CheckDate4.Purchase_Return_ReferenceDataTable
        Dim rowOut_PR_Ref As DseImport_CheckDate4.Purchase_Return_ReferenceRow
        Dim rowD As DataGridViewRow
        Dim Index As Integer = 0
        Dim PrimaryGRNMax As Double = -1
        Dim PrimaryGRNMin As Double = -1
        Dim NoFakturPajak As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(Tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, 0, TipePR) > 0 Then
                For Each row In Tbl.Rows
                    With row
                        If .Primary_GRN > PrimaryGRNMax Then
                            PrimaryGRNMax = .Primary_GRN
                        End If
                        If .Primary_GRN < PrimaryGRNMin Then
                            PrimaryGRNMin = -.Primary_GRN
                        End If
                    End With
                Next
                DTAOut_PR_Ref.FillBy_Primary(TBLOut_PR_Ref, PrimaryGRNMax, PrimaryGRNMin)

                For Each row In Tbl.Rows
                    TBLOut_PR_Ref.DefaultView.RowFilter = "[Primary GRN] = " + row.Primary_GRN.ToString
                    If TBLOut_PR_Ref.DefaultView.Count = 0 Then
                        NoFakturPajak = row.No_Nota_Retur_Pajak
                    Else
                        rowOut_PR_Ref = TBLOut_PR_Ref.DefaultView(0).Row
                        NoFakturPajak = rowOut_PR_Ref.Tax_Invoice_Number
                    End If

                    Me.TblRev.DefaultView.RowFilter = "[Tax Invoice Number]='" & NoFakturPajak & "'"
                    If Me.TblRev.DefaultView.Count > 0 Then
                        Me.RowRev = Me.TblRev.DefaultView(0).Row
                        With Me.RowRev
                            .Invoice_Number_SMESS = row.No_Transaksi
                            .Date_SMESS = Format(row.Tanggal_Transaksi, "dd MMM yyyy")
                            .No_Acuan = NoFakturPajak
                            .NPWP_SMESS = row.No_NPWP
                            If row.No_NPWP = "000000000000000" Then
                                .NPWP_SMESS = "00.000.000.0-000.000"
                            End If
                            .Customer_SMESS = row.Supplier_Name
                            .DPP_SMESS = Me.ObjCD.DigitGrouping(row.Sub_Total, DefCurr)
                            .PPN_SMESS = Me.ObjCD.DigitGrouping(row.Tax_Amount, DefCurr)
                        End With
                    Else
                        Me.RowRev = Me.TblRev.NewDataTable1Row
                        With Me.RowRev
                            .Invoice_Number_SMESS = row.No_Transaksi
                            .Date_SMESS = Format(row.Tanggal_Transaksi, "dd MMM yyyy")
                            .No_Acuan = NoFakturPajak
                            .NPWP_SMESS = row.No_NPWP
                            If row.No_NPWP = "000000000000000" Then
                                .NPWP_SMESS = "00.000.000.0-000.000"
                            End If
                            .Customer_SMESS = row.Supplier_Name
                            .DPP_SMESS = Me.ObjCD.DigitGrouping(row.Sub_Total, DefCurr)
                            .PPN_SMESS = Me.ObjCD.DigitGrouping(row.Tax_Amount, DefCurr)
                        End With
                        Me.TblRev.AddDataTable1Row(Me.RowRev)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetDataSalesReturn()
        Dim DTA As New DseImport_CheckDate4TableAdapters.tbltrans_salesreturnTableAdapter
        Dim Tbl As New DseImport_CheckDate4.tbltrans_salesreturnDataTable
        Dim row As DseImport_CheckDate4.tbltrans_salesreturnRow
        Dim rowD As DataGridViewRow
        Dim Index As Integer = 0
        Dim NoFakturPajak As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(Tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, 0, TipeSR) > 0 Then
                For Each row In Tbl.Rows
                    If row.Primary_Customer_Quick <> -1 Then
                        NoFakturPajak = row.Tax_Invoice_No_Quick
                    Else
                        NoFakturPajak = row.No_Seri_Faktur_Pajak & row.No_Nota_Retur_Pajak
                    End If
                    Me.TblRev.DefaultView.RowFilter = "[Tax Invoice Number]='" & NoFakturPajak & "'"
                    If Me.TblRev.DefaultView.Count > 0 Then
                        Me.RowRev = Me.TblRev.DefaultView(0).Row
                        With Me.RowRev
                            If row.Primary_Customer_Quick <> -1 Then
                                .Customer_SMESS = row.NPWP_Name
                                .Customer_Code_SMESS = row.Customer_Code_Quick
                                .Customer_Name_SMESS = row.Quick_Cust_Name
                                .No_Acuan = row.Tax_Invoice_No_Quick
                            Else
                                .Customer_SMESS = row.NPWP_Name
                                .Customer_Code_SMESS = row.Customer_Code_Quick
                                .Customer_Name_SMESS = row.Customer_Name
                                .No_Acuan = row.No_Nota_Retur_Pajak
                            End If
                            .Invoice_Number_SMESS = row.No_Transaksi
                            .Date_SMESS = Format(row.Tanggal_Transaksi, "dd MMM yyyy")
                            .NPWP_SMESS = row.No_NPWP
                            If row.No_NPWP = "000000000000000" Then
                                .NPWP_SMESS = "00.000.000.0-000.000"
                            End If
                            .DPP_SMESS = Me.ObjCD.DigitGrouping(row.Sub_Total, DefCurr)
                            .PPN_SMESS = Me.ObjCD.DigitGrouping(row.Tax_Amount, DefCurr)
                        End With
                    Else
                        Me.RowRev = Me.TblRev.NewDataTable1Row
                        With Me.RowRev
                            If row.Primary_Customer_Quick <> -1 Then
                                .Customer_SMESS = row.NPWP_Name
                                .Customer_Code_SMESS = row.Customer_Code_Quick
                                .Customer_Name_SMESS = row.Quick_Cust_Name
                                .No_Acuan = row.Tax_Invoice_No_Quick
                            Else
                                .Customer_SMESS = row.NPWP_Name
                                .Customer_Code_SMESS = row.Customer_Code_Quick
                                .Customer_Name_SMESS = row.Customer_Name
                                .No_Acuan = row.No_Nota_Retur_Pajak
                            End If
                            .Invoice_Number_SMESS = row.No_Transaksi
                            .Date_SMESS = Format(row.Tanggal_Transaksi, "dd MMM yyyy")
                            .NPWP_SMESS = row.No_NPWP
                            If row.No_NPWP = "000000000000000" Then
                                .NPWP_SMESS = "00.000.000.0-000.000"
                            End If
                            .DPP_SMESS = Me.ObjCD.DigitGrouping(row.Sub_Total, DefCurr)
                            .PPN_SMESS = Me.ObjCD.DigitGrouping(row.Tax_Amount, DefCurr)
                        End With
                        Me.TblRev.AddDataTable1Row(Me.RowRev)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Try

            Threading.Thread.Sleep(1)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.ObjCD.UpdateThreading()
            Me.IsiDGV()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError

    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim TblInitial As New DseImport_CheckDate4.DataTable2DataTable
    Private Sub AddRowTabelInitial(ByVal vNomorFaktur As String, ByVal vInitialInvoiceNo As String)
        Dim row As DseImport_CheckDate4.DataTable2Row
        Try
            row = TblInitial.NewDataTable2Row
            With row
                .Nomor_Faktur = vNomorFaktur
                .Initial_Tax_Invoice = vInitialInvoiceNo
            End With
            TblInitial.AddDataTable2Row(row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Verication_FakturPajakPengganti(ByRef TblDummy As DataTable)
        Dim rowDummy As DataRow
        Dim rowI As DseImport_CheckDate4.DataTable2Row
        Try
            '1. Cek apakah ada data dari excel tsb adalah faktur pajak pengganti?
            TblDummy.DefaultView.RowFilter = "[FG_PENGGANTI]=" & 1 'jika > 0 berarti ada.
            For i As Integer = 0 To TblDummy.DefaultView.Count - 1
                rowDummy = TblDummy.DefaultView(0).Row
                With rowDummy
                    Me.AddRowTabelInitial(.Item("NOMOR_FAKTUR"), "")
                End With

             
            Next

            'Isi data initial/sebelumnya
            For Each rowI In Me.TblInitial.Rows
                TblDummy.DefaultView.RowFilter = "[FG_PENGGANTI]=" & 0 & " AND [NOMOR_FAKTUR]=" & rowI.Nomor_Faktur
                If TblDummy.DefaultView.Count > 0 Then
                    rowDummy = TblDummy.DefaultView(0).Row
                    With rowDummy
                        rowI.Initial_Tax_Invoice = rowDummy.Item("KD_JENIS_TRANSAKSI") & rowDummy.Item("FG_PENGGANTI") & Me.SetFormatTaxNo(rowDummy.Item("NOMOR_FAKTUR").ToString.Trim)
                    End With

                    'hapus faktur yang lama. 
                    TblDummy.DefaultView.Delete(0)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetInitalTaxInvoiceNo(ByVal vNomorFaktur As String) As String
        Dim rowI As DseImport_CheckDate4.DataTable2Row
        Try
            TblInitial.DefaultView.RowFilter = "[Nomor Faktur]='" & vNomorFaktur & "'"
            If TblInitial.DefaultView.Count > 0 Then
                rowI = TblInitial.DefaultView(0).Row
                Return rowI.Initial_Tax_Invoice
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class