Option Explicit On
Option Strict Off

Public Class ClsMasterTransactionNumbering
    Private mPrimary As Double = -1
    Private mTipeTransaksi As EnumTipeTransaksiNumbering = EnumTipeTransaksiNumbering.Invoice
    Private mApakahAutoNumber As Integer = 0
    Private mDesripsi As String = ""
    Private mReset As EnumReset = EnumReset.Yearly

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Property Reset() As EnumReset
        Get
            Return Me.mReset
        End Get
        Set(ByVal value As EnumReset)
            mReset = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return Me.mDesripsi
        End Get
        Set(ByVal value As String)
            mDesripsi = value
        End Set
    End Property

    Public Property ApakahAutoNumber() As Boolean
        Get
            If Me.mApakahAutoNumber = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mApakahAutoNumber = 1
            Else
                Me.mApakahAutoNumber = 0
            End If
        End Set
    End Property

    Public Property TipeTransaksi() As EnumTipeTransaksiNumbering
        Get
            Return Me.mTipeTransaksi
        End Get
        Set(ByVal value As EnumTipeTransaksiNumbering)
            mTipeTransaksi = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property CreatedBy() As String
        Get
            Return Me.mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return Me.mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return Me.mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return Me.mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return Me.mTimestamp
        End Get
        Set(ByVal value As Date)
            Me.mTimestamp = value
        End Set
    End Property

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Public Enum EnumTipeTransaksiNumbering As Integer
        PO_Lokal
        PO_Impor
        GRN
        AOP_Lokal
        AOP_Impor
        Quotation
        SO
        _DO
        Invoice
        General_Jurnal
        Sales_Jurnal
        Purchase_Jurnal
        Closing_Jurnal
        CollectionLetter
        _MP
        _SF
        MO
        JO
        WO
        SPL
        Bundling
        Schedule_Downtime
        Credit_Note
        Debit_Note
    End Enum

    Public Enum EnumReset As Integer
        Yearly
        Monthly
    End Enum

    Public Enum EnumFormat As Integer
        Month_Romawi                    'I, II, III, IV, dst
        Month_1Digit                    '1,2,3,4, dst
        Month_2Digit                    '01, 02, 03, 04, dst
        Month_StringPendek_Indo         'Jan, Feb, Mar, Apr, Mei, dst
        Month_StringPanjang_Indo        'Januari, Februari, Maret, dst
        Month_StringPendek_English      'Jan, Feb, Mar, Apr, May, Jun, dst
        Month_StringPanjang_English     'January, February, March, dst
        Year_Pendek                     '01,02,03, dst
        Year_Panjang                    '2001, 2002, 2003, dst
        Running_2Digit                  'index : 9
        Running_3Digit
        Running_4Digit
        Running_5Digit
        Running_6Digit
        Running_7Digit
        Running_8Digit
        Running_9Digit
        Running_10Digit
        FixedText
    End Enum

    Public Enum EnumTipeKomponen As Integer
        FixedText
        Month
        Year
        RunningNumber
    End Enum

    'penomoran transaksi tidak bisa diedit jika:
    '1. sudah ada transaksi untuk item tersebut
    Public Function ApakahBolehEdit(ByVal vPrimary As Double) As Boolean
        Dim NamaTabel As String = ""
        Dim ObjMySQL As New ClsMySQLReference

        Try
            NamaTabel = Me.GetNamaTabelRujukan(vPrimary)
            If NamaTabel <> "" Then
                If CDbl(ObjMySQL.QueryExecuteScalar("SELECT COUNT(*) FROM " + NamaTabel + ";")) > 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return True
            End If

            'sampe sini maka boleh diedit
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetNamaTabelRujukan(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterNumberingTableAdapters.tblmaster_numberingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNamaTabelRujukan(vPrimary)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNamaTabelRujukan(vPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    'mengembalikan nilai true jika transaksinya merupakan autonumber, dan mengembalikan nilai false jika transaksinya bukan autonumber
    Public Function Apakah_Autonumber(ByVal vTipeTransaksi As EnumTipeTransaksiNumbering) As Boolean
        Dim DTA As New DseMasterNumberingTableAdapters.tblmaster_numberingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetAutonumber(vTipeTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetAutonumber(vTipeTransaksi).Value > 0 Then
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
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeKomponenInt(ByVal vTipeKomponenStr As String) As ClsMasterTransactionNumbering.EnumTipeKomponen
        Try
            If vTipeKomponenStr = "Fixed Text" Then
                Return ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText
            ElseIf vTipeKomponenStr = "Month" Then
                Return ClsMasterTransactionNumbering.EnumTipeKomponen.Month
            ElseIf vTipeKomponenStr = "Year" Then
                Return ClsMasterTransactionNumbering.EnumTipeKomponen.Year
            ElseIf vTipeKomponenStr = "Running Number" Then
                Return ClsMasterTransactionNumbering.EnumTipeKomponen.RunningNumber
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetTipeKomponenStr(ByVal vTipeKomponenInt As ClsMasterTransactionNumbering.EnumTipeKomponen) As String
        Try
            If vTipeKomponenInt = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText Then
                Return "Fixed Text"
            ElseIf vTipeKomponenInt = ClsMasterTransactionNumbering.EnumTipeKomponen.Month Then
                Return "Month"
            ElseIf vTipeKomponenInt = ClsMasterTransactionNumbering.EnumTipeKomponen.Year Then
                Return "Year"
            ElseIf vTipeKomponenInt = ClsMasterTransactionNumbering.EnumTipeKomponen.RunningNumber Then
                Return "Running Number"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetFormatInt(ByVal vFormatStr As String, ByVal vTipeKomponen As EnumTipeKomponen) As ClsMasterTransactionNumbering.EnumFormat
        Try
            If vFormatStr = "01, 02, 03, 04, ..." Then
                If vTipeKomponen = EnumTipeKomponen.Month Then
                    Return EnumFormat.Month_2Digit
                ElseIf vTipeKomponen = EnumTipeKomponen.Year Then
                    Return EnumFormat.Year_Pendek
                ElseIf vTipeKomponen = EnumTipeKomponen.RunningNumber Then
                    Return EnumFormat.Running_2Digit
                End If
            ElseIf vFormatStr = "001, 002, 003, 004, ..." Then
                Return EnumFormat.Running_3Digit
            ElseIf vFormatStr = "0001, 0002, 0003, 0004, ..." Then
                Return EnumFormat.Running_4Digit
            ElseIf vFormatStr = "00001, 00002, 00003, 00004, ..." Then
                Return EnumFormat.Running_5Digit
            ElseIf vFormatStr = "000001, 000002, 000003, 000004, ..." Then
                Return EnumFormat.Running_6Digit
            ElseIf vFormatStr = "0000001, 0000002, 0000003, 0000004, ..." Then
                Return EnumFormat.Running_7Digit
            ElseIf vFormatStr = "00000001, 00000002, 00000003, 00000004, ..." Then
                Return EnumFormat.Running_8Digit
            ElseIf vFormatStr = "000000001, 000000002, 000000003, 000000004, ..." Then
                Return EnumFormat.Running_9Digit
            ElseIf vFormatStr = "0000000001, 0000000002, 0000000003, 0000000004, ..." Then
                Return EnumFormat.Running_10Digit
            ElseIf vFormatStr = "" Then
                Return EnumFormat.FixedText
            ElseIf vFormatStr = "1, 2, 3, 4, ..." Then
                Return EnumFormat.Month_1Digit
            ElseIf vFormatStr = "Jan, Feb, Mar, Apr, Mei, ..." Then
                Return EnumFormat.Month_StringPendek_Indo
            ElseIf vFormatStr = "Januari, Februari, Maret, ..." Then
                Return EnumFormat.Month_StringPanjang_Indo
            ElseIf vFormatStr = "Jan, Feb, Mar, Apr, May, ..." Then
                Return EnumFormat.Month_StringPendek_English
            ElseIf vFormatStr = "January, February, March, ..." Then
                Return EnumFormat.Month_StringPanjang_English
            ElseIf vFormatStr = "I, II, III, IV, ..." Then
                Return EnumFormat.Month_Romawi
            ElseIf vFormatStr = "2001, 2002, 2003, 2004, ..." Then
                Return EnumFormat.Year_Panjang
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetFormatStr(ByVal vFormatInt As ClsMasterTransactionNumbering.EnumFormat) As String
        Try
            If vFormatInt = EnumFormat.FixedText Then
                Return ""
            ElseIf vFormatInt = EnumFormat.Month_1Digit Then
                Return "1, 2, 3, 4, ..."
            ElseIf vFormatInt = EnumFormat.Month_2Digit Then
                Return "01, 02, 03, 04, ..."
            ElseIf vFormatInt = EnumFormat.Month_Romawi Then
                Return "I, II, III, IV, ..."
            ElseIf vFormatInt = EnumFormat.Month_StringPanjang_English Then
                Return "January, February, March, ..."
            ElseIf vFormatInt = EnumFormat.Month_StringPanjang_Indo Then
                Return "Jan, Feb, Mar, Apr, May, ..."
            ElseIf vFormatInt = EnumFormat.Month_StringPendek_English Then
                Return "January, February, March, ..."
            ElseIf vFormatInt = EnumFormat.Month_StringPendek_Indo Then
                Return "Jan, Feb, Mar, Apr, Mei, ..."
            ElseIf vFormatInt = EnumFormat.Running_10Digit Then
                Return "00000001, 00000002, 00000003, 0000000004, ..."
            ElseIf vFormatInt = EnumFormat.Running_2Digit Then
                Return "01, 02, 03, 04, ..."
            ElseIf vFormatInt = EnumFormat.Running_3Digit Then
                Return "001, 002, 003, 004, ..."
            ElseIf vFormatInt = EnumFormat.Running_4Digit Then
                Return "0001, 0002, 0003, 0004, ..."
            ElseIf vFormatInt = EnumFormat.Running_5Digit Then
                Return "00001, 00002, 00003, 00004, ..."
            ElseIf vFormatInt = EnumFormat.Running_6Digit Then
                Return "000001, 000002, 000003, 000004, ..."
            ElseIf vFormatInt = EnumFormat.Running_7Digit Then
                Return "0000001, 0000002, 0000003, 0000004, ..."
            ElseIf vFormatInt = EnumFormat.Running_8Digit Then
                Return "00000001, 00000002, 00000003, 00000004, ..."
            ElseIf vFormatInt = EnumFormat.Running_9Digit Then
                Return "000000001, 000000002, 000000003, 000000004, ..."
            ElseIf vFormatInt = EnumFormat.Year_Panjang Then
                Return "2001, 2002, 2003, 2004, ..."
            ElseIf vFormatInt = EnumFormat.Year_Pendek Then
                Return "01, 02, 03, 04, ..."
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vTipeTransaksi As EnumTipeTransaksiNumbering) As Boolean
        Dim DTA As New DseMasterNumberingTableAdapters.tblmaster_numberingTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByTipeTransaksi(vTipeTransaksi)

            If RetVal.HasValue = False Then
                Return False
            Else
                If RetVal.Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'kalo bukan autonumber, maka dianggap reset per tahun
    Public Function GetResetType(ByVal vTipeNumbering As EnumTipeTransaksiNumbering) As EnumReset
        Dim DTA As New DseMasterNumberingTableAdapters.tblmaster_numberingTableAdapter
        Dim RetVal As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetTipeResetAutonumber(vTipeNumbering)

            If RetVal.HasValue = False Then
                Return EnumReset.Yearly
            Else
                If RetVal.Value = 0 Then
                    Return EnumReset.Yearly
                Else
                    Return EnumReset.Monthly
                End If
            End If

            Return EnumReset.Yearly
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumReset.Yearly
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vDesc As String) As Double
        Dim DTA As New DseMasterNumberingTableAdapters.tblmaster_numberingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vDesc).HasValue = False Then
                Return -1
            Else
                Return DTA.ScalarQuery_GetPrimary(vDesc).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiCbo(ByVal vCbo As ComboBox, ByVal vTBL As DseMasterNumbering.tblmaster_numberingDataTable)
        Dim DTA As New DseMasterNumberingTableAdapters.tblmaster_numberingTableAdapter

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.Fill(vTBL)
            vTBL.DefaultView.Sort = "deskripsi"
            vCbo.DataSource = vTBL.DefaultView
            vCbo.DisplayMember = "deskripsi"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Public Overridable Function SimpanPerubahan(ByVal vDGV As DataGridView, ByVal i As FrmMasterNumbering, ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Transaksi_Numbering_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vDGV, i, vUserCode)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjAntri = Nothing
            SW = Nothing
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Try
            'data salah jika: --> semunya di handle pada form
            '1. jika auto number, maka minimal harus ada 1 running number, 1 bulan, 1 tahun, 7an: supaya ga ada double, supaya autonumber
            '2. tipe auto number tetapi tidak memliki item
            '2. tipenya bukan auto number tapi:
            '       a. memliki item
            '       b. 
            '3. tiap komponen yg tipenya fixed text harus ada isinya

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vDGV As DataGridView, ByVal i As FrmMasterNumbering, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterNumberingTableAdapters.tblmaster_numberingTableAdapter
        Dim DTAItem As New DseMasterNumberingTableAdapters.tblmaster_numbering_itemTableAdapter
        Dim DGV As DataGridView
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            If Me.ApakahDataBenar = True Then
                'update data header
                DTA.UpdateQuery(Me.mApakahAutoNumber, Me.mReset, vUserCode, Now, Me.mPrimary, Me.mTimestamp)

                'simpan data item
                DTAItem.DeleteQuery(Me.mPrimary)

                DGV = i.dgvItem
                For Each rowDGV In DGV.Rows
                    DTAItem.InsertQuery(Me.mPrimary, rowDGV.Cells(i.ColTipe_Int.Index).Value _
                    , rowDGV.Cells(i.ColFormatInt.Index).Value, rowDGV.Cells(i.ColValue.Index).Value)
                Next

                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAItem.Dispose()
        End Try
    End Function

    Public Sub ResetField()
        Try
            Me.mPrimary = -1
            Me.mApakahAutoNumber = 0
            Me.mCreatedBy = ""
            Me.mCreatedDate = Today
            Me.mDesripsi = ""
            Me.mModifiedBy = ""
            Me.mModifiedDate = Today
            Me.mReset = EnumReset.Monthly
            Me.mTimestamp = Now
            Me.mTipeTransaksi = EnumTipeTransaksiNumbering.PO_Lokal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahItemCocok(ByVal vDGV As DataGridView, ByVal i As FrmMasterNumbering, ByVal vPrimary As Double) As Boolean
        Dim DTAItem As New DseMasterNumberingTableAdapters.tblmaster_numbering_itemTableAdapter
        Dim TBLItem As New DseMasterNumbering.tblmaster_numbering_itemDataTable
        Dim rowItem As DseMasterNumbering.tblmaster_numbering_itemRow
        Dim rowDGV As DataGridViewRow
        Dim ij As Integer = 0

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.FillBy_primary(TBLItem, vPrimary)

            If TBLItem.Rows.Count <> vDGV.Rows.Count Then
                Return False
            End If

            For Each rowItem In TBLItem.Rows
                rowDGV = vDGV.Rows(ij)

                If rowDGV.Cells(i.ColTipe_Int.Index).Value <> rowItem.tipe_komponen Then
                    Return False
                End If
                If rowDGV.Cells(i.ColFormatInt.Index).Value <> rowItem.format Then
                    Return False
                End If
                If rowDGV.Cells(i.ColTipe_Int.Index).Value = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText Then
                    If rowDGV.Cells(i.ColValue.Index).Value.ToString <> rowItem.value Then
                        Return False
                    End If
                End If

                ij += 1
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Public Sub GetTanggalDariSampai(ByRef vTanggalDari As Date, ByRef vTanggalSampai As Date, _
    ByVal vTipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering, ByVal vReferenceDate As Date)
        Dim DTA As New DseMasterNumberingTableAdapters.tblmaster_numberingTableAdapter
        Dim TBL As New DseMasterNumbering.tblmaster_numberingDataTable
        Dim row As DseMasterNumbering.tblmaster_numberingRow
        Dim ResetMode As EnumReset = EnumReset.Monthly
        Dim ObjCD As New ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_TipeTransaksi(TBL, vTipeTransaksi)

            If TBL.Rows.Count = 0 Then
                vTanggalDari = #1/1/1950#
                vTanggalSampai = #1/1/1950#
            Else
                row = TBL.Rows(0)
                ResetMode = row.reset_numbering

                If ResetMode = EnumReset.Monthly Then
                    vTanggalDari = ObjCD.GetFirstDateofTheMonth(vReferenceDate)
                    vTanggalSampai = ObjCD.GetLasDateOfTheMonth(vReferenceDate)
                ElseIf ResetMode = EnumReset.Yearly Then
                    vTanggalDari = ObjCD.GetFisrtDateOfTheYear(vReferenceDate)
                    vTanggalSampai = ObjCD.GetLastDateOfTheYear(vReferenceDate)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            TBL.Dispose()
        End Try
    End Sub

    Public Sub CariDetail(ByVal vPrimary As Double, ByVal vDGV As DataGridView, ByVal i As FrmMasterNumbering)
        Dim DTA As New DseMasterNumberingTableAdapters.tblmaster_numberingTableAdapter
        Dim TBL As New DseMasterNumbering.tblmaster_numberingDataTable
        Dim row As DseMasterNumbering.tblmaster_numberingRow
        Dim DTAItem As New DseMasterNumberingTableAdapters.tblmaster_numbering_itemTableAdapter
        Dim TBLItem As New DseMasterNumbering.tblmaster_numbering_itemDataTable
        Dim rowItem As DseMasterNumbering.tblmaster_numbering_itemRow
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.FillBy_Primary(TBL, vPrimary)
            DTAItem.FillBy_primary(TBLItem, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Me.mPrimary = row.primary

                With row
                    Me.mPrimary = .primary
                    Me.mApakahAutoNumber = row.auto_number
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mDesripsi = .deskripsi
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mReset = .reset_numbering
                    Me.mTimestamp = .timestamp
                    Me.mTipeTransaksi = .tipe_transaksi
                End With

                'isikan data item
                vDGV.Rows.Clear()
                For Each rowItem In TBLItem.Rows
                    With rowItem
                        vDGV.Rows.Add()
                        rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)

                        rowDGV.Cells(i.ColFormatInt.Index).Value = .format
                        rowDGV.Cells(i.ColTipeKomponen.Index).Value = Me.GetTipeKomponenStr(.tipe_komponen)
                        rowDGV.Cells(i.ColFormat.Index).Value = Me.GetFormatStr(.format)
                        rowDGV.Cells(i.ColValue.Index).Value = .value
                    End With
                Next
            Else
                Me.ResetField()
                vDGV.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Sub

    Public Sub SimpanBaru(ByVal vDeskripsi As String, ByVal vTipeTransaksi As EnumTipeTransaksiNumbering, ByVal vTabelRujukan As String, ByVal vUserCode As String)
        Dim DTA As New DseMasterNumberingTableAdapters.tblmaster_numberingTableAdapter

        Try
            'simpan bukan autonumber, reset yearly
            DTA.InsertQuery(vDeskripsi, CInt(vTipeTransaksi), 0, 0, vUserCode, Now, vUserCode, Now, Now, vTabelRujukan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetNomorTransaksiBaru1(ByVal vReferenceDate As Date, ByVal vReferenceUrutan As Double, _
    ByVal vTipeTranskasi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering) As String
        Dim Dummy As String = ""
        Dim TipeKomponen As ClsMasterTransactionNumbering.EnumTipeKomponen
        Dim Format As ClsMasterTransactionNumbering.EnumFormat
        Dim DTA As New DseMasterNumberingTableAdapters.tblmaster_numberingTableAdapter
        Dim TBL As New DseMasterNumbering.tblmaster_numberingDataTable
        Dim row As DseMasterNumbering.tblmaster_numberingRow
        Dim DTAItem As New DseMasterNumberingTableAdapters.tblmaster_numbering_itemTableAdapter
        Dim TBLItem As New DseMasterNumbering.tblmaster_numbering_itemDataTable
        Dim rowItem As DseMasterNumbering.tblmaster_numbering_itemRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.FillBy_TipeTransaksi(TBL, vTipeTranskasi)

            If TBL.Rows.Count = 0 Then
                Return ""
            Else
                row = TBL.Rows(0)

                If row.auto_number = 0 Then
                    Return ""
                Else
                    DTAItem.FillBy_primary(TBLItem, row.primary)
                    For Each rowItem In TBLItem.Rows
                        TipeKomponen = rowItem.tipe_komponen
                        Format = rowItem.format

                        If TipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText Then
                            If Dummy = "" Then
                                Dummy = rowItem.value
                            Else
                                Dummy = Dummy + rowItem.value
                            End If
                        Else
                            If Dummy = "" Then
                                Dummy = Me.GetValue(TipeKomponen, Format, vReferenceDate, vReferenceUrutan)
                            Else
                                Dummy = Dummy + Me.GetValue(TipeKomponen, Format, vReferenceDate, vReferenceUrutan)
                            End If
                        End If
                    Next

                    Return Dummy
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Public Function GetJmlRunningNumber(ByVal vTipeTransaksiAsal As ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering) As Integer
        Dim TipeKomp As ClsMasterTransactionNumbering.EnumTipeKomponen = EnumTipeKomponen.RunningNumber
        Dim DTAItem As New DseMasterNumberingTableAdapters.tblmaster_numbering_itemTableAdapter
        Dim TBLItem As New DseMasterNumbering.tblmaster_numbering_itemDataTable
        Dim rowItem As DseMasterNumbering.tblmaster_numbering_itemRow
        Dim Retval As Integer = 0
        Dim Format As ClsMasterTransactionNumbering.EnumFormat = EnumFormat.Running_10Digit

        Try
            DTAItem.FillBy_TipeTransaksi(TBLItem, CInt(vTipeTransaksiAsal))
            For Each rowItem In TBLItem.Rows
                If rowItem.tipe_komponen = TipeKomp Then
                    If rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_10Digit Then
                        Return 10
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_2Digit Then
                        Return 2
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_3Digit Then
                        Return 3
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_4Digit Then
                        Return 4
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_5Digit Then
                        Return 5
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_6Digit Then
                        Return 6
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_7Digit Then
                        Return 7
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_8Digit Then
                        Return 8
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_9Digit Then
                        Return 9
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Public Function GetValue(ByVal vTipeKomponen As ClsMasterTransactionNumbering.EnumTipeKomponen, ByVal vFormat As ClsMasterTransactionNumbering.EnumFormat, _
    ByVal vRefenceDate As Date, ByVal vReferenceUrutan As Double) As String
        Dim Dummy As String = ""
        Dim Lenght As Integer = 0

        Try
            If vTipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText Then
                Return ""
            ElseIf vTipeKomponen = ClsMasterTransactionNumbering.EnumTipeKomponen.Month Then
                If vFormat = ClsMasterTransactionNumbering.EnumFormat.Month_1Digit Then
                    Return vRefenceDate.Month
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Month_2Digit Then
                    Dummy = vRefenceDate.Month
                    If Dummy.Length < 2 Then
                        Dummy = "0" + Dummy
                    End If
                    Return Dummy
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Month_Romawi Then
                    If vRefenceDate.Month = 1 Then
                        Return "I"
                    ElseIf vRefenceDate.Month = 2 Then
                        Return "II"
                    ElseIf vRefenceDate.Month = 3 Then
                        Return "III"
                    ElseIf vRefenceDate.Month = 4 Then
                        Return "IV"
                    ElseIf vRefenceDate.Month = 5 Then
                        Return "V"
                    ElseIf vRefenceDate.Month = 6 Then
                        Return "VI"
                    ElseIf vRefenceDate.Month = 7 Then
                        Return "VII"
                    ElseIf vRefenceDate.Month = 8 Then
                        Return "VIII"
                    ElseIf vRefenceDate.Month = 9 Then
                        Return "IX"
                    ElseIf vRefenceDate.Month = 10 Then
                        Return "X"
                    ElseIf vRefenceDate.Month = 11 Then
                        Return "XI"
                    ElseIf vRefenceDate.Month = 12 Then
                        Return "XII"
                    End If
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Month_StringPanjang_English Then
                    If vRefenceDate.Month = 1 Then
                        Return "January"
                    ElseIf vRefenceDate.Month = 2 Then
                        Return "February"
                    ElseIf vRefenceDate.Month = 3 Then
                        Return "March"
                    ElseIf vRefenceDate.Month = 4 Then
                        Return "April"
                    ElseIf vRefenceDate.Month = 5 Then
                        Return "May"
                    ElseIf vRefenceDate.Month = 6 Then
                        Return "June"
                    ElseIf vRefenceDate.Month = 7 Then
                        Return "July"
                    ElseIf vRefenceDate.Month = 8 Then
                        Return "August"
                    ElseIf vRefenceDate.Month = 9 Then
                        Return "September"
                    ElseIf vRefenceDate.Month = 10 Then
                        Return "October"
                    ElseIf vRefenceDate.Month = 11 Then
                        Return "November"
                    ElseIf vRefenceDate.Month = 12 Then
                        Return "December"
                    End If
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Month_StringPanjang_Indo Then
                    If vRefenceDate.Month = 1 Then
                        Return "Januari"
                    ElseIf vRefenceDate.Month = 2 Then
                        Return "Februari"
                    ElseIf vRefenceDate.Month = 3 Then
                        Return "Maret"
                    ElseIf vRefenceDate.Month = 4 Then
                        Return "April"
                    ElseIf vRefenceDate.Month = 5 Then
                        Return "Mei"
                    ElseIf vRefenceDate.Month = 6 Then
                        Return "Juni"
                    ElseIf vRefenceDate.Month = 7 Then
                        Return "Juli"
                    ElseIf vRefenceDate.Month = 8 Then
                        Return "Agustus"
                    ElseIf vRefenceDate.Month = 9 Then
                        Return "September"
                    ElseIf vRefenceDate.Month = 10 Then
                        Return "Oktober"
                    ElseIf vRefenceDate.Month = 11 Then
                        Return "November"
                    ElseIf vRefenceDate.Month = 12 Then
                        Return "Desember"
                    End If
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Month_StringPendek_English Then
                    If vRefenceDate.Month = 1 Then
                        Return "Jan"
                    ElseIf vRefenceDate.Month = 2 Then
                        Return "Feb"
                    ElseIf vRefenceDate.Month = 3 Then
                        Return "Mar"
                    ElseIf vRefenceDate.Month = 4 Then
                        Return "Apr"
                    ElseIf vRefenceDate.Month = 5 Then
                        Return "May"
                    ElseIf vRefenceDate.Month = 6 Then
                        Return "Jun"
                    ElseIf vRefenceDate.Month = 7 Then
                        Return "Jul"
                    ElseIf vRefenceDate.Month = 8 Then
                        Return "Aug"
                    ElseIf vRefenceDate.Month = 9 Then
                        Return "Sep"
                    ElseIf vRefenceDate.Month = 10 Then
                        Return "Oct"
                    ElseIf vRefenceDate.Month = 11 Then
                        Return "Nov"
                    ElseIf vRefenceDate.Month = 12 Then
                        Return "Dec"
                    End If
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Month_StringPendek_Indo Then
                    If vRefenceDate.Month = 1 Then
                        Return "Jan"
                    ElseIf vRefenceDate.Month = 2 Then
                        Return "Feb"
                    ElseIf vRefenceDate.Month = 3 Then
                        Return "Mar"
                    ElseIf vRefenceDate.Month = 4 Then
                        Return "Apr"
                    ElseIf vRefenceDate.Month = 5 Then
                        Return "Mei"
                    ElseIf vRefenceDate.Month = 6 Then
                        Return "Jun"
                    ElseIf vRefenceDate.Month = 7 Then
                        Return "Jul"
                    ElseIf vRefenceDate.Month = 8 Then
                        Return "Agt"
                    ElseIf vRefenceDate.Month = 9 Then
                        Return "Sep"
                    ElseIf vRefenceDate.Month = 10 Then
                        Return "Okt"
                    ElseIf vRefenceDate.Month = 11 Then
                        Return "Nov"
                    ElseIf vRefenceDate.Month = 12 Then
                        Return "Des"
                    End If
                End If
            ElseIf vTipeKomponen = EnumTipeKomponen.RunningNumber Then
                If vFormat = ClsMasterTransactionNumbering.EnumFormat.Running_10Digit Then
                    Dummy = vReferenceUrutan
                    Lenght = 10
                    While Dummy.Length < Lenght
                        Dummy = "0" + Dummy
                    End While
                    Return Dummy
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Running_2Digit Then
                    Dummy = vReferenceUrutan
                    Lenght = 2
                    While Dummy.Length < Lenght
                        Dummy = "0" + Dummy
                    End While
                    Return Dummy
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Running_3Digit Then
                    Dummy = vReferenceUrutan
                    Lenght = 3
                    While Dummy.Length < Lenght
                        Dummy = "0" + Dummy
                    End While
                    Return Dummy
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Running_4Digit Then
                    Dummy = vReferenceUrutan
                    Lenght = 4
                    While Dummy.Length < Lenght
                        Dummy = "0" + Dummy
                    End While
                    Return Dummy
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Running_5Digit Then
                    Dummy = vReferenceUrutan
                    Lenght = 5
                    While Dummy.Length < Lenght
                        Dummy = "0" + Dummy
                    End While
                    Return Dummy
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Running_6Digit Then
                    Dummy = vReferenceUrutan
                    Lenght = 6
                    While Dummy.Length < Lenght
                        Dummy = "0" + Dummy
                    End While
                    Return Dummy
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Running_7Digit Then
                    Dummy = vReferenceUrutan
                    Lenght = 7
                    While Dummy.Length < Lenght
                        Dummy = "0" + Dummy
                    End While
                    Return Dummy
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Running_8Digit Then
                    Dummy = vReferenceUrutan
                    Lenght = 8
                    While Dummy.Length < Lenght
                        Dummy = "0" + Dummy
                    End While
                    Return Dummy
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Running_9Digit Then
                    Dummy = vReferenceUrutan
                    Lenght = 9
                    While Dummy.Length < Lenght
                        Dummy = "0" + Dummy
                    End While
                    Return Dummy
                End If
            ElseIf vTipeKomponen = EnumTipeKomponen.Year Then
                If vFormat = ClsMasterTransactionNumbering.EnumFormat.Year_Panjang Then
                    Dummy = Format(vRefenceDate, "yyyy")
                ElseIf vFormat = ClsMasterTransactionNumbering.EnumFormat.Year_Pendek Then
                    Dummy = Format(vRefenceDate, "yy")
                End If
                Return Dummy
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetUrutan(ByVal vNoTrans As String, ByVal vTipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering, _
    ByVal vTanggalTransaksi As Date) As Double
        Dim DTAItem As New DseMasterNumberingTableAdapters.tblmaster_numbering_itemTableAdapter
        Dim TBLItem As New DseMasterNumbering.tblmaster_numbering_itemDataTable
        Dim rowItem As DseMasterNumbering.tblmaster_numbering_itemRow
        Dim StartIndex As Integer = 0
        Dim Length As Integer = 0
        Dim Dummy As String = ""

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.FillBy_TipeTransaksi(TBLItem, vTipeTransaksi)

            For Each rowItem In TBLItem.Rows
                If rowItem.tipe_komponen = ClsMasterTransactionNumbering.EnumTipeKomponen.RunningNumber Then
                    If rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_10Digit Then
                        Length = 10
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_2Digit Then
                        Length = 2
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_3Digit Then
                        Length = 3
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_4Digit Then
                        Length = 4
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_5Digit Then
                        Length = 5
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_6Digit Then
                        Length = 6
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_7Digit Then
                        Length = 7
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_8Digit Then
                        Length = 8
                    ElseIf rowItem.format = ClsMasterTransactionNumbering.EnumFormat.Running_9Digit Then
                        Length = 9
                    End If

                    If vNoTrans.Length >= Length + StartIndex Then
                        Dummy = vNoTrans.Substring(StartIndex, Length)

                        If IsNumeric(Dummy) = True Then
                            Return CDbl(Dummy)
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                Else
                    If rowItem.tipe_komponen = ClsMasterTransactionNumbering.EnumTipeKomponen.FixedText Then
                        Dummy = rowItem.value
                    Else
                        Dummy = Me.GetValue(rowItem.tipe_komponen, rowItem.format, vTanggalTransaksi, 1)
                    End If
                    StartIndex += Dummy.Length
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Public Sub New()
        Me.DTAInOUT.Connection.ConnectionString = Core_Module.GetConstring
    End Sub

    Public Function TypeResetAutonumber(ByVal vTipeTransaksi As EnumTipeTransaksiNumbering) As EnumReset
        Dim DTA As New DseMasterNumberingTableAdapters.tblmaster_numberingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTipeResetAutonumber(vTipeTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetTipeResetAutonumber(vTipeTransaksi).Value > 0 Then
                    Return EnumReset.Monthly
                Else
                    Return EnumReset.Yearly
                End If
            Else
                Return EnumReset.Monthly
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function UpdateNumbering(ByVal vTblName As String, ByVal vTransName As String, ByVal vTransType As EnumTipeTransaksiNumbering) As Boolean
        Dim myConnection As New MySql.Data.MySqlClient.MySqlConnection
        Dim myCommand As New MySql.Data.MySqlClient.MySqlCommand
        Dim myObj As New Object

        Try
            myConnection.ConnectionString = Core_Module.GetConstring
            myConnection.Open()
            myCommand.Connection = myConnection

            myCommand.CommandText = "select count(*) as `jml` from " & vTblName
            myObj = myCommand.ExecuteScalar()
            If IsNothing(myObj) = False Then
                If CDec(myObj) > 0 Then
                    Return False
                End If
            End If

            myCommand.CommandText = "select count(*) as `jml` from tblmaster_numbering where `tabel rujukan` = '" & vTblName & "'"
            myObj = myCommand.ExecuteScalar()
            If IsNothing(myObj) = False Then
                If CDec(myObj) > 0 Then
                    Return False
                End If
            End If

            myCommand.CommandText = "INSERT INTO tblmaster_numbering (`deskripsi`, `tipe transaksi`, `auto number`, `reset numbering`, `created by`, `created date`, `modified by`, `modified date`, `timestamp`, `tabel rujukan`) VALUES ('" & vTransName & "', " & vTransType & ", '0', '0', 'USER', '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "', 'USER', '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "', '" & vTblName & "')"
            myCommand.ExecuteNonQuery()

            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            myConnection.Dispose()
            myCommand.Dispose()
            myObj = Nothing
        End Try
    End Function
End Class
