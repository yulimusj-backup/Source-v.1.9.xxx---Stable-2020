Option Explicit On
Option Strict Off

Public Class ClsTrans_TaxRecalculate
    Dim mPrimary As Double = -1
    Dim mDocNo As String = ""
    Dim mUrutan As Integer = -1
    Dim mDate As Date = #1/1/1950#
    Dim mPeriod As Date = #1/1/1950#
    Dim mDifference As Decimal = 0
    Dim mRemarks As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now

    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

#Region "Property"
    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property DocNo() As String
        Get
            Return mDocNo
        End Get
        Set(ByVal value As String)
            mDocNo = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Property Urutan() As Integer
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Integer)
            mUrutan = value
        End Set
    End Property

    Public Property Date_() As Date
        Get
            Return mDate
        End Get
        Set(ByVal value As Date)
            mDate = value
        End Set
    End Property

    Public Property Period() As Date
        Get
            Return Me.mPeriod
        End Get
        Set(ByVal value As Date)
            mPeriod = value
        End Set
    End Property

    Public Property Difference() As Decimal
        Get
            Return Me.mDifference
        End Get
        Set(ByVal value As Decimal)
            mDifference = value
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
#End Region

    Public Function ApakahBolehedit(ByVal vUserCode As String, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            '1.strict edit
            Dim UserCode As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If UserCode <> vUserCode Then
                MsgBox("Only user with ID:" + UserCode + " can edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vUserCode As String, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            '1.strict edit
            Dim UserCode As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If UserCode <> vUserCode Then
                MsgBox("Only user with ID:" + UserCode + " can delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryByPeriod(ByVal vPeriod As Date) As Double
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryByPeriod(vPeriod)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function ApakahDataBenar(ByVal FrmTR As FrmTrans_TaxRecalculate, ByVal vBaru As Boolean) As Boolean
        Dim PrimaryRef As Double = -1

        Try
            '1.tidak boleh double
            If Me.ApakahSudahAda(Me.DocNo) = True Then
                If vBaru = True Then
                    Me.DocNo = Me.GetNoBaru(Me.mDate)
                Else
                    If Me.GetPrimary(Me.DocNo) <> Me.mPrimary Then
                        Me.DocNo = Me.GetNoBaru(Me.mDate)
                    End If
                End If
            End If

            '2.Minimal ada 1 item
            If FrmTR.dgvItem.RowCount = 0 Then
                MsgBox("Please specify at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'periode kompensasi tidak boleh desember
            If Me.Period.Month = 12 Then
                MsgBox("We can't recalculate for december.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1 periode kompensasi hanya bisa di kompensasi 1x
            PrimaryRef = Me.GetPrimaryByPeriod(Me.mPeriod)
            If PrimaryRef <> -1 Then
                If Me.mPrimary <> PrimaryRef Then
                    MsgBox("There's tax recalculation for period " + Format(Me.mPeriod, "MMM yyyy") + " process stoped.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim JmlDb As Nullable(Of Double)
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            JmlDb = DTA.ScalarQuery_GetMaxUrutan(TanggalEnd, TanggalStart)
            If JmlDb.HasValue = True Then
                NoUrut = CStr(CInt(JmlDb.Value)) + 1
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            'pastikan no urut terdiri atas 6 huruf
            While NoUrut.Length < 6
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return NoUrut + "/" + "TR" + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmTrans_TaxRecalculate, _
   ByVal vDgvItem As Windows.Forms.DataGridView, ByVal vTBLItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable) As Boolean
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter
        Dim TBL As New DseTrans_Tax_Recalculate.tbltrans_tax_recalculateDataTable
        Dim row As DseTrans_Tax_Recalculate.tbltrans_tax_recalculateRow
        Dim ObjItem As New ClsTrans_TaxRecalculateItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mDocNo = .no_transaksi
                    Me.mUrutan = .urutan
                    Me.mDate = .tanggal_transaksi
                    Me.mPeriod = .period
                    Me.mDifference = .Difference
                    Me.mRemarks = .remarks
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp

                    ObjItem.IsiDGVItem(vDgvItem, .primary, vFrm, vTBLItemTax)
                End With
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vDocNo As String) As Boolean
        Dim dta As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter
        Dim dummy As Nullable(Of Long)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJmlByNoTrans(vDocNo)
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

    Public Function GetDocNoByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetNoTransksi(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTimeStampByPrimary(ByVal vPrimary As Double) As Date
        Dim dta As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetTimestamp(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return "1/1/1950"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vDGV As Windows.Forms.DataGridView, _
    ByVal frm As FrmTrans_TaxRecalculate, ByVal vTblItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_TaxRecalculate_SimpanBaru, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If CInt(Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value) = CInt(Me.vPrimary_Antrian) Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, vDGV, frm, vTblItemTax)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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
            SW = Nothing
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vDGV As Windows.Forms.DataGridView, ByVal frm As FrmTrans_TaxRecalculate, _
    ByVal vTBLItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable) As Boolean
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter
        Dim ObjItem As New ClsTrans_TaxRecalculateItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(frm, True) = True Then
                Me.mUrutan = Me.mDocNo.Substring(0, 6)

                If DTA.Insert(Me.mDocNo, Me.mUrutan, Me.mDate, Me.mPeriod, Me.mRemarks, vUserCode, Now, vUserCode, _
                 Now, Now, Me.mDifference) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mDocNo)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan item
                    ObjItem.SimpanItem(Me.Primary, vDGV, frm, vTBLItemTax)

                    'update ke tabel salary
                    Dim ij As New FrmTrans_TaxRecalculate_Update
                    ij.TipeTransaksi = FrmTrans_TaxRecalculate_Update.EnumTipeTransaksi.SimpanBaru
                    ij.PrimaryTransaksi = Me.mPrimary
                    ij.vFrm = frm
                    ij.ShowDialog()
                    ij = Nothing

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vDGV As Windows.Forms.DataGridView, _
   ByVal frm As FrmTrans_TaxRecalculate, ByVal vTBLItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_TaxRecalculate_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If CInt(Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value) = CInt(Me.vPrimary_Antrian) Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode, vDGV, frm, vTBLItemTax)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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
            SW = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal vDGV As Windows.Forms.DataGridView, ByVal frm As FrmTrans_TaxRecalculate, _
    ByVal vTBLItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable) As Boolean
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter
        Dim ObjItem As New ClsTrans_TaxRecalculateItem
        Dim ListSalary As New List(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(frm, False) = True Then
                Me.mUrutan = Me.mDocNo.Substring(0, 6)

                If DTA.Update(Me.mDocNo, Me.mUrutan, Me.mDate, Me.mPeriod, Me.mRemarks, vUserCode, Now, Now, _
                Me.mDifference, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus item
                    ObjItem.HapusItem(Me.mPrimary)

                    'simpan item
                    ObjItem.SimpanItem(Me.Primary, vDGV, frm, vTBLItemTax)
                    Dim ij As New FrmTrans_TaxRecalculate_Update
                    ij.TipeTransaksi = FrmTrans_TaxRecalculate_Update.EnumTipeTransaksi.SimpanPerubahan
                    ij.PrimaryTransaksi = Me.mPrimary
                    ij.vFrm = frm
                    ij.ShowDialog()
                    ij = Nothing

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ListSalary = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, ByVal frm As FrmTrans_TaxRecalculate) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_TaxRecalculate_Hapus, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If CInt(Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value) = CInt(Me.vPrimary_Antrian) Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, vUserCode, frm)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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
            SW = Nothing
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal frm As FrmTrans_TaxRecalculate) As Boolean
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter
        Dim IsBeginingBalance As Boolean = False
        Dim ListPrimarySalary As New List(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vUserCode, vPrimary) = True Then
                'kembalikan ke kondisi semua
                Dim ij As New FrmTrans_TaxRecalculate_Update
                ij.TipeTransaksi = FrmTrans_TaxRecalculate_Update.EnumTipeTransaksi.Hapus
                ij.PrimaryTransaksi = Me.mPrimary
                ij.vFrm = frm
                ij.ShowDialog()
                ij = Nothing

                'hapus item
                DTA.Delete(vPrimary)

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ListPrimarySalary = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, ByVal vDGV As Windows.Forms.DataGridView, ByVal vDTA As DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculate1TableAdapter, _
    ByVal vTBL As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate1DataTable, ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As Windows.Forms.ToolStripLabel, ByVal vTanggalDari As Date, _
    ByVal vTanggalSampai As Date, ByVal vFrm As FrmTrans_TaxRecalculate_Cari, ByVal ObjCD As Core_Module.ClsCekData) As Boolean
        Dim QueryResult As Double
        Dim StartT As Date
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Dim PayrollConventionDate As Date = ObjClosing.GetPayrollConvertionDate()

        Try
            vDTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
            vDGV.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = vDTA.FillBy_NoTransaksiLIKE(vTBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDate Then
                QueryResult = vDTA.FillBy_TglSampaiDari(vTBL, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                QueryResult = vDTA.FillBy_RemarksLIKE(vTBL, Criteria + "%")
            End If

            vDGV.DataSource = vTBL.DefaultView
            vDGV.Refresh()
            If vDGV.ColumnCount > 0 Then
                vDGV.Columns(0).Visible = False
                vDGV.Columns(0).Frozen = True
                'vDGV.Columns(1).Frozen = True
                vDGV.Columns(1).DefaultCellStyle.Format = "MMM yyyy"
                vDGV.Columns(3).DefaultCellStyle.Format = "dd MMM yyyy"
                vDGV.Columns(4).DefaultCellStyle.Format = "N0"
                vDGV.Columns(4).DefaultCellStyle.Alignment = Windows.Forms.DataGridViewContentAlignment.MiddleRight
            End If

            'update vDGV
            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If
            vDGV.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.Arrow
            ObjClosing = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vCriteria As String, ByVal vDGV As Windows.Forms.DataGridView, ByVal vSearchCriteria As Core_Module.EnumSearchCriteria, _
    ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date) As Double
        Dim Dta As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculate1TableAdapter
        Dim Dummy As Nullable(Of Long)

        Try

            Dta.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchCriteria = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                Dummy = Dta.ScalarQuery_GetJmlByNoTrans(vCriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = Core_Module.EnumSearchCriteria.ByTanggalTransaksi Then
                Dummy = Dta.ScalarQuery_GetJmlByTglTransaksi(vTanggalSampai, vTanggalDari)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                Dummy = Dta.ScalarQuery_GetJmlByRemarks(vCriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vDocumentNo As String) As Double
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimary(vDocumentNo)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksi_ByPrimarySalary(ByVal vPrimarySalary As Double) As String
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculateTableAdapter
        Dim RetVal As String

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetNoTransaksi1(vPrimarySalary)
            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub Cetak(ByVal vPrimaryLoan As Double, ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
