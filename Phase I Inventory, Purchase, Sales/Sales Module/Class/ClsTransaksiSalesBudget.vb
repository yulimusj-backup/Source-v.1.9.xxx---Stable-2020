Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesBudget
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mDocumentNo As String = ""
    Dim mUrutan As String = ""
    Dim mTanggalTransaksi As Date = Now
    Dim mPeriod As Date = Today
    Dim mTipe As Integer = 0
    Dim mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Enum EnumTipeBudget As Integer
        SO
        Invoice
        Invoice_GrossMargin
    End Enum

    Public Property Period() As Date
        Get
            Return Me.mPeriod
        End Get
        Set(ByVal value As Date)
            mPeriod = value
        End Set
    End Property

    Public Property TipeBudget() As EnumTipeBudget
        Get
            Return Me.mTipe
        End Get
        Set(ByVal value As EnumTipeBudget)
            mTipe = CInt(value)
        End Set
    End Property

    Public Property Urutan() As Double
        Get
            Return Me.mUrutan
        End Get
        Set(ByVal value As Double)
            mUrutan = value
        End Set
    End Property

    Public Property DocumentNo() As String
        Get
            Return Me.mDocumentNo
        End Get
        Set(ByVal value As String)
            mDocumentNo = value
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

    Public Property TanggalTransaksi() As Date
        Get
            Return Me.mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
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
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetJml(vNoTransaksi).Value > 0 Then
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

    Public Sub Cetak(ByVal vPrimary As Double, ByVal vFrm As FrmTransaksiSalesBudget_Cetak, ByVal report As CrpCetakTransaksiSalesBudget, ByVal DS As DseTransaksiSalesBudget)
        Dim crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.DataTable1TableAdapter
        Dim DTA1 As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter
        Dim TBL1 As New DseTransaksiSalesBudget.tbltrans_salesbudgetDataTable
        Dim row As DseTransaksiSalesBudget.tbltrans_salesbudgetRow
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            crv = vFrm.crv1
            DTA1.FillBy_Primary(TBL1, vPrimary)
            If TBL1.Rows.Count > 0 Then
                row = TBL1.Rows(0)

                DTA.FillBy_Primary_Cetak(DS.DataTable1, vPrimary)
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_DocNo.ParameterFieldName, row.document_no)
                report.SetParameterValue(report.Parameter_Period.ParameterFieldName, row.period.Year.ToString)
                report.SetParameterValue(report.Parameter_Tanggal.ParameterFieldName, Format(row.tanggal, "dd MMM yyyy"))
                report.SetParameterValue(report.Parameter_Tipe_Budget.ParameterFieldName, row.Tipe_Str)
                report.SetParameterValue(report.Parameter_DefPrec.ParameterFieldName, ObjCur.GetDecimal_DefaultCurrency)

                crv.ReportSource = report
            Else
                crv.ReportSource = Nothing
            End If
            crv.Refresh()
            crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
            DTA.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItem As DataGridView, ByVal vFrm As FrmTransaksiSalesBudget) As Boolean
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter
        Dim TBL As New DseTransaksiSalesBudget.tbltrans_salesbudgetDataTable
        Dim row As DseTransaksiSalesBudget.tbltrans_salesbudgetRow
        Dim ObjItem As New ClsTransaksiSalesBudget_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mDocumentNo = .document_no
                    Me.mPrimary = .primary
                    Me.mTipe = .tipe
                    Me.mUrutan = .urutan
                    Me.mRemarks = .remarks
                    Me.mTanggalTransaksi = .tanggal
                    Me.mPeriod = .period

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
                'salin data item 
                ObjItem.IsiDGV(vDgvItem, Me.mPrimary, vFrm)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, ByVal vStrictEdit As Boolean, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter
        Dim vCreatedBy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'tidak boleh edit jika strict edit dan user codenya tidak cocok

            If vStrictEdit = True Then
                vCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If vCreatedBy <> vUserCode Then
                    MsgBox("Only user: " + vCreatedBy + " can edit this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, ByVal vStrictDelete As Boolean, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter
        Dim vCreatedBy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'tidak boleh edit jika strict edit dan user codenya tidak cocok

            If vStrictDelete = True Then
                vCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If vCreatedBy <> vUserCode Then
                    MsgBox("Only user: " + vCreatedBy + " can delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vTipeBudget As EnumTipeBudget, ByVal vPeriod As Date) As Double
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJml2(CInt(vTipeBudget), vPeriod)

            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vDocumentNo As String) As Double
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter
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

    Public Function GetPrimary(ByVal vTipeBudget As EnumTipeBudget, ByVal vPeriod As Date) As Double
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimary1(CInt(vTipeBudget), vPeriod)

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

    Public Function GetPrimary(ByVal vTipeBudget As EnumTipeBudget, ByVal vPeriod As Date, ByVal vPrimaryTransaksiException As Double) As Double
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimary2(CInt(vTipeBudget), vPeriod, vPrimaryTransaksiException)

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

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetUrutanMax(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetUrutanMax(TanggalEnd, TanggalStart).Value + 1)
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
            Return NoUrut + "/" + "SB" + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetNoTransaksi(vPrimary)

            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vDGVItem As DataGridView, ByVal vFrm As FrmTransaksiSalesBudget, _
    ByVal vPeriod As Date, ByVal vTipeBudget As EnumTipeBudget, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim vPrimaryAcuan As Double = -1
        Dim rowDGV As DataGridViewRow
        Dim ListSalesPName As String = ""
        Dim DummyInt As Integer = 0

        Try
            'msgbox("belum")
            'data salah bila:
            '1. tidak ada 1 pun item 
            If vDGVItem.RowCount = 0 Then
                MsgBox("Please specify at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            '2. sudah ada data yang sama period dan tipe budgetnya
            If vBaru = True Then
                If Me.GetJml(vTipeBudget, vPeriod) > 0 Then
                    MsgBox("Data duplication, the transaction: " + Me.GetNoTransaksi(Me.GetPrimary(CInt(vTipeBudget), vPeriod)) + _
                    " is also " + Me.GetTipeStr(vTipeBudget) + " data. You cannot save this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                vPrimaryAcuan = Me.GetPrimary(vTipeBudget, vPeriod, vPrimaryTransaksi)
                If vPrimaryAcuan <> -1 Then
                    MsgBox("Data duplication, the transaction: " + Me.GetNoTransaksi(vPrimaryAcuan) + _
                    " is also " + Me.GetTipeStr(vTipeBudget) + " data. You cannot save this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '3. batasan untuk no transaksi
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mDocumentNo) = True Then
                    Me.mDocumentNo = Me.GetNoBaru(Me.mTanggalTransaksi)
                End If
            Else
                If Me.GetPrimary(Me.mDocumentNo) <> Me.mPrimary Then
                    Me.mDocumentNo = Me.GetNoBaru(Me.mTanggalTransaksi)
                End If
            End If

            '4. untuk setiap baris, minimal ada 1 bulan yang tidak nol
            For Each rowDGV In vDGVItem.Rows
                If CDec(rowDGV.Cells(vFrm.ColJan.Index).Value) = 0 And CDec(rowDGV.Cells(vFrm.ColFeb.Index).Value) = 0 And _
                CDec(rowDGV.Cells(vFrm.ColMar.Index).Value) = 0 And CDec(rowDGV.Cells(vFrm.ColApr.Index).Value) = 0 And _
                CDec(rowDGV.Cells(vFrm.ColMay.Index).Value) = 0 And CDec(rowDGV.Cells(vFrm.ColJun.Index).Value) = 0 And _
                CDec(rowDGV.Cells(vFrm.ColJul.Index).Value) = 0 And CDec(rowDGV.Cells(vFrm.ColAug.Index).Value) = 0 And _
                CDec(rowDGV.Cells(vFrm.ColSep.Index).Value) = 0 And CDec(rowDGV.Cells(vFrm.ColOct.Index).Value) = 0 And _
                CDec(rowDGV.Cells(vFrm.ColNov.Index).Value) = 0 And CDec(rowDGV.Cells(vFrm.ColDec.Index).Value) = 0 Then
                    DummyInt += 1
                    If ListSalesPName = "" Then
                        ListSalesPName = DummyInt.ToString + "." + rowDGV.Cells(vFrm.ColSalesPersonName.Index).Value.ToString
                    Else
                        ListSalesPName = ListSalesPName + vbCrLf + DummyInt.ToString + "." + rowDGV.Cells(vFrm.ColSalesPersonName.Index).Value.ToString
                    End If
                End If
            Next
            If ListSalesPName <> "" Then
                MsgBox("The following sales person(s) has invalid data: " + vbCrLf + ListSalesPName, MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
ByVal vFrm As FrmTransaksiSalesBudget) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
           Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_SalesBudget_SimpanBaru, vUserCode)

            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()

            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vdgvItem, vUserCode, vFrm)
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
            SW = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru1(ByVal vdgvItem As DataGridView, _
    ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesBudget) As Boolean
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter
        Dim ObjItem As New ClsTransaksiSalesBudget_Item

        Try
            If Me.ApakahDataBenar(True, vdgvItem, vFrm, Me.mPeriod, Me.mTipe, Me.mPrimary) = True Then
                Me.mUrutan = Me.mDocumentNo.Substring(0, 6)
                'insert query
                If DTA.Insert(Me.mDocumentNo, Me.mUrutan, Me.mTanggalTransaksi, _
                Me.mPeriod, Me.mRemarks, Me.mTipe, vUserCode, _
                Now, vUserCode, Now, Now, Me.GetTipeStr(Me.mTipe)) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mDocumentNo).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, Me.mPrimary, vFrm)

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

    Private Function GetTipeStr(ByVal vTipe As EnumTipeBudget) As String
        Try
            If vTipe = EnumTipeBudget.Invoice Then
                Return "Invoice Budget"
            ElseIf vTipe = EnumTipeBudget.Invoice_GrossMargin Then
                Return "Gross Margin Budget"
            ElseIf vTipe = EnumTipeBudget.SO Then
                Return "SO Budget"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesBudget) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
             Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_SalesBudget_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(dgvItem, vUserCode, vFrm)

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
            SW = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesBudget) As Boolean
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter
        Dim ObjItem As New ClsTransaksiSalesBudget_Item

        Try
            If Me.ApakahDataBenar(False, dgvItem, vFrm, Me.mPeriod, Me.mTipe, Me.mPrimary) = True Then
                Me.mUrutan = Me.DocumentNo.Substring(0, 6)
                If DTA.Update(Me.mDocumentNo, Me.mUrutan, Me.mTanggalTransaksi, Me.mPeriod, _
                Me.mRemarks, Me.mTipe, vUserCode, Now, Now, Me.GetTipeStr(Me.mTipe), Me.mPrimary, _
                Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mPrimary, vFrm)
                    'nilai balikan
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

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vStrictDelete As Boolean, ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, vPrimaryTransaksi, _
               Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_SalesBudget_Hapus, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, vStrictDelete, vUserCode)
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
            SW = Nothing
        End Try
    End Function

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, ByVal vStrictDelete As Boolean, ByVal vUserCode As String) As Boolean
        Dim Obj As New ClsTransaksiSalesBudget_Item
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudgetTableAdapter

        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, vStrictDelete, vUserCode) = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudget1TableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                RetVal = DTA.ScalarQuery_GetJmlByDocNoLIKE(Criteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                RetVal = DTA.ScalarQuery_GetJmlByTgl(vTanggalSampai, vTanggalDari)
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                RetVal = DTA.ScalarQuery_GetJmlByRemarksLIKE(Criteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonCode Then
                RetVal = DTA.ScalarQuery_GetJmlBySalesPCodeLIKE(Criteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonName Then
                RetVal = DTA.ScalarQuery_GetJmlBySalesPName(Criteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSalesBudget.tbltrans_salesbudget1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudget1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_DocNoLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_Tanggal(TBL, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonCode Then
                QueryResult = DTA.FillBy_SalesPersonCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonName Then
                QueryResult = DTA.FillBy_SalesPersonNameLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(3).DefaultCellStyle.Format = "yyyy"
            End If

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
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now

            Me.mPrimary = -1
            Me.mRemarks = ""
            Me.mTanggalTransaksi = Now
            Me.mPeriod = Today
            Me.mTimestamp = Now
            Me.mTipe = 0

            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
