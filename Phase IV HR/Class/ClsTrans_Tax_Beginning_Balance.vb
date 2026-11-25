Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class ClsTrans_Tax_Beginning_Balance
    Private mPrimary As Double = -1
    Private mDocumentNo As String = ""
    Private mUrutan As Integer = 0
    Private mTanggalTransaksi As Date = #1/1/1950#
    Private mPeriod As Date = #1/1/1950#
    Private mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Property Urutan() As Integer
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Integer)
            mUrutan = value
        End Set
    End Property

    Public Property DocumentNo() As String
        Get
            Return mDocumentNo
        End Get
        Set(ByVal value As String)
            mDocumentNo = value
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

    Public Property TanggalTransaksi() As Date
        Get
            Return mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property

    Public Property Period() As Date
        Get
            Return mPeriod
        End Get
        Set(ByVal value As Date)
            mPeriod = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
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

    Public Enum GolonganPajak As Integer
        TK0
        TK1
        TK2
        TK3
        K0
        K1
        K2
        K3
    End Enum

    Public Sub CetakTaxBalance(ByVal vPrimary As Double, ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer)
        Dim Report As New CrpCetakTransTaxBeginning
        Dim DS As New DseTrans_Tax_Beginning_Balance
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.TblCetakTaxTableAdapter
        Dim TBL As DseTrans_Tax_Beginning_Balance.TblCetakTaxDataTable = DS.TblCetakTax
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryTax(TBL, vPrimary)

            Report.SetDataSource(DS)
            vCrv.ReportSource = Report
            vCrv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDocumentNo(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetDocumentNo(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByDocumentNo(ByVal vDocumentNo As String) As Double
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByDocumentNo(vDocumentNo)
            If Dummy.hasvalue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vDocumentNo As String) As Double
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimary(vDocumentNo)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTanggalTransaksi(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTimestamp(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetMaxUrutan(ByVal vStartDate As Date, ByVal vEndDate As Date) As Double
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim Dummy As System.Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetMaxUrutan(vEndDate, vStartDate)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    '#Region "Standar Coding"
    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vObjCD As Core_Module.ClsCekData) As String
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim InOut As String = "TBB"
        NoUrut = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = vObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = vObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetMaxUrutan(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetMaxUrutan(TanggalEnd, TanggalStart).Value + 1)
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

            Return NoUrut + "/" + InOut + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vObjCD As Core_Module.ClsCekData, _
        ByVal vDGVItem As DataGridView, ByVal vDef_Prec As Double, ByVal i As FrmTrans_BeginingBalanceTax) As Boolean
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim TBL As New DseTrans_Tax_Beginning_Balance.tbltrans_tax_beginning_balanceDataTable
        Dim row As DseTrans_Tax_Beginning_Balance.tbltrans_tax_beginning_balanceRow
        Dim ObjItem As New ClsTrans_Tax_Beginning_Balance_Item
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mDocumentNo = .Document_No
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    Me.mUrutan = .Urutan
                    Me.mPeriod = .Period
                    Me.mRemarks = .Remarks

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'cari data item
                ObjItem.IsiDGV(Primary, vDGVItem, vDef_Prec, i)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As Core_Module.EnumSearchCriteria, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Double
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balance1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByDocumentNoLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByCode Then
                Return DTA.ScalarQuery_GetJmlByEmployeeCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByDate Then
                Return DTA.ScalarQuery_GetJmlByTanggalDariSampai(vEndDate, vStartDate).Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByFullName Then
                Return DTA.ScalarQuery_GetJmlByEmployeeNameLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByEmployeeNIK_LIKE(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTrans_Tax_Beginning_Balance.tbltrans_tax_beginning_balance1DataTable, _
    ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Boolean
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balance1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDate Then
                QueryResult = DTA.FillBy_TanggalDariSampai(TBL, vEndDate, vStartDate)
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_EmployeeCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_Remarks_LIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByFullName Then
                QueryResult = DTA.FillBy_EmployeeNameLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_EmployeeNIK_LIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(3).DefaultCellStyle.Format = "MMM yyyy"
                'DGV.Columns(2).HeaderText = "Description"
                'DGV.Columns(3).HeaderText = "City"
                'DGV.Columns(4).HeaderText = "Remarks"
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > TimeSpan.TicksPerSecond Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, _
                    0, MidpointRounding.AwayFromZero)) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vPrimary As Double, ByVal vdgvItem As DataGridView, ByVal ij As FrmTrans_BeginingBalanceTax) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Tax_Beginning_Balance_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, vObjCD, vPrimary, vdgvItem, ij)
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

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vPrimary As Double, ByVal vdgvItem As DataGridView, ByVal ij As FrmTrans_BeginingBalanceTax) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Tax_Beginning_Balance_SimpanPerubahan, vUserCode)

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
                            retval = Me.SimpanPerubahan1(vUserCode, vObjCD, vPrimary, vdgvItem, ij)
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
        End Try
    End Function

    Public Function Hapus(ByVal vUserCode As String, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vPrimary As Double, ByVal vdgvItem As DataGridView, ByVal ij As FrmTrans_BeginingBalanceTax, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, vPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Tax_Beginning_Balance_Hapus, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > ObjAntri.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If CInt(Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value) = CInt(Me.vPrimary_Antrian) Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimary, vUserCode, True, TampilkanPesan, vdgvItem, ij)
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

    Public Function SimpanBaru1(ByVal vUserCode As String, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vPrimary As Double, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_BeginingBalanceTax) As Boolean
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim ObjItem As New ClsTrans_Tax_Beginning_Balance_Item
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, vPrimary, vObjCD, i) = True Then
                Me.mUrutan = CDbl(Me.mDocumentNo.Substring(0, 6))
                'Insert Data
                If DTA.Insert(Me.mDocumentNo, Me.mTanggalTransaksi, Me.mPeriod, Me.mRemarks, _
                vUserCode, Now, vUserCode, Now, Now, Me.mUrutan) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mDocumentNo)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, vDGVItem, i)

                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

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
            ObjItem = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUserCode As String, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vPrimary As Double, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_BeginingBalanceTax) As Boolean
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim ObjItem As New ClsTrans_Tax_Beginning_Balance_Item
        Try
            If Me.ApakahDataBenar(False, vPrimary, vObjCD, i) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                Me.mUrutan = CDbl(Me.mDocumentNo.Substring(0, 6))
                If DTA.Update(Me.mDocumentNo, Me.mTanggalTransaksi, Me.mPeriod, Me.mRemarks, vUserCode, Now, Now, _
                Me.mUrutan, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, vDGVItem, i)

                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

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

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_BeginingBalanceTax) As Boolean
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim ObjItem As New ClsTrans_Tax_Beginning_Balance_Item
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, vStrictDelete, vShowMsg, i) = True Then
                'Hapus item
                ObjItem.HapusItem(vPrimary)
                'Hapus
                DTA.Delete(vPrimary)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mDocumentNo = ""
            mTanggalTransaksi = #1/1/1950#
            mUrutan = 0
            mPeriod = #1/1/1950#
            mRemarks = ""

            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today

            'Item
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, _
      ByVal vStrictEdit As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim pCreatedBy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictEdit = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If

                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        MsgBox("User can edit this data: " + pCreatedBy, _
                        MsgBoxStyle.Information, "Information")
                    End If
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

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, ByVal i As FrmTrans_BeginingBalanceTax) As Boolean
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim DTASalary As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_salaryTableAdapter
        Dim TBLSalary As New DseTrans_Tax_Beginning_Balance.tbltrans_salaryDataTable
        Dim pCreatedBy As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjItem As New ClsTrans_Tax_Beginning_Balance_Item
        Dim InfoData As String = ""
        Dim Nomor As Integer = 1
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Dim ConventionDate_Dari As String = "#1/1/1950#"
        Dim ConventionDate_Sampai As String = "1/1/1950#"
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictDelete = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If
                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        MsgBox("User that can delete this data is: " + pCreatedBy, _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
            '2. tidak boleh dihapus jika salah satu karyawan yang telah diinput, punya transaksi salary di bulan 12
            If i.dgvItem.Rows.Count > 0 Then
                DTASalary.Fill(TBLSalary)
                ConventionDate_Sampai = Format(ObjCD.GetLastDateOfTheYear(ObjClosing.GetPayrollConvertionDate()), "MM/dd/yyyy").ToString
                ConventionDate_Dari = Format(ObjCD.GetFirstDateofTheMonth(ConventionDate_Dari), "MM/dd/yyyy").ToString
                i.dgvItem.AllowUserToAddRows = False
                For Each RowDGV In i.dgvItem.Rows
                    Filter = "[Primary Employee] = " + RowDGV.Cells(i.ColPrimaryKaryawan.Index).Value.ToString _
                     + " AND [Tanggal Transaksi] >= #" + ConventionDate_Dari.ToString _
                     + "# AND [Tanggal Transaksi] <= #" + ConventionDate_Sampai.ToString + "#"
                    TBLSalary.DefaultView.RowFilter = Filter
                    If TBLSalary.DefaultView.Count > 0 Then
                        If String.IsNullOrEmpty(InfoData) = True Then
                            InfoData = Nomor.ToString + " " + RowDGV.Cells(i.ColNamaKaryawan.Index).Value.ToString
                        Else
                            InfoData += vbNewLine + Nomor.ToString + " " + RowDGV.Cells(i.ColNamaKaryawan.Index).Value.ToString
                        End If
                        Nomor += 1
                    End If
                Next
                i.dgvItem.AllowUserToAddRows = True
                If String.IsNullOrEmpty(InfoData) = False Then
                    MsgBox("The following employee(s) have salary calculation in December" + CDate(ConventionDate_Dari).Year.ToString + " :" _
                    + vbNewLine + InfoData, MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTASalary.Dispose()
            TBLSalary.Dispose()
            ObjItem = Nothing
            ObjClosing = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vPrimary As Double, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal i As FrmTrans_BeginingBalanceTax) As Boolean
        Dim ObjItem As New ClsTrans_Tax_Beginning_Balance_Item
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Dim CoventionDate As Date = #1/1/1950#
        Try
            i.dgvItem.AllowUserToAddRows = False

            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mDocumentNo) = True Then
                    Me.mDocumentNo = Me.GetNoBaru(Me.mTanggalTransaksi, vObjCD)
                End If
            Else
                Dim NoTransaksiLama As String = Me.GetDocumentNo(Me.mPrimary)
                If NoTransaksiLama <> Me.mDocumentNo Then
                    If Me.ApakahSudahAda(Me.mDocumentNo) = True Then
                        Me.mDocumentNo = Me.GetNoBaru(Me.mDocumentNo, vObjCD)
                    End If
                End If
            End If

            If i.dgvItem.Rows.Count <= 0 Then
                MsgBox("Please input item data.", MsgBoxStyle.Information, "Infomation")
                i.dgvItem.AllowUserToAddRows = True
                Return False
            End If

            If String.IsNullOrEmpty(i.TxtRemarks.Text) = False Then
                If ObjCD.CekDataKosong(i.TxtRemarks.Text, False) = False Then
                    MsgBox("Data cannot contains character : ' , % , or * .", MsgBoxStyle.Information, "Information")
                    i.dgvItem.AllowUserToAddRows = True
                    Return False
                End If
            End If

            CoventionDate = ObjClosing.GetPayrollConvertionDate()
            If CoventionDate.Year > Me.mPeriod.Year Then
                MsgBox("Please make transaction period on " + CoventionDate.Month.ToString + CoventionDate.Year.ToString, MsgBoxStyle.Information, "Information")
                i.dgvItem.AllowUserToAddRows = True
                Return False
            End If

            If Me.mPeriod.Month = 12 Then
                MsgBox("No tax beginning balance is needed.", MsgBoxStyle.Information, "Information")
                i.dgvItem.AllowUserToAddRows = True
                Return False
            End If

            'If Me.mPeriod.Year <> CoventionDate.Year Then
            '    'MsgBox("Please make transaction period  " + CoventionDate.Year.ToString, MsgBoxStyle.Information, "Information")
            '    'i.dgvItem.AllowUserToAddRows = True
            '    'Return False
            '    If MsgBox("Payroll convertion date is: " + CoventionDate.Year.ToString + ", continue anyway?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
            '        i.dgvItem.AllowUserToAddRows = True
            '        Return False
            '    End If
            'End If

            'If Me.mPeriod.Month = 12 Then
            '    If Me.mPeriod.Year <> CoventionDate.Year Then
            '        'MsgBox("Maximum period is november " + CoventionDate.Year.ToString, MsgBoxStyle.Information, "Information")
            '        MsgBox("You can not enter the data outside of " + CoventionDate.Year.ToString + _
            '        " and maximum period is november " + CoventionDate.Year.ToString, MsgBoxStyle.Information, "Information")
            '        i.dgvItem.AllowUserToAddRows = True
            '        Return False
            '    End If
            'Else
            '    If Me.mPeriod.Year <> CoventionDate.Year Then
            '        MsgBox("You can not enter the data outside of " + CoventionDate.Year.ToString, MsgBoxStyle.Information, "Information")
            '        i.dgvItem.AllowUserToAddRows = True
            '        Return False
            '    End If
            'End If

            Dim StringKet As String = Me.Cek2TransBeginningBalance(vBaru, Me.mPeriod, vObjCD, i)
            If String.IsNullOrEmpty(StringKet) = False Then
                MsgBox("The following employee(s) have tax beginning balance calculation: " + vbNewLine + StringKet, MsgBoxStyle.Information, "Information")
                i.dgvItem.AllowUserToAddRows = True
                Return False
            End If


            i.dgvItem.AllowUserToAddRows = True
            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjItem = Nothing
            ObjCD = Nothing
            ObjClosing = Nothing
        End Try
    End Function

    Public Function Cek2TransBeginningBalance(ByVal vBaru As Boolean, ByVal vPeriodDate As Date, ByVal vObjCD As Core_Module.ClsCekData, _
     ByVal i As FrmTrans_BeginingBalanceTax) As String
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.DataTable1_1TableAdapter
        Dim TBL As New DseTrans_Tax_Beginning_Balance.DataTable1_1DataTable
        Dim Row As DseTrans_Tax_Beginning_Balance.DataTable1_1Row
        Dim EndDate As Date = vObjCD.GetLasDateOfTheMonth(vPeriodDate)
        Dim StartDate As Date = vObjCD.GetFirstDateofTheMonth(EndDate)
        Dim TBL_A As New DseTrans_Tax_Beginning_Balance.DataTable1_2DataTable
        Dim Row_A As DseTrans_Tax_Beginning_Balance.DataTable1_2Row
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim Filter_MB As String = ""
        Dim Retval As String = ""
        Dim Nomor As Integer = 1
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vBaru = True Then
                DTA.Fill(TBL, EndDate, StartDate)
            Else
                DTA.FillBy_TidakSamaDenganPrimaryEdit(TBL, EndDate, StartDate, Me.mPrimary)
            End If
            If i.dgvItem.Rows.Count > 0 Then
                For Each RowDGV In i.dgvItem.Rows
                    Filter = "[Primary karyawan] = " + RowDGV.Cells(i.ColPrimaryKaryawan.Index).Value.ToString + ""
                    TBL.DefaultView.RowFilter = Filter
                    If TBL.DefaultView.Count > 0 Then
                        Row = TBL.DefaultView(0).Row
                        Filter_MB = "[No Trans] = '" + Row.Document_No.ToString + "'"
                        TBL_A.DefaultView.RowFilter = Filter_MB
                        If TBL_A.DefaultView.Count > 0 Then
                            Row_A = TBL_A.DefaultView(0).Row
                            Row_A.Nama = Row.Full_Name
                        Else
                            Row_A = TBL_A.NewDataTable1_2Row
                            Row_A.No_Trans = Row.Document_No
                            Row_A.Nama = Row.Full_Name
                            TBL_A.Rows.Add(Row_A)
                        End If
                    End If
                Next
                For Each Row_A In TBL_A.Rows
                    If Retval = "" Then
                        Retval = Nomor.ToString + ". " + Row_A.No_Trans.ToString + " : " + Row_A.Nama.ToString
                    Else
                        Retval += vbNewLine + Nomor.ToString + ". " + Row_A.No_Trans.ToString + " : " + Row_A.Nama.ToString
                    End If
                    Nomor += 1
                Next
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
            TBL_A.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vDocumentNo As String) As Boolean
        Try
            If Me.GetJmlByDocumentNo(vDocumentNo) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    '#End Region

    Public Function HitungPKP(ByVal vNilaiPKPSetahun As Decimal) As Decimal
        Dim DTAPersen As New DseTrans_Tax_Beginning_BalanceTableAdapters.tblmaster_tarif_pphTableAdapter
        Dim TBLPersen As New DseTrans_Tax_Beginning_Balance.tblmaster_tarif_pphDataTable
        Dim RowPersen_After As DseTrans_Tax_Beginning_Balance.tblmaster_tarif_pphRow
        Dim RowPersen_Before As DseTrans_Tax_Beginning_Balance.tblmaster_tarif_pphRow
        Dim Retval As Decimal = 0
        Dim Calculate As Decimal = 0
        Dim NilaiSisa As Decimal = 0
        Try
            DTAPersen.Connection.ConnectionString = Core_Module.GetConstring
            DTAPersen.Fill(TBLPersen)
            If TBLPersen.Rows.Count > 0 Then
                For ij As Integer = 0 To TBLPersen.Count - 1
                    RowPersen_After = TBLPersen.Rows(ij)
                    If ij = 0 Then
                        Calculate = vNilaiPKPSetahun - RowPersen_After.maximum_income
                        If Calculate <= 0 Then
                            Retval = RowPersen_After.tarif / 100 * vNilaiPKPSetahun
                            NilaiSisa = 0
                            Exit For
                        Else
                            Retval = RowPersen_After.tarif / 100 * RowPersen_After.maximum_income
                            NilaiSisa = Calculate
                        End If
                    Else
                        RowPersen_Before = TBLPersen.Rows(ij - 1)
                        Calculate = NilaiSisa - (RowPersen_After.maximum_income - RowPersen_Before.maximum_income)
                        If Calculate <= 0 Then
                            Retval += RowPersen_After.tarif / 100 * NilaiSisa
                            NilaiSisa = 0
                            Exit For
                        Else
                            Retval += RowPersen_After.tarif / 100 * (RowPersen_After.maximum_income - RowPersen_Before.maximum_income)
                            NilaiSisa = Calculate
                        End If
                    End If
                Next
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPersen.Dispose()
            TBLPersen.Dispose()
        End Try
    End Function

    Public Function GetValue(ByVal vTBL As DseTrans_Tax_Beginning_Balance.tbltrans_tax_beginning_balance_item1DataTable, ByVal vStartPeriod As Date, _
    ByVal vEndPeriode As Date, ByVal vPrimaryKaryawan As Double)
        Dim DTA As New DseTrans_Tax_Beginning_BalanceTableAdapters.tbltrans_tax_beginning_balance_item1TableAdapter
        Dim row As DseTrans_Tax_Beginning_Balance.tbltrans_tax_beginning_balance_item1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vTBL.Rows.Clear()
            DTA.Fill(vTBL, vEndPeriode, vStartPeriod, vPrimaryKaryawan)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
