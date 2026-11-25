Option Explicit On
Option Strict Off

Imports System.Windows.Forms, CrystalDecisions.Windows.Forms

Public Class ClsTransaksiLoanPayment
    Private mPrimary As Double = -1
    Private mDocumentNo As String = ""
    Private mJournalNo As String = ""
    Private mJournalDesc As String = ""
    Private mUrutan As Integer = 0
    Private mDate As Date = Now
    Private mPrimarySalary As Double = -1
    Private mPrimaryJurnal As Double = -1
    Private mPrimaryKaryawan As Double = -1
    Private mPrimaryCash As Double = -1
    Private mPrimaryGain As Double = -1
    Private mPrimaryLoss As Double = -1
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

    Dim ObjJurnal As New AccModule.ClsTransaksiJournal

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
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

    Public Property JournalNo() As String
        Get
            Return mJournalNo
        End Get
        Set(ByVal value As String)
            mJournalNo = value
        End Set
    End Property

    Public Property JournalDesc() As String
        Get
            Return mJournalDesc
        End Get
        Set(ByVal value As String)
            mJournalDesc = value
        End Set
    End Property

    Public Property TanggalTransaksi() As Date
        Get
            Return mDate
        End Get
        Set(ByVal value As Date)
            mDate = value
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

    Public Property PrimarySalary() As Double
        Get
            Return mPrimarySalary
        End Get
        Set(ByVal value As Double)
            mPrimarySalary = value
        End Set
    End Property

    Public Property PrimaryJurnal() As Double
        Get
            Return mPrimaryJurnal
        End Get
        Set(ByVal value As Double)
            mPrimaryJurnal = value
        End Set
    End Property

    Public Property PrimaryKaryawan() As Double
        Get
            Return mPrimaryKaryawan
        End Get
        Set(ByVal value As Double)
            mPrimaryKaryawan = value
        End Set
    End Property

    Public Property PrimaryCash() As Double
        Get
            Return mPrimaryCash
        End Get
        Set(ByVal value As Double)
            mPrimaryCash = value
        End Set
    End Property

    Public Property PrimaryGain() As Double
        Get
            Return mPrimaryGain
        End Get
        Set(ByVal value As Double)
            mPrimaryGain = value
        End Set
    End Property

    Public Property PrimaryLoss() As Double
        Get
            Return mPrimaryLoss
        End Get
        Set(ByVal value As Double)
            mPrimaryLoss = value
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
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
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

    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
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

    Enum EnumJurnalOtomatis
        Simpan
        Edit
        Hapus
    End Enum

    Public Enum EnumFrmTransLoanPaymentDari As Integer
        FrmTransLoanPayment
        FrmTransLoanPayment_Beginning_Balance
    End Enum

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vDGVItem As DataGridView, ByVal vDef_Prec As Double, ByVal i As FrmTransaksiLoan_Payment) As Boolean
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim TBL As New DseTransaksiLoanPayment.tbltrans_loan_paymentDataTable
        Dim row As DseTransaksiLoanPayment.tbltrans_loan_paymentRow
        Dim ObjItem As New ClsTransaksiLoanPayment_Item
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mDocumentNo = .Document_No
                    Me.mDate = ._Date
                    Me.mUrutan = .Urutan
                    Me.mPrimarySalary = .Primary_Salary
                    Me.mPrimaryJurnal = .Primary_Jurnal
                    Me.mPrimaryKaryawan = .Primary_Karyawan
                    Me.mPrimaryCash = .Primary_Cash_In_Account
                    Me.mPrimaryGain = .Primary_Gain_Account
                    Me.mPrimaryLoss = .Primary_Loss_Account
                    Me.mRemarks = .Remarks
                    Me.mJournalNo = Me.ObjJurnal.GetRefCode(.Primary_Jurnal)
                    Me.mJournalDesc = Me.ObjJurnal.GetDescription(.Primary_Jurnal)

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

    Public Function GetJml_LoanPayment(ByVal vKriteria As String, ByVal vSearchBy As Core_Module.EnumSearchCriteria, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Double
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.TblCariTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByDocumentNoLIKE_N(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByCode Then
                Return DTA.ScalarQuery_GetJmlByEmployeeCodeLIKE_N(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByDate Then
                Return DTA.ScalarQuery_GetJmlByTanggalDariSampai_N(vEndDate, vStartDate).Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByFullName Then
                Return DTA.ScalarQuery_GetJmlByEmployeeName_N(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE_N(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByLoanNumber Then
                Return DTA.ScalarQuery_GetJmlByLoanNumberLIKE_N(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As Core_Module.EnumSearchCriteria, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Double
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.TblCariTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByDocumentNoLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByCode Then
                Return DTA.ScalarQuery_GetJmlByEmployeeCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByDate Then
                Return DTA.ScalarQuery_GetJmlByTanggalDariSampai(vEndDate, vStartDate).Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByFullName Then
                Return DTA.ScalarQuery_GetJmlByEmployeeName(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByLoanNumber Then
                Return DTA.ScalarQuery_GetJmlByLoanNumberLIKE(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV_LoanPayment(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiLoanPayment.TblCariDataTable, _
    ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Boolean
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.TblCariTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_DocumentNoLIKE_N(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDate Then
                QueryResult = DTA.FillBy_TanggalDariSampai_N(TBL, vEndDate, vStartDate)
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_EmployeeCodeLIKE_N(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByFullName Then
                QueryResult = DTA.FillBy_EmployeeNameLIKE_N(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE_N(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByLoanNumber Then
                QueryResult = DTA.FillBy_LoanNumberLIKE_N(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                'DGV.Columns(1).HeaderText = "Code"
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

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiLoanPayment.TblCariDataTable, _
    ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Boolean
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.TblCariTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_DocumentNoLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDate Then
                QueryResult = DTA.FillBy_TanggalDariSampai(TBL, vEndDate, vStartDate)
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_EmployeeCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByFullName Then
                QueryResult = DTA.FillBy_EmployeeNameLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByLoanNumber Then
                QueryResult = DTA.FillBy_LoanNumberLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(3).DefaultCellStyle.Format = "dd MMM yyyy"
                'DGV.Columns(1).HeaderText = "Code"
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
    ByVal vPrimary As Double, ByVal vdgvItem As DataGridView, ByVal ij As FrmTransaksiLoan_Payment) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_Payment_SimpanBaru, vUserCode)
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
    ByVal vPrimary As Double, ByVal vdgvItem As DataGridView, ByVal ij As FrmTransaksiLoan_Payment) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_Payment_SimpanPerubahan, vUserCode)

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
                            retval = Me.SimpanPerubahan1(vUserCode, vPrimary, vObjCD, vdgvItem, ij)
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
    ByVal vPrimary As Double, ByVal vdgvItem As DataGridView, ByVal ij As FrmTransaksiLoan_Payment, _
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
            Me.vPrimary_Antrian, vPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_Payment_Hapus, vUserCode)

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

    'Public Sub GetNoAntrian(ByRef vIdentifier As String, ByVal ObjRnd As Random, _
    'ByVal DTAInOUT As Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter, ByVal vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian, _
    'ByRef vPrimary_Antrian As Double, ByVal vPrimaryTransaksiAsal As Double, ByVal vTipeTransaksiAsal As Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian, _
    'ByVal vNamaUser As String)
    '    Dim NamaKomputer As String = ""
    '    Dim IP_Add As String = ""
    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Try
    '        NamaKomputer = My.Computer.Name
    '        IP_Add = ObjCD.GetIpAddress_Local
    '        DTAInOUT.Connection.ConnectionString = Core_Module.GetConstring

    '        vIdentifier = ObjRnd.Next(1000000)
    '        'jangan sampai sama dengan identifier yang sudah ada
    '        While DTAInOUT.ScalarQuery_GetJmlIdentifier(vTipeAntrian, vIdentifier).Value > 0
    '            vIdentifier = ObjRnd.Next(1000000)
    '        End While
    '        'masukan data baru ke antrian
    '        If System.Net.Dns.GetHostAddresses("").Length > 2 Then
    '            DTAInOUT.Insert(vTipeAntrian, vIdentifier, vPrimaryTransaksiAsal, vTipeTransaksiAsal, vNamaUser, My.Computer.Name + "-" + _
    '            System.Net.Dns.GetHostAddresses("")(2).ToString)
    '        Else
    '            DTAInOUT.Insert(vTipeAntrian, vIdentifier, vPrimaryTransaksiAsal, vTipeTransaksiAsal, vNamaUser, My.Computer.Name + "- localhost")
    '        End If

    '        'cek jangan sampai ada 2 data yang sama
    '        While DTAInOUT.ScalarQuery_GetJmlIdentifier(vTipeAntrian, vIdentifier).Value > 1
    '            'bila ada yang sama maka harus masukkan data baru
    '            DTAInOUT.DeleteQuery(vTipeAntrian, vIdentifier)
    '            'minta identifier baru
    '            vIdentifier = ObjRnd.Next(1000000)
    '            'jangan sampai sama dengan identifier yang sudah ada
    '            While DTAInOUT.ScalarQuery_GetJmlIdentifier(vTipeAntrian, vIdentifier).Value > 0
    '                vIdentifier = ObjRnd.Next(1000000)
    '                'masukan data baru ke antrian
    '                DTAInOUT.Insert(vTipeAntrian, vIdentifier, vPrimaryTransaksiAsal, vTipeTransaksiAsal, vNamaUser, My.Computer.Name)
    '            End While
    '        End While

    '        'ambil no antrian
    '        vPrimary_Antrian = DTAInOUT.Scalar_Query_GetPrimary(vIdentifier, vTipeAntrian).Value
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjCD = Nothing
    '    End Try
    'End Sub

    Public Function GetJmlByDocumentNo(ByVal vDocumentNo As String) As Double
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByDocumentNo(vDocumentNo)
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

    Public Function GetMAXUrutan(ByVal vEndDate As Date, ByVal vStartDate As Date) As Double
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim Dummy As System.Nullable(Of Long)
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

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vObjCD As Core_Module.ClsCekData) As String
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim InOut As String = "LOANPAYMENT"
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

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNoTransaksi(vPrimary)
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

    Public Function GetPrimary(ByVal vDocumentNo As String) As Double
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimary(vDocumentNo)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vPrimary As Double, _
    ByVal vShowMsg As Boolean, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTransaksiLoan_Payment) As Boolean
        Dim RowDGV As DataGridViewRow
        Dim ObjItem As New ClsTransaksiLoanPayment_Item
        Dim InfoString As String = ""
        Dim ij As Integer = 0
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Dim PayrollConventionDate As Date = ObjClosing.GetPayrollConvertionDate()
        Try
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mDocumentNo) = True Then
                    Me.mDocumentNo = Me.GetNoBaru(Me.mDate, vObjCD)
                End If
            Else
                Dim NoTransaksiLama As String = Me.GetNoTransaksi(Me.mPrimary)
                If NoTransaksiLama <> Me.mDocumentNo Then
                    If Me.ApakahSudahAda(Me.mDocumentNo) = True Then
                        Me.mDocumentNo = Me.GetNoBaru(Me.mDate, vObjCD)
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. kode employee code
            If Me.mPrimaryKaryawan = "-1" Then
                If vShowMsg = True Then
                    MsgBox("Please input employee code.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '2. kode cash account
            If Me.mPrimaryCash = "-1" Then
                If vShowMsg = True Then
                    MsgBox("Please input cash out account.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '3. kode gain account
            If Me.mPrimaryGain = "-1" Then
                If vShowMsg = True Then
                    MsgBox("Please input gain account.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '4. kode loss account
            If Me.mPrimaryLoss = "-1" Then
                If vShowMsg = True Then
                    MsgBox("Please input loss account.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '5. Remarks
            If vObjCD.CekDataKosong(Me.mRemarks, False) = False Then
                If vShowMsg = True Then
                    MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '2.Harus input Convention date HR
            If payrollConventionDate = #1/1/1950# Then
                MsgBox("Please input payroll convention date in global setting data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If i.EnumDipanggilDari = EnumFrmTransLoanPaymentDari.FrmTransLoanPayment Then
                '3.Harus dibuat setelah konvention date HR
                If Me.mDate < PayrollConventionDate Then
                    MsgBox("Please make sure date in transaction loan payment, loan payment's date must be bigger than payroll convention date.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf i.EnumDipanggilDari = EnumFrmTransLoanPaymentDari.FrmTransLoanPayment_Beginning_Balance Then
                '3.Harus dibuat setelah konvention date HR
                If Me.mDate >= PayrollConventionDate Then
                    MsgBox("Please make sure date in transaction beginning balance loan payment, loan payment's date must be bigger than payroll convention date.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'Cek apakah transaksi loan masih ada
            If i.dgvItem.Rows.Count > 0 Then
                For Each RowDGV In i.dgvItem.Rows
                    With RowDGV
                        If ObjItem.GetJmlLoanItemByIdentifier(RowDGV.Cells(i.ColIdentifier.Index).Value) <= 0 Then
                            ij += 1
                            If String.IsNullOrEmpty(InfoString) = False Then
                                InfoString += vbCrLf + ij.ToString + ". Loan Number : " + RowDGV.Cells(i.ColLoanNo.Index).Value.ToString + "."
                            Else
                                InfoString = ij.ToString + ". Loan Number : " + RowDGV.Cells(i.ColLoanNo.Index).Value.ToString + "."
                            End If
                        End If
                    End With
                Next
            End If
            If String.IsNullOrEmpty(InfoString) = False Then
                If vShowMsg = True Then
                    MsgBox("You cannot save this data because data item : " + vbCrLf + InfoString + vbCrLf + "is deleted.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjItem = Nothing
            ObjClosing = Nothing
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vUserCode As String, _
       ByVal vObjCD As Core_Module.ClsCekData, ByVal vPrimary As Double, _
       ByVal vDGVItem As DataGridView, ByVal i As FrmTransaksiLoan_Payment) As Boolean
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim DtaJournalHeader As New DseTransaksiLoanPaymentTableAdapters.tbltrans_jurnalTableAdapter
        Dim ObjItem As New ClsTransaksiLoanPayment_Item
        Dim ObjLoan As New ClsTransaksiLoan
        Dim RowStatusLoan As DseTransaksiLoanPayment.TBLUpdateStatus_LoanRow
        Dim BufferDoc As String = Me.mDocumentNo
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DtaJournalHeader.Connection.ConnectionString = DTA.Connection.ConnectionString

            If Me.ApakahDataBenar(True, vPrimary, True, vObjCD, vDGVItem, i) = True Then
                Me.mUrutan = CDbl(Me.mDocumentNo.Substring(0, 6))
                If DTA.Insert(Me.mDocumentNo, Me.mUrutan, Me.mDate, Me.mPrimarySalary, _
                Me.mPrimaryJurnal, Me.mPrimaryKaryawan, Me.mPrimaryCash, Me.mPrimaryGain, _
                Me.mPrimaryLoss, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mDocumentNo)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, vDGVItem, i)

                    'Update Payment Loan
                    ObjItem.UpdatePaymentLoan_Item(True, i.dtpDate.Value, i.TBLLoan, ClsTransaksiLoanPayment_Item.EnumNilaiUpdate.Update, i)

                    'Update Status Loan Transaksi
                    For Each RowStatusLoan In i.TBLLoanStatus.Rows
                        ObjLoan.UpdateStatusLoan(RowStatusLoan.Posisi, RowStatusLoan.Loan_Number)
                    Next

                    'journal otomatis, hanya jika primary salary = -1 dan show from salary = true
                    If Me.mPrimarySalary = -1 Then
                        If i.ShowFromSalary = False Then
                            If i.EnumDipanggilDari = EnumFrmTransLoanPaymentDari.FrmTransLoanPayment Then
                                Me.UpdateJournalNumber(BufferDoc)
                                Me.JurnalOtomatis(vUserCode, EnumJurnalOtomatis.Simpan, vDGVItem, i)
                            End If
                        End If
                    End If

                    If Me.mPrimarySalary = -1 Then
                        If i.ShowFromSalary = True Then
                            Me.mPrimaryJurnal = -1
                        Else
                            If i.EnumDipanggilDari = EnumFrmTransLoanPaymentDari.FrmTransLoanPayment Then
                                Me.mPrimaryJurnal = DtaJournalHeader.ScalarQuery_GetPrimaryByDocNo(Me.mJournalNo)
                            Else
                                Me.mPrimaryJurnal = -1
                            End If
                        End If
                    Else
                        Me.mPrimaryJurnal = -1
                    End If

                    DTA.UpdateQuery_PrimaryJurnal(Me.mPrimaryJurnal, Me.mPrimary)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)

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
            DtaJournalHeader.Dispose()
            ObjItem = Nothing
            ObjLoan = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUserCode As String, _
         ByVal vPrimary As Double, ByVal vObjCD As Core_Module.ClsCekData, _
         ByVal vDGVItem As DataGridView, ByVal i As FrmTransaksiLoan_Payment) As Boolean
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim DTAItem As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_payment_temTableAdapter
        Dim TBLItem As New DseTransaksiLoanPayment.tbltrans_loan_payment_temDataTable
        Dim RowItem As DseTransaksiLoanPayment.tbltrans_loan_payment_temRow
        Dim RowStatusLoan As DseTransaksiLoanPayment.TBLUpdateStatus_LoanRow
        Dim DTAUpdate As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_itemTableAdapter
        Dim ObjItem As New ClsTransaksiLoanPayment_Item
        Dim ObjLoan As New ClsTransaksiLoan
        Dim DTALoanItem As New DseTransaksiLoanPaymentTableAdapters.DataTable1TableAdapter
        Dim TBLLoanItem As New DseTransaksiLoanPayment.DataTable1DataTable
        Dim RowLoanItem As DseTransaksiLoanPayment.DataTable1Row
        Dim DtaJournalHeader As New DseTransaksiLoanPaymentTableAdapters.tbltrans_jurnalTableAdapter
        Dim BufferDoc As String = Me.mDocumentNo

        Try
            If Me.ApakahDataBenar(False, vPrimary, True, vObjCD, vDGVItem, i) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTAItem.Connection.ConnectionString = Core_Module.GetConstring
                DTALoanItem.Connection.ConnectionString = DTA.Connection.ConnectionString
                DtaJournalHeader.Connection.ConnectionString = DTA.Connection.ConnectionString

                Me.mUrutan = CDbl(Me.mDocumentNo.Substring(0, 6))
                If DTA.Update(Me.mDocumentNo, Me.mUrutan, Me.mDate, _
                              Me.mPrimarySalary, Me.mPrimaryJurnal, Me.mPrimaryKaryawan, Me.mPrimaryCash, _
                              Me.mPrimaryGain, Me.mPrimaryLoss, Me.mRemarks, _
                              vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'Update Payment Loan ---> NilaiDefault
                    DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimary)
                    If TBLItem.Rows.Count > 0 Then
                        For Each RowItem In TBLItem.Rows
                            DTAUpdate.UpdateQuery_NilaiPayment(#1/1/1950#, 0, RowItem.Identifier_Loan)
                        Next
                    End If
                    'Update Status Loan (Data Lama)
                    If i.TBLLoanStatus.Rows.Count > 0 Then
                        DTALoanItem.FillBy_PrimaryTransaksi(TBLLoanItem, vPrimary)
                        If TBLLoanItem.Rows.Count > 0 Then
                            For Each RowLoanItem In TBLLoanItem.Rows
                                If ObjLoan.GetInstallmentTo(RowLoanItem.Loan_Number) = RowLoanItem.Installment_At Then
                                    ObjLoan.UpdateStatusLoan(ClsTransaksiLoan.EnumStatusLoan.OnGoing, RowLoanItem.Loan_Number)
                                End If
                            Next
                        End If
                    End If

                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, vDGVItem, i)

                    'Update Payment Loan (Data Baru)
                    ObjItem.UpdatePaymentLoan_Item(False, i.dtpDate.Value, i.TBLLoan, ClsTransaksiLoanPayment_Item.EnumNilaiUpdate.Update, i)

                    'Update Status Loan Transaksi
                    For Each RowStatusLoan In i.TBLLoanStatus.Rows
                        ObjLoan.UpdateStatusLoan(RowStatusLoan.Posisi, RowStatusLoan.Loan_Number)
                    Next

                    'journal otomatis, jika bukan dipanggil dari salary, dan bukan begining balance
                    If Me.mPrimarySalary = -1 Then
                        If i.ShowFromSalary = False Then
                            If i.EnumDipanggilDari = EnumFrmTransLoanPaymentDari.FrmTransLoanPayment Then
                                Me.UpdateJournalNumber(BufferDoc)
                                Me.JurnalOtomatis(vUserCode, EnumJurnalOtomatis.Edit, vDGVItem, i, Me.mPrimaryJurnal)
                            End If
                        End If
                    End If

                    If Me.mPrimarySalary = -1 Then
                        If i.ShowFromSalary = True Then
                            Me.mPrimaryJurnal = -1
                        Else
                            If i.EnumDipanggilDari = EnumFrmTransLoanPaymentDari.FrmTransLoanPayment Then
                                Me.mPrimaryJurnal = DtaJournalHeader.ScalarQuery_GetPrimaryByDocNo(Me.mJournalNo)
                            Else
                                Me.mPrimaryJurnal = -1
                            End If
                        End If
                    Else
                        Me.mPrimaryJurnal = -1
                    End If

                    DTA.UpdateQuery_PrimaryJurnal(Me.mPrimaryJurnal, Me.mPrimary)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)

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
            DTAItem.Dispose()
            TBLItem.Dispose()
            DTAUpdate.Dispose()
            DTALoanItem.Dispose()
            TBLLoanItem.Dispose()
            ObjItem = Nothing
            ObjLoan = Nothing
            DtaJournalHeader.Dispose()
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTransaksiLoan_Payment) As Boolean
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim DTAUpdate As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_itemTableAdapter
        Dim ObjItem As New ClsTransaksiLoanPayment_Item
        Dim RowDGV As DataGridViewRow
        Dim ObjLoan As New ClsTransaksiLoan
        Dim DTALoanItem As New DseTransaksiLoanPaymentTableAdapters.DataTable1TableAdapter
        Dim TBLLoanItem As New DseTransaksiLoanPayment.DataTable1DataTable
        Dim RowLoanItem As DseTransaksiLoanPayment.DataTable1Row
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, vStrictDelete, vShowMsg, vDGVItem, i, i.vFromSalaryCalc, i.EnumDipanggilDari) = True Then
                'Update Nilai Paymet ---> Nilai Default
                For Each RowDGV In vDGVItem.Rows
                    DTAUpdate.UpdateQuery_NilaiPayment(#1/1/1950#, 0, RowDGV.Cells(i.ColIdentifier.Index).Value.ToString)
                Next
                'Update Status Loan (Data Lama)
                DTALoanItem.FillBy_PrimaryTransaksi(TBLLoanItem, vPrimary)
                If TBLLoanItem.Rows.Count > 0 Then
                    For Each RowLoanItem In TBLLoanItem.Rows
                        If ObjLoan.GetInstallmentTo(RowLoanItem.Loan_Number) = RowLoanItem.Installment_At Then
                            ObjLoan.UpdateStatusLoan(ClsTransaksiLoan.EnumStatusLoan.OnGoing, RowLoanItem.Loan_Number)
                        End If
                    Next
                End If

                'Hapus item
                ObjItem.HapusItem(vPrimary)
                'Hapus
                DTA.Delete(vPrimary)

                'journal otomatis
                If Me.mPrimarySalary = -1 Then
                    If i.EnumDipanggilDari = EnumFrmTransLoanPaymentDari.FrmTransLoanPayment Then

                        Me.JurnalOtomatis(vUserCode, EnumJurnalOtomatis.Hapus, vDGVItem, i, Me.ObjJurnal.GetPrimary(i.txtJournalNo.Text))
                    End If
                End If

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAUpdate.Dispose()
            DTALoanItem.Dispose()
            TBLLoanItem.Dispose()
            ObjItem = Nothing
            ObjLoan = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mDocumentNo = ""
            mUrutan = 0
            mDate = Now
            mPrimarySalary = -1
            mPrimaryJurnal = -1
            mPrimaryKaryawan = -1
            mPrimaryCash = -1
            mPrimaryGain = -1
            mPrimaryLoss = -1
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
    ByVal vStrictEdit As Boolean, ByVal vShowMsg As Boolean, ByVal vFromSalary As Boolean, ByVal vSourceForm As EnumFrmTransLoanPaymentDari) As Boolean
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim pCreatedBy As String = ""
        Dim vPrimarySalary As Double = -1
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Dim PayrollConvDate As Date = Today
        Dim vTanggalTransaksi As Date = Today

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

            '2. Berkaitan dengan transaksi salary
            'jika primary transaksi salary <> -1 and dipanggil bukan dari form salary -> dipanggil dari begining balance/payment normal, tidak boleh edit
            vPrimarySalary = Me.GetPrimarySalary(vPrimary)
            If vPrimarySalary <> -1 And vFromSalary = False Then
                If vShowMsg = True Then
                    Dim ObjSalaryCalc As New ClsTransaksiSalary
                    MsgBox("Please use salary calculation : " + ObjSalaryCalc.GetNoTransaksi(vPrimary) + " to edit this data.", MsgBoxStyle.Information, "Information")
                    ObjSalaryCalc = Nothing
                End If
                Return False
            End If

            '3. Jika dibuat dari begining balance tetapi di edit dari form normal
            PayrollConvDate = ObjClosing.GetPayrollConvertionDate
            vTanggalTransaksi = Me.GetTanggalTransaksi(vPrimary)
            If vSourceForm = EnumFrmTransLoanPaymentDari.FrmTransLoanPayment Then
                If vTanggalTransaksi.Date < PayrollConvDate.Date Then
                    If vShowMsg = True Then
                        Dim ObjSalaryCalc As New ClsTransaksiSalary
                        MsgBox("Please use loan payment form to edit this data.", MsgBoxStyle.Information, "Information")
                        ObjSalaryCalc = Nothing
                    End If
                    Return False
                End If
            Else
                '4. Jika dibuat dari form normal tetapi diedit dari beginning balance
                If vTanggalTransaksi.Date >= PayrollConvDate.Date Then
                    If vShowMsg = True Then
                        Dim ObjSalaryCalc As New ClsTransaksiSalary
                        If vPrimarySalary = -1 Then
                            MsgBox("Please use loan payment begining balance form to edit this data.", MsgBoxStyle.Information, "Information")
                        Else
                            MsgBox("Please use salary calculation : " + ObjSalaryCalc.GetNoTransaksi(vPrimary) + " to edit this data.", MsgBoxStyle.Information, "Information")
                        End If
                        ObjSalaryCalc = Nothing
                    End If
                    Return False
                End If
            End If


            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjClosing = Nothing
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTransaksiLoan_Payment, ByVal vFromSalary As Boolean, ByVal vSourceForm As EnumFrmTransLoanPaymentDari) As Boolean
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim RowDGV As DataGridViewRow
        Dim pCreatedBy As String = ""
        Dim PrimaryLoanMin As Double = Double.MaxValue
        Dim ObjItem As New ClsTransaksiLoanPayment_Item
        Dim InfoData As String = ""
        Dim PayrollConvDate As Date = Today
        Dim vTanggalTransaksi As Date = Today
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Dim vPrimarySalary As Double = -1

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

            '2. Berkaitan dengan transaksi salary
            'jika primary transaksi salary <> -1 and dipanggil bukan dari form salary -> dipanggil dari begining balance/payment normal, tidak boleh edit
            vPrimarySalary = Me.GetPrimarySalary(vPrimary)
            If vPrimarySalary <> -1 And vFromSalary = False Then
                If vShowMsg = True Then
                    Dim ObjSalaryCalc As New ClsTransaksiSalary
                    MsgBox("Please use salary calculation : " + ObjSalaryCalc.GetNoTransaksi(vPrimary) + " to edit this data.", MsgBoxStyle.Information, "Information")
                    ObjSalaryCalc = Nothing
                End If
                Return False
            End If

            '3. Jika dibuat dari begining balance tetapi di edit dari form normal
            PayrollConvDate = ObjClosing.GetPayrollConvertionDate
            vTanggalTransaksi = Me.GetTanggalTransaksi(vPrimary)
            If vSourceForm = EnumFrmTransLoanPaymentDari.FrmTransLoanPayment Then
                If vTanggalTransaksi.Date < PayrollConvDate.Date Then
                    If vShowMsg = True Then
                        Dim ObjSalaryCalc As New ClsTransaksiSalary
                        MsgBox("Please use loan payment begining balance form to edit this data.", MsgBoxStyle.Information, "Information")
                        ObjSalaryCalc = Nothing
                    End If
                    Return False
                End If
            Else
                '4. Jika dibuat dari form normal tetapi diedit dari beginning balance
                If vTanggalTransaksi.Date >= PayrollConvDate.Date Then
                    If vShowMsg = True Then
                        Dim ObjSalaryCalc As New ClsTransaksiSalary
                        MsgBox("Please use loan payment form to edit this data.", MsgBoxStyle.Information, "Information")
                        ObjSalaryCalc = Nothing
                    End If
                    Return False
                End If
            End If

            '5. Cek masing-masing item dengan loan apakah pinjaman yang terakhir
            If i.dgvItem.Rows.Count > 0 Then
                For Each RowDGV In i.dgvItem.Rows
                    If RowDGV.Cells(i.ColPrimaryLoan.Index).Value < PrimaryLoanMin Then
                        PrimaryLoanMin = RowDGV.Cells(i.ColPrimaryLoan.Index).Value
                    End If
                Next
                If PrimaryLoanMin <> Double.MaxValue Then
                    InfoData = ObjItem.ApakahLoanNumberBolehDiHapus(PrimaryLoanMin, i)
                End If
                If InfoData <> "" Then
                    MsgBox("Transaction can't delete because some loan isn't last installment, such as :" + vbCrLf + InfoData, MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjClosing = Nothing
        End Try
    End Function

    Public Sub IsiTBLCashCode(ByVal CBO As ComboBox, ByVal vPrimaryCurr_IDR As Double, _
    ByVal TBL As DseTransaksiLoanPayment.tblmaster_akunDataTable, ByVal ij As FrmTransaksiLoan_Payment, _
    Optional ByVal vTipeAkun1 As AccModule.ClsMasterAkun.EnumTipeAkun = AccModule.ClsMasterAkun.EnumTipeAkun.Asset, _
    Optional ByVal vTipeAkun2 As AccModule.ClsMasterAkun.EnumTipeAkun2 = AccModule.ClsMasterAkun.EnumTipeAkun2.Cash_And_Bank)
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            CBO.DataSource = Nothing
            CBO.Items.Clear()
            TBL.Rows.Clear()

            DTA.FillBy_GetAccountCode(TBL, vPrimaryCurr_IDR, vTipeAkun1, vTipeAkun2)

            CBO.DataSource = TBL.DefaultView
            TBL.DefaultView.Sort = "[Account Code] ASC"
            CBO.DisplayMember = "Account Code"
            ij.TxtDescCash.Text = TBL(CBO.SelectedIndex).Description
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
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

    Public Function GetTanggalTransaksi(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTanggalTransaksi(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Fungsi ini dipanggil setelah salin ke field jika baru = false 
    Public Sub IsiTBLLoanPayment_OnGoing(ByVal vTBLLoan_OnGoing As DseTransaksiLoanPayment.DataTable2DataTable, _
    ByVal vPrimaryKaryawan As Double, ByVal vBaru As Boolean, ByVal vPrimaryLoanPayment As Double, ByVal vDGVItem As DataGridView, _
    ByVal ij As FrmTransaksiLoan_Payment)
        Dim Dummy_TBLLoan_OnGoing As New DseTransaksiLoanPayment.DataTable2DataTable
        Dim Dummy_RowLoan_OnGoing As DseTransaksiLoanPayment.DataTable2Row
        Dim RowLoan_OnGoing As DseTransaksiLoanPayment.DataTable2Row
        Dim DTA_All As New DseTransaksiLoanPaymentTableAdapters.DataTable2TableAdapter
        Dim TBL_All As New DseTransaksiLoanPayment.DataTable2DataTable
        Dim Row_All As DseTransaksiLoanPayment.DataTable2Row
        Dim DTAPayment As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_payment_temTableAdapter
        Dim TBLPayment As New DseTransaksiLoanPayment.tbltrans_loan_payment_temDataTable
        Dim RowPayment As DseTransaksiLoanPayment.tbltrans_loan_payment_temRow
        Dim RowDGV As DataGridViewRow
        Dim ListPrimaryLoan_GaAda As New List(Of Double)
        Dim ListDummy As New List(Of Double)
        Dim ListIdentifierLoan_GaAda As New List(Of String)
        Dim ListPaymentAmount_GaAda As New List(Of Decimal)
        Dim AdaPrimaryLoan_GaAda As Boolean = False
        Dim PrimaryLoan_GaAda As Double = -1
        Dim IdentifierLoan_GaAda As String = ""
        Dim NoList As Integer = -1
        Dim DV As Data.DataRowView
        Dim Filter As String = ""
        Try
            'Connection string 
            DTA_All.Connection.ConnectionString = Core_Module.GetConstring
            DTAPayment.Connection.ConnectionString = Core_Module.GetConstring

            'Bersihkan Object
            vTBLLoan_OnGoing.Rows.Clear()

            'Fill data Loan berdasarkan status = on going AND Primary Karyawan tertentu
            Me.DataLoan_OnGoing(vTBLLoan_OnGoing, vPrimaryKaryawan)
            DTA_All.FillBy_StatusOnGoing(TBL_All)

            'Bandingkan dengan database Payment
            If vBaru = False Then
                ListPrimaryLoan_GaAda.Clear()
                ListDummy.Clear()
                ListIdentifierLoan_GaAda.Clear()
                ListPaymentAmount_GaAda.Clear()
                DTAPayment.FillBy_PrimaryTransaksi(TBLPayment, vPrimaryLoanPayment)
                If TBLPayment.Rows.Count > 0 Then
                    For Each RowPayment In TBLPayment.Rows
                        Filter = "[Primary] = " + RowPayment.Primary_Loan.ToString + " AND [Identifier] = '" + RowPayment.Identifier_Loan.ToString + "'"
                        vTBLLoan_OnGoing.DefaultView.RowFilter = Filter
                        If vTBLLoan_OnGoing.DefaultView.Count > 0 Then
                            RowLoan_OnGoing = vTBLLoan_OnGoing.DefaultView(0).Row
                            RowLoan_OnGoing.Payment_Amount = RowPayment.Payment
                        Else
                            If ListIdentifierLoan_GaAda.Contains(RowPayment.Identifier_Loan) = False Then
                                ListIdentifierLoan_GaAda.Add(RowPayment.Identifier_Loan)
                                ListPaymentAmount_GaAda.Add(RowPayment.Payment)
                                ListPrimaryLoan_GaAda.Add(RowPayment.Primary_Loan)
                            End If
                        End If
                    Next
                    vTBLLoan_OnGoing.DefaultView.RowFilter = ""
                End If
                'masukan data loan yang tidak ditemukan
                If ListPrimaryLoan_GaAda.Count > 0 Then
                    AdaPrimaryLoan_GaAda = True
                    For Each PrimaryLoan_GaAda In ListPrimaryLoan_GaAda
                        If ListDummy.Contains(PrimaryLoan_GaAda) = False Then
                            ListDummy.Add(PrimaryLoan_GaAda)
                            Filter = "[Primary] = " + PrimaryLoan_GaAda.ToString + ""
                            TBL_All.DefaultView.RowFilter = Filter
                            If TBL_All.DefaultView.Count > 0 Then
                                For Each DV In TBL_All.DefaultView
                                    Row_All = DV.Row
                                    RowLoan_OnGoing = vTBLLoan_OnGoing.NewDataTable2Row
                                    For idx As Integer = 0 To vTBLLoan_OnGoing.Columns.Count - 1
                                        RowLoan_OnGoing.Item(idx) = Row_All.Item(idx)
                                    Next
                                    vTBLLoan_OnGoing.Rows.Add(RowLoan_OnGoing)
                                Next
                            End If
                        End If
                    Next
                    TBL_All.DefaultView.RowFilter = ""
                    'masukan nilai amount yang tidak ditemukan
                    For Each IdentifierLoan_GaAda In ListIdentifierLoan_GaAda
                        NoList += 1
                        Filter = "[Primary] = " + ListPrimaryLoan_GaAda(NoList).ToString + " AND [Identifier] = '" + IdentifierLoan_GaAda.ToString + "'"
                        vTBLLoan_OnGoing.DefaultView.RowFilter = Filter
                        If vTBLLoan_OnGoing.DefaultView.Count > 0 Then
                            RowLoan_OnGoing = vTBLLoan_OnGoing.DefaultView(0).Row
                            RowLoan_OnGoing.Payment_Amount = ListPaymentAmount_GaAda(NoList)
                        End If
                    Next
                    vTBLLoan_OnGoing.DefaultView.RowFilter = ""
                End If
            End If

            'Bandingkan dengan DGVItem
            ListPrimaryLoan_GaAda.Clear()
            ListDummy.Clear()
            ListIdentifierLoan_GaAda.Clear()
            ListPaymentAmount_GaAda.Clear()
            If vDGVItem.Rows.Count > 0 Then
                For Each RowDGV In vDGVItem.Rows
                    Filter = "[Primary] = " + RowDGV.Cells(ij.ColPrimaryLoan.Index).Value.ToString + " AND [Identifier] = '" + RowDGV.Cells(ij.ColIdentifier.Index).Value.ToString + "'"
                    vTBLLoan_OnGoing.DefaultView.RowFilter = Filter
                    If vTBLLoan_OnGoing.DefaultView.Count > 0 Then
                        RowLoan_OnGoing = vTBLLoan_OnGoing.DefaultView(0).Row
                        RowLoan_OnGoing.Payment_Amount = RowDGV.Cells(ij.ColPayment.Index).Value
                    Else
                        If ListIdentifierLoan_GaAda.Contains(RowDGV.Cells(ij.ColIdentifier.Index).Value) = False Then
                            ListIdentifierLoan_GaAda.Add(RowDGV.Cells(ij.ColIdentifier.Index).Value)
                            ListPaymentAmount_GaAda.Add(RowDGV.Cells(ij.ColPayment.Index).Value)
                            ListPrimaryLoan_GaAda.Add(RowDGV.Cells(ij.ColPrimaryLoan.Index).Value)
                        End If
                    End If
                Next
                vTBLLoan_OnGoing.DefaultView.RowFilter = ""
            End If
            'masukan data loan DGV yang tidak ditemukan
            If ListPrimaryLoan_GaAda.Count > 0 Then
                AdaPrimaryLoan_GaAda = True
                For Each PrimaryLoan_GaAda In ListPrimaryLoan_GaAda
                    If ListDummy.Contains(PrimaryLoan_GaAda) = False Then
                        ListDummy.Add(PrimaryLoan_GaAda)
                        Filter = "[Primary] = " + PrimaryLoan_GaAda.ToString + ""
                        TBL_All.DefaultView.RowFilter = Filter
                        If TBL_All.DefaultView.Count > 0 Then
                            For Each DV In TBL_All.DefaultView
                                Row_All = DV.Row
                                RowLoan_OnGoing = vTBLLoan_OnGoing.NewDataTable2Row
                                For idx As Integer = 0 To vTBLLoan_OnGoing.Columns.Count - 1
                                    RowLoan_OnGoing.Item(idx) = Row_All.Item(idx)
                                Next
                                vTBLLoan_OnGoing.Rows.Add(RowLoan_OnGoing)
                            Next
                        End If
                    End If
                Next
                TBL_All.DefaultView.RowFilter = ""
                'masukan nilai amount DGV yang tidak ditemukan
                For Each IdentifierLoan_GaAda In ListIdentifierLoan_GaAda
                    NoList += 1
                    Filter = "[Primary] = " + ListPrimaryLoan_GaAda.Item(NoList).ToString + " AND [Identifier] = '" + IdentifierLoan_GaAda.ToString + "'"
                    vTBLLoan_OnGoing.DefaultView.RowFilter = Filter
                    If vTBLLoan_OnGoing.DefaultView.Count > 0 Then
                        RowLoan_OnGoing = vTBLLoan_OnGoing.DefaultView(0).Row
                        RowLoan_OnGoing.Payment_Amount = ListPaymentAmount_GaAda(NoList)
                    End If
                Next
                vTBLLoan_OnGoing.DefaultView.RowFilter = ""
            End If

            'Sorting Data bila diperlukan
            If AdaPrimaryLoan_GaAda = True Then
                'Sorting data By Primary Loan and Identifier Loan
                Dummy_TBLLoan_OnGoing.Rows.Clear()
                vTBLLoan_OnGoing.DefaultView.Sort = "[Primary], [Identifier]"
                For Each DV In vTBLLoan_OnGoing.DefaultView
                    RowLoan_OnGoing = DV.Row
                    Dummy_RowLoan_OnGoing = Dummy_TBLLoan_OnGoing.NewDataTable2Row
                    For idx As Integer = 0 To Dummy_TBLLoan_OnGoing.Columns.Count - 1
                        Dummy_RowLoan_OnGoing.Item(idx) = RowLoan_OnGoing.Item(idx)
                    Next
                    Dummy_TBLLoan_OnGoing.AddDataTable2Row(Dummy_RowLoan_OnGoing)
                Next
                'Kembalikan data Sort
                vTBLLoan_OnGoing.Rows.Clear()
                vTBLLoan_OnGoing.DefaultView.RowFilter = ""
                vTBLLoan_OnGoing.DefaultView.Sort = ""
                For Each Dummy_RowLoan_OnGoing In Dummy_TBLLoan_OnGoing.Rows
                    RowLoan_OnGoing = vTBLLoan_OnGoing.NewDataTable2Row
                    For idx As Integer = 0 To vTBLLoan_OnGoing.Columns.Count - 1
                        RowLoan_OnGoing.Item(idx) = Dummy_RowLoan_OnGoing.Item(idx)
                    Next
                    vTBLLoan_OnGoing.AddDataTable2Row(RowLoan_OnGoing)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPayment.Dispose()
            TBLPayment.Dispose()
            DTA_All.Dispose()
            TBL_All.Dispose()
            Dummy_TBLLoan_OnGoing.Dispose()
            ListPrimaryLoan_GaAda = Nothing
            ListDummy = Nothing
            ListIdentifierLoan_GaAda = Nothing
            ListPaymentAmount_GaAda = Nothing
        End Try
    End Sub

    Private Sub DataLoan_OnGoing(ByVal vTBLLoan_OnGoing As DseTransaksiLoanPayment.DataTable2DataTable, _
    ByVal vPrimaryKaryawan As Double)
        Dim DTALoan_OnGoing As New DseTransaksiLoanPaymentTableAdapters.DataTable2TableAdapter
        Try
            DTALoan_OnGoing.Connection.ConnectionString = Core_Module.GetConstring
            vTBLLoan_OnGoing.Rows.Clear()
            DTALoan_OnGoing.Fill(vTBLLoan_OnGoing, vPrimaryKaryawan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTALoan_OnGoing.Dispose()
        End Try
    End Sub

    Public Sub IsiCBOLoanPayment(ByVal vTBLLoan_OnGoing As DseTransaksiLoanPayment.DataTable2DataTable, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vDefPrec As Double, ByVal ij As FrmTransaksiLoan_Payment, _
    ByRef TBLCboItem As DseTransaksiLoanPayment.DataTable2DataTable)
        Dim RowLoan As DseTransaksiLoanPayment.DataTable2Row
        Dim RowCBO As DseTransaksiLoanPayment.DataTable2Row
        Dim ListPrimaryLoan As New List(Of Double)
        Dim ListLoanNo As New List(Of String)
        Dim PrimaryLoan As Double
        Dim Filter As String = ""
        Dim Payment As Decimal = 0
        Dim NoList As Integer = -1
        Try
            'Bersihkan Object
            TBLCboItem.Rows.Clear()
            ij.TxtIdentifier.Text = ""
            ij.TxtPrimaryLoan.Text = "-1"
            ij.TxtInstallmentAt.Text = "0"
            ij.TxtBase.Text = "0"
            ij.TxtInterest.Text = "0"
            ij.TxtTotalInstallment.Text = "0"
            ij.TxtPayment.Text = "0"
            ij.TxtPrimaryRev.Text = "-1"
            ij.CboLoanNumber.DataSource = Nothing

            'Input Data CBOItem
            For Each RowLoan In vTBLLoan_OnGoing.Rows
                If ListPrimaryLoan.Contains(RowLoan.Primary) = False Then
                    ListPrimaryLoan.Add(RowLoan.Primary)
                    ListLoanNo.Add(RowLoan.Loan_No)
                End If
            Next
            For Each PrimaryLoan In ListPrimaryLoan
                NoList += 1
                Filter = "[Primary] = " + PrimaryLoan.ToString + " AND [Payment Amount] = " + Payment.ToString + ""
                vTBLLoan_OnGoing.DefaultView.RowFilter = Filter
                vTBLLoan_OnGoing.DefaultView.Sort = "[Installment Period]"
                'vTBLLoan_OnGoing.DefaultView.Sort = "[Primary], [Installment Period]"
                If vTBLLoan_OnGoing.DefaultView.Count > 0 Then
                    RowLoan = vTBLLoan_OnGoing.DefaultView(0).Row
                    RowCBO = TBLCboItem.NewRow
                    RowCBO.Primary = PrimaryLoan
                    RowCBO.Loan_No = ListLoanNo(NoList)
                    RowCBO.Installment_To = RowLoan.Installment_To
                    RowCBO.Base = RowLoan.Base
                    RowCBO.Interest = RowLoan.Interest
                    RowCBO.Payment_Amount = RowLoan.Payment_Amount
                    RowCBO.Identifier = RowLoan.Identifier
                    RowCBO.Installment_Period = RowLoan.Installment_Period
                    RowCBO.Primary_Receivable_Account = RowLoan.Primary_Receivable_Account
                    RowCBO.Date_Schedule = RowLoan.Date_Schedule
                    TBLCboItem.Rows.Add(RowCBO)
                End If
            Next
            If TBLCboItem.DefaultView.Count > 0 Then
                ij.CboLoanNumber.DataSource = TBLCboItem.DefaultView
                ij.CboLoanNumber.DisplayMember = "Loan No"
                ij.CboLoanNumber.SelectedIndex = 0
                RowCBO = TBLCboItem.DefaultView(0).Row
                ij.TxtIdentifier.Text = RowCBO.Identifier 'TBLCboItem(ij.CboLoanNumber.SelectedIndex).Identifier
                ij.TxtPrimaryLoan.Text = TBLCboItem(ij.CboLoanNumber.SelectedIndex).Primary
                ij.TxtInstallmentAt.Text = vObjCD.DigitGrouping(RowCBO.Installment_Period) 'TBLCboItem(ij.CboLoanNumber.SelectedIndex).Installment_Period)
                ij.TxtBase.Text = vObjCD.DigitGrouping(RowCBO.Base, vDefPrec) 'TBLCboItem(ij.CboLoanNumber.SelectedIndex).Base, vDefPrec)
                ij.TxtInterest.Text = vObjCD.DigitGrouping(RowCBO.Interest, vDefPrec) 'TBLCboItem(ij.CboLoanNumber.SelectedIndex).Interest, vDefPrec)
                ij.TxtTotalInstallment.Text = vObjCD.DigitGrouping(RowCBO.Base + RowCBO.Interest, vDefPrec) 'TBLCboItem(ij.CboLoanNumber.SelectedIndex).Base + TBLCboItem(ij.CboLoanNumber.SelectedIndex).Interest, vDefPrec)
                ij.TxtPrimaryRev.Text = RowCBO.Primary_Receivable_Account
            End If

            'Default value
            vTBLLoan_OnGoing.DefaultView.RowFilter = ""
            vTBLLoan_OnGoing.DefaultView.Sort = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryLoan = Nothing
            ListLoanNo = Nothing
        End Try
    End Sub

    'Dipakai di Event UserDeletingRow
    Public Function BatasanHapusLoanPayment_Row_Database(ByVal vPrimaryLoan As Double, _
    ByVal vPrimaryTransPayment As Double, ByVal vPembandingInstallment As Double, ByRef NoDoc_Payment As String, _
    ByRef NoLoan As String, ByRef Installment_At As Double) As Boolean
        Dim DTAItem1 As New DseTransaksiLoanPaymentTableAdapters.TblBatasan_Delete_Item1TableAdapter
        Dim TBLItem1 As New DseTransaksiLoanPayment.TblBatasan_Delete_Item1DataTable
        Dim RowItem1 As DseTransaksiLoanPayment.TblBatasan_Delete_Item1Row
        Try
            DTAItem1.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem1.Fill(TBLItem1, vPrimaryLoan, vPrimaryTransPayment)
            If TBLItem1.Rows.Count > 0 Then
                RowItem1 = TBLItem1.Rows(TBLItem1.Rows.Count - 1)
                NoDoc_Payment = RowItem1.Document_No
                NoLoan = RowItem1.No_Loan
                Installment_At = Me.GetInstallmentPeriod(RowItem1.Identifier_Loan)
                If Installment_At > vPembandingInstallment Then
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem1.Dispose()
            TBLItem1.Dispose()
        End Try
    End Function

    Public Function GetLoanNumber(ByVal vPrimaryLoan As Double) As String
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.DataTable2TableAdapter
        Dim Dummy As String
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetLoanNumber(vPrimaryLoan)
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

    Public Function GetInstallmentPeriod(ByVal vIdentifierLoan As String) As Double
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.TblBatasan_Delete_Item1TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetInstallmentLoanByIdentifier(vIdentifierLoan)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub JurnalOtomatis(ByVal vUserCode As String, ByVal vJournalOtomatis As EnumJurnalOtomatis, ByVal vDgv As DataGridView, _
    ByVal vFrm As FrmTransaksiLoan_Payment, Optional ByVal vPrimaryJournal As Double = -1)
        Dim PrimaryJournal As Double = -1
        Dim ObjAcc As New Core_Module.ClsMasterAkun1
        Dim DtaJournalHeader As New DseTransaksiLoanPaymentTableAdapters.tbltrans_jurnalTableAdapter
        Dim DtaJournalItem As New DseTransaksiLoanPaymentTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim TblDummy As New DseTransaksiLoanPayment.Dummy_TableDataTable
        Dim rowDummy As DseTransaksiLoanPayment.Dummy_TableRow
        Dim Dta As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim rowDgv As DataGridViewRow
        Dim cash As Decimal = 0
        Dim rev As Decimal = 0
        Dim gain As Decimal = 0
        Dim loss As Decimal = 0
        Dim PrimaryRev As Double = -1
        Dim desc As String = ""
        Try
            desc = Me.JournalDesc
            'desc = vFrm.TxtDocumentNo.Text + "-" + vFrm.TxtEmployeeCode1.Text + "-" + vFrm.TxtEmployeeName.Text
            If vJournalOtomatis = EnumJurnalOtomatis.Simpan Then
                'baru
                'simpan header
                DtaJournalHeader.Insert(Me.JournalNo, Me.TanggalTransaksi, 0, AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, Me.Primary, _
                 Me.PrimaryKaryawan, desc, vUserCode, Now, vUserCode, Now, Now, -1, -1, 0, 1, 0, 0, 0, 0, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Loan_Payment, _
                ObjAcc.GetPrimaryCurrency(Me.PrimaryCash), AccModule.ClsTransaksiJournal.EnumTipeInOut._In, -1, -1)
                PrimaryJournal = DtaJournalHeader.ScalarQuery_GetPrimaryByDocNo(Me.JournalNo)
                'simpan item
                'pastiin table dummy kosong
                TblDummy.Rows.Clear()
                For Each rowDgv In vDgv.Rows
                    cash = rowDgv.Cells(vFrm.ColPayment.Index).Value
                    rev = rowDgv.Cells(vFrm.ColBase.Index).Value
                    gain = cash - rev
                    loss = rev - cash
                    PrimaryRev = rowDgv.Cells(vFrm.ColPrimaryRev.Index).Value
                    Me.CekJournalItem(gain, cash, rev, loss, TblDummy, rowDummy, PrimaryRev)
                Next
                'simpan ke database
                For Each rowDummy In TblDummy.Rows
                    DtaJournalItem.Insert(PrimaryJournal, rowDummy.Primary_Akun, rowDummy.Debet, rowDummy.Kredit, "", _
                    rowDummy.Debet, rowDummy.Kredit, 1, -1, -1)
                Next
                'update primary jurnal di tabel loan
                Dta.UpdateQuery_PrimaryJurnal(PrimaryJournal, Me.Primary)
                Me.mPrimaryJurnal = PrimaryJournal
            ElseIf vJournalOtomatis = EnumJurnalOtomatis.Edit Then
                'edit
                'simpan header
                DtaJournalHeader.Update(Me.JournalNo, Me.TanggalTransaksi, 0, AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, Me.Primary, _
                 Me.PrimaryKaryawan, desc, vUserCode, Now, Now, -1, -1, 0, 1, 0, 0, 0, 0, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Loan_Payment, _
                ObjAcc.GetPrimaryCurrency(Me.PrimaryCash), AccModule.ClsTransaksiJournal.EnumTipeInOut._In, -1, -1, vPrimaryJournal)
                'hapus item
                DtaJournalItem.Delete(vPrimaryJournal)
                'simpan item
                'pastiin table dummy kosong
                TblDummy.Rows.Clear()
                For Each rowDgv In vDgv.Rows
                    cash = rowDgv.Cells(vFrm.ColPayment.Index).Value
                    rev = rowDgv.Cells(vFrm.ColBase.Index).Value
                    gain = cash - rev
                    loss = rev - cash
                    PrimaryRev = rowDgv.Cells(vFrm.ColPrimaryRev.Index).Value
                    Me.CekJournalItem(gain, cash, rev, loss, TblDummy, rowDummy, PrimaryRev)
                Next
                'simpan ke database
                For Each rowDummy In TblDummy.Rows
                    DtaJournalItem.Insert(vPrimaryJournal, rowDummy.Primary_Akun, rowDummy.Debet, rowDummy.Kredit, "", _
                    rowDummy.Debet, rowDummy.Kredit, 1, -1, -1)
                Next
            ElseIf vJournalOtomatis = EnumJurnalOtomatis.Hapus Then
                'hapus header
                DtaJournalHeader.Delete(vPrimaryJournal)
                'hapus item
                DtaJournalItem.Delete(vPrimaryJournal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaJournalHeader.Dispose()
            DtaJournalItem.Dispose()
            Dta.Dispose()
            ObjAcc = Nothing
        End Try
    End Sub

    Public Sub CekJournalItem(ByVal vGain As Decimal, ByVal vCash As Decimal, ByVal vRev As Decimal, ByVal vLoss As Decimal, _
    ByVal vTblDummy As DseTransaksiLoanPayment.Dummy_TableDataTable, ByVal vRowDummy As DseTransaksiLoanPayment.Dummy_TableRow, _
    ByVal vPrimaryRev As Double)
        Try
            If vGain = 0 Then
                'kasus 1
                'cash = rev
                'cari debet yang sama
                vTblDummy.DefaultView.RowFilter = "[Primary Akun] = " + Me.PrimaryCash.ToString
                If vTblDummy.DefaultView.Count > 0 Then
                    vRowDummy = vTblDummy.DefaultView(0).Row
                    vRowDummy.Debet += vCash
                Else
                    vRowDummy = vTblDummy.NewRow
                    vRowDummy.Primary_Akun = Me.PrimaryCash
                    vRowDummy.Debet = vCash
                    vRowDummy.Kredit = 0
                    vTblDummy.AddDummy_TableRow(vRowDummy)
                End If
                'cari kredit yang sama
                vTblDummy.DefaultView.RowFilter = "[Primary Akun] = " + vPrimaryRev.ToString
                If vTblDummy.DefaultView.Count > 0 Then
                    vRowDummy = vTblDummy.DefaultView(0).Row
                    vRowDummy.Kredit += vRev
                Else
                    vRowDummy = vTblDummy.NewRow
                    vRowDummy.Primary_Akun = vPrimaryRev
                    vRowDummy.Debet = 0
                    vRowDummy.Kredit = vRev
                    vTblDummy.AddDummy_TableRow(vRowDummy)
                End If
            ElseIf vGain > 0 Then
                'cari debet yang sama
                vTblDummy.DefaultView.RowFilter = "[Primary Akun] = " + Me.PrimaryCash.ToString
                If vTblDummy.DefaultView.Count > 0 Then
                    vRowDummy = vTblDummy.DefaultView(0).Row
                    vRowDummy.Debet += vCash
                Else
                    vRowDummy = vTblDummy.NewRow
                    vRowDummy.Primary_Akun = Me.PrimaryCash
                    vRowDummy.Debet = vCash
                    vRowDummy.Kredit = 0
                    vTblDummy.AddDummy_TableRow(vRowDummy)
                End If
                'cari kredit yang sama (rev)
                vTblDummy.DefaultView.RowFilter = "[Primary Akun] = " + vPrimaryRev.ToString
                If vTblDummy.DefaultView.Count > 0 Then
                    vRowDummy = vTblDummy.DefaultView(0).Row
                    vRowDummy.Kredit += vRev
                Else
                    vRowDummy = vTblDummy.NewRow
                    vRowDummy.Primary_Akun = vPrimaryRev
                    vRowDummy.Debet = 0
                    vRowDummy.Kredit = vRev
                    vTblDummy.AddDummy_TableRow(vRowDummy)
                End If
                'cari kredit yang sama (gain)
                vTblDummy.DefaultView.RowFilter = "[Primary Akun] = " + Me.PrimaryGain.ToString
                If vTblDummy.DefaultView.Count > 0 Then
                    vRowDummy = vTblDummy.DefaultView(0).Row
                    vRowDummy.Kredit += vGain
                Else
                    vRowDummy = vTblDummy.NewRow
                    vRowDummy.Primary_Akun = Me.PrimaryGain
                    vRowDummy.Debet = 0
                    vRowDummy.Kredit = vGain
                    vTblDummy.AddDummy_TableRow(vRowDummy)
                End If
            ElseIf vGain < 0 Then
                'kasus 3
                'cash < rev
                'cari debet yang sama(cash)
                vTblDummy.DefaultView.RowFilter = "[Primary Akun] = " + Me.PrimaryCash.ToString
                If vTblDummy.DefaultView.Count > 0 Then
                    vRowDummy = vTblDummy.DefaultView(0).Row
                    vRowDummy.Debet += vCash
                Else
                    vRowDummy = vTblDummy.NewRow
                    vRowDummy.Primary_Akun = Me.PrimaryCash
                    vRowDummy.Debet = vCash
                    vRowDummy.Kredit = 0
                    vTblDummy.AddDummy_TableRow(vRowDummy)
                End If
                'cari debet yang sama (loss)
                vTblDummy.DefaultView.RowFilter = "[Primary Akun] = " + Me.PrimaryLoss.ToString
                If vTblDummy.DefaultView.Count > 0 Then
                    vRowDummy = vTblDummy.DefaultView(0).Row
                    vRowDummy.Debet += vLoss
                Else
                    vRowDummy = vTblDummy.NewRow
                    vRowDummy.Primary_Akun = Me.PrimaryLoss
                    vRowDummy.Debet = vLoss
                    vRowDummy.Kredit = 0
                    vTblDummy.AddDummy_TableRow(vRowDummy)
                End If
                'cari kredit yang sama
                vTblDummy.DefaultView.RowFilter = "[Primary Akun] = " + vPrimaryRev.ToString
                If vTblDummy.DefaultView.Count > 0 Then
                    vRowDummy = vTblDummy.DefaultView(0).Row
                    vRowDummy.Kredit += vRev
                Else
                    vRowDummy = vTblDummy.NewRow
                    vRowDummy.Primary_Akun = vPrimaryRev
                    vRowDummy.Debet = 0
                    vRowDummy.Kredit = vRev
                    vTblDummy.AddDummy_TableRow(vRowDummy)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakaJurnalNonIDR(ByVal vPrimaryAkun As Double) As Boolean
        Dim DTA As New AccModule.DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlNonIDR(Core_Module.PrimaryBaseCurrency, vPrimaryAkun).HasValue = True Then
                If DTA.ScalarQuery_GetJmlNonIDR(Core_Module.PrimaryBaseCurrency, vPrimaryAkun).Value <= 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CetakJournal(ByVal vPrimaryJournal As Double, _
    ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vDGV As DataGridView, _
    ByVal vFrm As FrmTransaksiLoan_Payment) As Boolean
        Dim Ds As New DseTransaksiLoanPayment
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.DataTable3TableAdapter
        Dim TBL As DseTransaksiLoanPayment.DataTable3DataTable = Ds.DataTable3
        Dim FormatDecimal As Integer = 0
        Dim Negatif As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Report As New CrpCetakJurnalDetail
        'Dim Report2 As New CrystalReport1
        Dim Report2 As New CrpCetakJurnalDetail2
        Dim Report3 As New CrpCetakJurnalDetail3
        Dim CetakJurnalPortrait As String = ""
        Dim KodeCur As String = ""
        Dim row As DseTransaksiLoanPayment.DataTable3Row
        Dim rowDGV As DataGridViewRow
        Dim rowDetail As DseTransaksiLoanPayment.TblDetailLoanRow
        Dim InstallmentAt_String As String = ""
        Try
            'isikan data detail invoice
            Ds.TblDetailLoan.Rows.Clear()
            For Each rowDGV In vDGV.Rows
                'If CDec(rowDGV.Cells(vFrm.ColReceivementOriginal.Index).Value) <> 0 Then
                rowDetail = Ds.TblDetailLoan.NewTblDetailLoanRow()

                rowDetail.Loan_Number = rowDGV.Cells(vFrm.ColLoanNo.Index).Value
                rowDetail.Installment_at = rowDGV.Cells(vFrm.ColInstallementAt.Index).Value
                rowDetail.Base = rowDGV.Cells(vFrm.ColBase.Index).Value
                rowDetail.Rate = rowDGV.Cells(vFrm.ColInterest.Index).Value
                rowDetail.Payment = rowDGV.Cells(vFrm.ColPayment.Index).Value
                Ds.TblDetailLoan.Rows.Add(rowDetail)
                'End If
                If InstallmentAt_String = "" Then
                    InstallmentAt_String = "Installment at "
                    InstallmentAt_String += rowDetail.Installment_at.ToString
                Else
                    InstallmentAt_String += ", " + rowDetail.Installment_at.ToString
                End If
            Next

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryJournal)
            'Tambahan untuk description 
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    row.Desc = vFrm.TxtDocumentNo.Text + " - " + vFrm.TxtEmployeeCode1.Text + " - " + vFrm.TxtEmployeeName.Text + " - " + InstallmentAt_String
                Next
            End If

            KodeCur = ObjCur.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)

            CetakJurnalPortrait = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CetakJurnalPortrait)

            If CetakJurnalPortrait = "" Or CetakJurnalPortrait = "0" Then
                If Me.ApakaJurnalNonIDR(vPrimaryJournal) = False Then
                    'cetak lanscape tanda currency no idr
                    Report2.SetDataSource(Ds)
                    Report2.SetParameterValue(Report2.Parameter_FormatDecimal.ParameterFieldName, ObjCur.GetDecimal_DefaultCurrency)
                    If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                        Negatif = 1
                    Else
                        Negatif = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                        Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                    End If
                    Report2.SetParameterValue(Report2.Parameter_Negatif.ParameterFieldName, Negatif)
                    Report2.SetParameterValue(Report2.Parameter_KodeBaseCurrency.ParameterFieldName, KodeCur)

                    Crv.ReportSource = Report2
                    'Crv.Zoom(1)
                    Crv.Show()
                Else
                    Report3.SetDataSource(Ds)
                    For Each row In Ds.DataTable3.Rows
                        If row.Credit_Original > 0 Then
                            row.Debet_Original = row.Credit_Original
                        End If
                    Next
                    Report3.SetParameterValue(Report3.Parameter_FormatDecimal.ParameterFieldName, ObjCur.GetDecimal_DefaultCurrency)
                    If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                        Negatif = 1
                    Else
                        Negatif = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                        Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                    End If
                    Report3.SetParameterValue(Report3.Parameter_Negatif.ParameterFieldName, Negatif)
                    Report3.SetParameterValue(Report3.Parameter_KodeBaseCurrency.ParameterFieldName, KodeCur)

                    Crv.ReportSource = Report3
                    'Crv.Zoom(1)
                    Crv.Show()
                End If
            Else
                'cetak portrait
                Report.SetDataSource(DS)
                Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, ObjCur.GetDecimal_DefaultCurrency)
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                    Negatif = 1
                Else
                    Negatif = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                End If
                Report.SetParameterValue(Report.Parameter_Negatif.ParameterFieldName, Negatif)
                Report.SetParameterValue(Report.Parameter_KodeBaseCurrency.ParameterFieldName, KodeCur)

                Crv.ReportSource = Report
                Crv.Show()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Ds.Dispose()
            DTA.Dispose()
            ObjCur = Nothing
            ObjSetting = Nothing
        End Try
    End Function

    Public Function GetPrimaryByPrimarySalary(ByVal vPrimarySalary As Double) As Double
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryByPrimarySalary(vPrimarySalary)
            If Dummy.HasValue Then
                Return Dummy.Value
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

    Public Function GetPrimarySalary(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim Dummy As System.Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimarySalary(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
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

    Public Sub UpdatePrimarySalary(ByVal vPrimarySalary As Double, ByVal vPrimary As Double)
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetPrimarySalaryDiTransLoanPayment(vPrimarySalary, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryJournal(ByVal vPrimaryLoanPayment As Double) As Double
        Dim d As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim n As New Nullable(Of Double)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            n = d.ScalarQuery_GetPrimaryJournal(vPrimaryLoanPayment)
            If n.HasValue Then Return n.Value
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    'fungsi ini untuk update no jurnal saat no document ke ubah karena antrian simpan berbarengan
    'selain update no document update juga no jurnal jika no jurnal sama dengan nilai buffer no document
    Private Sub UpdateJournalNumber(ByVal vBufferDoc As String)
        Try
            If Me.JournalNo = vBufferDoc Then
                Me.JournalNo = Me.DocumentNo
            Else
                If Me.ObjJurnal.ApakahSudahAda(Me.JournalNo) = True Then
                    Me.JournalNo = Me.DocumentNo
                End If
            End If
            If Me.JournalDesc.Contains(vBufferDoc) = True Then
                Me.JournalDesc = Me.JournalDesc.Replace(vBufferDoc, Me.DocumentNo)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
