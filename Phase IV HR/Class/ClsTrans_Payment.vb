Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing
Public Class ClsTrans_Payment

    Private mPrimary As Double = -1
    Private mDocumentNo As String = ""
    Private mJournalNo As String = ""
    Private mJournalDesc As String = ""
    Private mNoUrut As Integer = 0
    Private mTanggalTransaksi As Date = #1/1/1950#
    Private mStartDate As Date = #1/1/1950#
    Private mEndDate As Date = #1/1/1950#
    Private mRemarks As String = ""
    Private mPrimaryJurnalRef As Double = -1

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

    Public Property NoUrut() As Integer
        Get
            Return mNoUrut
        End Get
        Set(ByVal value As Integer)
            mNoUrut = value
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

    Public Property StartDate() As Date
        Get
            Return mStartDate
        End Get
        Set(ByVal value As Date)
            mStartDate = value
        End Set
    End Property

    Public Property EndDate() As Date
        Get
            Return mEndDate
        End Get
        Set(ByVal value As Date)
            mEndDate = value
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

    Public Property PrimaryJurnalRef() As Double
        Get
            Return mPrimaryJurnalRef
        End Get
        Set(ByVal value As Double)
            mPrimaryJurnalRef = value
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

    Public Function GetDocumentNo_ByPrimarySalary(ByVal vPrimarySalary As Double) As String
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_payment_item_employeeTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetDocumentNo_ByPrimarySalary(vPrimarySalary)
            If String.IsNullOrEmpty(Dummy) = True Then
                Return ""
            Else
                Return Dummy
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryPayment_ByPrimarySalary(ByVal vPrimarySalary As Double) As Double
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_payment_item_employeeTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryPayment1(vPrimarySalary)
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

    Public Function GetPrimary(ByVal vDocumentNo As String) As Double
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
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
        End Try
    End Function

    Public Function GetPrimaryJurnalRef(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryJournalRef(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
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
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
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

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
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
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
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

    Public Function GetMAXUrutan(ByVal vEndDate As Date, ByVal vStartDate As Date) As Integer
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNoUrutMAX(vEndDate, vStartDate)
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

    Public Function GetJmlByDocumentNo(ByVal vDocumentNo As String) As Double
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByNoTransaksi(vDocumentNo)
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

#Region "Standar Coding"
    Public Function CariDetail(ByVal vPrimary As Double, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vDGVSalary As DataGridView, ByVal vDGVPayment As DataGridView, ByVal vDGVJurnalItem As DataGridView, _
    ByVal vDef_Prec As Double, ByVal i As FrmTransaksiPayment) As Boolean
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
        Dim TBL As New DseTransPayment.tbltrans_paymentDataTable
        Dim row As DseTransPayment.tbltrans_paymentRow
        Dim ObjItem As New ClsTrans_Payment_Item
        Dim IndexSalaryRow As Integer = -1
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mDocumentNo = .Document_No
                    Me.mTanggalTransaksi = ._date
                    Me.mNoUrut = .NoUrut
                    Me.mPrimaryJurnalRef = .primary_jurnal_ref
                    Me.mStartDate = .start_date
                    Me.mEndDate = .end_date
                    Me.mRemarks = .remarks
                    Me.mJournalNo = Me.ObjJurnal.GetRefCode(.primary_jurnal_ref)
                    Me.mJournalDesc = Me.ObjJurnal.GetDescription(.primary_jurnal_ref)

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
                'cari data item
                ObjItem.IsiDGVItemSalary(Primary, vDGVSalary, vDef_Prec, i)
                If vDGVSalary.Rows.Count > 0 Then
                    If vDGVSalary.SelectedCells.Count > 0 Then
                        IndexSalaryRow = vDGVSalary.SelectedCells(0).RowIndex
                    Else
                        IndexSalaryRow = -1
                    End If
                End If
                ObjItem.IsiDGVPayment(Primary, vDGVPayment, vDGVSalary, vDef_Prec, IndexSalaryRow, i)
                ObjItem.IsiDGVItemJurnal(Primary, vDGVJurnalItem, vDef_Prec, i, Me)
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
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_payment1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByDocumentNoLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByEmployeeCode Then
                Return DTA.ScalarQuery_GetJmlByEmployeeCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByEmployeeName Then
                Return DTA.ScalarQuery_GetJmlByFullNameLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByDate Then
                Return DTA.ScalarQuery_GetJmlByTanggalDariSampai(vEndDate, vStartDate).Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.BySalaryNo Then
                Return DTA.ScalarQuery_GetJmlBySalaryNoLIKE(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransPayment.tbltrans_payment1DataTable, _
    ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Boolean
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_payment1TableAdapter
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
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByEmployeeCode Then
                QueryResult = DTA.FillBy_EmployeeCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByEmployeeName Then
                QueryResult = DTA.FillBy_EmployeeNameLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.BySalaryNo Then
                QueryResult = DTA.FillBy_SalaryNoLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(3).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(4).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(5).HeaderText = "Remarks"
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
    ByVal vPrimary As Double, ByVal vDGVSalary As DataGridView, ByVal vDGVPayment As DataGridView, _
    ByVal vDGVJurnalItem As DataGridView, ByVal ij As FrmTransaksiPayment) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Payment_SimpanBaru, vUserCode)
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
                            retval = Me.SimpanBaru1(vUserCode, vObjCD, vPrimary, vDGVSalary, vDGVPayment, vDGVJurnalItem, ij)
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
    ByVal vPrimary As Double, ByVal vDGVSalary As DataGridView, ByVal vDGVPayment As DataGridView, _
    ByVal vDGVJurnalItem As DataGridView, ByVal ij As FrmTransaksiPayment) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Payment_SimpanPerubahan, vUserCode)

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
                            retval = Me.SimpanPerubahan1(vUserCode, vObjCD, vPrimary, vDGVSalary, vDGVPayment, vDGVJurnalItem, ij)
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
    ByVal vPrimary As Double, ByVal vDGVSalary As DataGridView, ByVal vDGVPayment As DataGridView, _
    ByVal vDGVJurnal As DataGridView, ByVal ij As FrmTransaksiPayment, _
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
            Me.vPrimary_Antrian, vPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Payment_Hapus, vUserCode)

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
                            retval = Me.Hapus1(vPrimary, vUserCode, vDGVSalary, True, TampilkanPesan, ij)
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
    ByVal vDGVSalary As DataGridView, ByVal vDGVPayment As DataGridView, _
    ByVal vDGVJurnalItem As DataGridView, ByVal i As FrmTransaksiPayment) As Boolean
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
        Dim ObjItem As New ClsTrans_Payment_Item
        Dim ObjLoanPay As New ClsTransaksiLoanPayment
        Dim ObjSalary As New ClsTransaksiSalary
        Dim RowDGV As DataGridViewRow
        Dim BufferDoc As String = Me.mDocumentNo
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, vPrimary, True, vObjCD, vDGVSalary, vDGVPayment, i) = True Then
                Me.mNoUrut = CDec(Me.mDocumentNo.Substring(0, 6))
                'Insert Data
                If DTA.Insert(Me.mTanggalTransaksi, Me.mStartDate, Me.mEndDate, Me.mRemarks, _
                Me.mPrimaryJurnalRef, vUserCode, Now, vUserCode, Now, Now, Me.mDocumentNo, Me.mNoUrut) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mDocumentNo)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, vDGVSalary, i)
                    vDGVSalary.AllowUserToAddRows = False
                    If vDGVSalary.Rows.Count > 0 Then
                        For Each RowDGV In vDGVSalary.Rows
                            ObjSalary.UpdatePayment(Me.mTanggalTransaksi, ClsTransaksiSalary.EnumPaymentStatus.paid, _
                            Me.mPrimary, CDec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value))
                        Next
                    End If
                    vDGVSalary.AllowUserToAddRows = True
                    ObjItem.SimpanItemPayment(Me.mPrimary, i.mTBLItem_Employee_d, i)
                    ObjItem.SimpanItemJurnal(Me.mPrimary, vDGVJurnalItem, i)

                    'Journal Otomatis
                    Me.UpdateJournalNumber(BufferDoc)
                    Me.JurnalOtomatis(vUserCode, EnumJurnalOtomatis.Simpan, i)

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
            ObjSalary = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUserCode As String, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vPrimary As Double, _
    ByVal vDGVSalary As DataGridView, ByVal vDGVPayment As DataGridView, _
    ByVal vDGVJurnalItem As DataGridView, ByVal i As FrmTransaksiPayment) As Boolean
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
        Dim DTAItem_Salary As New DseTransPaymentTableAdapters.tbltrans_payment_item_employeeTableAdapter
        Dim TBLItem_Salary As New DseTransPayment.tbltrans_payment_item_employeeDataTable
        Dim RowItem_Salary As DseTransPayment.tbltrans_payment_item_employeeRow
        Dim ObjItem As New ClsTrans_Payment_Item
        Dim ObjSalary As New ClsTransaksiSalary
        Dim RowDGV As DataGridViewRow
        Dim BufferDoc As String = Me.mDocumentNo

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem_Salary.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False, vPrimary, True, vObjCD, vDGVSalary, vDGVPayment, i) = True Then
                Me.mNoUrut = CDec(Me.mDocumentNo.Substring(0, 6))
                If DTA.Update(Me.mTanggalTransaksi, Me.mStartDate, _
                Me.mEndDate, Me.mRemarks, Me.mPrimaryJurnalRef, vUserCode, Now, Now, _
                Me.mDocumentNo, Me.mNoUrut, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    DTAItem_Salary.FillBy_PrimaryTransaksi(TBLItem_Salary, vPrimary)
                    If TBLItem_Salary.Rows.Count > 0 Then
                        For Each RowItem_Salary In TBLItem_Salary.Rows
                            ObjSalary.UpdatePayment(#1/1/1950#, ClsTransaksiSalary.EnumPaymentStatus.unpaid, _
                            -1, RowItem_Salary.primary_transaksi_salary)
                        Next
                    End If
                    ObjItem.HapusItem(Me.mPrimary)
                    ObjItem.HapusItemPayment(Me.mPrimary)
                    ObjItem.HapusItemJurnal(Me.mPrimary)

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, vDGVSalary, i)
                    vDGVSalary.AllowUserToAddRows = False
                    If vDGVSalary.Rows.Count > 0 Then
                        For Each RowDGV In vDGVSalary.Rows
                            ObjSalary.UpdatePayment(Me.mTanggalTransaksi, ClsTransaksiSalary.EnumPaymentStatus.paid, _
                            Me.mPrimary, CDec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value))
                        Next
                    End If
                    vDGVSalary.AllowUserToAddRows = True
                    ObjItem.SimpanItemPayment(Me.mPrimary, i.mTBLItem_Employee_d, i)
                    ObjItem.SimpanItemJurnal(Me.mPrimary, vDGVJurnalItem, i)

                    'Journal Otomatis
                    Me.UpdateJournalNumber(BufferDoc)
                    Me.JurnalOtomatis(vUserCode, EnumJurnalOtomatis.Edit, i)

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
            ObjSalary = Nothing
            DTAItem_Salary.Dispose()
            TBLItem_Salary.Dispose()
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vDGVSalary As DataGridView, ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal i As FrmTransaksiPayment) As Boolean
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
        Dim ObjItem As New ClsTrans_Payment_Item
        Dim ObjSalary As New ClsTransaksiSalary
        Dim RowDGV As DataGridViewRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, vStrictDelete, vShowMsg, vDGVSalary, i) = True Then
                'Update status = unpaid, primary payment = -1, Tanggal Paid = 1/1/1950, dimasing-masing transaksi salary
                vDGVSalary.AllowUserToAddRows = False
                If vDGVSalary.Rows.Count > 0 Then
                    For Each RowDGV In vDGVSalary.Rows
                        ObjSalary.UpdatePayment(#1/1/1950#, ClsTransaksiSalary.EnumPaymentStatus.unpaid, _
                        -1, RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value)
                    Next
                End If
                vDGVSalary.AllowUserToAddRows = True
                'Hapus item
                ObjItem.HapusItem(vPrimary)
                ObjItem.HapusItemPayment(vPrimary)
                ObjItem.HapusItemJurnal(vPrimary)
                'Hapus
                DTA.Delete(vPrimary)
                'Journal Otomatis
                If Me.PrimaryJurnalRef <> -1 Then
                    Me.JurnalOtomatis(vUserCode, EnumJurnalOtomatis.Hapus, i, Me.PrimaryJurnalRef)
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjSalary = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvSalary As DataGridView, ByVal vDGVPayment As DataGridView, _
    ByVal vDGVJurnal As DataGridView)
        Try
            mPrimary = -1
            mDocumentNo = ""
            mTanggalTransaksi = #1/1/1950#
            mNoUrut = 0
            mStartDate = #1/1/1950#
            mEndDate = #1/1/1950#
            mRemarks = ""
            mPrimaryJurnalRef = -1

            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today

            'Item
            vDgvSalary.Rows.Clear()
            vDGVPayment.Rows.Clear()
            vDGVJurnal.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictEdit As Boolean, ByVal vShowMsg As Boolean, ByVal vDGVSalary As DataGridView, ByVal i As FrmTransaksiPayment) As Boolean
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
        Dim ObjJurnal As New AccModule.ClsTransaksiJournal
        Dim pCreatedBy As String = ""
        'Dim RowDGV As DataGridViewRow
        Dim Str As String = ""
        Dim StrNo As Integer = 1
        Dim PrimaryMinSalary2 As Double = -1
        Dim ObjSalary2 As New HR.ClsTransaksiSalary
        Dim TBLSalary2 As New DseTransaksiSalary.tbltrans_salary2DataTable
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
            '2. Berkaitan dengan transaksi jurnal
            If ObjJurnal.ApakahBolehEdit(Me.mPrimaryJurnalRef, vUserCode, True, True) = False Then
                Return False
            End If

            '3. Pengecekan Status = Payment
            'Str = ""
            'StrNo = 1
            'vDGVSalary.AllowUserToAddRows = False
            'If vDGVSalary.Rows.Count > 0 Then
            '    For idx As Integer = 0 To vDGVSalary.Rows.Count - 1
            '        RowDGV = vDGVSalary.Rows(idx)
            '        If idx = 0 Then
            '            PrimaryMinSalary2 = cdec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value)
            '        Else
            '            If cdec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value) < PrimaryMinSalary2 Then
            '                PrimaryMinSalary2 = cdec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value)
            '            End If
            '        End If
            '    Next
            '    ObjSalary2.GetTBLPengecekanStatusPayment(TBLSalary2, PrimaryMinSalary2)
            '    For Each RowDGV In vDGVSalary.Rows
            '        If ObjSalary2.GetPaymentStatus(TBLSalary2, cdec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value)) = ClsTransaksiSalary.EnumPaymentStatus.paid Then
            '            If String.IsNullOrEmpty(Str) = True Then
            '                Str = StrNo.ToString + ". " + RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value.ToString + " - " _
            '                + RowDGV.Cells(i.ColEmp_EmployeeCode.Index).Value.ToString + " - " + _
            '                RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString + "."
            '            Else
            '                Str += vbNewLine + StrNo.ToString + ". " + RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value.ToString + " - " _
            '                + RowDGV.Cells(i.ColEmp_EmployeeCode.Index).Value.ToString + " - " + _
            '                RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString + "."
            '            End If
            '        StrNo += 1
            '        End If
            '    Next
            '    If String.IsNullOrEmpty(Str) = False Then
            '        MsgBox("The following salary have already paid : " + vbNewLine + Str, MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If
            'vDGVSalary.AllowUserToAddRows = True

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjJurnal = Nothing
            ObjSalary2 = Nothing
            TBLSalary2.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vDGVSalary As DataGridView, ByVal i As FrmTransaksiPayment) As Boolean
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
        Dim pCreatedBy As String = ""
        Dim ObjItem As New ClsTransaksiLoanPayment_Item
        Dim ObjJurnal As New AccModule.ClsTransaksiJournal
        Dim InfoData As String = ""
        'Dim RowDGV As DataGridViewRow
        Dim Str As String = ""
        Dim StrNo As Integer = 1
        Dim PrimaryMinSalary2 As Double = -1
        Dim ObjSalary2 As New HR.ClsTransaksiSalary
        Dim TBLSalary2 As New DseTransaksiSalary.tbltrans_salary2DataTable
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

            '2. Berkaitan dengan Transaksi Jurnal
            If Me.mPrimaryJurnalRef <> -1 Then
                If ObjJurnal.ApakahBolehHapus(Me.mPrimaryJurnalRef, True, vUserCode, True) = False Then
                    Return False
                End If
            End If

            '3. Pengecekan Status = Payment
            'Str = ""
            'StrNo = 1
            'vDGVSalary.AllowUserToAddRows = False
            'If vDGVSalary.Rows.Count > 0 Then
            '    For idx As Integer = 0 To vDGVSalary.Rows.Count - 1
            '        RowDGV = vDGVSalary.Rows(idx)
            '        If idx = 0 Then
            '            PrimaryMinSalary2 = cdec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value)
            '        Else
            '            If cdec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value) < PrimaryMinSalary2 Then
            '                PrimaryMinSalary2 = cdec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value)
            '            End If
            '        End If
            '    Next
            '    ObjSalary2.GetTBLPengecekanStatusPayment(TBLSalary2, PrimaryMinSalary2)
            '    For Each RowDGV In vDGVSalary.Rows
            '        If ObjSalary2.GetPaymentStatus(TBLSalary2, cdec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value)) = ClsTransaksiSalary.EnumPaymentStatus.paid Then
            '            If String.IsNullOrEmpty(Str) = True Then
            '                Str = StrNo.ToString + ". " + RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value.ToString + " - " _
            '                + RowDGV.Cells(i.ColEmp_EmployeeCode.Index).Value.ToString + " - " + _
            '                RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString + "."
            '            Else
            '                Str += vbNewLine + StrNo.ToString + ". " + RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value.ToString + " - " _
            '                + RowDGV.Cells(i.ColEmp_EmployeeCode.Index).Value.ToString + " - " + _
            '                RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString + "."
            '            End If
            '            StrNo += 1
            '        End If
            '    Next
            '    If String.IsNullOrEmpty(Str) = False Then
            '        MsgBox("The following salary have already paid : " + vbNewLine + Str, MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If
            'vDGVSalary.AllowUserToAddRows = True

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjJurnal = Nothing
            ObjSalary2 = Nothing
            TBLSalary2.Dispose()
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vPrimary As Double, _
        ByVal vShowMsg As Boolean, ByVal vObjCD As Core_Module.ClsCekData, _
        ByVal vDGVSalary As DataGridView, ByVal vDGVPayment As DataGridView, ByVal i As FrmTransaksiPayment) As Boolean
        Dim ObjItem As New ClsTransaksiSalary_Item
        Dim ij As Integer = 0
        Dim ListString As New List(Of String)
        Dim TanggalStart As Date = i.dtpPeriod.Value
        Dim TanggalEnd As Date = i.DtpTo.Value
        Dim FrmAtt As New FrmTrans_Gaji_ShowAttendance
        'Dim RowDGV As DataGridViewRow
        Dim Str As String = ""
        Dim StrNo As Integer = 1
        Dim PrimaryMinSalary2 As Double = -1
        Dim ObjSalary2 As New HR.ClsTransaksiSalary
        Dim TBLSalary2 As New DseTransaksiSalary.tbltrans_salary2DataTable
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Dim ObjJurnal As New AccModule.ClsTransaksiJournal
        Dim ObjCD As New Core_Module.ClsCekData
        Try
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
                        Me.mDocumentNo = Me.GetNoBaru(Me.mTanggalTransaksi, vObjCD)
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. DGV Salary
            vDGVSalary.AllowUserToAddRows = False
            If vDGVSalary.Rows.Count <= 0 Then
                If vShowMsg = True Then
                    MsgBox("Please calculate salary, minimal 1 item salary.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            vDGVSalary.AllowUserToAddRows = True
            '2. Remarks
            If vObjCD.CekDataKosong(Me.mRemarks, False) = False Then
                If vShowMsg = True Then
                    MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '3. Batasan jika jurnal telah diclosing 
            'jika sudah ada closing jurnal untuk tahun ybs maka tidak bisa input data baru
            If ObjJurnal.GetJml(AccModule.ClsTransaksiJournal.EnumTipeJournal.ClosingJournal, ObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi), _
            ObjCD.GetLastDateOfTheYear(Me.mTanggalTransaksi)) > 0 Then
                MsgBox("There are closing jurnal for this year you cannot save this data", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '4. Pengecekan Status = Payment
            'Str = ""
            'StrNo = 1
            'vDGVSalary.AllowUserToAddRows = False
            'If vDGVSalary.Rows.Count > 0 Then
            '    For idx As Integer = 0 To vDGVSalary.Rows.Count - 1
            '        RowDGV = vDGVSalary.Rows(idx)
            '        If idx = 0 Then
            '            PrimaryMinSalary2 = cdec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value)
            '        Else
            '            If cdec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value) < PrimaryMinSalary2 Then
            '                PrimaryMinSalary2 = cdec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value)
            '            End If
            '        End If
            '    Next
            '    ObjSalary2.GetTBLPengecekanStatusPayment(TBLSalary2, PrimaryMinSalary2)
            '    For Each RowDGV In vDGVSalary.Rows
            '        If ObjSalary2.GetPaymentStatus(TBLSalary2, cdec(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value)) = ClsTransaksiSalary.EnumPaymentStatus.paid Then
            '            If String.IsNullOrEmpty(Str) = True Then
            '                Str = StrNo.ToString + ". " + RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value.ToString + " - " _
            '                + RowDGV.Cells(i.ColEmp_EmployeeCode.Index).Value.ToString + " - " + _
            '                RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString + "."
            '            Else
            '                Str += vbNewLine + StrNo.ToString + ". " + RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value.ToString + " - " _
            '                + RowDGV.Cells(i.ColEmp_EmployeeCode.Index).Value.ToString + " - " + _
            '                RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString + "."
            '            End If
            '            StrNo += 1
            '        End If
            '    Next
            '    If String.IsNullOrEmpty(Str) = False Then
            '        MsgBox("The following salary have already paid : " + vbNewLine + Str, MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If
            'vDGVSalary.AllowUserToAddRows = True

            '5. Batasan Load Ulang Data Salary
            Dim PrimaryMin As Double = Double.MaxValue
            Dim PrimaryMax As Double = Double.MinValue
            i.GetNilaiPrimarySalary(PrimaryMin, PrimaryMax)
            If Me.PengecekanLoadUlangSalary(PrimaryMin, PrimaryMax, vDGVSalary, i) = False Then
                Return False
            End If

            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjItem = Nothing
            ListString = Nothing
            ObjJurnal = Nothing
            ObjCD = Nothing
            ObjSalary2 = Nothing
            FrmAtt.Dispose()
            TBLSalary2.Dispose()
        End Try
    End Function

    Public Function GetJmlByTanggalTransaksi(ByVal vEndDate As Date, ByVal vStartDate As Date) As Double
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByTanggalDariSampai(vEndDate, vStartDate)
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
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim InOut As String = "PAY"
        NoUrut = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = vObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = vObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetNoUrutMAX(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetNoUrutMAX(TanggalEnd, TanggalStart).Value + 1)
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
#End Region

    Public Function PengecekanLoadUlangSalary(ByVal vPrimarySalaryMin As Double, ByVal vPrimarySalaryMax As Double, _
    ByVal vDGVSalary As DataGridView, ByVal i As FrmTransaksiPayment) As Boolean
        Dim DTASalary As New DseTransPaymentTableAdapters.tbltrans_salaryTableAdapter
        Dim TBLSalary As New DseTransPayment.tbltrans_salaryDataTable
        Dim RowSalary As DseTransPayment.tbltrans_salaryRow
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim Retval As String = ""
        Dim Ketemu As Boolean = False
        Dim Counter As Integer = 1
        Try
            DTASalary.Connection.ConnectionString = Core_Module.GetConstring
            i.dgvItem.AllowUserToAddRows = False
            DTASalary.FillBy_PrimaryMinMax(TBLSalary, vPrimarySalaryMin, vPrimarySalaryMax)
            If TBLSalary.Rows.Count > 0 Then
                For Each RowDGV In vDGVSalary.Rows
                    Filter = "[Primary] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLSalary.DefaultView.RowFilter = Filter
                    Ketemu = False
                    If TBLSalary.DefaultView.Count > 0 Then
                        RowSalary = TBLSalary.DefaultView(0).Row
                        'Cek satu persatu nilai-nya
                        If RowSalary.Income <> Convert.ToDecimal(RowDGV.Cells(i.ColEmp_Income.Index).Value) Then
                            Ketemu = True
                        End If
                        If RowSalary.Deduction <> Convert.ToDecimal(RowDGV.Cells(i.ColEmp_Deduction.Index).Value) Then
                            Ketemu = True
                        End If
                        If RowSalary.Gross_Salary <> Convert.ToDecimal(RowDGV.Cells(i.ColEmp_GrossSalary.Index).Value) Then
                            Ketemu = True
                        End If
                        If RowSalary.Income_Tax <> Convert.ToDecimal(RowDGV.Cells(i.ColEmp_IncomeTax.Index).Value) Then
                            Ketemu = True
                        End If
                        If RowSalary.Loan_Installment <> Convert.ToDecimal(RowDGV.Cells(i.ColEmp_LoanInstallment.Index).Value) Then
                            Ketemu = True
                        End If
                        If RowSalary.Other_Deduction <> Convert.ToDecimal(RowDGV.Cells(i.ColEmp_OtherDeduction.Index).Value) Then
                            Ketemu = True
                        End If
                        If RowSalary.Salary_Received <> Convert.ToDecimal(RowDGV.Cells(i.ColEmp_TakeHomePay.Index).Value) Then
                            Ketemu = True
                        End If
                    Else
                        'Berarti data salary sudah dihapus, jadi nilai-nilainya sama dengan nol
                    End If
                    If Ketemu = True Then
                        If String.IsNullOrEmpty(Retval) = False Then
                            Retval += vbNewLine + Counter.ToString + ", " + RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value.ToString + "."
                        Else
                            Retval = Counter.ToString + ", " + RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value.ToString + "."
                        End If
                        Counter += 1
                    End If
                Next
            End If
            i.dgvItem.AllowUserToAddRows = True
            If String.IsNullOrEmpty(Retval) = False Then
                MsgBox("The following salary document is invalid : " + vbNewLine + Retval + vbNewLine + "Re-calculating data.", MsgBoxStyle.Information, "Information")
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahDataPembayaranAmountNyaSesuaiDenganTakeHomePay(ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vDGVSalary As DataGridView, ByVal mTBLTrans_Payment_Employee_d As DseTransPayment.tbltrans_payment_item_employee_dDataTable, _
    ByVal i As FrmTransaksiPayment) As Boolean
        Dim Filter As String = ""
        Dim Nmr As Integer = 1
        Dim Retval As String = ""
        Dim TotalAmount As Decimal = 0
        Dim RowView As Data.DataRowView
        Dim RowDGV As DataGridViewRow
        Dim RowPay As DseTransPayment.tbltrans_payment_item_employee_dRow
        Try
            vDGVSalary.AllowUserToAddRows = False
            If vDGVSalary.Rows.Count > 0 Then
                For Each RowDGV In vDGVSalary.Rows
                    Filter = "[identifier] = '" + RowDGV.Cells(i.ColIdentifier.Index).Value.ToString + "'"
                    mTBLTrans_Payment_Employee_d.DefaultView.RowFilter = Filter
                    TotalAmount = 0
                    If mTBLTrans_Payment_Employee_d.DefaultView.Count > 0 Then
                        For Each RowView In mTBLTrans_Payment_Employee_d.DefaultView
                            RowPay = RowView.Row
                            TotalAmount += RowPay.amount
                        Next
                        If CDec(RowDGV.Cells(i.ColEmp_TakeHomePay.Index).Value) <> CDec(TotalAmount) Then
                            If String.IsNullOrEmpty(Retval) = False Then
                                Retval += vbNewLine + Nmr.ToString + ". Salary no : " + RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value.ToString + ", take home pay : " + RowDGV.Cells(i.ColEmp_TakeHomePay.Index).Value.ToString + ", payment : " + vObjCD.DigitGrouping(TotalAmount, i.Def_Prec).ToString + "."
                            Else
                                Retval = Nmr.ToString + ". Salary no : " + RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value.ToString + ", take home pay : " + RowDGV.Cells(i.ColEmp_TakeHomePay.Index).Value.ToString + ", payment : " + vObjCD.DigitGrouping(TotalAmount, i.Def_Prec).ToString + "."
                            End If
                            Nmr += 1
                        End If
                    Else
                        Retval = Nmr.ToString + ". Salary no : " + RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value.ToString + ", take home pay : " + RowDGV.Cells(i.ColEmp_TakeHomePay.Index).Value.ToString + ", payment : " + vObjCD.DigitGrouping(TotalAmount, i.Def_Prec).ToString + "."
                    End If
                Next
            End If
            mTBLTrans_Payment_Employee_d.DefaultView.RowFilter = ""
            vDGVSalary.AllowUserToAddRows = True
            If String.IsNullOrEmpty(Retval) = False Then
                MsgBox("The following payment isn't suitable from take home pay value : " + vbNewLine + Retval, MsgBoxStyle.Information, "Information")
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimarySalaryMIN(ByVal vPrimaryPayment As Double) As Double
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_payment_item_employeeTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryMinSalary(vPrimaryPayment)
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

    Public Function GetPrimarySalaryMAX(ByVal vPrimaryPayment As Double) As Double
        Dim DTA As New DseTransPaymentTableAdapters.tbltrans_payment_item_employeeTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryMAXSalary(vPrimaryPayment)
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

    Public Function ApakahBolehInputSalary(ByVal vPrimaryPayroll As Double, _
    ByVal i As FrmTransaksiPayment) As Boolean
        Dim DTAUser As New DseTransPaymentTableAdapters.TblPayrollAccessTableAdapter
        Dim TBLUser As New DseTransPayment.TblPayrollAccessDataTable
        Dim ObjSalary As New ClsTransaksiSalary
        Dim PrimaryFormula As Double = -1
        Dim TanggalSalary As Date
        Dim Salah As Boolean = False
        Dim Filter As String = ""
        Dim Retval As String = ""
        Try
            DTAUser.Connection.ConnectionString = Core_Module.GetConstring
            DTAUser.FillBy_KodeUser(TBLUser, i.UserCode)
            PrimaryFormula = ObjSalary.GetPrimarySalaryFormula(vPrimaryPayroll)
            Filter = "[Primary Salary Formula] = " + PrimaryFormula.ToString + ""
            TBLUser.DefaultView.RowFilter = Filter
            If TBLUser.Rows.Count <= 0 Then
                '1. Tidak boleh input jika tidak punya payroll akses atas salary formula.
                Retval = "Salary no " & ObjSalary.GetNoTransaksi(vPrimaryPayroll).ToString & " isn't payroll access."
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                MsgBox(Retval, MsgBoxStyle.Information, "Information")
                Return False
            End If
            '2. Harus sesuai dengan Period from dan period to.
            TanggalSalary = ObjSalary.GetTanggalTransaksi(vPrimaryPayroll)
            If TanggalSalary > i.DtpTo.Value.Date Then
                Salah = True
            Else
                If TanggalSalary < i.dtpPeriod.Value.Date Then
                    Salah = True
                End If
            End If
            If Salah = True Then
                MsgBox("Please input salary data between period from " + Format(i.dtpPeriod.Value.Date, "dd MMM yyyy") + " to " _
                + Format(i.DtpTo.Value.Date, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAUser.Dispose()
            TBLUser.Dispose()
            ObjSalary = Nothing
        End Try
    End Function

    Public Function ApakahPunyaPayrollAkses(ByVal vPrimaryPayroll As Double, ByVal i As FrmTransaksiPayment) As Boolean
        Dim DTAUser As New DseTransPaymentTableAdapters.TblPayrollAccessTableAdapter
        Dim TBLUser As New DseTransPayment.TblPayrollAccessDataTable
        Dim ObjSalary As New ClsTransaksiSalary
        Dim PrimaryFormula As Double = -1
        Dim Filter As String = ""
        Dim Retval As String = ""
        Try
            DTAUser.Connection.ConnectionString = Core_Module.GetConstring
            DTAUser.FillBy_KodeUser(TBLUser, i.UserCode)
            PrimaryFormula = ObjSalary.GetPrimarySalaryFormula(vPrimaryPayroll)
            Filter = "[Primary Salary Formula] = " + PrimaryFormula.ToString + ""
            TBLUser.DefaultView.RowFilter = Filter
            If TBLUser.Rows.Count <= 0 Then
                Retval = "Salary no " & ObjSalary.GetNoTransaksi(vPrimaryPayroll).ToString & " isn't payroll access."
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAUser.Dispose()
            TBLUser.Dispose()
            ObjSalary = Nothing
        End Try
    End Function

    Public Function ApakahBolehFindDataPayment(ByVal vPrimaryPayment As Double, ByVal vUserCode As String) As Boolean
        Dim DTAItem As New DseTransPaymentTableAdapters.tbltrans_payment_item_employeeTableAdapter
        Dim TBLItem As New DseTransPayment.tbltrans_payment_item_employeeDataTable
        Dim RowItem As DseTransPayment.tbltrans_payment_item_employeeRow
        Dim DTAUser As New DseTransPaymentTableAdapters.TblPayrollAccessTableAdapter
        Dim TBLUser As New DseTransPayment.TblPayrollAccessDataTable
        Dim DTASalary As New DseTransPaymentTableAdapters.tbltrans_salaryTableAdapter
        Dim TBLSalary As New DseTransPayment.tbltrans_salaryDataTable
        Dim RowSalary As DseTransPayment.tbltrans_salaryRow
        Dim PrimarySalaryMIN As Double = -1
        Dim PrimarySalaryMAX As Double = -1
        Dim PrimarySalaryFormula As Double = -1
        Dim SalaryNo As String = ""
        Dim Filter As String = ""
        Dim Retval As String = ""
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAUser.Connection.ConnectionString = Core_Module.GetConstring
            DTASalary.Connection.ConnectionString = Core_Module.GetConstring
            DTAUser.FillBy_KodeUser(TBLUser, vUserCode)
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimaryPayment)
            PrimarySalaryMIN = Me.GetPrimarySalaryMIN(vPrimaryPayment)
            PrimarySalaryMAX = Me.GetPrimarySalaryMAX(vPrimaryPayment)
            DTASalary.FillBy_PrimaryMinMax(TBLSalary, PrimarySalaryMIN, PrimarySalaryMAX)
            If TBLItem.Rows.Count > 0 Then
                For Each RowItem In TBLItem.Rows
                    PrimarySalaryFormula = -1
                    SalaryNo = ""
                    Filter = "[Primary] = " + RowItem.primary_transaksi_salary.ToString + ""
                    TBLSalary.DefaultView.RowFilter = Filter
                    If TBLSalary.DefaultView.Count > 0 Then
                        RowSalary = TBLSalary.DefaultView(0).Row
                        PrimarySalaryFormula = RowSalary.Primary_Salary_Formula
                        SalaryNo = RowSalary.No_Transaksi
                    End If
                    Filter = "[Primary Salary Formula] = " + PrimarySalaryFormula.ToString + ""
                    TBLUser.DefaultView.RowFilter = Filter
                    If TBLUser.DefaultView.Count <= 0 Then
                        If String.IsNullOrEmpty(Retval) = False Then
                            Retval += ", " + SalaryNo
                        Else
                            Retval = SalaryNo
                        End If
                    End If
                Next
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                MsgBox("The following salary transaction isn't payroll access : " + vbNewLine + Retval + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            DTAUser.Dispose()
            TBLUser.Dispose()
            DTASalary.Dispose()
            TBLSalary.Dispose()
        End Try
    End Function

    Public Enum EnumItemJurnalPayment As Integer
        SalaryExpense
        OtherDeduction
        Tax
        LoanInstallment
        Payment
        CompanyExpense
        EmployeeExpense
    End Enum

    Public Sub BuatFormatDataSetCetakanPayment(ByVal vPrimarySalaryMin As Double, ByVal vPrimarySalaryMax As Double, _
    ByVal mTBLPaymentAkun As DseTransPayment.tbltrans_payment_item_employee_dDataTable, _
    ByRef TBLHasil As DseTransPayment.DataTable1DataTable, ByRef TBLCetak As DseTransPayment_Cetak.TblCetakDataTable, _
    ByVal vDef_Prec As Double, ByVal i As FrmTransaksiPayment)
        Dim DTAPPh As New DseTransPaymentTableAdapters.tbltrans_salaryTableAdapter
        Dim TBLPPh As New DseTransPayment.tbltrans_salaryDataTable
        Dim DTAItem As New DseTransPaymentTableAdapters.tbltrans_salary_itemTableAdapter
        Dim TBLItemIncome As New DseTransPayment.tbltrans_salary_itemDataTable
        Dim TBLItemCompanyExpense As New DseTransPayment.tbltrans_salary_itemDataTable
        Dim TBLItemEmployeeExpense As New DseTransPayment.tbltrans_salary_itemDataTable
        Dim TBLItemDeduction As New DseTransPayment.tbltrans_salary_itemDataTable
        Dim RowPaymentAkun As DseTransPayment.tbltrans_payment_item_employee_dRow
        Dim RowHasil As DseTransPayment.DataTable1Row
        Dim RowCetak As DseTransPayment_Cetak.TblCetakRow
        Dim DTALoanPay As New DseTransPaymentTableAdapters.DataTable_LoanPaymentTableAdapter
        Dim TBLLoanPay As New DseTransPayment.DataTable_LoanPaymentDataTable
        Dim DTAOtherDeduction As New DseTransPaymentTableAdapters.tbltrans_salary_item_otherdeductionTableAdapter
        Dim TBLOtherDeduction As New DseTransPayment.tbltrans_salary_item_otherdeductionDataTable
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjAkun As New AccModule.ClsMasterAkun
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim PrimaryCurrency As Double = 1
        Dim StrCurr As String = ObjCurr.GetCurrencyCode(PrimaryCurrency)
        Dim RowDGV As DataGridViewRow
        Dim RV As Data.DataRowView
        Dim Filter As String = ""
        Dim ListStringAkun As New List(Of String)
        Dim DebetKredit As Integer = -1
        Dim DebetKredit_2 As Integer = -1
        Dim KetemuLoanPayment As Boolean = False
        Dim TipeSalaryKomponen As Core_Module.FrmMasterSalaryFormula_Component.EnumType = Core_Module.FrmMasterSalaryFormula_Component.EnumType.CompanyExpense

        Try
            'Kumpulkan data
            DTAPPh.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTALoanPay.Connection.ConnectionString = Core_Module.GetConstring
            DTAOtherDeduction.Connection.ConnectionString = Core_Module.GetConstring

            TBLHasil.Rows.Clear()
            TBLCetak.Rows.Clear()
            If i.DGVJournalItem.AllowUserToAddRows = True Then
                i.DGVJournalItem.AllowUserToAddRows = False
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTAPPh.FillBy_PrimaryMinMax(TBLPPh, vPrimarySalaryMin, vPrimarySalaryMax)
            DTAItem.FillBy_PrimaryTransaksi_Income_Akun(TBLItemIncome, vPrimarySalaryMin, vPrimarySalaryMax)
            DTAItem.FillBy_PrimaryTransaksi_Deduction_Akun(TBLItemDeduction, vPrimarySalaryMin, vPrimarySalaryMax)
            DTAOtherDeduction.FillBy_PrimarySalary_Akun(TBLOtherDeduction, vPrimarySalaryMin, vPrimarySalaryMax)
            DTALoanPay.FillBy_PrimarySalaryMinMax(TBLLoanPay, vPrimarySalaryMin, vPrimarySalaryMax)
            TipeSalaryKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.CompanyExpense
            DTAItem.FillBy_TipeSalaryKomponen(TBLItemCompanyExpense, vPrimarySalaryMin, CInt(TipeSalaryKomponen), vPrimarySalaryMax)
            TipeSalaryKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense
            DTAItem.FillBy_TipeSalaryKomponen(TBLItemEmployeeExpense, vPrimarySalaryMin, CInt(TipeSalaryKomponen), vPrimarySalaryMax)

            i.dgvItem.AllowUserToAddRows = False
            If i.dgvItem.Rows.Count > 0 Then
                For Each RowDGV In i.dgvItem.Rows
                    Filter = "[Primary Transaksi] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLItemIncome.DefaultView.RowFilter = Filter
                    If TBLItemIncome.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_Income(TBLHasil, TBLItemIncome, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    Filter = "[Primary Transaksi] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLItemDeduction.DefaultView.RowFilter = Filter
                    If TBLItemDeduction.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_Deduction(TBLHasil, TBLItemDeduction, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    'company expense
                    Filter = "[Primary Transaksi] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLItemCompanyExpense.DefaultView.RowFilter = Filter
                    If TBLItemCompanyExpense.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_CompanyExpense(TBLHasil, TBLItemCompanyExpense, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    'employee expense
                    Filter = "[Primary Transaksi] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLItemEmployeeExpense.DefaultView.RowFilter = Filter
                    If TBLItemEmployeeExpense.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_EmployeeExpense(TBLHasil, TBLItemEmployeeExpense, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    Filter = "[Primary Transaksi] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLOtherDeduction.DefaultView.RowFilter = Filter
                    If TBLOtherDeduction.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_OtherDeduction(TBLHasil, TBLOtherDeduction, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    Filter = "[Primary Salary] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLLoanPay.DefaultView.RowFilter = Filter
                    If TBLLoanPay.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_LoanPayment(TBLHasil, TBLLoanPay, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    Filter = "[Primary] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLPPh.DefaultView.RowFilter = Filter
                    If TBLPPh.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_PPh21(TBLHasil, TBLPPh, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    Filter = "[identifier] = '" + RowDGV.Cells(i.ColIdentifier.Index).Value.ToString + "'"
                    mTBLPaymentAkun.DefaultView.RowFilter = Filter
                    If mTBLPaymentAkun.DefaultView.Count > 0 Then
                        For Each RV In mTBLPaymentAkun.DefaultView
                            RowPaymentAkun = RV.Row
                            RowHasil = TBLHasil.NewDataTable1Row
                            RowHasil.Enum_Tipe_Jurnal = ClsTrans_Payment.EnumItemJurnalPayment.Payment
                            RowHasil.Primary_Salary = Convert.ToDouble(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value)
                            RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.Payment
                            RowHasil.Amount_1 = RowPaymentAkun.amount
                            RowHasil.Amount_2 = 0
                            RowHasil.Primary_Akun_1 = RowPaymentAkun.primary_akun_payment
                            RowHasil.Primary_Akun_2 = -1
                            RowHasil.Description_Salary = "Payment"
                            RowHasil.Employee_Name = RowDGV.Cells(i.ColEmp_EmployeeCode.Index).Value
                            RowHasil.Salary_Number = RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value
                            TBLHasil.Rows.Add(RowHasil)
                        Next
                    End If
                    mTBLPaymentAkun.DefaultView.RowFilter = ""
                Next

                'Koding untuk pecahan per-salary dan employee name
                If TBLHasil.Rows.Count > 0 Then
                    For Each RowHasil In TBLHasil.Rows
                        KetemuLoanPayment = False
                        'Debet = 0, kredit = 1 
                        If RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.Income Then
                            If ObjCD.CekBilanganPositif(RowHasil.Amount_1) = False Then
                                DebetKredit = 1
                            Else
                                DebetKredit = 0
                            End If
                        ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.Deduction _
                        Or RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.OtherDeduction Then
                            If ObjCD.CekBilanganPositif(RowHasil.Amount_1) = False Then
                                DebetKredit = 0
                            Else
                                DebetKredit = 1
                            End If
                        ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.PPh21 Then
                            DebetKredit = 1
                        ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.Payment Then
                            DebetKredit = 1
                        ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.LoanInstallment Then
                            KetemuLoanPayment = True
                        End If
                        If KetemuLoanPayment = False Then
                            If RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.CompanyExpense Then
                                'expense account-> debet
                                DebetKredit = 0
                                If ListStringAkun.Contains(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                RowHasil.Primary_Akun_1.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString) = False Then
                                    ListStringAkun.Add(RowHasil.Primary_Akun_1.ToString + "-" + _
                                    RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString)
                                    'Buat Baris Baru
                                    Me.BuatBarisCetakBaru(TBLCetak, DebetKredit, RowHasil, ObjAkun, RowHasil.Primary_Akun_1)
                                Else
                                    'Cari baris mana yang sesuai
                                    RowCetak = TBLCetak.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                    RowHasil.Primary_Akun_1.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString))

                                    If DebetKredit = 1 Then
                                        RowCetak.Kredit = RowCetak.Kredit + Math.Abs(RowHasil.Amount_1)
                                    ElseIf DebetKredit = 0 Then
                                        RowCetak.Debet = RowCetak.Debet + Math.Abs(RowHasil.Amount_1)
                                    End If
                                End If

                                'payable account -> kredit
                                DebetKredit = 1
                                If ListStringAkun.Contains(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                RowHasil.Primary_Akun_1.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString) = False Then
                                    ListStringAkun.Add(RowHasil.Primary_Akun_2.ToString + "-" + _
                                    RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString)
                                    'Buat Baris Baru
                                    Me.BuatBarisCetakBaru(TBLCetak, DebetKredit, RowHasil, ObjAkun, RowHasil.Primary_Akun_2)
                                Else
                                    RowCetak = TBLCetak.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                    RowHasil.Primary_Akun_1.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString))
                                    If DebetKredit = 1 Then
                                        RowCetak.Kredit = RowCetak.Kredit + Math.Abs(RowHasil.Amount_1)
                                    ElseIf DebetKredit = 0 Then
                                        RowCetak.Debet = RowCetak.Debet + Math.Abs(RowHasil.Amount_1)
                                    End If
                                End If
                            ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.EmployeeExpense Then
                                'payable account -> credit
                                DebetKredit = 1
                                If ListStringAkun.Contains(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                RowHasil.Primary_Akun_1.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString) = False Then
                                    ListStringAkun.Add(RowHasil.Primary_Akun_2.ToString + "-" + _
                                    RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString)
                                    'Buat Baris Baru
                                    Me.BuatBarisCetakBaru(TBLCetak, DebetKredit, RowHasil, ObjAkun, RowHasil.Primary_Akun_2)
                                Else
                                    'Cari baris mana yang sesuai
                                    RowCetak = TBLCetak.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                    RowHasil.Primary_Akun_1.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString))
                                    If DebetKredit = 1 Then
                                        RowCetak.Kredit = RowCetak.Kredit + Math.Abs(RowHasil.Amount_1)
                                    ElseIf DebetKredit = 0 Then
                                        RowCetak.Debet = RowCetak.Debet + Math.Abs(RowHasil.Amount_1)
                                    End If
                                End If
                            Else
                                If ListStringAkun.Contains(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                RowHasil.Primary_Akun_1.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString) = False Then
                                    ListStringAkun.Add(RowHasil.Primary_Akun_1.ToString + "-" + _
                                    RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString)
                                    'Buat Baris Baru
                                    Me.BuatBarisCetakBaru(TBLCetak, DebetKredit, RowHasil, ObjAkun, RowHasil.Primary_Akun_1)
                                Else
                                    'Cari baris mana yang sesuai
                                    RowCetak = TBLCetak.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                    RowHasil.Primary_Akun_1.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString))
                                    If DebetKredit = 1 Then
                                        RowCetak.Kredit = RowCetak.Kredit + Math.Abs(RowHasil.Amount_1)
                                    ElseIf DebetKredit = 0 Then
                                        RowCetak.Debet = RowCetak.Debet + Math.Abs(RowHasil.Amount_1)
                                    End If
                                End If
                            End If
                        ElseIf KetemuLoanPayment = True Then
                            If RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.LoanInstallment Then
                                If RowHasil.Primary_Akun_1 <> -1 Then
                                    'Pertama
                                    DebetKredit = 1
                                End If
                                If RowHasil.Primary_Akun_2 <> -1 Then
                                    'Kedua
                                    If ObjCD.CekBilanganPositif(RowHasil.Amount_2) = False Then
                                        DebetKredit_2 = 0
                                    Else
                                        DebetKredit_2 = 1
                                    End If
                                Else
                                    'nilai default
                                    DebetKredit_2 = 0
                                End If
                            End If
                            'Cek 2 kali
                            'Pertama
                            If ListStringAkun.Contains(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                              RowHasil.Primary_Akun_1.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString) = False Then
                                ListStringAkun.Add(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                                   RowHasil.Primary_Akun_1.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString)
                                'Buat Baris Baru
                                RowCetak = TBLCetak.NewTblCetakRow
                                RowCetak.No_Document = RowHasil.Salary_Number
                                RowCetak.Employee_Name = RowHasil.Employee_Name
                                RowCetak.Description = RowHasil.Description_Salary
                                RowCetak.Account_Description = ObjAkun.GetDescription_ByPrimaryAcc(RowHasil.Primary_Akun_1)
                                RowCetak.Primary_Salary = RowHasil.Primary_Salary
                                If DebetKredit = 0 Then
                                    RowCetak.Debet = RowHasil.Amount_1
                                    RowCetak.Kredit = 0
                                    RowCetak.Position_Debet_Kredit = DebetKredit
                                ElseIf DebetKredit = 1 Then
                                    RowCetak.Debet = 0
                                    RowCetak.Kredit = RowHasil.Amount_1
                                    RowCetak.Position_Debet_Kredit = DebetKredit
                                End If
                                TBLCetak.Rows.Add(RowCetak)
                            Else
                                'Cari baris mana yang sesuai
                                RowCetak = TBLCetak.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                RowHasil.Primary_Akun_1.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString))
                                If DebetKredit = 1 Then
                                    RowCetak.Kredit = RowCetak.Kredit + Math.Abs(RowHasil.Amount_1)
                                ElseIf DebetKredit = 0 Then
                                    RowCetak.Debet = RowCetak.Debet + Math.Abs(RowHasil.Amount_1)
                                End If
                            End If
                            'kedua
                            If RowHasil.Primary_Akun_2 <> -1 Then
                                If ListStringAkun.Contains(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                                  RowHasil.Primary_Akun_2.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit_2.ToString) = False Then
                                    ListStringAkun.Add(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                                   RowHasil.Primary_Akun_2.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit_2.ToString)
                                    'Buat Baris Baru
                                    RowCetak = TBLCetak.NewTblCetakRow
                                    RowCetak.No_Document = RowHasil.Salary_Number
                                    RowCetak.Employee_Name = RowHasil.Employee_Name
                                    RowCetak.Description = RowHasil.Description_Salary
                                    RowCetak.Account_Description = ObjAkun.GetDescription_ByPrimaryAcc(RowHasil.Primary_Akun_2)
                                    RowCetak.Primary_Salary = RowHasil.Primary_Salary
                                    If DebetKredit = 0 Then
                                        RowCetak.Debet = RowHasil.Amount_2
                                        RowCetak.Kredit = 0
                                        RowCetak.Position_Debet_Kredit = DebetKredit_2
                                    ElseIf DebetKredit = 1 Then
                                        RowCetak.Debet = 0
                                        RowCetak.Kredit = RowHasil.Amount_2
                                        RowCetak.Position_Debet_Kredit = DebetKredit_2
                                    End If
                                    TBLCetak.Rows.Add(RowCetak)
                                Else
                                    'Cari baris mana yang sesuai
                                    RowCetak = TBLCetak.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Salary.ToString + "-" + RowHasil.Employee_Name.ToString + "-" + _
                                                                 RowHasil.Primary_Akun_2.ToString + "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit_2.ToString))
                                    If DebetKredit = 1 Then
                                        RowCetak.Kredit = RowCetak.Kredit + Math.Abs(RowHasil.Amount_2)
                                    ElseIf DebetKredit = 0 Then
                                        RowCetak.Debet = RowCetak.Debet + Math.Abs(RowHasil.Amount_2)
                                    End If
                                End If
                                ''
                            End If
                        End If
                    Next
                End If
            End If
            '--------------------------------------------------
            'Cek dulu apakah data DGV ada yang netto / gross up
            Dim KetemuNettoAtoGrossUp As Boolean = False
            If i.dgvItem.Rows.Count > 0 Then
                For Each RowDGV In i.dgvItem.Rows
                    If RowDGV.Cells(i.ColEmp_TipeTaxIncome_Enum.Index).Value = ClsTransaksiSalary.EnumTax.Gross_Up Or _
                    RowDGV.Cells(i.ColEmp_TipeTaxIncome_Enum.Index).Value = ClsTransaksiSalary.EnumTax.Netto Then
                        KetemuNettoAtoGrossUp = True
                        Exit For
                    End If
                Next
            End If
            'Tambahan koding untuk ketentuan Balance journal tax jika netto ato gross up
            If TBLCetak.Rows.Count > 0 Then
                If KetemuNettoAtoGrossUp = True Then
                    For Each RowDGV In i.dgvItem.Rows
                        If RowDGV.Cells(i.ColEmp_TipeTaxIncome_Enum.Index).Value = ClsTransaksiSalary.EnumTax.Gross_Up Or _
                           RowDGV.Cells(i.ColEmp_TipeTaxIncome_Enum.Index).Value = ClsTransaksiSalary.EnumTax.Netto Then
                            Filter = "[Enum Tipe Jurnal] = " & CInt(ClsTrans_Payment.EnumItemJurnalPayment.Tax).ToString _
                                      & " AND [Primary Salary] = " & RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString & ""
                            TBLHasil.DefaultView.RowFilter = Filter
                            If TBLHasil.DefaultView.Count > 0 Then
                                RowHasil = TBLHasil.DefaultView(0).Row
                                RowCetak = TBLCetak.NewTblCetakRow
                                RowCetak.No_Document = RowHasil.Salary_Number
                                RowCetak.Employee_Name = RowHasil.Employee_Name
                                RowCetak.Description = RowHasil.Description_Salary
                                RowCetak.Account_Description = ObjAkun.GetDescription_ByPrimaryAcc(RowHasil.Primary_Akun_1)
                                RowCetak.Primary_Salary = RowHasil.Primary_Salary
                                RowCetak.Debet = RowHasil.Amount_1
                                RowCetak.Kredit = 0
                                RowCetak.Position_Debet_Kredit = 0
                                TBLCetak.Rows.Add(RowCetak)
                            End If
                            TBLHasil.DefaultView.RowFilter = ""
                        End If
                    Next
                End If
            End If
            i.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPPh.Dispose()
            TBLPPh.Dispose()
            DTAItem.Dispose()
            TBLItemDeduction.Dispose()
            TBLItemIncome.Dispose()
            DTALoanPay.Dispose()
            TBLLoanPay.Dispose()
            ListStringAkun = Nothing
            ObjCD = Nothing
            ObjAkun = Nothing
            ObjCurr = Nothing
            TBLItemCompanyExpense = Nothing
            TBLItemEmployeeExpense = Nothing
        End Try
    End Sub

    Private Sub BuatBarisCetakBaru(ByRef TBLCetak As DseTransPayment_Cetak.TblCetakDataTable, ByVal DebetKredit As Integer, _
    ByVal RowHasil As DseTransPayment.DataTable1Row, ByVal ObjAkun As AccModule.ClsMasterAkun, ByVal vprimaryAkun As Double)
        Dim RowCetak As DseTransPayment_Cetak.TblCetakRow

        Try
            RowCetak = TBLCetak.NewTblCetakRow
            RowCetak.No_Document = RowHasil.Salary_Number
            RowCetak.Employee_Name = RowHasil.Employee_Name
            RowCetak.Description = RowHasil.Description_Salary
            RowCetak.Account_Description = ObjAkun.GetDescription_ByPrimaryAcc(vprimaryAkun)
            RowCetak.Primary_Salary = RowHasil.Primary_Salary
            If DebetKredit = 0 Then
                RowCetak.Debet = RowHasil.Amount_1
                RowCetak.Kredit = 0
                RowCetak.Position_Debet_Kredit = DebetKredit
            ElseIf DebetKredit = 1 Then
                RowCetak.Debet = 0
                RowCetak.Kredit = RowHasil.Amount_1
                RowCetak.Position_Debet_Kredit = DebetKredit
            End If
            TBLCetak.Rows.Add(RowCetak)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanBarisJurnalBaru(ByVal i As FrmTransaksiPayment, ByVal DebetKredit As Integer, _
    ByVal RowHasil As DseTransPayment.DataTable1Row, ByVal ObjAkun As AccModule.ClsMasterAkun, _
    ByVal strCurr As String, ByVal PrimaryCurrency As Double, ByVal ObjCD As Core_Module.ClsCekData, _
    ByVal vDef_Prec As Integer, ByVal vPrimaryAkun As Double, Optional ByVal vAbs As Boolean = True)
        Dim rowdgv As DataGridViewRow

        Try
            i.DGVJournalItem.Rows.Add()
            rowdgv = i.DGVJournalItem.Rows(i.DGVJournalItem.Rows.Count - 1)
            rowdgv.Cells(i.ColJurnal_TipeJurnal_dbl.Index).Value = RowHasil.Enum_Tipe_Jurnal
            rowdgv.Cells(i.ColJurnal_TipeJurnal_str.Index).Value = Me.GetStringTipeJurnal(RowHasil.Enum_Tipe_Jurnal)
            rowdgv.Cells(i.ColJurnal_Account.Index).Value = ObjAkun.GetAccountCode(vPrimaryAkun)
            rowdgv.Cells(i.ColJurnal_DescAccount.Index).Value = ObjAkun.GetDescription_ByPrimaryAcc(vPrimaryAkun)
            rowdgv.Cells(i.ColJurnal_ExchangeRate.Index).Value = "1"
            rowdgv.Cells(i.ColJurnal_OriginalCurrency.Index).Value = strCurr
            rowdgv.Cells(i.ColJurnal_PrimaryCurrency.Index).Value = PrimaryCurrency
            rowdgv.Cells(i.ColJurnal_Remarks.Index).Value = ""
            rowdgv.Cells(i.ColJurnal_PrimaryAkun.Index).Value = vPrimaryAkun
            rowdgv.Cells(i.ColJurnal_Job.Index).Value = ""
            rowdgv.Cells(i.ColJurnal_Dept.Index).Value = ""
            rowdgv.Cells(i.ColJurnal_PrimaryDept.Index).Value = "-1"
            rowdgv.Cells(i.ColJurnal_PrimaryJob.Index).Value = "-1"
            If vAbs = True Then
                If DebetKredit = 0 Then
                    rowdgv.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(Math.Abs(RowHasil.Amount_1), vDef_Prec)
                    rowdgv.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                    rowdgv.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(Math.Abs(RowHasil.Amount_1), vDef_Prec)
                    rowdgv.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                ElseIf DebetKredit = 1 Then
                    rowdgv.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(Math.Abs(RowHasil.Amount_1), vDef_Prec)
                    rowdgv.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                    rowdgv.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(Math.Abs(RowHasil.Amount_1), vDef_Prec)
                    rowdgv.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                End If
            Else
                If DebetKredit = 0 Then
                    rowdgv.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(RowHasil.Amount_1, vDef_Prec)
                    rowdgv.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                    rowdgv.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(RowHasil.Amount_1, vDef_Prec)
                    rowdgv.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                ElseIf DebetKredit = 1 Then
                    rowdgv.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(RowHasil.Amount_1, vDef_Prec)
                    rowdgv.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                    rowdgv.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(RowHasil.Amount_1, vDef_Prec)
                    rowdgv.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub BuatJournalViewPayment(ByVal vPrimarySalaryMin As Double, ByVal vPrimarySalaryMax As Double, _
    ByVal mTBLPaymentAkun As DseTransPayment.tbltrans_payment_item_employee_dDataTable, _
    ByRef TBLHasil As DseTransPayment.DataTable1DataTable, ByVal vDef_Prec As Double, ByVal i As FrmTransaksiPayment)
        Dim DTAPPh As New DseTransPaymentTableAdapters.tbltrans_salaryTableAdapter
        Dim TBLPPh As New DseTransPayment.tbltrans_salaryDataTable
        Dim DTAItem As New DseTransPaymentTableAdapters.tbltrans_salary_itemTableAdapter
        Dim TBLItemIncome As New DseTransPayment.tbltrans_salary_itemDataTable
        Dim TBLItemDeduction As New DseTransPayment.tbltrans_salary_itemDataTable
        Dim TBLItemCompanyExpense As New DseTransPayment.tbltrans_salary_itemDataTable
        Dim TBLItemEmployeeExpense As New DseTransPayment.tbltrans_salary_itemDataTable
        Dim RowPaymentAkun As DseTransPayment.tbltrans_payment_item_employee_dRow
        Dim RowHasil As DseTransPayment.DataTable1Row
        Dim DTALoanPay As New DseTransPaymentTableAdapters.DataTable_LoanPaymentTableAdapter
        Dim TBLLoanPay As New DseTransPayment.DataTable_LoanPaymentDataTable
        Dim DTAOtherDeduction As New DseTransPaymentTableAdapters.tbltrans_salary_item_otherdeductionTableAdapter
        Dim TBLOtherDeduction As New DseTransPayment.tbltrans_salary_item_otherdeductionDataTable
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjAkun As New AccModule.ClsMasterAkun
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim PrimaryCurrency As Double = 1
        Dim StrCurr As String = ObjCurr.GetCurrencyCode(PrimaryCurrency)
        Dim Filter As String = ""
        Dim ListStringAkun As New List(Of String)
        Dim RowDGV As DataGridViewRow
        Dim RV As Data.DataRowView
        Dim DebetKredit As Integer = -1
        Dim DebetKredit_2 As Integer = -1
        Dim KetemuLoanPayment As Boolean = False
        Dim TipeSalaryKomponen As Core_Module.FrmMasterSalaryFormula_Component.EnumType = Core_Module.FrmMasterSalaryFormula_Component.EnumType.CompanyExpense

        Try
            DTAPPh.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTALoanPay.Connection.ConnectionString = Core_Module.GetConstring
            DTAOtherDeduction.Connection.ConnectionString = Core_Module.GetConstring

            TBLHasil.Rows.Clear()
            If i.DGVJournalItem.AllowUserToAddRows = True Then
                i.DGVJournalItem.AllowUserToAddRows = False
            End If
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTAPPh.FillBy_PrimaryMinMax(TBLPPh, vPrimarySalaryMin, vPrimarySalaryMax)
            DTAItem.FillBy_PrimaryTransaksi_Income_Akun(TBLItemIncome, vPrimarySalaryMin, vPrimarySalaryMax)
            DTAItem.FillBy_PrimaryTransaksi_Deduction_Akun(TBLItemDeduction, vPrimarySalaryMin, vPrimarySalaryMax)
            DTAOtherDeduction.FillBy_PrimarySalary_Akun(TBLOtherDeduction, vPrimarySalaryMin, vPrimarySalaryMax)
            DTALoanPay.FillBy_PrimarySalaryMinMax(TBLLoanPay, vPrimarySalaryMin, vPrimarySalaryMax)
            TipeSalaryKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.CompanyExpense
            DTAItem.FillBy_TipeSalaryKomponen(TBLItemCompanyExpense, vPrimarySalaryMin, CInt(TipeSalaryKomponen), vPrimarySalaryMax)
            TipeSalaryKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense
            DTAItem.FillBy_TipeSalaryKomponen(TBLItemEmployeeExpense, vPrimarySalaryMin, CInt(TipeSalaryKomponen), vPrimarySalaryMax)

            i.dgvItem.AllowUserToAddRows = False
            If i.dgvItem.Rows.Count > 0 Then
                For Each RowDGV In i.dgvItem.Rows
                    'income
                    Filter = "[Primary Transaksi] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLItemIncome.DefaultView.RowFilter = Filter
                    If TBLItemIncome.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_Income(TBLHasil, TBLItemIncome, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    'deduction
                    Filter = "[Primary Transaksi] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLItemDeduction.DefaultView.RowFilter = Filter
                    If TBLItemDeduction.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_Deduction(TBLHasil, TBLItemDeduction, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    'company expense
                    Filter = "[Primary Transaksi] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLItemCompanyExpense.DefaultView.RowFilter = Filter
                    If TBLItemCompanyExpense.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_CompanyExpense(TBLHasil, TBLItemCompanyExpense, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    'employee expense
                    Filter = "[Primary Transaksi] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLItemEmployeeExpense.DefaultView.RowFilter = Filter
                    If TBLItemEmployeeExpense.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_EmployeeExpense(TBLHasil, TBLItemEmployeeExpense, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    'other deduction
                    Filter = "[Primary Transaksi] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLOtherDeduction.DefaultView.RowFilter = Filter
                    If TBLOtherDeduction.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_OtherDeduction(TBLHasil, TBLOtherDeduction, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    'loan payment
                    Filter = "[Primary Salary] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLLoanPay.DefaultView.RowFilter = Filter
                    If TBLLoanPay.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_LoanPayment(TBLHasil, TBLLoanPay, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    'pph 21
                    Filter = "[Primary] = " + RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString + ""
                    TBLPPh.DefaultView.RowFilter = Filter
                    If TBLPPh.DefaultView.Count > 0 Then
                        Me.InputkeDataTableHasil_PPh21(TBLHasil, TBLPPh, RowDGV.Cells(i.ColEmp_Name.Index).Value.ToString, _
                        RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value)
                    End If
                    'take home pay
                    Filter = "[identifier] = '" + RowDGV.Cells(i.ColIdentifier.Index).Value.ToString + "'"
                    mTBLPaymentAkun.DefaultView.RowFilter = Filter
                    If mTBLPaymentAkun.DefaultView.Count > 0 Then
                        For Each RV In mTBLPaymentAkun.DefaultView
                            RowPaymentAkun = RV.Row
                            RowHasil = TBLHasil.NewDataTable1Row
                            RowHasil.Enum_Tipe_Jurnal = ClsTrans_Payment.EnumItemJurnalPayment.Payment
                            RowHasil.Primary_Salary = Convert.ToDouble(RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value)
                            RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.Payment
                            RowHasil.Amount_1 = RowPaymentAkun.amount
                            RowHasil.Amount_2 = 0
                            RowHasil.Primary_Akun_1 = RowPaymentAkun.primary_akun_payment
                            RowHasil.Primary_Akun_2 = -1
                            TBLHasil.Rows.Add(RowHasil)
                        Next
                    End If
                    mTBLPaymentAkun.DefaultView.RowFilter = ""
                Next
            End If
            i.DGVJournalItem.Rows.Clear()

            'Buat fungsi translate ke journal
            If TBLHasil.Rows.Count > 0 Then
                For Each RowHasil In TBLHasil.Rows
                    KetemuLoanPayment = False
                    'Debet = 0, kredit = 1 
                    If RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.Income Then
                        'bagian ini di override di bawah, abaikan saja
                        If ObjCD.CekBilanganPositif(RowHasil.Amount_1) = False Then
                            DebetKredit = 1
                        Else
                            DebetKredit = 0
                        End If
                    ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.Deduction _
                    Or RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.OtherDeduction Then
                        'untuk bagian deduction, abaikan saja akan di override di bawah
                        If ObjCD.CekBilanganPositif(RowHasil.Amount_1) = False Then
                            DebetKredit = 0
                        Else
                            DebetKredit = 1
                        End If
                    ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.PPh21 Then
                        DebetKredit = 1
                    ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.Payment Then
                        DebetKredit = 1
                    ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.LoanInstallment Then
                        KetemuLoanPayment = True
                    ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.CompanyExpense Or _
                    RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.EmployeeExpense Then
                        DebetKredit = -1
                    End If

                    If KetemuLoanPayment = False Then
                        If RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.CompanyExpense Then
                            'expense account-> debet
                            DebetKredit = 0
                            If ListStringAkun.Contains(RowHasil.Primary_Akun_1.ToString + "-" + _
                            RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString) = False Then
                                ListStringAkun.Add(RowHasil.Primary_Akun_1.ToString + "-" + _
                                RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString)
                                'Buat Baris Baru
                                Me.TambahkanBarisJurnalBaru(i, DebetKredit, RowHasil, ObjAkun, StrCurr, PrimaryCurrency, ObjCD, vDef_Prec, RowHasil.Primary_Akun_1)
                            Else
                                'Cari baris mana yang sesuai
                                RowDGV = i.DGVJournalItem.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Akun_1.ToString + _
                                "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString))
                                RowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                            End If

                            'payable account -> kredit
                            DebetKredit = 1
                            If ListStringAkun.Contains(RowHasil.Primary_Akun_2.ToString + "-" + _
                            RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString) = False Then
                                ListStringAkun.Add(RowHasil.Primary_Akun_2.ToString + "-" + _
                                RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString)
                                'Buat Baris Baru
                                Me.TambahkanBarisJurnalBaru(i, DebetKredit, RowHasil, ObjAkun, StrCurr, PrimaryCurrency, ObjCD, vDef_Prec, RowHasil.Primary_Akun_2)
                            Else
                                'Cari baris mana yang sesuai
                                RowDGV = i.DGVJournalItem.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Akun_2.ToString + _
                                "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString))
                                RowDGV.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Kredit.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                            End If
                        ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.EmployeeExpense Then
                            'payable account -> credit
                            DebetKredit = 1
                            If ListStringAkun.Contains(RowHasil.Primary_Akun_2.ToString + "-" + _
                            RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString) = False Then
                                ListStringAkun.Add(RowHasil.Primary_Akun_2.ToString + "-" + _
                                RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString)
                                'Buat Baris Baru
                                Me.TambahkanBarisJurnalBaru(i, DebetKredit, RowHasil, ObjAkun, StrCurr, PrimaryCurrency, ObjCD, vDef_Prec, RowHasil.Primary_Akun_2)
                            Else
                                'Cari baris mana yang sesuai
                                RowDGV = i.DGVJournalItem.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Akun_2.ToString + _
                                "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString))
                                RowDGV.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Kredit.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                            End If
                        ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.Income Then
                            'posisi pasti debet, nilanya sesuai
                            DebetKredit = 0
                            If ListStringAkun.Contains(RowHasil.Primary_Akun_1.ToString + "-" + _
                            RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString) = False Then
                                ListStringAkun.Add(RowHasil.Primary_Akun_1.ToString + "-" + _
                                RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString)
                                Me.TambahkanBarisJurnalBaru(i, DebetKredit, RowHasil, ObjAkun, StrCurr, _
                                PrimaryCurrency, ObjCD, vDef_Prec, RowHasil.Primary_Akun_1, False)
                            Else
                                'Cari baris mana yang sesuai
                                RowDGV = i.DGVJournalItem.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Akun_1.ToString + _
                                "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString))
                                RowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet.Index).Value) + RowHasil.Amount_1, vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value) + RowHasil.Amount_1, vDef_Prec)
                            End If
                        ElseIf RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.Deduction Then
                            'posisi pasti debet, nilanya sesuai *-1
                            DebetKredit = 0
                            RowHasil.Amount_1 *= -1
                            If ListStringAkun.Contains(RowHasil.Primary_Akun_1.ToString + "-" + _
                            RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString) = False Then
                                ListStringAkun.Add(RowHasil.Primary_Akun_1.ToString + "-" + _
                                RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString)
                                Me.TambahkanBarisJurnalBaru(i, DebetKredit, RowHasil, ObjAkun, StrCurr, _
                                PrimaryCurrency, ObjCD, vDef_Prec, RowHasil.Primary_Akun_1, False)
                            Else
                                'Cari baris mana yang sesuai
                                RowDGV = i.DGVJournalItem.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Akun_1.ToString + _
                                "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString))
                                RowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet.Index).Value) + RowHasil.Amount_1, vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value) + RowHasil.Amount_1, vDef_Prec)
                            End If
                        Else
                            If ListStringAkun.Contains(RowHasil.Primary_Akun_1.ToString + "-" + _
                            RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString) = False Then
                                ListStringAkun.Add(RowHasil.Primary_Akun_1.ToString + "-" + _
                                RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString)
                                'Buat Baris Baru
                                Me.TambahkanBarisJurnalBaru(i, DebetKredit, RowHasil, ObjAkun, StrCurr, PrimaryCurrency, ObjCD, vDef_Prec, RowHasil.Primary_Akun_1, False)
                            Else
                                'Cari baris mana yang sesuai
                                RowDGV = i.DGVJournalItem.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Akun_1.ToString + _
                                "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString))
                                If RowHasil.Amount_1 < 0 Then
                                    If DebetKredit = 1 Then
                                        RowDGV.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Kredit.Index).Value) - Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                        RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value) - Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                    ElseIf DebetKredit = 0 Then
                                        RowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet.Index).Value) - Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                        RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value) - Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                    End If
                                Else
                                    If DebetKredit = 1 Then
                                        RowDGV.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Kredit.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                        RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                    ElseIf DebetKredit = 0 Then
                                        RowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                        RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                    End If
                                End If
                            End If
                        End If
                    ElseIf KetemuLoanPayment = True Then
                        If RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.LoanInstallment Then
                            If RowHasil.Primary_Akun_1 <> -1 Then
                                'Pertama
                                DebetKredit = 1
                            End If
                            If RowHasil.Primary_Akun_2 <> -1 Then
                                'Kedua
                                If ObjCD.CekBilanganPositif(RowHasil.Amount_2) = False Then
                                    DebetKredit_2 = 0
                                Else
                                    DebetKredit_2 = 1
                                End If
                            Else
                                'nilai default
                                DebetKredit_2 = 0
                            End If
                        End If
                        'Cek 2 kali
                        'Pertama
                        If ListStringAkun.Contains(RowHasil.Primary_Akun_1.ToString + "-" + _
                        RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString) = False Then
                            ListStringAkun.Add(RowHasil.Primary_Akun_1.ToString + "-" + _
                                                   RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString)
                            'Buat Baris Baru
                            i.DGVJournalItem.Rows.Add()
                            RowDGV = i.DGVJournalItem.Rows(i.DGVJournalItem.Rows.Count - 1)
                            RowDGV.Cells(i.ColJurnal_TipeJurnal_dbl.Index).Value = RowHasil.Enum_Tipe_Jurnal
                            RowDGV.Cells(i.ColJurnal_TipeJurnal_str.Index).Value = Me.GetStringTipeJurnal(RowHasil.Enum_Tipe_Jurnal)
                            RowDGV.Cells(i.ColJurnal_Account.Index).Value = ObjAkun.GetAccountCode(RowHasil.Primary_Akun_1)
                            RowDGV.Cells(i.ColJurnal_DescAccount.Index).Value = ObjAkun.GetDescription_ByPrimaryAcc(RowHasil.Primary_Akun_1)
                            RowDGV.Cells(i.ColJurnal_ExchangeRate.Index).Value = "1"
                            RowDGV.Cells(i.ColJurnal_OriginalCurrency.Index).Value = StrCurr
                            RowDGV.Cells(i.ColJurnal_PrimaryCurrency.Index).Value = PrimaryCurrency
                            RowDGV.Cells(i.ColJurnal_Remarks.Index).Value = ""
                            RowDGV.Cells(i.ColJurnal_PrimaryAkun.Index).Value = RowHasil.Primary_Akun_1
                            RowDGV.Cells(i.ColJurnal_Job.Index).Value = ""
                            RowDGV.Cells(i.ColJurnal_Dept.Index).Value = ""
                            RowDGV.Cells(i.ColJurnal_PrimaryDept.Index).Value = "-1"
                            RowDGV.Cells(i.ColJurnal_PrimaryJob.Index).Value = "-1"
                            If DebetKredit = 0 Then
                                RowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                            ElseIf DebetKredit = 1 Then
                                RowDGV.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                            End If
                        Else
                            'Cari baris mana yang sesuai
                            RowDGV = i.DGVJournalItem.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Akun_1.ToString + _
                            "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit.ToString))
                            If DebetKredit = 1 Then
                                RowDGV.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Kredit.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                            ElseIf DebetKredit = 0 Then
                                RowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                                RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value) + Math.Abs(RowHasil.Amount_1), vDef_Prec)
                            End If
                        End If
                        'kedua
                        If RowHasil.Primary_Akun_2 <> -1 Then
                            If ListStringAkun.Contains(RowHasil.Primary_Akun_2.ToString + "-" + _
                            RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit_2.ToString) = False Then
                                ListStringAkun.Add(RowHasil.Primary_Akun_2.ToString + "-" + _
                                                   RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit_2.ToString)
                                'Buat Baris Baru
                                i.DGVJournalItem.Rows.Add()
                                RowDGV = i.DGVJournalItem.Rows(i.DGVJournalItem.Rows.Count - 1)
                                RowDGV.Cells(i.ColJurnal_TipeJurnal_dbl.Index).Value = RowHasil.Enum_Tipe_Jurnal
                                RowDGV.Cells(i.ColJurnal_TipeJurnal_str.Index).Value = Me.GetStringTipeJurnal(RowHasil.Enum_Tipe_Jurnal)
                                RowDGV.Cells(i.ColJurnal_Account.Index).Value = ObjAkun.GetAccountCode(RowHasil.Primary_Akun_2)
                                RowDGV.Cells(i.ColJurnal_DescAccount.Index).Value = ObjAkun.GetDescription_ByPrimaryAcc(RowHasil.Primary_Akun_2)
                                RowDGV.Cells(i.ColJurnal_ExchangeRate.Index).Value = "1"
                                RowDGV.Cells(i.ColJurnal_OriginalCurrency.Index).Value = StrCurr
                                RowDGV.Cells(i.ColJurnal_PrimaryCurrency.Index).Value = PrimaryCurrency
                                RowDGV.Cells(i.ColJurnal_Remarks.Index).Value = ""
                                RowDGV.Cells(i.ColJurnal_PrimaryAkun.Index).Value = RowHasil.Primary_Akun_2
                                RowDGV.Cells(i.ColJurnal_Job.Index).Value = ""
                                RowDGV.Cells(i.ColJurnal_Dept.Index).Value = ""
                                RowDGV.Cells(i.ColJurnal_PrimaryDept.Index).Value = "-1"
                                RowDGV.Cells(i.ColJurnal_PrimaryJob.Index).Value = "-1"
                                If DebetKredit_2 = 0 Then
                                    RowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(Math.Abs(RowHasil.Amount_2), vDef_Prec)
                                    RowDGV.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                                    RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(Math.Abs(RowHasil.Amount_2), vDef_Prec)
                                    RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                                ElseIf DebetKredit_2 = 1 Then
                                    RowDGV.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(Math.Abs(RowHasil.Amount_2), vDef_Prec)
                                    RowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                                    RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(Math.Abs(RowHasil.Amount_2), vDef_Prec)
                                    RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping("0", vDef_Prec)
                                End If
                            Else
                                'Cari baris mana yang sesuai
                                RowDGV = i.DGVJournalItem.Rows(ListStringAkun.IndexOf(RowHasil.Primary_Akun_2.ToString + _
                                "-" + RowHasil.Enum_Tipe_Jurnal.ToString + "-" + DebetKredit_2.ToString))
                                If DebetKredit_2 = 1 Then
                                    RowDGV.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Kredit.Index).Value) + Math.Abs(RowHasil.Amount_2), vDef_Prec)
                                    RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value) + Math.Abs(RowHasil.Amount_2), vDef_Prec)
                                ElseIf DebetKredit_2 = 0 Then
                                    RowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet.Index).Value) + Math.Abs(RowHasil.Amount_2), vDef_Prec)
                                    RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value) + Math.Abs(RowHasil.Amount_2), vDef_Prec)
                                End If
                            End If
                            ''
                        End If
                    End If
                Next
            End If
            '--------------------------------------------------
            'Cek dulu apakah data DGV ada yang netto / gross up
            Dim KetemuNettoAtoGrossUp As Boolean = False
            If i.dgvItem.Rows.Count > 0 Then
                For Each RowDGV In i.dgvItem.Rows
                    If RowDGV.Cells(i.ColEmp_TipeTaxIncome_Enum.Index).Value = ClsTransaksiSalary.EnumTax.Gross_Up Or _
                    RowDGV.Cells(i.ColEmp_TipeTaxIncome_Enum.Index).Value = ClsTransaksiSalary.EnumTax.Netto Then
                        KetemuNettoAtoGrossUp = True
                        Exit For
                    End If
                Next
            End If

            'Tambahan koding untuk ketentuan Balance journal tax jika netto ato gross up -> tax payable
            Dim BeforePosisiDebetKredit As Double = -1
            Dim NowPosisiDebetKredit As Double = -1
            Dim DestinationIndex As Double = -1
            Dim AngkaIncomeTaxDebet As Decimal = 0
            If i.DGVJournalItem.Rows.Count > 0 Then
                If KetemuNettoAtoGrossUp = True Then
                    For Each RowDGV In i.dgvItem.Rows
                        If RowDGV.Cells(i.ColEmp_TipeTaxIncome_Enum.Index).Value = ClsTransaksiSalary.EnumTax.Gross_Up Or _
                           RowDGV.Cells(i.ColEmp_TipeTaxIncome_Enum.Index).Value = ClsTransaksiSalary.EnumTax.Netto Then
                            Filter = "[Enum Tipe Jurnal] = " & CInt(ClsTrans_Payment.EnumItemJurnalPayment.Tax).ToString _
                                    & " AND [Primary Salary] = " & RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value.ToString & ""
                            TBLHasil.DefaultView.RowFilter = Filter
                            If TBLHasil.DefaultView.Count > 0 Then
                                RowHasil = TBLHasil.DefaultView(0).Row
                                AngkaIncomeTaxDebet += RowHasil.Amount_1
                            End If
                        End If
                    Next
                    If i.DGVJournalItem.Rows.Count > 0 Then
                        For Each RowDGV In i.DGVJournalItem.Rows
                            'Set apakah debet ato kredit
                            If Convert.ToDecimal(RowDGV.Cells(i.ColJurnal_TipeJurnal_dbl.Index).Value) <> CInt(ClsTrans_Payment.EnumItemJurnalPayment.Tax) Then
                                NowPosisiDebetKredit = 0
                                BeforePosisiDebetKredit = 0
                            Else
                                NowPosisiDebetKredit = 1
                                If BeforePosisiDebetKredit = 0 Then
                                    DestinationIndex = RowDGV.Index - 1
                                    Exit For
                                End If
                                BeforePosisiDebetKredit = 1
                            End If
                        Next
                    End If
                    'masukin angka income tax salary dari yang netto dan Gross Up
                    If DestinationIndex > -1 Then
                        Me.InsertBarisTengahTengahDGVJournal1(DestinationIndex, AngkaIncomeTaxDebet, i)
                    End If
                    ''
                End If
            End If

            Me.TukerTaxMinus(i.DGVJournalItem, i)

            i.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPPh.Dispose()
            TBLPPh.Dispose()
            DTAItem.Dispose()
            TBLItemDeduction.Dispose()
            TBLItemIncome.Dispose()
            TBLItemCompanyExpense.Dispose()
            TBLItemEmployeeExpense.Dispose()
            DTALoanPay.Dispose()
            TBLLoanPay.Dispose()
            ListStringAkun = Nothing
            ObjCD = Nothing
            ObjAkun = Nothing
            ObjCurr = Nothing
        End Try
    End Sub

    Private Sub TukerTaxMinus(ByVal vDGV As DataGridView, ByVal i As FrmTransaksiPayment)
        Dim rowDGV As DataGridViewRow
        Dim AdaMinus As Boolean = False
        Dim vTipeJurnal As ClsTrans_Payment.EnumItemJurnalPayment = EnumItemJurnalPayment.Tax
        Dim Value As Decimal = 0
        Dim DefPrec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            For Each rowDGV In vDGV.Rows
                If rowDGV.Cells(i.ColJurnal_TipeJurnal_dbl.Index).Value = vTipeJurnal Then
                    If CDec(rowDGV.Cells(i.ColJurnal_Debet.Index).Value) < 0 Then
                        AdaMinus = True
                        Exit For
                    ElseIf CDec(rowDGV.Cells(i.ColJurnal_Kredit.Index).Value) < 0 Then
                        AdaMinus = True
                        Exit For
                    End If
                End If
            Next

            If AdaMinus = True Then
                DefPrec = ObjCur.GetDecimal_DefaultCurrency

                For Each rowDGV In vDGV.Rows
                    If rowDGV.Cells(i.ColJurnal_TipeJurnal_dbl.Index).Value = vTipeJurnal Then
                        If CDec(rowDGV.Cells(i.ColJurnal_Debet.Index).Value) <> 0 Then
                            rowDGV.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(Math.Abs(CDec(rowDGV.Cells(i.ColJurnal_Debet.Index).Value)), DefPrec)
                            rowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = rowDGV.Cells(i.ColJurnal_Kredit.Index).Value
                            rowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(0, DefPrec)
                            rowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = rowDGV.Cells(i.ColJurnal_Debet.Index).Value
                        ElseIf CDec(rowDGV.Cells(i.ColJurnal_Kredit.Index).Value) <> 0 Then
                            rowDGV.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(Math.Abs(CDec(rowDGV.Cells(i.ColJurnal_Kredit.Index).Value)), DefPrec)
                            rowDGV.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = rowDGV.Cells(i.ColJurnal_Debet.Index).Value
                            rowDGV.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping(0, DefPrec)
                            rowDGV.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = rowDGV.Cells(i.ColJurnal_Kredit.Index).Value
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
            ObjCD = Nothing
        End Try
    End Sub

    Private Sub InsertBarisTengahTengahDGVJournal1(ByVal vIndexRow As Integer, ByVal vAmountTax As Decimal, ByVal i As FrmTransaksiPayment)
        Dim RowDGV_Source As DataGridViewRow = i.DGVJournalItem.Rows(vIndexRow)
        Dim RowDGV_Destination As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim ObjAkun As New AccModule.ClsMasterAkun
        Dim Primary_Akun As Double = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
        Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryTaxExpenseAccount)

        Try
            If vAmountTax <> 0 Then
                i.DGVJournalItem.Rows.InsertCopy(vIndexRow, vIndexRow + 1)
                RowDGV_Destination = i.DGVJournalItem.Rows(vIndexRow + 1)
                RowDGV_Destination.Cells(i.ColJurnal_TipeJurnal_dbl.Index).Value = CInt(ClsTrans_Payment.EnumItemJurnalPayment.Tax)
                RowDGV_Destination.Cells(i.ColJurnal_TipeJurnal_str.Index).Value = "Tax"
                RowDGV_Destination.Cells(i.ColJurnal_Account.Index).Value = ObjAkun.GetAccountCode(Primary_Akun)
                RowDGV_Destination.Cells(i.ColJurnal_DescAccount.Index).Value = ObjAkun.GetDescription_ByPrimaryAcc(Primary_Akun)
                RowDGV_Destination.Cells(i.ColJurnal_ExchangeRate.Index).Value = "1"
                RowDGV_Destination.Cells(i.ColJurnal_OriginalCurrency.Index).Value = i.Def_Prec
                RowDGV_Destination.Cells(i.ColJurnal_PrimaryCurrency.Index).Value = i.PrimaryCurrency
                RowDGV_Destination.Cells(i.ColJurnal_Remarks.Index).Value = ""
                RowDGV_Destination.Cells(i.ColJurnal_PrimaryAkun.Index).Value = Primary_Akun
                RowDGV_Destination.Cells(i.ColJurnal_Job.Index).Value = ""
                RowDGV_Destination.Cells(i.ColJurnal_Dept.Index).Value = ""
                RowDGV_Destination.Cells(i.ColJurnal_PrimaryDept.Index).Value = "-1"
                RowDGV_Destination.Cells(i.ColJurnal_PrimaryJob.Index).Value = "-1"
                RowDGV_Destination.Cells(i.ColJurnal_Kredit.Index).Value = ObjCD.DigitGrouping("0", i.Def_Prec)
                RowDGV_Destination.Cells(i.ColJurnal_Debet.Index).Value = ObjCD.DigitGrouping(Math.Abs(vAmountTax), i.Def_Prec)
                RowDGV_Destination.Cells(i.ColJurnal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping("0", i.Def_Prec)
                RowDGV_Destination.Cells(i.ColJurnal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(Math.Abs(vAmountTax), i.Def_Prec)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            ObjAkun = Nothing
            ObjSetting = Nothing
        End Try
    End Sub

    Public Function GetStringTipeJurnal(ByVal vTipeJurnal As ClsTrans_Payment.EnumItemJurnalPayment) As String
        Try
            If vTipeJurnal = EnumItemJurnalPayment.SalaryExpense Then
                Return "Salary expense"
            ElseIf vTipeJurnal = EnumItemJurnalPayment.OtherDeduction Then
                Return "Other deduction"
            ElseIf vTipeJurnal = EnumItemJurnalPayment.LoanInstallment Then
                Return "Loan installment"
            ElseIf vTipeJurnal = EnumItemJurnalPayment.Tax Then
                Return "Tax"
            ElseIf vTipeJurnal = EnumItemJurnalPayment.Payment Then
                Return "Payment"
            ElseIf vTipeJurnal = EnumItemJurnalPayment.CompanyExpense Then
                Return "Company expense"
            ElseIf vTipeJurnal = EnumItemJurnalPayment.EmployeeExpense Then
                Return "Employee expense"
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub InputkeDataTableHasil_PPh21(ByVal TblHasil As DseTransPayment.DataTable1DataTable, _
    ByVal TBLSalary As DseTransPayment.tbltrans_salaryDataTable, ByVal vEmployeeName As String, _
    ByVal vSalaryNo As String)
        Dim RowHasil As DseTransPayment.DataTable1Row
        Dim RowSalary As DseTransPayment.tbltrans_salaryRow
        Dim RV As Data.DataRowView
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting

        Try
            For Each RV In TBLSalary.DefaultView
                RowSalary = RV.Row
                RowHasil = TblHasil.NewDataTable1Row
                RowHasil.Enum_Tipe_Jurnal = ClsTrans_Payment.EnumItemJurnalPayment.Tax
                RowHasil.Primary_Salary = RowSalary.Primary
                RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.PPh21

                If RowSalary.Tanggal_Transaksi.Month = 12 Then
                    RowHasil.Amount_1 = RowSalary.Kurang_Bayar_Tax
                Else
                    RowHasil.Amount_1 = RowSalary.Income_Tax
                End If
                RowHasil.Amount_2 = 0
                If String.IsNullOrEmpty(ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryTaxPayableAccount1)) = True Then
                    RowHasil.Primary_Akun_1 = -1
                Else
                    RowHasil.Primary_Akun_1 = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryTaxPayableAccount1)
                End If
                RowHasil.Primary_Akun_2 = -1
                RowHasil.Description_Salary = "PPh Pasal 21"
                RowHasil.Employee_Name = vEmployeeName
                RowHasil.Salary_Number = vSalaryNo
                If RowHasil.Amount_1 <> 0 Then
                    TblHasil.Rows.Add(RowHasil)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSetting = Nothing
        End Try
    End Sub

    Private Sub InputkeDataTableHasil_LoanPayment(ByVal TblHasil As DseTransPayment.DataTable1DataTable, _
    ByVal TBLLoanPay As DseTransPayment.DataTable_LoanPaymentDataTable, ByVal vEmployeeName As String, _
    ByVal vSalaryNo As String)
        Dim RowHasil As DseTransPayment.DataTable1Row
        Dim RowLoanPay As DseTransPayment.DataTable_LoanPaymentRow
        Dim RV As Data.DataRowView
        Try
            For Each RV In TBLLoanPay.DefaultView
                RowLoanPay = RV.Row
                RowHasil = TblHasil.NewDataTable1Row
                RowHasil.Enum_Tipe_Jurnal = ClsTrans_Payment.EnumItemJurnalPayment.LoanInstallment
                RowHasil.Primary_Salary = RowLoanPay.Primary
                RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.LoanInstallment
                RowHasil.Amount_1 = RowLoanPay.Base
                RowHasil.Amount_2 = RowLoanPay.Payment - RowLoanPay.Base
                RowHasil.Primary_Akun_1 = RowLoanPay.Primary_Receivable_Account
                If RowHasil.Amount_2 <> 0 Then
                    If RowHasil.Amount_2 < 0 Then
                        RowHasil.Primary_Akun_2 = RowLoanPay.Primary_Loss_Account
                    Else
                        RowHasil.Primary_Akun_2 = RowLoanPay.Primary_Gain_Account
                    End If
                    'RowHasil.Amount_2 = Math.Abs(RowLoanPay.Payment - RowLoanPay.Installment_Amount)
                Else
                    RowHasil.Primary_Akun_2 = -1
                    RowHasil.Amount_2 = 0
                End If
                RowHasil.Description_Salary = "Loan Installment"
                RowHasil.Employee_Name = vEmployeeName
                RowHasil.Salary_Number = vSalaryNo
                TblHasil.Rows.Add(RowHasil)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputkeDataTableHasil_OtherDeduction(ByVal TblHasil As DseTransPayment.DataTable1DataTable, _
    ByVal TBLOtherDeduction As DseTransPayment.tbltrans_salary_item_otherdeductionDataTable, ByVal vEmployeeName As String, _
    ByVal vSalaryNo As String)
        Dim RowHasil As DseTransPayment.DataTable1Row
        Dim RowOtherDeduction As DseTransPayment.tbltrans_salary_item_otherdeductionRow
        Dim RV As Data.DataRowView
        Try
            For Each RV In TBLOtherDeduction.DefaultView
                RowOtherDeduction = RV.Row
                RowHasil = TblHasil.NewDataTable1Row
                RowHasil.Enum_Tipe_Jurnal = ClsTrans_Payment.EnumItemJurnalPayment.OtherDeduction
                RowHasil.Primary_Salary = RowOtherDeduction.Primary_Transaksi
                RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.OtherDeduction
                RowHasil.Amount_1 = RowOtherDeduction.Amount
                RowHasil.Amount_2 = 0
                RowHasil.Primary_Akun_1 = RowOtherDeduction.Primary_Expense_Account
                RowHasil.Primary_Akun_2 = -1
                RowHasil.Description_Salary = "Other Deduction"
                RowHasil.Employee_Name = vEmployeeName
                RowHasil.Salary_Number = vSalaryNo
                TblHasil.Rows.Add(RowHasil)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputkeDataTableHasil_Deduction(ByVal TblHasil As DseTransPayment.DataTable1DataTable, _
    ByVal TBLDeduction As DseTransPayment.tbltrans_salary_itemDataTable, ByVal vEmployeeName As String, _
    ByVal vSalaryNo As String)
        Dim RowHasil As DseTransPayment.DataTable1Row
        Dim RowDeduction As DseTransPayment.tbltrans_salary_itemRow
        Dim RV As Data.DataRowView
        Try
            For Each RV In TBLDeduction.DefaultView
                RowDeduction = RV.Row
                RowHasil = TblHasil.NewDataTable1Row
                RowHasil.Enum_Tipe_Jurnal = ClsTrans_Payment.EnumItemJurnalPayment.SalaryExpense
                RowHasil.Primary_Salary = RowDeduction.Primary_Transaksi
                RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.Deduction
                RowHasil.Amount_1 = RowDeduction.Amount
                RowHasil.Amount_2 = 0
                RowHasil.Primary_Akun_1 = RowDeduction.Primary_Expense_Account
                RowHasil.Primary_Akun_2 = -1
                RowHasil.Description_Salary = "Deduction-" + RowDeduction.Description
                RowHasil.Employee_Name = vEmployeeName
                RowHasil.Salary_Number = vSalaryNo
                TblHasil.Rows.Add(RowHasil)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputkeDataTableHasil_Income(ByVal TblHasil As DseTransPayment.DataTable1DataTable, _
    ByVal TBLIncome As DseTransPayment.tbltrans_salary_itemDataTable, ByVal vEmployeeName As String, _
    ByVal vSalaryNo As String)
        Dim RowHasil As DseTransPayment.DataTable1Row
        Dim RowIncome As DseTransPayment.tbltrans_salary_itemRow
        Dim RV As Data.DataRowView
        Try
            For Each RV In TBLIncome.DefaultView
                RowIncome = RV.Row
                RowHasil = TblHasil.NewDataTable1Row
                RowHasil.Enum_Tipe_Jurnal = ClsTrans_Payment.EnumItemJurnalPayment.SalaryExpense
                RowHasil.Primary_Salary = RowIncome.Primary_Transaksi
                RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.Income
                RowHasil.Amount_1 = RowIncome.Amount
                RowHasil.Amount_2 = 0
                RowHasil.Primary_Akun_1 = RowIncome.Primary_Expense_Account
                RowHasil.Primary_Akun_2 = -1
                RowHasil.Description_Salary = "Income-" + RowIncome.Description
                RowHasil.Employee_Name = vEmployeeName
                RowHasil.Salary_Number = vSalaryNo
                TblHasil.Rows.Add(RowHasil)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputkeDataTableHasil_EmployeeExpense(ByVal TblHasil As DseTransPayment.DataTable1DataTable, _
    ByVal TblEmployeeExpense As DseTransPayment.tbltrans_salary_itemDataTable, ByVal vEmployeeName As String, _
    ByVal vSalaryNo As String)
        Dim RowHasil As DseTransPayment.DataTable1Row
        Dim RowEmployeeExpense As DseTransPayment.tbltrans_salary_itemRow
        Dim RV As Data.DataRowView

        Try
            For Each RV In TblEmployeeExpense.DefaultView
                RowEmployeeExpense = RV.Row
                RowHasil = TblHasil.NewDataTable1Row
                RowHasil.Enum_Tipe_Jurnal = ClsTrans_Payment.EnumItemJurnalPayment.EmployeeExpense
                RowHasil.Primary_Salary = RowEmployeeExpense.Primary_Transaksi
                RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.EmployeeExpense
                RowHasil.Amount_1 = RowEmployeeExpense.Amount
                RowHasil.Amount_2 = 0
                RowHasil.Primary_Akun_1 = -1
                RowHasil.Primary_Akun_2 = RowEmployeeExpense.Primary_Payment_Account
                RowHasil.Description_Salary = "Company Expense-" + RowEmployeeExpense.Description
                RowHasil.Employee_Name = vEmployeeName
                RowHasil.Salary_Number = vSalaryNo
                TblHasil.Rows.Add(RowHasil)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputkeDataTableHasil_CompanyExpense(ByVal TblHasil As DseTransPayment.DataTable1DataTable, _
    ByVal TblCompanyExpense As DseTransPayment.tbltrans_salary_itemDataTable, ByVal vEmployeeName As String, _
    ByVal vSalaryNo As String)
        Dim RowHasil As DseTransPayment.DataTable1Row
        Dim RowCompanyExpense As DseTransPayment.tbltrans_salary_itemRow
        Dim RV As Data.DataRowView

        Try
            For Each RV In TblCompanyExpense.DefaultView
                RowCompanyExpense = RV.Row
                RowHasil = TblHasil.NewDataTable1Row
                RowHasil.Enum_Tipe_Jurnal = ClsTrans_Payment.EnumItemJurnalPayment.CompanyExpense
                RowHasil.Primary_Salary = RowCompanyExpense.Primary_Transaksi
                RowHasil.Tipe_Income_Deduction = ClsTrans_Payment.EnumLinkLabel.CompanyExpense
                RowHasil.Amount_1 = RowCompanyExpense.Amount
                RowHasil.Amount_2 = 0
                RowHasil.Primary_Akun_1 = RowCompanyExpense.Primary_Expense_Account
                RowHasil.Primary_Akun_2 = RowCompanyExpense.Primary_Payment_Account
                RowHasil.Description_Salary = "Company Expense-" + RowCompanyExpense.Description
                RowHasil.Employee_Name = vEmployeeName
                RowHasil.Salary_Number = vSalaryNo
                TblHasil.Rows.Add(RowHasil)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Enum EnumLinkLabel As Integer
        Income
        Deduction
        OtherDeduction
        LoanInstallment
        PPh21
        Payment
        CompanyExpense
        EmployeeExpense
    End Enum

    Public Sub FungsiBtnHitungPayment_LinkLabel( _
    ByVal vEnum As ClsTrans_Payment.EnumLinkLabel, ByVal vPrimarySalary As Double, _
    ByVal TBLSalary As DseTransPayment.tbltrans_salaryDataTable, _
    ByVal TBLItemSalary_Income As DseTransPayment.tbltrans_salary_itemDataTable, _
    ByVal TBLItemSalary_Deduction As DseTransPayment.tbltrans_salary_itemDataTable, _
    ByVal TBLItemSalary_Other As DseTransPayment.tbltrans_salary_item_otherdeductionDataTable, _
    ByVal TBLLoanPayment As DseTransPayment.TBLTransLoanPaymentDataTable, _
    ByVal TBLLoanPaymentItem As DseTransPayment.TBLTransLoanPaymentItemDataTable)
        Dim DTASalary As New DseTransPaymentTableAdapters.tbltrans_salaryTableAdapter
        Dim DTAItemSalary_Income As New DseTransPaymentTableAdapters.tbltrans_salary_itemTableAdapter
        Dim DTAItemSalary_Deduction As New DseTransPaymentTableAdapters.tbltrans_salary_itemTableAdapter
        Dim DTAItemSalary_Other As New DseTransPaymentTableAdapters.tbltrans_salary_item_otherdeductionTableAdapter
        Dim DTALoanPayment As New DseTransPaymentTableAdapters.TBLTransLoanPaymentTableAdapter
        Dim DTALoanPaymentItem As New DseTransPaymentTableAdapters.TBLTransLoanPaymentItemTableAdapter
        Dim ObjLoanPayment As New ClsTransaksiLoanPayment
        Dim PrimaryLoanPayment As Double = -1
        Try
            'Connection string
            DTASalary.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemSalary_Deduction.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemSalary_Income.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemSalary_Other.Connection.ConnectionString = Core_Module.GetConstring
            DTALoanPayment.Connection.ConnectionString = Core_Module.GetConstring
            DTALoanPaymentItem.Connection.ConnectionString = Core_Module.GetConstring
            'Barsihkan Objek
            TBLSalary.Rows.Clear()
            TBLItemSalary_Deduction.Rows.Clear()
            TBLItemSalary_Income.Rows.Clear()
            TBLItemSalary_Other.Rows.Clear()
            TBLLoanPayment.Rows.Clear()
            TBLLoanPaymentItem.Rows.Clear()
            'Fill data
            If vEnum = EnumLinkLabel.Deduction Then
                DTAItemSalary_Deduction.FillBy_PrimaryTransaksi_Deduction1(TBLItemSalary_Deduction, vPrimarySalary)
            ElseIf vEnum = EnumLinkLabel.Income Then
                DTAItemSalary_Income.FillBy_PrimaryTransaksi_Income1(TBLItemSalary_Income, vPrimarySalary)
            ElseIf vEnum = EnumLinkLabel.OtherDeduction Then
                DTAItemSalary_Other.FillBy_PrimarySalary(TBLItemSalary_Other, vPrimarySalary)
            ElseIf vEnum = EnumLinkLabel.LoanInstallment Then
                DTALoanPayment.FillBy_PrimarySalary(TBLLoanPayment, vPrimarySalary)
                PrimaryLoanPayment = ObjLoanPayment.GetPrimaryByPrimarySalary(vPrimarySalary)
                DTALoanPaymentItem.FillBy_PrimaryTransaksiLoanPayment(TBLLoanPaymentItem, PrimaryLoanPayment)
            ElseIf vEnum = EnumLinkLabel.PPh21 Then
                'Fill Salary PPh 21
                DTASalary.FillBy_PrimarySalary(TBLSalary, vPrimarySalary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASalary.Dispose()
            DTAItemSalary_Deduction.Dispose()
            DTAItemSalary_Income.Dispose()
            DTAItemSalary_Other.Dispose()
            ObjLoanPayment = Nothing
        End Try
    End Sub

    Public Sub FungsiBtnHitungPayment_KumpulkanData(ByVal TBLSalary As DseTransPayment.tbltrans_salaryDataTable, _
    ByVal TBLItem_Employee_d As DseTransPayment.tbltrans_payment_item_employee_dDataTable, _
    ByVal vDGVSalary As DataGridView, ByVal vDGVPayment As DataGridView, _
    ByVal vDGVJurnalItem As DataGridView, ByVal i As FrmTransaksiPayment)
        Dim DTASalary As New DseTransPaymentTableAdapters.tbltrans_salaryTableAdapter
        Dim RowSalary As DseTransPayment.tbltrans_salaryRow
        Dim DTAUser As New DseTransPaymentTableAdapters.TblPayrollAccessTableAdapter
        Dim TBLUser As New DseTransPayment.TblPayrollAccessDataTable
        Dim Filter As String = ""
        Dim idx As Integer = 0
        Try
            'Connection string
            DTASalary.Connection.ConnectionString = Core_Module.GetConstring
            DTAUser.Connection.ConnectionString = Core_Module.GetConstring
            'Settingan Row DGV item
            vDGVSalary.AllowUserToAddRows = False
            vDGVPayment.AllowUserToAddRows = False
            'Barsihkan Objek
            vDGVPayment.Rows.Clear()
            vDGVSalary.Rows.Clear()
            vDGVJurnalItem.Rows.Clear()
            'TBL-Nya
            TBLSalary.Rows.Clear()
            TBLItem_Employee_d.Rows.Clear()
            'Textbox
            i.TxtDeduction.Text = "0"
            i.TxtIncome.Text = "0"
            i.TxtGrossSalary.Text = "0"
            i.TxtLoanInstallment.Text = "0"
            i.TxtIncomeTax.Text = "0"
            i.TxtOtherDedution.Text = "0"
            i.TxtSalaryReceived.Text = "0"
            'Fill TBLUser, dan Salary
            DTAUser.FillBy_KodeUser(TBLUser, i.UserCode)
            DTASalary.FillBy_TanggalTransaksi(TBLSalary, i.DtpTo.Value, i.dtpPeriod.Value, i.TxtPrimary.Text)
            'Kumpulkan data
            If TBLSalary.Rows.Count > 0 Then
                idx = 0
                While idx <= TBLSalary.Rows.Count - 1
                    RowSalary = TBLSalary.Rows(idx)
                    'Cek Payroll Akses
                    Filter = "[Primary Salary Formula] = " + RowSalary.Primary_Salary_Formula.ToString + ""
                    TBLUser.DefaultView.RowFilter = Filter
                    If TBLUser.DefaultView.Count <= 0 Then
                        TBLSalary.Rows.Remove(RowSalary)
                    Else
                        'Lakukan pengecekan dengan TBLSalary, jika primary payment di TBLSalary tidak sama dengan -1, 
                        'maka lakukan pengecekan dengan primary payment yang di textbox, jika tidak sama maka dihapus barisnya 
                        If RowSalary.Primary_Payment <> -1 Then
                            If Convert.ToDouble(i.TxtPrimary.Text) <> Convert.ToDouble(RowSalary.Primary_Payment) Then
                                TBLSalary.Rows.Remove(RowSalary)
                            Else
                                idx += 1
                            End If
                        Else
                            idx += 1
                        End If
                    End If
                End While
            End If
            'Settingan Row DGV item
            vDGVSalary.AllowUserToAddRows = True
            vDGVPayment.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASalary.Dispose()
            DTAUser.Dispose()
            TBLUser.Dispose()
        End Try
    End Sub

    'Public Sub FungsiBtnHitungPayment_KumpulkanData(ByVal TBLSalary As DseTransPayment.tbltrans_salaryDataTable, _
    'ByVal TBLItemSalary_Income As DseTransPayment.tbltrans_salary_itemDataTable, _
    'ByVal TBLItemSalary_Deduction As DseTransPayment.tbltrans_salary_itemDataTable, _
    'ByVal TBLItemSalary_Other As DseTransPayment.tbltrans_salary_item_otherdeductionDataTable, _
    'ByVal TBLLoanPayment As DseTransPayment.TBLTransLoanPaymentDataTable, _
    'ByVal TBLLoanPaymentItem As DseTransPayment.TBLTransLoanPaymentItemDataTable, _
    'ByVal vDGVSalary As DataGridView, ByVal vDGVPayment As DataGridView, _
    'ByVal vDGVJurnalItem As DataGridView, ByVal i As FrmTransaksiPayment)
    '    Dim DTASalary As New DseTransPaymentTableAdapters.tbltrans_salaryTableAdapter
    '    Dim RowSalary As DseTransPayment.tbltrans_salaryRow
    '    Dim DTAUser As New DseTransPaymentTableAdapters.TblPayrollAccessTableAdapter
    '    Dim TBLUser As New DseTransPayment.TblPayrollAccessDataTable
    '    Dim DTAItemSalary_Income As New DseTransPaymentTableAdapters.tbltrans_salary_itemTableAdapter
    '    Dim DTAItemSalary_Deduction As New DseTransPaymentTableAdapters.tbltrans_salary_itemTableAdapter
    '    Dim DTAItemSalary_Other As New DseTransPaymentTableAdapters.tbltrans_salary_item_otherdeductionTableAdapter
    '    Dim DTALoanPayment As New DseTransPaymentTableAdapters.TBLTransLoanPaymentTableAdapter
    '    Dim DTALoanPaymentItem As New DseTransPaymentTableAdapters.TBLTransLoanPaymentItemTableAdapter
    '    Dim ObjLoanPayment As New ClsTransaksiLoanPayment
    '    Dim idx As Integer = 0
    '    Dim Filter As String = ""
    '    Dim PrimaryLoanPayment As Double = -1
    '    Try
    '        'Connection string
    '        DTASalary.Connection.ConnectionString = Core_Module.GetConstring
    '        DTAUser.Connection.ConnectionString = Core_Module.GetConstring
    '        DTAItemSalary_Deduction.Connection.ConnectionString = Core_Module.GetConstring
    '        DTAItemSalary_Income.Connection.ConnectionString = Core_Module.GetConstring
    '        DTAItemSalary_Other.Connection.ConnectionString = Core_Module.GetConstring
    '        DTALoanPayment.Connection.ConnectionString = Core_Module.GetConstring
    '        DTALoanPaymentItem.Connection.ConnectionString = Core_Module.GetConstring
    '        'Settingan Row DGV item
    '        vDGVSalary.AllowUserToAddRows = False
    '        vDGVPayment.AllowUserToAddRows = False
    '        'Fill Setting
    '        DTAItemSalary_Deduction.ClearBeforeFill = False
    '        DTAItemSalary_Income.ClearBeforeFill = False
    '        DTAItemSalary_Other.ClearBeforeFill = False
    '        DTALoanPayment.ClearBeforeFill = False
    '        DTALoanPaymentItem.ClearBeforeFill = False
    '        'Barsihkan Objek
    '        TBLSalary.Rows.Clear()
    '        TBLItemSalary_Deduction.Rows.Clear()
    '        TBLItemSalary_Income.Rows.Clear()
    '        TBLItemSalary_Other.Rows.Clear()
    '        TBLLoanPayment.Rows.Clear()
    '        TBLLoanPaymentItem.Rows.Clear()
    '        vDGVSalary.Rows.Clear()
    '        vDGVPayment.Rows.Clear()
    '        vDGVJurnalItem.Rows.Clear()
    '        i.TxtDeduction.Text = "0"
    '        i.TxtIncome.Text = "0"
    '        i.TxtGrossSalary.Text = "0"
    '        i.TxtLoanInstallment.Text = "0"
    '        i.TxtIncomeTax.Text = "0"
    '        i.TxtOtherDedution.Text = "0"
    '        i.TxtSalaryReceived.Text = "0"
    '        'Fill TBLUser, dan Salary
    '        DTAUser.FillBy_KodeUser(TBLUser, i.UserCode)
    '        DTASalary.FillBy_TanggalTransaksi(TBLSalary, i.DtpTo.Value, i.dtpPeriod.Value)
    '        'Kumpulkan data
    '        If TBLSalary.Rows.Count > 0 Then
    '            idx = 0
    '            While idx <= TBLSalary.Rows.Count - 1
    '                RowSalary = TBLSalary.Rows(idx)
    '                PrimaryLoanPayment = -1
    '                'Cek Payroll Akses
    '                Filter = "[Primary Salary Formula] = " + RowSalary.Primary_Salary_Formula.ToString + ""
    '                TBLUser.DefaultView.RowFilter = Filter
    '                If TBLUser.DefaultView.Count <= 0 Then
    '                    TBLSalary.Rows.Remove(RowSalary)
    '                Else
    '                    'isi TBLItemSalary, Primary-nya sama dengan Primary Salary.
    '                    DTAItemSalary_Deduction.FillBy_PrimaryTransaksi_Deduction(TBLItemSalary_Deduction, RowSalary.Primary)
    '                    DTAItemSalary_Income.FillBy_PrimaryTransaksi_Income(TBLItemSalary_Income, RowSalary.Primary)
    '                    DTAItemSalary_Other.FillBy_PrimarySalary(TBLItemSalary_Other, RowSalary.Primary)
    '                    DTALoanPayment.FillBy_PrimarySalary(TBLLoanPayment, RowSalary.Primary)
    '                    PrimaryLoanPayment = ObjLoanPayment.GetPrimaryByPrimarySalary(RowSalary.Primary)
    '                    DTALoanPaymentItem.FillBy_PrimaryTransaksiLoanPayment(TBLLoanPaymentItem, PrimaryLoanPayment)
    '                    idx += 1
    '                End If
    '            End While
    '        End If
    '        vDGVSalary.AllowUserToAddRows = True
    '        vDGVPayment.AllowUserToAddRows = True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTASalary.Dispose()
    '        DTAUser.Dispose()
    '        TBLUser.Dispose()
    '        DTAItemSalary_Deduction.Dispose()
    '        DTAItemSalary_Income.Dispose()
    '        DTAItemSalary_Other.Dispose()
    '        ObjLoanPayment = Nothing
    '    End Try
    'End Sub

    Public Sub FungsiBtnHitungPayment_IsiUlangDGVSalary(ByVal TBLSalary As DseTransPayment.tbltrans_salaryDataTable, _
    ByVal vDGVSalary As DataGridView, ByVal vDef_Prec As Double, ByVal vBaru As Boolean, ByVal vPrimaryPayment As Double, _
    ByVal i As FrmTransaksiPayment)
        Dim RowSalary As DseTransPayment.tbltrans_salaryRow
        Dim RowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjItem As New ClsTrans_Payment_Item
        Dim CboDGV As DataGridViewComboBoxCell
        Try
            vDGVSalary.AllowUserToAddRows = False
            vDGVSalary.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVSalary.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            If TBLSalary.Rows.Count > 0 Then
                For Each RowSalary In TBLSalary.Rows
                    'Identifer item
                    Dim Identifier As String = ""
                    Identifier = ObjItem.GetIdentifierUrutan _
                    (vDGVSalary, i.ColIdentifier.Index, vBaru, vPrimaryPayment)
                    'Input Salary Item
                    vDGVSalary.Rows.Add()
                    RowDGV = vDGVSalary.Rows(vDGVSalary.Rows.Count - 1)
                    RowDGV.Cells(i.ColEmp_No.Index).Value = ""
                    RowDGV.Cells(i.ColEmp_SalaryDoc.Index).Value = RowSalary.No_Transaksi
                    RowDGV.Cells(i.ColEmp_EmployeeCode.Index).Value = RowSalary.Employee_Code
                    RowDGV.Cells(i.ColEmp_Name.Index).Value = RowSalary.Employee_Name
                    RowDGV.Cells(i.ColEmp_Income.Index).Value = ObjCD.DigitGrouping(RowSalary.Income, vDef_Prec)
                    RowDGV.Cells(i.ColEmp_Deduction.Index).Value = ObjCD.DigitGrouping(RowSalary.Deduction, vDef_Prec)
                    RowDGV.Cells(i.ColEmp_GrossSalary.Index).Value = ObjCD.DigitGrouping(RowSalary.Gross_Salary, vDef_Prec)
                    RowDGV.Cells(i.ColEmp_IncomeTax.Index).Value = ObjCD.DigitGrouping(RowSalary.Income_Tax, vDef_Prec)
                    RowDGV.Cells(i.ColEmp_LoanInstallment.Index).Value = ObjCD.DigitGrouping(RowSalary.Loan_Installment, vDef_Prec)
                    RowDGV.Cells(i.ColEmp_OtherDeduction.Index).Value = ObjCD.DigitGrouping(RowSalary.Other_Deduction, vDef_Prec)
                    RowDGV.Cells(i.ColEmp_TakeHomePay.Index).Value = ObjCD.DigitGrouping(RowSalary.Salary_Received, vDef_Prec)
                    RowDGV.Cells(i.ColEmp_CompanyExpense.Index).Value = ObjCD.DigitGrouping(RowSalary.Company_Expense, vDef_Prec)
                    RowDGV.Cells(i.ColEmp_EmployeeExpense.Index).Value = ObjCD.DigitGrouping(RowSalary.Employee_Expense, vDef_Prec)

                    RowDGV.Cells(i.ColEmp_PrimarySalary.Index).Value = RowSalary.Primary
                    RowDGV.Cells(i.ColEmp_PrimaryEmployee.Index).Value = RowSalary.Primary_Employee
                    RowDGV.Cells(i.ColEmp_Remarks.Index).Value = RowSalary.Remarks

                    RowDGV.Cells(i.ColIdentifier.Index).Value = Identifier
                    RowDGV.Cells(i.ColEmp_SalaryDate.Index).Value = Format(RowSalary.Tanggal_Transaksi, "dd MMM yyyy")
                    CboDGV = RowDGV.Cells(i.ColEmp_TipeTaxIncome.Index)
                    RowDGV.Cells(i.ColEmp_TipeTaxIncome.Index).Value = CboDGV.Items.Item(RowSalary.Tax)
                    RowDGV.Cells(i.ColEmp_TipeTaxIncome_Enum.Index).Value = RowSalary.Tax
                Next
            End If
            ObjCD.SetNumber(vDGVSalary)
            vDGVSalary.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVSalary.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGVSalary.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            ObjItem = Nothing
        End Try
    End Sub

    Public Sub FungsiBtnHitungPayment_IsiTxtSalary(ByVal TBLSalary As DseTransPayment.tbltrans_salaryDataTable, ByVal vIndex As Integer, _
    ByVal vDGVSalary As DataGridView, ByVal vDef_Prec As Double, ByVal i As FrmTransaksiPayment)
        'Dim Filter As String = ""
        'Dim RowSalary As DseTransPayment.tbltrans_salaryRow
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            If String.IsNullOrEmpty(vDGVSalary.Item(i.ColEmp_PrimarySalary.Index, vIndex).Value) = False And _
            CDec(vDGVSalary.Item(i.ColEmp_PrimarySalary.Index, vIndex).Value) <> -1 Then
                'Filter = "[Primary] = " + cdec(vDGVSalary.Item(i.ColEmp_PrimarySalary.Index, vIndex).Value).ToString + ""
                'TBLSalary.DefaultView.RowFilter = Filter
                'If TBLSalary.DefaultView.Count > 0 Then
                '    RowSalary = TBLSalary.DefaultView(0).Row
                '    i.TxtDeduction.Text = ObjCD.DigitGrouping(RowSalary.Deduction, vDef_Prec)
                '    i.TxtIncome.Text = ObjCD.DigitGrouping(RowSalary.Income, vDef_Prec)
                '    i.TxtGrossSalary.Text = ObjCD.DigitGrouping(RowSalary.Gross_Salary, vDef_Prec)
                '    i.TxtLoanInstallment.Text = ObjCD.DigitGrouping(RowSalary.Loan_Installment, vDef_Prec)
                '    i.TxtIncomeTax.Text = ObjCD.DigitGrouping(RowSalary.Income_Tax, vDef_Prec)
                '    i.TxtOtherDedution.Text = ObjCD.DigitGrouping(RowSalary.Other_Deduction, vDef_Prec)
                '    i.TxtSalaryReceived.Text = ObjCD.DigitGrouping(RowSalary.Salary_Received, vDef_Prec)
                'End If
                'TBLSalary.DefaultView.RowFilter = ""
                If vDGVSalary.Rows.Count > 0 Then
                    i.TxtDeduction.Text = vDGVSalary.Item(i.ColEmp_Deduction.Index, vIndex).Value
                    i.TxtIncome.Text = vDGVSalary.Item(i.ColEmp_Income.Index, vIndex).Value
                    i.TxtGrossSalary.Text = vDGVSalary.Item(i.ColEmp_GrossSalary.Index, vIndex).Value
                    i.TxtLoanInstallment.Text = vDGVSalary.Item(i.ColEmp_LoanInstallment.Index, vIndex).Value
                    i.TxtIncomeTax.Text = vDGVSalary.Item(i.ColEmp_IncomeTax.Index, vIndex).Value
                    i.TxtOtherDedution.Text = vDGVSalary.Item(i.ColEmp_OtherDeduction.Index, vIndex).Value
                    i.TxtSalaryReceived.Text = vDGVSalary.Item(i.ColEmp_TakeHomePay.Index, vIndex).Value
                End If
            Else
                i.TxtDeduction.Text = "0"
                i.TxtIncome.Text = "0"
                i.TxtGrossSalary.Text = "0"
                i.TxtLoanInstallment.Text = "0"
                i.TxtIncomeTax.Text = "0"
                i.TxtOtherDedution.Text = "0"
                i.TxtSalaryReceived.Text = "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Sub FungsiBtnHitungPayment_IsiDGVPayment_NilaiAccountDefault(ByVal vobjSetting As Core_Module.ClsTransaksiSetting, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vDGVPayment As DataGridView, ByVal vdef_Prec As Double, _
    ByVal IndexSalaryRow As Integer, ByVal i As FrmTransaksiPayment)
        Dim PrimaryAkun As String = ""
        Dim AkunCode As String = ""
        Dim AkunDesc As String = ""
        Dim Filter As String = ""
        Dim ObjAkun As New AccModule.ClsMasterAkun
        Dim RowDGV As DataGridViewRow
        Dim RowView As Data.DataRowView
        Dim RowItem As DseTransPayment.tbltrans_payment_item_employee_dRow
        Dim PrimaryAkunDefault As String = "-1"
        Dim AkunCodeDefault As String = ""
        Dim AkunDescDefault As String = ""
        Dim DTA As New DseTransPaymentTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        Dim TBL As New DseTransPayment.tblmaster_salary_formula_component_detailemployeeDataTable
        Dim row As DseTransPayment.tblmaster_salary_formula_component_detailemployeeRow
        Dim PrimaryAkunDummy As String = ""

        Try
            i.mTBLItem_Employee_d.Rows.Clear()
            PrimaryAkunDefault = vobjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryCashAccount)
            If PrimaryAkunDefault = "" Then
                PrimaryAkunDefault = "-1"
            End If
            AkunCodeDefault = ObjAkun.GetAccountCode(PrimaryAkunDefault)
            AkunDescDefault = ObjAkun.GetDescription(AkunCodeDefault)

            'isikan default payment account
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)

            If String.IsNullOrEmpty(PrimaryAkun) = False Or PrimaryAkun <> "-1" Then
                vDGVPayment.AllowUserToAddRows = False
                i.dgvItem.AllowUserToAddRows = False
                For Each RowDGV In i.dgvItem.Rows
                    TBL.DefaultView.RowFilter = "[Primary Karyawan] = " + RowDGV.Cells(i.ColEmp_PrimaryEmployee.Index).Value.ToString
                    If TBL.DefaultView.Count = 0 Then
                        PrimaryAkun = PrimaryAkunDefault
                        AkunCode = AkunCodeDefault
                        AkunDesc = AkunDescDefault
                    Else
                        row = TBL.DefaultView(0).Row
                        PrimaryAkun = row.Primary_Default_Payment_Account
                        If PrimaryAkun <> PrimaryAkunDummy Then
                            AkunCode = ObjAkun.GetAccountCode(PrimaryAkun)
                            AkunDesc = ObjAkun.GetDescription(AkunCode)
                            PrimaryAkunDummy = PrimaryAkun
                        End If
                    End If
                    RowItem = i.mTBLItem_Employee_d.Newtbltrans_payment_item_employee_dRow
                    RowItem.primary = -1
                    RowItem.primary_transaksi = CDec(i.TxtPrimary.Text)
                    RowItem.identifier = RowDGV.Cells(i.ColIdentifier.Index).Value
                    RowItem.primary_akun_payment = PrimaryAkun
                    RowItem.akun_payment_code = AkunCode
                    RowItem.akun_payment_desc = AkunDesc
                    RowItem.amount = vObjCD.DigitGrouping(CDec(RowDGV.Cells(i.ColEmp_TakeHomePay.Index).Value), vdef_Prec)
                    RowItem.remarks = ""
                    RowItem.Index_Row_Salary = RowDGV.Index
                    i.mTBLItem_Employee_d.Rows.Add(RowItem)
                Next

                'Setting sesuai dengan index salary, payment item-nya
                vDGVPayment.Rows.Clear()
                vDGVPayment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                vDGVPayment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Filter = "[identifier] = '" + i.dgvItem.Item(i.ColIdentifier.Index, IndexSalaryRow).Value.ToString + "'"
                i.mTBLItem_Employee_d.DefaultView.RowFilter = Filter
                If i.mTBLItem_Employee_d.DefaultView.Count > 0 Then
                    For Each RowView In i.mTBLItem_Employee_d.DefaultView
                        RowItem = RowView.Row
                        vDGVPayment.Rows.Add()
                        RowDGV = vDGVPayment.Rows(vDGVPayment.Rows.Count - 1)
                        RowDGV.Cells(i.ColDetailPayment_AccountCode.Index).Value = RowItem.akun_payment_code
                        RowDGV.Cells(i.ColDetailPayment_AccountDesc.Index).Value = RowItem.akun_payment_desc
                        RowDGV.Cells(i.ColDetailPayment_PrimaryAkun.Index).Value = RowItem.primary_akun_payment
                        RowDGV.Cells(i.ColDetailPayment_Identifier.Index).Value = RowItem.identifier
                        RowDGV.Cells(i.ColDetailPayment_Amount.Index).Value = vObjCD.DigitGrouping(RowItem.amount, vdef_Prec)
                        RowDGV.Cells(i.ColDetailPayment_Remarks.Index).Value = RowItem.remarks
                        RowDGV.Cells(i.ColDetailPayment_IndexRowSalary.Index).Value = RowItem.Index_Row_Salary
                    Next
                End If
                i.mTBLItem_Employee_d.DefaultView.RowFilter = ""
                vDGVPayment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                vDGVPayment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                vDGVPayment.AllowUserToAddRows = True
                i.dgvItem.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjAkun = Nothing
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub FungsiBtnHitungPayment_IsiDGVPayment_NilaiAccountSelectionChanged(ByVal vDGVPayment As DataGridView, _
    ByVal IndexSalaryRow As Integer, ByVal vObjCD As Core_Module.ClsCekData, ByVal vDef_Prec As Double, ByVal i As FrmTransaksiPayment)
        Dim Filter As String = ""
        Dim RowItem As DseTransPayment.tbltrans_payment_item_employee_dRow
        Dim RowView As Data.DataRowView
        Dim RowDGV As DataGridViewRow
        Try
            vDGVPayment.Rows.Clear()
            vDGVPayment.AllowUserToAddRows = False
            vDGVPayment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVPayment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Filter = "[identifier] = '" + i.dgvItem.Item(i.ColIdentifier.Index, IndexSalaryRow).Value.ToString + "'"
            i.mTBLItem_Employee_d.DefaultView.RowFilter = Filter
            If i.mTBLItem_Employee_d.DefaultView.Count > 0 Then
                For Each RowView In i.mTBLItem_Employee_d.DefaultView
                    RowItem = RowView.Row
                    vDGVPayment.Rows.Add()
                    RowDGV = vDGVPayment.Rows(vDGVPayment.Rows.Count - 1)
                    RowDGV.Cells(i.ColDetailPayment_AccountCode.Index).Value = RowItem.akun_payment_code
                    RowDGV.Cells(i.ColDetailPayment_AccountDesc.Index).Value = RowItem.akun_payment_desc
                    RowDGV.Cells(i.ColDetailPayment_PrimaryAkun.Index).Value = RowItem.primary_akun_payment
                    RowDGV.Cells(i.ColDetailPayment_Identifier.Index).Value = RowItem.identifier
                    RowDGV.Cells(i.ColDetailPayment_Amount.Index).Value = vObjCD.DigitGrouping(RowItem.amount, vDef_Prec)
                    RowDGV.Cells(i.ColDetailPayment_Remarks.Index).Value = RowItem.remarks
                    RowDGV.Cells(i.ColDetailPayment_IndexRowSalary.Index).Value = RowItem.Index_Row_Salary
                Next
            End If
            i.mTBLItem_Employee_d.DefaultView.RowFilter = ""
            vDGVPayment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVPayment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGVPayment.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnHitungPayment_IsiDGVPayment_UserDeletingRow(ByVal TBLTrans_Payment_Employee_D As DseTransPayment.tbltrans_payment_item_employee_dDataTable, _
    ByVal RowDGVDelete As DataGridViewRow, ByVal vDGVPayment As DataGridView, ByVal ij As FrmTransaksiPayment)
        Dim Filter As String = ""
        Dim RowItem As DseTransPayment.tbltrans_payment_item_employee_dRow
        Try
            vDGVPayment.AllowUserToAddRows = False
            Filter = "[identifier] = '" + RowDGVDelete.Cells(ij.ColDetailPayment_Identifier.Index).Value.ToString + _
            "' AND [primary akun payment] = " + RowDGVDelete.Cells(ij.ColDetailPayment_PrimaryAkun.Index).Value.ToString + ""
            TBLTrans_Payment_Employee_D.DefaultView.RowFilter = Filter
            If TBLTrans_Payment_Employee_D.DefaultView.Count > 0 Then
                RowItem = TBLTrans_Payment_Employee_D.DefaultView(0).Row
                TBLTrans_Payment_Employee_D.Rows.Remove(RowItem)
            End If
            TBLTrans_Payment_Employee_D.DefaultView.RowFilter = ""
            vDGVPayment.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnHitungPayment_IsiDGVItem_UserDeletingRow(ByVal TBLSalary As DseTransPayment.tbltrans_salaryDataTable, _
    ByVal TBLTrans_Payment_Employee_D As DseTransPayment.tbltrans_payment_item_employee_dDataTable, _
    ByVal RowDGVDeleteSalary As DataGridViewRow, ByVal vDGVSalary As DataGridView, ByVal vDGVPayment As DataGridView, _
    ByVal ij As FrmTransaksiPayment)
        Dim Idn As Integer = 0
        Dim Filter As String = ""
        Dim RowItem As DseTransPayment.tbltrans_payment_item_employee_dRow
        Dim RowSalary As DseTransPayment.tbltrans_salaryRow
        Try
            vDGVSalary.AllowUserToAddRows = False
            vDGVPayment.AllowUserToAddRows = False
            'Data Salary Header
            Filter = "[Primary] = " + RowDGVDeleteSalary.Cells(ij.ColEmp_PrimarySalary.Index).Value.ToString + ""
            TBLSalary.DefaultView.RowFilter = Filter
            If TBLSalary.DefaultView.Count > 0 Then
                RowSalary = TBLSalary.DefaultView(0).Row
                TBLSalary.Rows.Remove(RowSalary)
            End If
            TBLSalary.DefaultView.RowFilter = ""
            'Data Payment
            Filter = "[identifier] = '" + RowDGVDeleteSalary.Cells(ij.ColIdentifier.Index).Value.ToString + _
            "' AND [Index Row Salary] = " + RowDGVDeleteSalary.Index.ToString + ""
            TBLTrans_Payment_Employee_D.DefaultView.RowFilter = Filter
            If TBLTrans_Payment_Employee_D.DefaultView.Count > 0 Then
                While Idn <= TBLTrans_Payment_Employee_D.DefaultView.Count - 1
                    RowItem = TBLTrans_Payment_Employee_D.DefaultView(Idn).Row
                    TBLTrans_Payment_Employee_D.Rows.Remove(RowItem)
                End While
            End If
            TBLTrans_Payment_Employee_D.DefaultView.RowFilter = ""
            vDGVPayment.AllowUserToAddRows = True
            vDGVSalary.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Enum EnumJurnalOtomatis As Integer
        Simpan
        Edit
        Hapus
    End Enum

    Public Sub JurnalOtomatis(ByVal vUserCode As String, ByVal vJournalOtomatis As EnumJurnalOtomatis, ByVal frm As FrmTransaksiPayment, _
    Optional ByVal vPrimaryJournal As Double = -1)
        Dim PrimaryJournal As Double = -1
        Dim ObjAcc As New Core_Module.ClsMasterAkun1
        Dim DtaJournalHeader As New DseTransaksiLoanTableAdapters.tbltrans_jurnalTableAdapter
        Dim DtaJournalItem As New DseTransaksiLoanTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim Dta As New DseTransPaymentTableAdapters.tbltrans_paymentTableAdapter
        Dim RowDGV As DataGridViewRow
        Dim desc As String = ""
        Dim PrimaryPihakAcuan As Double = -1
        Dim Def_Prec As Double = 0
        Try
            desc = frm.TxtJournalDesc.Text
            'desc = "Salary Payment - " + Format(Me.TanggalTransaksi, "MMM yyyy")
            PrimaryPihakAcuan = -1
            Def_Prec = frm.Def_Prec
            If vJournalOtomatis = EnumJurnalOtomatis.Simpan Then
                'baru
                'simpan header
                DtaJournalHeader.Insert(frm.txtJournalNo.Text, Me.TanggalTransaksi, 0, AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, Me.Primary, _
                PrimaryPihakAcuan, desc, vUserCode, Now, vUserCode, Now, Now, -1, -1, 0, 1, 0, 0, 0, 0, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Payment, _
                Def_Prec, AccModule.ClsTransaksiJournal.EnumTipeInOut._In, -1, -1)
                PrimaryJournal = DtaJournalHeader.ScalarQuery_GetPrimaryByDocNo(frm.txtJournalNo.Text)
                Me.mPrimaryJurnalRef = PrimaryJournal
                'Cek Baris Kosong di DGVJournalItem
                If frm.DGVJournalItem.AllowUserToAddRows = True Then
                    frm.DGVJournalItem.AllowUserToAddRows = False
                End If
                'simpan item
                If frm.DGVJournalItem.Rows.Count > 0 Then
                    For Each RowDGV In frm.DGVJournalItem.Rows
                        If Convert.ToDecimal(RowDGV.Cells(frm.ColJurnal_Debet.Index).Value) > 0 Then
                            'debet
                            DtaJournalItem.Insert(PrimaryJournal, RowDGV.Cells(frm.ColJurnal_PrimaryAkun.Index).Value, _
                            CDec(RowDGV.Cells(frm.ColJurnal_Debet.Index).Value), 0, "", CDec(RowDGV.Cells(frm.ColJurnal_Debet.Index).Value), 0, _
                            1, RowDGV.Cells(frm.ColJurnal_PrimaryDept.Index).Value, RowDGV.Cells(frm.ColJurnal_PrimaryJob.Index).Value)
                        Else
                            'kredit
                            DtaJournalItem.Insert(PrimaryJournal, RowDGV.Cells(frm.ColJurnal_PrimaryAkun.Index).Value, _
                            0, CDec(RowDGV.Cells(frm.ColJurnal_Kredit.Index).Value), "", 0, CDec(RowDGV.Cells(frm.ColJurnal_Kredit.Index).Value), _
                            1, RowDGV.Cells(frm.ColJurnal_PrimaryDept.Index).Value, RowDGV.Cells(frm.ColJurnal_PrimaryJob.Index).Value)
                        End If
                    Next
                End If
                'update primary jurnal di tabel loan
                Dta.UpdateQuery_PrimaryJurnalBuatPayment(PrimaryJournal, Me.Primary)
            ElseIf vJournalOtomatis = EnumJurnalOtomatis.Edit Then
                'edit
                'simpan header
                DtaJournalHeader.Update(frm.txtJournalNo.Text, Me.TanggalTransaksi, 0, AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, Me.Primary, _
                PrimaryPihakAcuan, desc, vUserCode, Now, Now, -1, -1, 0, 1, 0, 0, 0, 0, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Payment, _
                Def_Prec, AccModule.ClsTransaksiJournal.EnumTipeInOut._In, -1, -1, Me.mPrimaryJurnalRef)
                PrimaryJournal = Me.mPrimaryJurnalRef
                'hapus item
                DtaJournalItem.Delete(PrimaryJournal)
                'Cek Baris Kosong di DGVJournalItem
                If frm.DGVJournalItem.AllowUserToAddRows = True Then
                    frm.DGVJournalItem.AllowUserToAddRows = False
                End If
                'simpan item
                If frm.DGVJournalItem.Rows.Count > 0 Then
                    For Each RowDGV In frm.DGVJournalItem.Rows
                        If Convert.ToDecimal(RowDGV.Cells(frm.ColJurnal_Debet.Index).Value) > 0 Then
                            'debet
                            DtaJournalItem.Insert(PrimaryJournal, RowDGV.Cells(frm.ColJurnal_PrimaryAkun.Index).Value, _
                            CDec(RowDGV.Cells(frm.ColJurnal_Debet.Index).Value), 0, "", CDec(RowDGV.Cells(frm.ColJurnal_Debet.Index).Value), 0, _
                            1, RowDGV.Cells(frm.ColJurnal_PrimaryDept.Index).Value, RowDGV.Cells(frm.ColJurnal_PrimaryJob.Index).Value)
                        Else
                            'kredit
                            DtaJournalItem.Insert(PrimaryJournal, RowDGV.Cells(frm.ColJurnal_PrimaryAkun.Index).Value, _
                            0, CDec(RowDGV.Cells(frm.ColJurnal_Kredit.Index).Value), "", 0, CDec(RowDGV.Cells(frm.ColJurnal_Kredit.Index).Value), _
                            1, RowDGV.Cells(frm.ColJurnal_PrimaryDept.Index).Value, RowDGV.Cells(frm.ColJurnal_PrimaryJob.Index).Value)
                        End If
                    Next
                End If
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

    Public Function ApakahAkunPaymentCashNBankAtoHutang(ByVal vPrimaryAkun As Double) As Boolean
        Dim ObjAkun As New AccModule.ClsMasterAkun
        Dim TipeAkun2 As AccModule.ClsMasterAkun.EnumTipeAkun2
        Dim ListTipeAkun2 As New List(Of Integer)
        Dim EnumTipeAkun2 As Integer
        Dim Ketemu As Boolean = False

        Try
            TipeAkun2 = ObjAkun.GetTipeAkun2ByPrimaryAkun(vPrimaryAkun)
            ListTipeAkun2.Add(AccModule.ClsMasterAkun.EnumTipeAkun2.Cash_And_Bank)
            ListTipeAkun2.Add(AccModule.ClsMasterAkun.EnumTipeAkun2.Liabilities)
            ListTipeAkun2.Add(AccModule.ClsMasterAkun.EnumTipeAkun2.Long_Term_Liability)
            ListTipeAkun2.Add(AccModule.ClsMasterAkun.EnumTipeAkun2.Other_Current_Liability)
            ListTipeAkun2.Add(AccModule.ClsMasterAkun.EnumTipeAkun2.Account_Payable)
            ListTipeAkun2.Add(AccModule.ClsMasterAkun.EnumTipeAkun2.Other_Payable)
            For Each EnumTipeAkun2 In ListTipeAkun2
                If TipeAkun2 = EnumTipeAkun2 Then
                    Return True
                End If
            Next

            'Jika sampai sini bukan cash and bank atau liabilities
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjAkun = Nothing
            ListTipeAkun2 = Nothing
        End Try
    End Function

    Public Function CetakJournal(ByVal vPrimary As Double, _
    ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    Optional ByVal vTanggalPayment As Date = #1/1/1950#) As Boolean
        Dim Ds As New DseTransaksiLoan
        Dim DTA As New DseTransaksiLoanTableAdapters.DataTable3TableAdapter
        Dim TBL As DseTransaksiLoan.DataTable3DataTable = Ds.DataTable3
        Dim FormatDecimal As Integer = 0
        Dim Negatif As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim ObjJournal As New AccModule.ClsTransaksiJournal
        Dim Report As New CrpCetakJurnal
        Dim Report2 As New CrpCetakJurnal2
        Dim Report3 As New CrpCetakJurnal3
        Dim CetakJurnalPortrait As String = ""
        Dim KodeCur As String = ""
        Dim row As DseTransaksiLoan.DataTable3Row
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)
            'Koding Tambahan untuk description payment
            If vTanggalPayment <> #1/1/1950# Then
                If TBL.Rows.Count > 0 Then
                    For Each row In TBL.Rows
                        row.Desc = "Salary Payment On " + Format(vTanggalPayment, "MMMM yyyy").ToString
                    Next
                End If
            End If
            '-----------------------------------------
            KodeCur = ObjCur.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)

            CetakJurnalPortrait = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CetakJurnalPortrait)

            If CetakJurnalPortrait = "" Or CetakJurnalPortrait = "0" Then
                If ObjJournal.ApakaJurnalNonIDR(vPrimary) = False Then
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
                    Report2.SetParameterValue(Report2.Parameter_Judul.ParameterFieldName, "Payment")

                    Crv.ReportSource = Report2
                    Crv.Zoom(1)
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
                    Report3.SetParameterValue(Report3.Parameter_Judul.ParameterFieldName, "Payment")

                    Crv.ReportSource = Report3
                    Crv.Zoom(1)
                    Crv.Show()
                End If
            Else
                'cetak portrait
                Report.SetDataSource(Ds)
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
                Report.SetParameterValue(Report.Parameter_Judul.ParameterFieldName, "Payment")

                Crv.ReportSource = Report
                Crv.Show()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCur = Nothing
            ObjSetting = Nothing
            ObjJournal = Nothing
            Ds.Dispose()
            DTA.Dispose()
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
