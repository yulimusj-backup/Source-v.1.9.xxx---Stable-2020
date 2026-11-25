Option Explicit On
Option Strict Off

Imports system.Windows.Forms
Public Class ClsTrans_Gaji_MasaSblm
    Dim mPrimary As Double = -1
    Dim mPrimaryEmployee As Double = -1
    Dim mNoTransaksi As String = ""
    Dim mDate As Date = #1/1/1950#
    Dim mPeriodStart As Date = #1/1/1950#
    Dim mPeriodEnd As Date = #1/1/1950#
    Dim mEmployer As String = ""
    Dim mNettIncome As Decimal = 0
    Dim mIncomeTaxPaid As Decimal = 0
    Dim mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUser As New Core_Module.ClsUser

    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Enum EnumSearch As Integer
        By_NoTransaksi
        By_Date
        By_PeriodStart
        By_PeriodEnd
        By_EmployeeCode
        By_EmployeeName
        By_Employer
        By_Remarks
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property PrimaryEmployee() As Double
        Get
            Return mPrimaryEmployee
        End Get
        Set(ByVal value As Double)
            mPrimaryEmployee = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property PeriodStart() As Date
        Get
            Return mPeriodStart
        End Get
        Set(ByVal value As Date)
            mPeriodStart = value
        End Set
    End Property

    Public Property PeriodEnd() As Date
        Get
            Return mPeriodEnd
        End Get
        Set(ByVal value As Date)
            mPeriodEnd = value
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

    Public Property Employer() As String
        Get
            Return mEmployer
        End Get
        Set(ByVal value As String)
            mEmployer = value
        End Set
    End Property

    Public Property NettIncome() As Decimal
        Get
            Return mNettIncome
        End Get
        Set(ByVal value As Decimal)
            mNettIncome = value
        End Set
    End Property

    Public Property IncomeTaxPaid() As Decimal
        Get
            Return mIncomeTaxPaid
        End Get
        Set(ByVal value As Decimal)
            mIncomeTaxPaid = value
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

    Public Function ApakahBolehedit(ByVal vUserCode As String, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim Dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_incomeTableAdapter
        Dim Dummy As String = ""

        Try
            '1.strict edit
            Dim UserCode As String = Dta.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If UserCode <> vUserCode Then
                If Me.ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_Gaji_MasaSblm_EditUserLain") = False Then
                    MsgBox("Only user with ID:" + UserCode + " can edit this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'Jika sudah ada transaksi salary atas karyawan tersebut yang menggunakan doc tsb, maka tidak bisa hapus, tidak bisa edit, tidak bisa simpan perubahan
            Dummy = Me.ApakahDocumentSudahDigunakanTransaksiSalary1(Me.mPrimary)
            If Dummy <> "" Then
                MsgBox("Previous income data is reference by salary calculation: " + Dummy + ", you can't edit this data.", MsgBoxStyle.Information, "Information") 'msgbox english
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vUserCode As String, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim Dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_incomeTableAdapter
        Dim Dummy As String = ""

        Try
            '1.strict edit
            Dim UserCode As String = Dta.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If UserCode <> vUserCode Then
                If Me.ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_Gaji_MasaSblm_HapusUserLain") = False Then
                    MsgBox("Only user with ID:" + UserCode + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'Jika sudah ada transaksi salary atas karyawan tersebut yang menggunakan doc tsb, maka tidak bisa hapus, tidak bisa edit, tidak bisa simpan perubahan
            Dummy = Me.ApakahDocumentSudahDigunakanTransaksiSalary1(Me.mPrimary)
            If Dummy <> "" Then
                MsgBox("Previous income data is reference by salary calculation: " + Dummy + ", you can't delete this data.", MsgBoxStyle.Information, "Information") 'msgbox english
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Private Function ApakahDataBenar(ByVal Frm As FrmTrans_Gaji_MasaSblm, ByVal vBaru As Boolean) As Boolean
        Dim i As New FrmTrans_Gaji_MasaSblm
        Dim Dummy As String = ""

        Try
            '1.tidak boleh double
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.NoTransaksi) = True Then
                    Return False
                End If
            Else
                Dim NoTransaksiLama As String = Me.GetNoTransaksiByPrimary(Me.mPrimary)
                If NoTransaksiLama <> Me.mNoTransaksi Then
                    Return False
                End If

                'Jika sudah ada transaksi salary atas karyawan tersebut yang menggunakan doc tsb, maka tidak bisa hapus, tidak bisa edit, tidak bisa simpan perubahan
                Dummy = Me.ApakahDocumentSudahDigunakanTransaksiSalary1(Me.mPrimary)
                If Dummy <> "" Then
                    MsgBox("Previous income data is reference by salary calculation: " + Dummy + ", you can't save this data.", MsgBoxStyle.Information, "Information") 'msgbox english
                    Return False
                End If
            End If

            'satu employee, tidak bisa memiliki lebih dari 1 form A1 yang bersinggungan start period dan end period-nya
            If ApakahEmployeeSudahAdaInPeriod(Me.mPrimaryEmployee, Me.mPeriodStart, Me.mPeriodEnd, Me.mPrimary) = True Then
                Return False
            End If

            'Jika sudah ada transaksi salary yang resign atau akhir tahun atas karyawan tersebut, maka tidak bisa simpan baru.
            'Jika sudah ada transaksi salary atas karyawan ybs yg <= end period, maka tidak bisa simpan/simpan perubahan
            'Dim dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_item_prev_incomeTableAdapter
            'Dim tbl As New DseTrans_Gaji_MasaSblm.tbltrans_salary_item_prev_incomeDataTable
            'Dim row As DseTrans_Gaji_MasaSblm.tbltrans_salary_item_prev_incomeRow
            Dim dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salaryTableAdapter
            Dim tbl As New DseTrans_Gaji_MasaSblm.tbltrans_salaryDataTable
            Dim row As DseTrans_Gaji_MasaSblm.tbltrans_salaryRow

            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(tbl, Me.mPrimaryEmployee) > 0 Then
                For ij As Integer = 0 To tbl.Rows.Count - 1
                    row = tbl.Rows(ij)
                    With row
                        ''Jika sudah ada transaksi salary yang resign maka tidak bisa simpan
                        'If .salary_is_resign <> 0 Then
                        '    MsgBox("This employee has a resignation salary calculation transaction, you can't save this data.", MsgBoxStyle.Information, "Information") 'msgbox english
                        '    Return False
                        'End If
                        ''Jika sudah ada transaksi salary atas karyawan ybs yg <= end period (atau akhir tahun atas karyawan tersebut) maka tidak bisa simpan
                        'If .salary_end_period <= Me.mPeriodEnd Then
                        '    MsgBox("This employee has a salary calculation transaction period <= " & Format(Me.mPeriodEnd, "MMM yyyy") & ", you can't save this data.", MsgBoxStyle.Information, "Information") 'msgbox english
                        '    Return False
                        'End If
                    End With
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmTrans_Gaji_MasaSblm) As Boolean
        Dim DTA As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_incomeTableAdapter
        Dim TBL As New DseTrans_Gaji_MasaSblm.tbltrans_salary_prev_incomeDataTable
        Dim row As DseTrans_Gaji_MasaSblm.tbltrans_salary_prev_incomeRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mPrimaryEmployee = .primary_employee
                    Me.mNoTransaksi = .a1_form_no
                    Me.mDate = ._date
                    Me.mPeriodStart = .start_period
                    Me.mPeriodEnd = .end_period
                    Me.mEmployer = .employer
                    Me.mNettIncome = .nett_income
                    Me.mIncomeTaxPaid = .income_tax_paid
                    Me.mRemarks = .remarks

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
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

    Public Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_incomeTableAdapter
        Dim dummy As Nullable(Of Long)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJml(vNoTransaksi)
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

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim d As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_incomeTableAdapter
        Dim n As New Nullable(Of Double)

        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            n = d.ScalarQuery_GetPrimary(vNoTransaksi)
            If n.HasValue Then
                Return n.Value
            End If

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            d.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksiByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_incomeTableAdapter
        Dim Dummy As String

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_Get_No_Transaksi(vPrimary)
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
        Dim dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_incomeTableAdapter
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

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal frm As FrmTrans_Gaji_MasaSblm) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Salary_SimpanBaru, vUserCode)

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
                            retval = Me.SimpanBaru1(vUserCode, frm)
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal frm As FrmTrans_Gaji_MasaSblm) As Boolean
        Dim DTA As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_incomeTableAdapter
        Dim BufferDoc As String = Me.mNoTransaksi

        Try
            If Me.ApakahDataBenar(frm, True) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                If DTA.Insert(Me.mPrimaryEmployee, Me.mPeriodStart, Me.mPeriodEnd, Me.mNoTransaksi, Me.mEmployer, Me.mDate, Me.mNettIncome, _
                Me.mIncomeTaxPaid, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = Me.GetPrimary(Me.mNoTransaksi)
                    Me.mTimestamp = Me.GetTimeStampByPrimary(Me.mPrimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
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
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal frm As FrmTrans_Gaji_MasaSblm) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Salary_SimpanPerubahan, vUserCode)

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
                            retval = Me.SimpanPerubahan1(vUserCode, frm)
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

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal frm As FrmTrans_Gaji_MasaSblm) As Boolean
        Dim DTA As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_incomeTableAdapter
        Dim BufferDoc As String = Me.mNoTransaksi

        Try
            If Me.ApakahDataBenar(frm, False) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                If DTA.Update(Me.mPrimaryEmployee, Me.mPeriodStart, Me.mPeriodEnd, Me.mNoTransaksi, Me.mEmployer, Me.mDate, Me.mNettIncome, _
                Me.mIncomeTaxPaid, Me.mRemarks, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = Me.GetTimeStampByPrimary(Me.mPrimary)
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
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, ByVal frm As FrmTrans_Gaji_MasaSblm) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Salary_Hapus, vUserCode)

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

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal frm As FrmTrans_Gaji_MasaSblm) As Boolean
        Dim DTA As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_incomeTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahBolehHapus(vUserCode, vPrimary) = True Then
                DTA.Delete(vPrimary)

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

    Public Function IsiDGV(ByVal Criteria As String, ByVal vDGV As DataGridView, _
    ByVal vTBL As DseTrans_Gaji_MasaSblm.tbltrans_salary_prev_income_cariDataTable, ByVal SearchCriteria As EnumSearch, _
    ByVal tsl As ToolStripLabel, ByVal vTanggalDari As Date, _
    ByVal vTanggalSampai As Date, ByVal vFrm As FrmTrans_Gaji_MasaSblm, Optional ByVal vJml As Integer = 0) As Boolean
        Dim DTA As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_income_cariTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearch.By_NoTransaksi Then
                QueryResult = DTA.Fill(vTBL, "%", "%", "%", Criteria + "%", #1/1/1950#, #1/1/2200#, "%", "%")
            ElseIf SearchCriteria = EnumSearch.By_Date Then
                QueryResult = DTA.Fill(vTBL, "%", "%", "%", "%", vTanggalDari, vTanggalSampai, "%", "%")
            ElseIf SearchCriteria = EnumSearch.By_EmployeeCode Then
                QueryResult = DTA.Fill(vTBL, Criteria + "%", "%", "%", "%", #1/1/1950#, #1/1/2200#, "%", "%")
            ElseIf SearchCriteria = EnumSearch.By_EmployeeName Then
                QueryResult = DTA.Fill(vTBL, "%", "%", Criteria + "%", "%", #1/1/1950#, #1/1/2200#, "%", "%")
            ElseIf SearchCriteria = EnumSearch.By_Employer Then
                QueryResult = DTA.Fill(vTBL, "%", "%", "%", "%", #1/1/1950#, #1/1/2200#, Criteria + "%", "%")
            ElseIf SearchCriteria = EnumSearch.By_Remarks Then
                QueryResult = DTA.Fill(vTBL, "%", "%", "%", "%", #1/1/1950#, #1/1/2200#, "%", Criteria + "%")
            End If

            'update dgv
            vDGV.DataSource = vTBL.DefaultView
            vDGV.Refresh()
            If vDGV.ColumnCount > 0 Then
                vDGV.Columns(0).Visible = False
                vDGV.Columns(2).Visible = False
                vDGV.Columns(5).DefaultCellStyle.Format = "dd MMM yyyy"
                vDGV.Columns(7).DefaultCellStyle.Format = "MMM yyyy"
                vDGV.Columns(8).DefaultCellStyle.Format = "MMM yyyy"
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
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vJml = QueryResult

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, ByVal vDGV As DataGridView, ByVal SearchCriteria As EnumSearch, _
    ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date) As Double
        Dim Dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_income_cariTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try

            Dta.Connection.ConnectionString = Core_Module.GetConstring

            If SearchCriteria = EnumSearch.By_NoTransaksi Then
                Dummy = Dta.ScalarQuery_GetJmlLIKE("%", "%", "%", Criteria + "%", #1/1/1950#, #1/1/2200#, "%", "%")
            ElseIf SearchCriteria = EnumSearch.By_Date Then
                Dummy = Dta.ScalarQuery_GetJmlLIKE("%", "%", "%", "%", vTanggalDari, vTanggalSampai, "%", "%")
            ElseIf SearchCriteria = EnumSearch.By_EmployeeCode Then
                Dummy = Dta.ScalarQuery_GetJmlLIKE(Criteria + "%", "%", "%", "%", #1/1/1950#, #1/1/2200#, "%", "%")
            ElseIf SearchCriteria = EnumSearch.By_EmployeeName Then
                Dummy = Dta.ScalarQuery_GetJmlLIKE("%", "%", Criteria + "%", "%", #1/1/1950#, #1/1/2200#, "%", "%")
            ElseIf SearchCriteria = EnumSearch.By_Employer Then
                Dummy = Dta.ScalarQuery_GetJmlLIKE("%", "%", "%", "%", #1/1/1950#, #1/1/2200#, Criteria + "%", "%")
            ElseIf SearchCriteria = EnumSearch.By_Remarks Then
                Dummy = Dta.ScalarQuery_GetJmlLIKE("%", "%", "%", "%", #1/1/1950#, #1/1/2200#, "%", Criteria + "%")
            End If

            If Dummy.HasValue = True Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function ApakahEmployeeSudahAdaInPeriod(ByVal vPrimaryEmp As Double, ByVal vPeriodStart As Date, ByVal vPeriodTo As Date, ByVal vPrimary As Double) As Boolean
        Dim dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_incomeTableAdapter
        Dim Retval As Nullable(Of Long)
        Dim StringDate As String = ""

        Try
            'satu employee, tidak bisa memiliki lebih dari 1 form A1 yang bersinggungan start period dan end period-nya
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Retval = dta.ScalarQuery_GetJml_02(vPrimaryEmp, vPeriodTo, vPeriodStart, vPrimary)
            If Retval.HasValue Then
                If Retval.Value > 0 Then
                    If vPeriodStart <> vPeriodTo Then
                        StringDate = Format(vPeriodStart, "MMM yyyy") & " to " & Format(vPeriodTo, "MMM yyyy")
                    Else
                        StringDate = Format(vPeriodStart, "MMM yyyy")
                    End If
                    MsgBox("This employee already has a previous income transaction in the period range " & StringDate, MsgBoxStyle.Information, "Information") 'msgbox english
                    Return True
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahDocumentSudahDigunakanTransaksiSalary1(ByVal vPrimary As Double) As String
        Dim dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_item_prev_incomeTableAdapter
        Dim TBL As New DseTrans_Gaji_MasaSblm.tbltrans_salary_item_prev_incomeDataTable
        Dim row As DseTrans_Gaji_MasaSblm.tbltrans_salary_item_prev_incomeRow
        Dim Retval As Nullable(Of Long)

        Try
            'jika sudah ada transaksi salary atas karyawan tersebut yang menggunakan doc tsb, maka tidak bisa hapus, tidak bisa edit, tidak bisa simpan perubahan
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.salary_no
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

    Public Function Cetak(ByVal vPrimary As Double, ByVal vDateFrom As Date, ByVal vDateTo As Date, _
    ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean
        Dim report As New CrpTrans_Gaji_MasaSblm_Landscape
        Dim ds As New DseTrans_Gaji_MasaSblm
        Dim dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_prev_incomeTableAdapter
        Dim tbl As DseTrans_Gaji_MasaSblm.tbltrans_salary_prev_incomeDataTable = ds.tbltrans_salary_prev_income
        Dim row As DseTrans_Gaji_MasaSblm.tbltrans_salary_prev_incomeRow
        Dim Par_Period As String = Format(vDateFrom, "dd MMMM yyyy") & " to " & Format(vDateTo, "dd MMMM yyyy")

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.FillBy_DateStartEnd(tbl, vDateFrom, vDateTo)

            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_Param_Date.ParameterFieldName, Par_Period)
            vCrv.ReportSource = report
            vCrv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
