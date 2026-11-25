Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class ClsTrans_SPL_Setting
    Dim mPrimary As Double = -1
    Dim mCode As String = ""
    Dim mDesc As String = ""
    Dim mAllEmployee As Integer = EnumAllEmployee.No
    Dim mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now

    Dim ObjItem As New ClsTrans_SPL_Setting_Item
    Dim ObjCD As New Core_Module.ClsCekData

    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Enum EnumSearch As Integer
        Code
        Desc
        UserCode
        UserName
        Remarks
    End Enum

    Public Enum EnumAllEmployee As Integer
        No
        Yes
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property Code() As String
        Get
            Return mCode
        End Get
        Set(ByVal value As String)
            mCode = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
        End Set
    End Property

    Public Property AllEmployee() As Integer
        Get
            Return mAllEmployee
        End Get
        Set(ByVal value As Integer)
            mAllEmployee = value
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
        Dim Dta As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_scalarTableAdapter

        Try
            '1.strict edit
            Dim UserCode As String = Dta.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If UserCode <> vUserCode Then
                MsgBox("Only user with ID:" + UserCode + " can edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            ''2.status aktif tidak boleh di edit
            'Dim status As Integer = Dta.ScalarQuery_GetStatusByPrimary(vPrimaryTransaksi)
            'If status = ClsTransaksiLoan.EnumStatusLoan.Finished Then
            '    MsgBox("This document no : " + Me.NoTransaksi.ToString + " cannot edit because loan status is finished.", _
            '    MsgBoxStyle.Information, "Information")
            'End If

            '4.tidak boleh jika sudah ada di transaksi loan payment

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vUserCode As String, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim Dta As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_scalarTableAdapter

        Try
            '1.strict edit
            Dim UserCode As String = Dta.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If UserCode <> vUserCode Then
                MsgBox("Only user with ID:" + UserCode + " can delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            ''2.status aktif tidak boleh di edit
            'Dim status As Integer = Dta.ScalarQuery_GetStatusByPrimary(vPrimaryTransaksi)
            'If status = ClsTransaksiLoan.EnumStatusLoan.Finished Then
            '    MsgBox("This document no : " + Me.NoTransaksi.ToString + " cannot delete because loan status is finished.", _
            '    MsgBoxStyle.Information, "Information")
            'End If

            ''3.Jika sudah terjadi closing tidak boleh di edit
            'If vBeginingBalance = False Then
            '    Dim closing As Date = ObjClosing.GetClosingDate
            '    If Me.Date_ < closing Then
            '        MsgBox("You cannot delete this data because this data is closed in period : " + Format(closing, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If

            ''4. tidak boleh hapus jika sudah ada payment yang merujuk
            'If Me.GetJmlLoanPayment(vPrimaryTransaksi) > 0 Then
            '    MsgBox("You cannot delete this data, there are loan payment refering to this data", MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim dta As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_scalarTableAdapter
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

    Private Function ApakahDataBenar(ByVal Frm As FrmTrans_SPL_Setting, ByVal vBaru As Boolean, _
    ByVal vDGV_Scope As DataGridView, ByVal vDGV_User As DataGridView) As Boolean
        Dim i As New FrmTrans_SPL_Setting
        Dim SudahAda As Boolean = False

        Try
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mCode) = True Then
                    SudahAda = True
                End If
            Else
                Dim NoTransaksiLama As String = Me.GetCode(Me.mPrimary)
                If NoTransaksiLama <> Me.mCode Then
                    If Me.ApakahSudahAda(Me.mCode) = True Then
                        SudahAda = True
                    End If
                End If
            End If

            If SudahAda = True Then
                MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'dalam satu user hanya bisa ada di satu user group
            For Each row As DataGridViewRow In vDGV_User.Rows
                If row.Cells(i.ColUserPrimaryUser.Index).Value <> -1 And String.IsNullOrEmpty(row.Cells(i.ColUserPrimaryUser.Index).Value) = False Then
                    If Me.ApakahUserSudahDigunakan(Me.mPrimary, row.Cells(i.ColUserPrimaryUser.Index).Value) = True Then
                        MsgBox("Department (" & row.Cells(i.ColUserCode.Index).Value & " - " & row.Cells(i.ColUserName.Index).Value & ") has been used in period", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Function

    Public Function ApakahUserSudahDigunakan(ByVal vPrimaryTransaksi As Double, ByVal vPrimaryUser As Double) As Boolean
        Dim Dta As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_scalarTableAdapter
        Dim JmlUser As Nullable(Of Long)

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            JmlUser = Dta.ScalarQuery_GetJmlUser(vPrimaryUser, vPrimaryTransaksi)
            If JmlUser.HasValue = True Then
                If JmlUser.Value > 0 Then
                    Return True
                End If
            Else
                Return False
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmTrans_SPL_Setting, _
    ByVal vDgv_Scope As Windows.Forms.DataGridView, ByVal vDgv_User As Windows.Forms.DataGridView) As Boolean
        Dim DTA As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_groupTableAdapter
        Dim TBL As New DseTrans_SPL_Setting.tbltrans_spl_approval_groupDataTable
        Dim row As DseTrans_SPL_Setting.tbltrans_spl_approval_groupRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mCode = .Code
                    Me.mDesc = .Description
                    Me.mRemarks = .Remarks
                    Me.mAllEmployee = .All_Employee
                    Me.mRemarks = .Remarks

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp

                    Me.ObjItem.IsiDGVItem(vDgv_Scope, vDgv_User, .Primary, vFrm)
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

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim d As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_scalarTableAdapter
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

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim dta As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_scalarTableAdapter
        Dim Dummy As String

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetCode(vPrimary)
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

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim dta As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_scalarTableAdapter
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

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vDGV_Scope As DataGridView, ByVal vDGV_User As DataGridView, _
   ByVal frm As FrmTrans_SPL_Setting) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_SimpanBaru, vUserCode)

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
                            retval = Me.SimpanBaru1(vUserCode, vDGV_Scope, vDGV_User, frm)
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vDGV_Scope As DataGridView, ByVal vDGV_User As DataGridView, _
    ByVal frm As FrmTrans_SPL_Setting) As Boolean
        Dim DTA As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_groupTableAdapter
        Dim BufferDoc As String = Me.mCode

        Try
            If Me.ApakahDataBenar(frm, True, vDGV_Scope, vDGV_User) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring

                If DTA.Insert(Me.mCode, Me.mDesc, Me.mRemarks, Me.mAllEmployee, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = Me.GetPrimary(Me.mCode)
                    Me.mTimestamp = Me.GetTimeStamp(Me.mPrimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    Me.ObjItem.SimpanItemScope(Me.Primary, vDGV_Scope, frm)
                    Me.ObjItem.SimpanItemUser(Me.Primary, vDGV_User, frm)

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

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vDGV_Scope As DataGridView, _
    ByVal vDGV_User As DataGridView, ByVal frm As FrmTrans_SPL_Setting) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_SimpanPerubahan, vUserCode)

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
                            retval = Me.SimpanPerubahan1(vUserCode, vDGV_Scope, vDGV_User, frm)
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

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal vDGV_Scope As DataGridView, ByVal vDGV_User As DataGridView, _
    ByVal frm As FrmTrans_SPL_Setting) As Boolean
        Dim DTA As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_groupTableAdapter
        Dim BufferDoc As String = Me.mCode

        Try
            If Me.ApakahDataBenar(frm, False, vDGV_Scope, vDGV_User) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                If DTA.Update(Me.mCode, Me.mDesc, Me.mRemarks, Me.mAllEmployee, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = Me.GetTimeStamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    Me.ObjItem.HapusItemScope(Me.mPrimary)
                    Me.ObjItem.SimpanItemScope(Me.Primary, vDGV_Scope, frm)
                    Me.ObjItem.HapusItemUser(Me.mPrimary)
                    Me.ObjItem.SimpanItemUser(Me.Primary, vDGV_User, frm)

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

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, ByVal frm As FrmTrans_SPL_Setting) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_Hapus, vUserCode)

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

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal frm As FrmTrans_SPL_Setting) As Boolean
        Dim DTA As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahBolehHapus(vUserCode, vPrimary) = True Then
                DTA.Delete(vPrimary)
                Me.ObjItem.HapusItemScope(vPrimary)
                Me.ObjItem.HapusItemUser(vPrimary)

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
    ByVal vTBL As DseTrans_SPL_Setting.tbltrans_spl_approval_group_fillDataTable, ByVal SearchCriteria As EnumSearch, _
    ByVal tsl As ToolStripLabel, ByVal vTanggalDari As Date, _
    ByVal vTanggalSampai As Date, ByVal vFrm As FrmTrans_SPL_Setting, Optional ByVal vJml As Integer = 0) As Boolean
        Dim DTA As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_fillTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearch.Code Then
                QueryResult = DTA.Fill_Like(vTBL, Criteria + "%", "%", "%", "%", "%")
            ElseIf SearchCriteria = EnumSearch.Desc Then
                QueryResult = DTA.Fill_Like(vTBL, "%", Criteria + "%", "%", "%", "%")
            ElseIf SearchCriteria = EnumSearch.UserCode Then
                QueryResult = DTA.Fill_Like(vTBL, "%", "%", "%", Criteria + "%", "%")
            ElseIf SearchCriteria = EnumSearch.UserName Then
                QueryResult = DTA.Fill_Like(vTBL, "%", "%", "%", "%", Criteria + "%")
            ElseIf SearchCriteria = EnumSearch.Remarks Then
                QueryResult = DTA.Fill_Like(vTBL, "%", "%", Criteria + "%", "%", "%")
            End If

            'update dgv
            vDGV.DataSource = vTBL.DefaultView
            vDGV.Refresh()
            If vDGV.ColumnCount > 0 Then
                vDGV.Columns(0).Visible = False
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

    Public Function GetJml(ByVal vCriteria As String, ByVal vDGV As DataGridView, ByVal vSearchCriteria As EnumSearch, _
    ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date) As Double
        Dim Dta As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_fillTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try

            Dta.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchCriteria = EnumSearch.Code Then
                Dummy = Dta.ScalarQuery_GetJmlLike(vCriteria + "%", "%", "%", "%", "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSearch.Desc Then
                Dummy = Dta.ScalarQuery_GetJmlLike("%", vCriteria + "%", "%", "%", "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSearch.UserCode Then
                Dummy = Dta.ScalarQuery_GetJmlLike("%", "%", "%", vCriteria + "%", "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSearch.UserName Then
                Dummy = Dta.ScalarQuery_GetJmlLike("%", "%", "%", "%", vCriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSearch.Remarks Then
                Dummy = Dta.ScalarQuery_GetJmlLike("%", "%", vCriteria + "%", "%", "%")
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
End Class
