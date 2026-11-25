Option Explicit On
Option Strict Off

Public Class ClsMasterPayrollAccess
    Private mPrimary As Double = -1
    Private mCode As String = ""
    Private mDesc As String = ""
    Private mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

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

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_accessTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimary(vCode)
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

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_accessTableAdapter
        Dim Dummy As String
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetCode(vPrimary)
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

    Public Function GetJmlByCode(ByVal vCode As String) As Double
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_accessTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByCode(vCode)
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

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_accessTableAdapter
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

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_accessTableAdapter
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

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_accessTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByCode(vCode).Value > 0 Then
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

    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Dim ObjCD As New ClsCekData
        Try
            'tidak boleh ada kode yang sama
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mCode) = True Then
                    If vShowMsg = True Then
                        MsgBox("Duplication in payroll access code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.GetCode(Me.mPrimary)
                If CodeLama <> Me.mCode Then
                    If Me.ApakahSudahAda(Me.mCode) = True Then
                        If vShowMsg = True Then
                            MsgBox("Duplication in payroll access code, please change your code", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. kode
            If Me.mCode = "" Then
                If vShowMsg = True Then
                    MsgBox("Please input payroll access code.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '2. nama
            If Me.mDesc = "" Then
                If vShowMsg = True Then
                    MsgBox("Please input payroll access description.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vStrictEdit As Boolean, _
    ByVal vShowMsg As Boolean, ByVal i As FrmMasterPayrollAccess) As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vStrictDelete As Boolean, _
    ByVal vShowMsg As Boolean, ByVal i As FrmMasterPayrollAccess) As Boolean
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_access_item_salaryformulaTableAdapter
        Dim TBL0 As New DseMasterPayrollAkses.tblmaster_payroll_access_item_salaryformulaDataTable
        Dim TBL1 As New DseMasterPayrollAkses.tblmaster_payroll_access_item_salaryformulaDataTable
        'Dim Row0 As DseMasterPayrollAkses.tblmaster_payroll_access_item_salaryformulaRow
        Dim Row1 As DseMasterPayrollAkses.tblmaster_payroll_access_item_salaryformulaRow
        Dim Filter As String = ""
        Dim ErrorMsg As String = ""
        Dim ij As Integer = 1
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL0, vPrimary)
            DTA.FillBy_PrimaryMaster(TBL1, vPrimary)
            If TBL1.Rows.Count > 0 Then
                For Each Row1 In TBL1.Rows
                    Filter = "[Primary Salary Formula] = " + Row1.Primary_Salary_Formula.ToString + ""
                    TBL0.DefaultView.RowFilter = Filter
                    If TBL0.DefaultView.Count <= 0 Then
                        If String.IsNullOrEmpty(ErrorMsg) = True Then
                            ErrorMsg = ij.ToString + ". " + Row1.code.ToString + ""
                        Else
                            ErrorMsg += vbCrLf + ij.ToString + ". " + Row1.code.ToString + ""
                        End If
                        ij += 1
                    End If
                Next
                If String.IsNullOrEmpty(ErrorMsg) = False Then
                    If vShowMsg = True Then
                        MsgBox("The following salary formula will have no access : " + vbCrLf + ErrorMsg, MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL0.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterPayrollAkses.tblmaster_payroll_access1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_access1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescriptionLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(1).HeaderText = "Code"
                DGV.Columns(2).HeaderText = "Description"
                DGV.Columns(3).HeaderText = "Remarks"
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As EnumSearchCriteria) As Double
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_access1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                Return DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByDesciptionLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vTBL_SF As DseMasterPayrollAkses.tblmaster_salary_formula1DataTable, _
    ByVal vTBL_User As DseMasterPayrollAkses.tbluserDataTable, ByVal i As FrmMasterPayrollAccess) As Boolean
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_accessTableAdapter
        Dim TBL As New DseMasterPayrollAkses.tblmaster_payroll_accessDataTable
        Dim row As DseMasterPayrollAkses.tblmaster_payroll_accessRow
        Dim ObjItem As New ClsMasterPayrollAccess_Item
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mCode = .Code
                    Me.mDesc = .Description
                    Me.mRemarks = .Remarks

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
            End If
            ObjItem.IsiDGV(vPrimary, vTBL_SF, vTBL_User, i)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vCLB_SalaryFormula As CheckedListBox, ByVal vCLB_Users As CheckedListBox)
        Try
            mPrimary = -1
            mCode = ""
            mDesc = ""
            mRemarks = ""

            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today

            vCLB_SalaryFormula.DataSource = Nothing
            vCLB_Users.DataSource = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vPrimary As Double, ByVal ij As FrmMasterPayrollAccess) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterPayrollAccess_SimpanBaru, vUserCode)
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
                            retval = Me.SimpanBaru1(vUserCode, vObjCD, vPrimary, ij)
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
    ByVal vPrimary As Double, ByVal ij As FrmMasterPayrollAccess) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterPayrollAccess_SimpanPerubahan, vUserCode)

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
                            retval = Me.SimpanPerubahan1(vUserCode, vPrimary, vObjCD, ij)
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
    ByVal vPrimary As Double, ByVal ij As FrmMasterPayrollAccess, _
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
            Me.vPrimary_Antrian, vPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterPayrollAccess_Hapus, vUserCode)

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
                            retval = Me.Hapus1(vPrimary, vUserCode, True, TampilkanPesan, ij)
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
    ByVal ij As FrmMasterPayrollAccess) As Boolean
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_accessTableAdapter
        Dim ObjItem As New ClsMasterPayrollAccess_Item
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, True) = True Then
                If DTA.Insert(Me.mCode, Me.mDesc, Me.mRemarks, _
                vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCode)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, ij)

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
    ByVal vPrimary As Double, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal ij As FrmMasterPayrollAccess) As Boolean
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_accessTableAdapter
        Dim ObjItem As New ClsMasterPayrollAccess_Item
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False, True) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                If DTA.Update(Me.mCode, Me.mDesc, Me.mRemarks, _
                vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)
                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, ij)
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
    ByVal ij As FrmMasterPayrollAccess) As Boolean
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_accessTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, vStrictDelete, vShowMsg, ij) = True Then
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
        End Try
    End Function

    Public Function ApakahMemilikiPayrollAccess(ByVal vPrimarySalaryFormula As Double, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_access_item_salaryformulaTableAdapter
        Dim TBL As New DseMasterPayrollAkses.tblmaster_payroll_access_item_salaryformulaDataTable
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_UserCode(TBL, vUserCode, vPrimarySalaryFormula)
            If TBL.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub New()
        Me.DTAInOUT.Connection.ConnectionString = Core_Module.GetConstring
    End Sub
End Class
