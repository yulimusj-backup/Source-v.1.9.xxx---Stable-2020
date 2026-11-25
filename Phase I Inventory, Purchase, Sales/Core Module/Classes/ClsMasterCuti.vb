Option Explicit On
Option Strict Off

Imports Core_Module
Public Class ClsMasterCuti
    Dim mPrimary As Double = -1
    Dim mCode As String = ""
    Dim mDesc As String = ""
    Dim mPeriod As Date = #1/1/1990#
    Dim mQty As Double = 0
    Dim mRemark As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now

    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Dim ObjItem As New ClsMasterCuti_Item
    Dim ObjCd As New ClsCekData

    Enum EnumSearchCuti
        Code
        Desc
        Remarks
        KodeKaryawan
        NamaKaryawan
    End Enum

    Enum enumCariData
        dept
        Job
        Other
        AllEmp
    End Enum

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            Me.mPrimary = value
        End Set
    End Property

    Public Property Code() As String
        Get
            Return Me.mCode
        End Get
        Set(ByVal value As String)
            Me.mCode = value
        End Set
    End Property

    Public Property Desc() As String
        Get
            Return Me.mDesc
        End Get
        Set(ByVal value As String)
            Me.mDesc = value
        End Set
    End Property

    Public Property Period() As Date
        Get
            Return Me.mPeriod
        End Get
        Set(ByVal value As Date)
            Me.mPeriod = value
        End Set
    End Property

    Public Property Qty() As Double
        Get
            Return Me.mQty
        End Get
        Set(ByVal value As Double)
            Me.mQty = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return Me.mRemark
        End Get
        Set(ByVal value As String)
            Me.mRemark = value
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

    ''@rosadinnn
    ''@20150908 1336

    Private mIsDefault As Integer
    Public Property IsDefault() As Integer
        Get
            Return mIsDefault
        End Get
        Set(ByVal value As Integer)
            mIsDefault = value
        End Set
    End Property


    Public Function ApakahBolehEdit(ByVal vPrimary As Double) As Boolean
        'Dim DTAWorkshift As New DseMasterCutiTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        'Dim DTALeaveApp As New DseMasterCutiTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        'Dim DtaAtt As New DseMasterWorkShiftTableAdapters.tbl_trans_att_viewTableAdapter
        'Dim DummyLong As Nullable(Of Long)
        Try
            '    DTAWorkshift.Connection.ConnectionString = Core_Module.GetConstring
            '    DTALeaveApp.Connection.ConnectionString = Core_Module.GetConstring
            '    DtaAtt.Connection.ConnectionString = Core_Module.GetConstring

            '    DummyLong = DTAWorkshift.ScalarQuery_GetJmlPrimaryByPrimaryCuti(vPrimary)
            '    If DummyLong.HasValue = True Then
            '        If DummyLong.Value > 0 Then
            '            MsgBox("This leave data can't edit because this leave data is refering to master workshift.", _
            '            MsgBoxStyle.Information, "Information")
            '            Return False
            '        End If
            '    Else
            '        Return False
            '    End If

            '    DummyLong = DTALeaveApp.ScalarQuery_GetJmlPrimaryByPrimaryCuti(vPrimary)
            '    If DummyLong.HasValue = True Then
            '        If DummyLong.Value > 0 Then
            '            MsgBox("This leave data can't edit because this leave data is refering to leave application.", _
            '            MsgBoxStyle.Information, "Information")
            '            Return False
            '        End If
            '    Else
            '        Return False
            '    End If

            '    DummyLong = DtaAtt.ScalarQuery_GetJmlByPrimaryCuti(vPrimary)
            '    If DummyLong.HasValue = True Then
            '        If DummyLong.Value > 0 Then
            '            MsgBox("This leave data can't edit because this leave data is refering to attendance.", _
            '            MsgBoxStyle.Information, "Information")
            '            Return False
            '        End If
            '    Else
            '        Return False
            '    End If

            '    'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erorr")
        Finally
            '    DTAWorkshift.Dispose()
            '    DTALeaveApp.Dispose()
            '    DtaAtt.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean
        Dim DTAWorkshift As New DseMasterCutiTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim DTALeaveApp As New DseMasterCutiTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim DtaAtt As New DseMasterWorkShiftTableAdapters.tbl_trans_att_viewTableAdapter
        Dim DummyLong As Nullable(Of Long)

        Dim TipeParameter As ClsMasterSalaryFormula_item.EnumTipeConstraitVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi
        Dim ObjSal As New ClsMasterSalaryFormula

        Try
            DTAWorkshift.Connection.ConnectionString = Core_Module.GetConstring
            DTALeaveApp.Connection.ConnectionString = Core_Module.GetConstring
            DtaAtt.Connection.ConnectionString = Core_Module.GetConstring

            DummyLong = DTAWorkshift.ScalarQuery_GetJmlPrimaryByPrimaryCuti(vPrimary)
            If DummyLong.HasValue = True Then
                If DummyLong.Value > 0 Then
                    MsgBox("This leave data can't delete because this leave data is refering to master workshift.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                Return False
            End If

            DummyLong = DTALeaveApp.ScalarQuery_GetJmlPrimaryByPrimaryCuti(vPrimary)
            If DummyLong.HasValue = True Then
                If DummyLong.Value > 0 Then
                    MsgBox("Can't delete this leave data because this leave data is refered by leave application.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                Return False
            End If

            DummyLong = DtaAtt.ScalarQuery_GetJmlByPrimaryCuti(vPrimary)
            If DummyLong.HasValue = True Then
                If DummyLong.Value > 0 Then
                    MsgBox("This leave data can't delate because this leave data is refering to attendance.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                Return False
            End If

            If ObjSal.GetJmlConstraint(133, 3, ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode, vPrimary.ToString) > 0 Then
                MsgBox("There salary formula using this data, You can't delete this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erorr")
        Finally
            DTAWorkshift.Dispose()
            DTALeaveApp.Dispose()
            DtaAtt.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim dta As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim dummy As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJmlPrimaryByCode(vCode)
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

    Public Function GetCodeByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetCodeByPrimary(vPrimary)
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

    Public Function GetPeriodByPrimary(ByVal vPrimary As Double) As Date
        Dim dta As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim Dummy As Nullable(Of Date)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetPeriodByPrimary(vPrimary)
            If Dummy.HasValue = True Then
                Return Dummy
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetPrimaryByCode(ByVal vCode As String) As Double
        Dim dta As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetPrimaryByCode(vCode)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTimeStampByPrimary(ByVal vPrimary As Double) As Date
        Dim dta As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetTimeStampByPrimary(vPrimary)
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

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmMasterCuti, _
    ByVal vDgv As DataGridView, ByVal vObjCd As Core_Module.ClsCekData, ByVal vTblItemDesc As DseMasterCuti.tblmaster_cuti_item_descDataTable) As Boolean
        Dim DTA As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim TBL As New DseMasterCuti.tblmaster_cutiDataTable
        Dim row As DseMasterCuti.tblmaster_cutiRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mCode = .code
                    Me.mDesc = .desc
                    Me.mPeriod = .period
                    Me.mQty = .leave_qty
                    Me.Remarks = .remarks
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp

                    ''@rosadinnn
                    ''@20150908 1341
                    Me.mIsDefault = .is_default

                    Me.ObjItem.IsiDGVItem(vDgv, vPrimary, vFrm, vObjCd, vTblItemDesc)
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

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vDGV As DataGridView, _
    ByVal frm As FrmMasterCuti, ByVal vTblDesc As DseMasterCuti.tblmaster_cuti_item_descDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterCuti_SimpanBaru, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.SimpanBaru1(vUserCode, vDGV, frm, vTblDesc)
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
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vDGV As DataGridView, ByVal frm As FrmMasterCuti, ByVal vTblDesc As DseMasterCuti.tblmaster_cuti_item_descDataTable) As Boolean
        Dim DTA As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mCode, Me.mDesc, vUserCode, Now, vUserCode, Now, Me.mTimestamp, _
            Me.mPeriod, Me.mRemark, Me.mQty, Me.mIsDefault) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimaryByCode(Me.mCode)
                Me.mTimestamp = DTA.ScalarQuery_GetTimestampByPrimary(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan item
                Me.ObjItem.SimpanItem(vDGV, Me.mPrimary, frm, vTblDesc)

                'Set Balance By Primary Cuti
                Me.SetBalance(Me.mPrimary, Me.mPeriod)

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

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vDGV As DataGridView, _
    ByVal frm As FrmMasterCuti, ByVal vTblDesc As DseMasterCuti.tblmaster_cuti_item_descDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterCuti_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.SimpanPerubahan1(vUserCode, vDGV, frm, vTblDesc)
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
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUser As String, ByVal vDGV As DataGridView, _
    ByVal frm As FrmMasterCuti, ByVal vTblDesc As DseMasterCuti.tblmaster_cuti_item_descDataTable) As Boolean
        Dim DTA As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.mCode, Me.mDesc, vUser, Now, Now, _
            Me.mPeriod, Me.mRemark, Me.mQty, Me.mIsDefault, Me.mPrimary, Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestampByPrimary(Me.Primary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUser

                'hapus item
                Me.ObjItem.HapusItem(Me.mPrimary)

                'simpan item
                Me.ObjItem.SimpanItem(vDGV, Me.mPrimary, frm, vTblDesc)

                'Set Balance By Primary Cuti
                Me.SetBalance(Me.mPrimary, Me.mPeriod)

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

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUser As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterCuti_Hapus, vUser)
            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.Hapus1(vPrimaryTransaksi)
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
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary) = True Then
                DTA.Delete(vPrimary)
                Me.ObjItem.HapusItem(vPrimary)
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

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterCuti.tblmaster_cutiDataTable, _
    ByVal SearchCriteria As EnumSearchCuti, ByVal tsl As ToolStripLabel, _
    ByVal TBL1 As DseMasterCuti.tblmaster_cuti1DataTable) As Boolean
        Dim DTA As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim DTA1 As New DseMasterCutiTableAdapters.tblmaster_cuti1TableAdapter
        Dim row As DseMasterCuti.tblmaster_cutiRow
        Dim row1 As DseMasterCuti.tblmaster_cuti1Row
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCuti.Code Then
                QueryResult = DTA.FillBy_LikeCode(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCuti.Desc Then
                QueryResult = DTA.FillBy_LikeDesc(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCuti.Remarks Then
                QueryResult = DTA.FillBy_LikeRemarks(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCuti.KodeKaryawan Then
                QueryResult = DTA1.FillBy_LikeCodeKaryawan(TBL1, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCuti.NamaKaryawan Then
                QueryResult = DTA1.FillBy_LikeName(TBL1, Criteria + "%")
            End If

            'update dgv
            DGV.Rows.Clear()
            DGV.Refresh()
            If SearchCriteria = EnumSearchCuti.Code Or SearchCriteria = EnumSearchCuti.Desc Or _
            SearchCriteria = EnumSearchCuti.Remarks Then
                For Each row In TBL.Rows
                    DGV.Rows.Add(row.primary, row.code, row.desc, Format(row.period, "yyyy"), row.remarks)
                Next
            ElseIf SearchCriteria = EnumSearchCuti.KodeKaryawan Or SearchCriteria = EnumSearchCuti.NamaKaryawan Then
                For Each row1 In TBL1.Rows
                    DGV.Rows.Add(row1.primary, row1.code, row1.desc, Format(row1.period, "yyyy"), row1.remarks)
                Next
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > 10000000 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Ticks / TimeSpan.TicksPerSecond) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, 0, MidpointRounding.AwayFromZero)) + " miliseconds"
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
    EnumSearchCuti) As Double
        Dim DTA As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim DTA1 As New DseMasterCutiTableAdapters.tblmaster_cuti1TableAdapter
        Dim Dummy As Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            If vSearchBy = EnumSearchCuti.Code Then
                Dummy = DTA.ScalarQuery_GetJmlPrimaryByLikeCode(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCuti.Desc Then
                Dummy = DTA.ScalarQuery_GetJmlPrimaryByLikeDesc(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCuti.Remarks Then
                Dummy = DTA.ScalarQuery_GetJmlPrimaryByLikeRemarks(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCuti.KodeKaryawan Then
                Dummy = DTA1.ScalarQuery_GetJmlByLikeKodeKaryawan(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCuti.NamaKaryawan Then
                Dummy = DTA1.ScalarQuery_GetJmlByLikeName(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    'tambahan - kirim parameter tbl item desc karena jika replace data tbl item desc nya harus di clear
    Public Sub CariDataKaryawan(ByVal vCariBerdasakan As enumCariData, ByVal vPrimary As Double, ByVal frm As FrmMasterCuti, _
    ByRef vTblItemDesc As DseMasterCuti.tblmaster_cuti_item_descDataTable)
        Dim dta As New DseMasterCutiTableAdapters.tblmaster_karyawanTableAdapter
        Dim tbl As New DseMasterCuti.tblmaster_karyawanDataTable
        Dim row As DseMasterCuti.tblmaster_karyawanRow
        Dim dtaItem As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim tblItem As New DseMasterCuti.tblmaster_cuti_itemDataTable
        Dim rowItem As DseMasterCuti.tblmaster_cuti_itemRow
        Dim rowdgv As DataGridViewRow
        Dim FrmDialog As New FrmAddReplace
        Dim masukan As Boolean = False
        Dim ObjKaryawan As New ClsMasterKaryawan
        'pada saat fill jika hanya ingin menampilkan karyawan aktif saja maka set aktif1 dan aktif2 dengan nilai 0
        'jika hanya ingin menampilkan karyawan tidak aktif saja maka set aktif1 dan aktif2 dengan nilai 1
        'jika ingin menampilkan semua karyawan maka set aktif1 = 1 dan aktif2 = 0
        Dim Aktif1 As Integer = 0 'jika 0 maka status karyawan aktif
        Dim Aktif2 As Integer = 0 'jika 0 maka status karyawan aktif

        Try
            ''pilih add/replace
            'ok = Add
            'cancel = replace
            FrmDialog.ShowDialog()
            If FrmDialog.DialogResult = DialogResult.OK Then 'jika add 
                masukan = False
            ElseIf FrmDialog.DialogResult = DialogResult.Yes Then 'jika replace
                frm.dgvEmployee.Rows.Clear()
                masukan = True
                vTblItemDesc.Rows.Clear()
            ElseIf FrmDialog.DialogResult = DialogResult.Cancel Then 'jika cancel
                Exit Sub
            End If

            frm.dgvEmployee.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            frm.dgvEmployee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            dta.Connection.ConnectionString = Core_Module.GetConstring
            dtaItem.Connection.ConnectionString = dta.Connection.ConnectionString

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            ''filter data
            If vCariBerdasakan = enumCariData.Job Then
                dta.FillBy_PrimaryJob(tbl, vPrimary, Aktif1, Aktif2)
            ElseIf vCariBerdasakan = enumCariData.dept Then
                dta.FillBy_PrimaryDept(tbl, vPrimary, Aktif1, Aktif2)
            ElseIf vCariBerdasakan = enumCariData.Other Then
                dtaItem.FillBy_PrimaryCuti(tblItem, vPrimary)
            ElseIf vCariBerdasakan = enumCariData.AllEmp Then
                dta.Fill(tbl, Aktif1, Aktif2)
            End If

            ''pindahkan ke dgv
            If vCariBerdasakan = enumCariData.dept Or vCariBerdasakan = enumCariData.Job Or vCariBerdasakan = enumCariData.AllEmp Then
                'berdasarkan departemen atau job posisi
                For Each row In tbl.Rows
                    'cek data apakah sudah ada di dgv
                    'jika replace lewati pengecekan
                    If FrmDialog.DialogResult = DialogResult.OK Then
                        For Each rowdgv In frm.dgvEmployee.Rows
                            If rowdgv.Cells(frm.ColPrimary.Index).Value <> row.Primary Then
                                masukan = True
                            Else
                                masukan = False
                                Exit For
                            End If
                        Next
                    End If
                    If masukan = True Then
                        frm.dgvEmployee.Rows.Add()
                        If frm.dgvEmployee.Rows.Count > 1 Then
                            rowdgv = frm.dgvEmployee.Rows(frm.dgvEmployee.RowCount - 2)
                        Else
                            rowdgv = frm.dgvEmployee.Rows(frm.dgvEmployee.RowCount)
                        End If
                        rowdgv.Cells(frm.ColCode.Index).Value = row.Code
                        rowdgv.Cells(frm.ColName.Index).Value = row.Full_Name
                        rowdgv.Cells(frm.ColPrimary.Index).Value = row.Primary
                    End If
                Next
            ElseIf vCariBerdasakan = enumCariData.Other Then
                'berdasarkan data cuti yang sudah ada
                For Each rowItem In tblItem.Rows
                    'cek data apakah sudah ada di dgv
                    'jika replace lewati pengecekan
                    If FrmDialog.DialogResult = DialogResult.OK Then
                        For Each rowdgv In frm.dgvEmployee.Rows
                            If rowdgv.Cells(frm.ColPrimary.Index).Value <> rowItem.primary_karyawan Then
                                masukan = True
                            Else
                                masukan = False
                                Exit For
                            End If
                        Next
                    End If
                    If masukan = True Then
                        frm.dgvEmployee.Rows.Add()
                        If frm.dgvEmployee.Rows.Count > 1 Then
                            rowdgv = frm.dgvEmployee.Rows(frm.dgvEmployee.RowCount - 2)
                        Else
                            rowdgv = frm.dgvEmployee.Rows(frm.dgvEmployee.RowCount)
                        End If
                        rowdgv.Cells(frm.ColCode.Index).Value = ObjKaryawan.GetCodeByPrimary(rowItem.primary_karyawan)
                        rowdgv.Cells(frm.ColName.Index).Value = ObjKaryawan.GetFullNameByPrimary(rowItem.primary_karyawan)
                        rowdgv.Cells(frm.ColRemarks.Index).Value = rowItem.remarks
                        rowdgv.Cells(frm.ColPrimary.Index).Value = rowItem.primary_karyawan
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            tbl.Dispose()
            dtaItem.Dispose()
            tblItem.Dispose()
            FrmDialog.Dispose()
            frm.dgvEmployee.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            frm.dgvEmployee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    'Dipanggil di Leave Approval dan Leave Application
    Public Sub SetBalance(ByVal vPrimaryKaryawan As Double, ByVal vPrimaryCuti As Double, ByVal vPeriod As Date)
        Dim DtaBeginingQty As New DseMasterCuti_SetBalanceTableAdapters.tblmaster_cuti_beginningTableAdapter
        Dim TblBeginingQty As New DseMasterCuti_SetBalance.tblmaster_cuti_beginningDataTable
        Dim RowBeginingQty As DseMasterCuti_SetBalance.tblmaster_cuti_beginningRow
        Dim DtaMaxDate As New DseMasterCuti_SetBalanceTableAdapters.tblMAX_DateTableAdapter
        Dim TblMaxDate As New DseMasterCuti_SetBalance.tblMAX_DateDataTable
        Dim RowMaxdate As DseMasterCuti_SetBalance.tblMAX_DateRow
        Dim DtaCutiPengajuan As New DseMasterCuti_SetBalanceTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim TblCutiPengajuan As New DseMasterCuti_SetBalance.tbltrans_cuti_pengajuanDataTable
        Dim RowCutiPengajuan As DseMasterCuti_SetBalance.tbltrans_cuti_pengajuanRow
        Dim DtaAttend As New DseMasterCuti_SetBalanceTableAdapters.TblTrans_AttendanceTableAdapter
        Dim TblAttend As New DseMasterCuti_SetBalance.TblTrans_AttendanceDataTable
        Dim rowAttend As DseMasterCuti_SetBalance.TblTrans_AttendanceRow

        Try
            DtaBeginingQty.Connection.ConnectionString = Core_Module.GetConstring
            DtaAttend.Connection.ConnectionString = DtaBeginingQty.Connection.ConnectionString
            DtaCutiPengajuan.Connection.ConnectionString = DtaBeginingQty.Connection.ConnectionString
            DtaMaxDate.Connection.ConnectionString = DtaBeginingQty.Connection.ConnectionString

            'load data begining qty cuti
            DtaBeginingQty.FillBy_PrimaryKaryawan(TblBeginingQty, vPrimaryCuti, vPrimaryKaryawan)
            'load data untuk menentukan kapan attendance terakhir
            DtaMaxDate.FillBy_PrimaryKaryawan(TblMaxDate, ObjCd.GetLastDateOfTheYear(vPeriod), vPrimaryKaryawan)
            'load data cuti pengajuan
            DtaCutiPengajuan.FillBy_PrimaryKaryawan(TblCutiPengajuan, ObjCd.GetLastDateOfTheYear(vPeriod), ObjCd.GetFisrtDateOfTheYear(vPeriod), vPrimaryCuti, vPrimaryKaryawan)
            'load data attendance
            DtaAttend.FillBy_PrimaryKaryawan(TblAttend, ObjCd.GetLastDateOfTheYear(vPeriod), ObjCd.GetFisrtDateOfTheYear(vPeriod), vPrimaryCuti, vPrimaryKaryawan)

            If TblMaxDate.Rows.Count = 0 Then
                RowMaxdate = TblMaxDate.NewtblMAX_DateRow
                RowMaxdate.primary_karyawan = vPrimaryKaryawan
                RowMaxdate.MAX_Date = #1/1/1950#
                TblMaxDate.Rows.Add(RowMaxdate)
            End If
            RowMaxdate = TblMaxDate.Rows(0)

            For Each RowCutiPengajuan In TblCutiPengajuan.Rows
                RowCutiPengajuan.TglMAX = RowMaxdate.MAX_Date
            Next

            For Each RowBeginingQty In TblBeginingQty.Rows
                'dikurangi dengan item pengajuan cuti yang masih terbuka atau sudah di approve, yang belum ada attendance nya
                For Each RowCutiPengajuan In TblCutiPengajuan.Rows
                    If RowCutiPengajuan._date > RowCutiPengajuan.TglMAX Then
                        RowBeginingQty.Cuti_belum_terjadi += RowCutiPengajuan.duration
                    End If
                Next

                'dikurangi cuti pada item attendance yang terjadi
                For Each rowAttend In TblAttend.Rows
                    RowBeginingQty.Cuti_yang_terjadi += rowAttend.is_cuti
                Next


                'Balance = Qty (berdasarkan primary) - Qty Deduction + Qty Addition - Cuti yang Terjadi - Cuti pengajuan yg blm terjadi(status open/approved)
                RowBeginingQty.Balance = RowBeginingQty.Begining_Qty - RowBeginingQty.Cuti_belum_terjadi - RowBeginingQty.Cuti_yang_terjadi

                'update balance
                DtaBeginingQty.UpdateQuery_SetBalance(RowBeginingQty.Balance, RowBeginingQty.primary_leave, RowBeginingQty.primary_karyawan)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaBeginingQty.Dispose()
            TblBeginingQty.Dispose()
            DtaMaxDate.Dispose()
            TblMaxDate.Dispose()
            DtaCutiPengajuan.Dispose()
            TblCutiPengajuan.Dispose()
            DtaAttend.Dispose()
            TblAttend.Dispose()
        End Try
    End Sub

    Private Function GetDuration(ByVal vPrimaryWS As Double, ByVal vDay_of_week As Integer, ByVal vTBL As DseMasterCuti_SetBalance.tblmaster_workshift_itemDataTable) As Decimal
        Dim row As DseMasterCuti_SetBalance.tblmaster_workshift_itemRow

        Try
            vTBL.DefaultView.RowFilter = "[Primary Workshift] = " + vPrimaryWS.ToString + " AND [day of week] = " + vDay_of_week.ToString
            If vTBL.DefaultView.Count > 0 Then
                row = vTBL.DefaultView(0).Row
                If row.duratioin = ClsMasterWorkShift.enumDuration.fullday Then
                    Return 1
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    'Dipanggil oleh Transaksi Attendance
    Public Sub SetBalance(ByVal vPrimaryKaryawan As Double, ByVal vListPrimarycuti As List(Of Double), ByVal vPeriod As Date)
        Dim PrimaryCuti As Double
        Try
            'Balance = Qty (berdasarkan primary) - Qty Deduction + Qty Addition - Cuti yang Terjadi - Cuti pengajuan yg blm terjadi(status open/approved) 
            For Each PrimaryCuti In vListPrimarycuti
                Me.SetBalance(vPrimaryKaryawan, PrimaryCuti, vPeriod)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'mengembalikan primary cuti yang secara default ada cuti tahunan untuk periode yang bersangkutan
    Public Function GetPrimaryDefaultYearlyLeave(ByVal vRefDate As Date) As Double
        Dim DTA As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim TBL As New DseMasterCuti.tblmaster_cutiDataTable
        Dim row As DseMasterCuti.tblmaster_cutiRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim StartDate, EndDate As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartDate = ObjCD.GetFisrtDateOfTheYear(vRefDate)
            EndDate = ObjCD.GetLastDateOfTheYear(vRefDate)
            DTA.FillBy_DefaultCutiTahunan(TBL, EndDate, StartDate)
            If TBL.Rows.Count = 0 Then
                Return -1
            Else
                row = TBL.Rows(0)
                Return row.primary
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetBalance(ByVal vPrimaryKary As Double, ByVal vPrimaryCuti As Double, ByVal vRefDate As Date) As Decimal
        Dim DTA As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim Dummy As System.Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNilaiBalance(vPrimaryCuti, vPrimaryKary)
            If Dummy.HasValue Then
                Return Dummy.Value
                'Balance = Qty (berdasarkan primary) - Qty Deduction + Qty Addition - Cuti yang Terjadi - Cuti pengajuan yg blm terjadi(status open/approved) 
                'berdasarkan rumus diatas, maka balance cuti untuk tanggal tertentu adalah:
                'balance akhir + cuti yang terjadi setelah ref date - 
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetBalance(ByVal vPrimaryKary As Double, ByVal vPrimaryCuti As Double) As Decimal
        Dim DTA As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim Dummy As System.Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNilaiBalance(vPrimaryCuti, vPrimaryKary)
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

    Public Function GetBalanceByCutOff(ByVal vPrimaryKary As Double, ByVal vPrimaryCuti As Double, ByVal vEndDateCutOff As Date, _
    Optional ByRef GetTotalBeginning As Decimal = 0, Optional ByRef GetTotalUsed As Decimal = 0) As Decimal
        Dim DTA As New DseMasterCutiTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim DummyBeginning As System.Nullable(Of Double)
        Dim DummyUsed As System.Nullable(Of Double)
        Dim Balance As Decimal = 0
        Dim StartDate As Date = Me.ObjCd.GetFisrtDateOfTheYear(vEndDateCutOff)
        Dim LeaveQty As Decimal = Me.GetQtyLeave(vPrimaryCuti)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'filter pada dse PrimaryCuti (=), PrimaryKary (=), EndDate (<=), StartDate (>=)
            DummyBeginning = DTA.ScalarQuery_GetTotalBeginningByCutOff(vPrimaryCuti, vPrimaryKary, vEndDateCutOff)
            If DummyBeginning.HasValue Then
                GetTotalBeginning = DummyBeginning.Value
            End If
            GetTotalBeginning += LeaveQty
            DummyUsed = DTA.ScalarQuery_GetTotalUsedByCutOff(vEndDateCutOff, StartDate, vPrimaryCuti, vPrimaryKary)
            If DummyUsed.HasValue Then
                GetTotalUsed = DummyUsed.Value
            End If
            Balance = GetTotalBeginning - GetTotalUsed

            Return Balance
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Dipanggil oleh master cuti
    'Dibawah simpan baru dengan simpan perubahan
    'Tambahkan Batasan untuk Leave Qty, Addition, dan Deduction tidak boleh menyebabkan balance minus
    'End Edit 
    Public Sub SetBalance(ByVal vPrimaryCuti As Double, ByVal vPeriod As Date)
        Dim DtaBeginingQty As New DseMasterCuti_SetBalanceTableAdapters.tblmaster_cuti_beginningTableAdapter
        Dim TblBeginingQty As New DseMasterCuti_SetBalance.tblmaster_cuti_beginningDataTable
        Dim RowBeginingQty As DseMasterCuti_SetBalance.tblmaster_cuti_beginningRow
        Dim DtaMaxDate As New DseMasterCuti_SetBalanceTableAdapters.tblMAX_DateTableAdapter
        Dim TblMaxDate As New DseMasterCuti_SetBalance.tblMAX_DateDataTable
        Dim RowMaxdate As DseMasterCuti_SetBalance.tblMAX_DateRow
        Dim DtaCutiPengajuan As New DseMasterCuti_SetBalanceTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim TblCutiPengajuan As New DseMasterCuti_SetBalance.tbltrans_cuti_pengajuanDataTable
        Dim RowCutiPengajuan As DseMasterCuti_SetBalance.tbltrans_cuti_pengajuanRow
        Dim DtaAttend As New DseMasterCuti_SetBalanceTableAdapters.TblTrans_AttendanceTableAdapter
        Dim TblAttend As New DseMasterCuti_SetBalance.TblTrans_AttendanceDataTable
        Dim rowAttend As DseMasterCuti_SetBalance.TblTrans_AttendanceRow
        Dim ObjCd As New ClsCekData
        Dim filter As String = ""

        Try
            DtaBeginingQty.Connection.ConnectionString = Core_Module.GetConstring
            DtaAttend.Connection.ConnectionString = DtaBeginingQty.Connection.ConnectionString
            DtaCutiPengajuan.Connection.ConnectionString = DtaBeginingQty.Connection.ConnectionString
            DtaMaxDate.Connection.ConnectionString = DtaBeginingQty.Connection.ConnectionString

            DtaBeginingQty.Fill(TblBeginingQty, vPrimaryCuti)
            DtaMaxDate.Fill(TblMaxDate, ObjCd.GetLastDateOfTheYear(vPeriod))
            DtaCutiPengajuan.Fill(TblCutiPengajuan, ObjCd.GetLastDateOfTheYear(vPeriod), ObjCd.GetFisrtDateOfTheYear(vPeriod), vPrimaryCuti)
            DtaAttend.Fill(TblAttend, ObjCd.GetLastDateOfTheYear(vPeriod), ObjCd.GetFisrtDateOfTheYear(vPeriod), vPrimaryCuti)

            For Each RowCutiPengajuan In TblCutiPengajuan.Rows
                filter = "[primary karyawan] = " + RowCutiPengajuan.primary_karyawan.ToString
                TblMaxDate.DefaultView.RowFilter = filter
                If TblMaxDate.DefaultView.Count > 0 Then
                    RowMaxdate = TblMaxDate.DefaultView(0).Row
                    RowCutiPengajuan.TglMAX = RowMaxdate.MAX_Date
                Else
                    RowCutiPengajuan.TglMAX = #1/1/1950#
                End If
            Next


            For Each RowBeginingQty In TblBeginingQty.Rows
                'filter data cuti pengajuan berdasarkan primary karyawan dan cuti
                filter = "[primary karyawan] = " + RowBeginingQty.primary_karyawan.ToString + _
                "AND [primary cuti] = " + RowBeginingQty.primary_leave.ToString
                TblCutiPengajuan.DefaultView.RowFilter = filter
                If TblCutiPengajuan.DefaultView.Count > 0 Then
                    For i As Integer = 1 To TblCutiPengajuan.DefaultView.Count
                        RowCutiPengajuan = TblCutiPengajuan.DefaultView(i - 1).Row
                        If RowCutiPengajuan._date > RowCutiPengajuan.TglMAX Then
                            RowBeginingQty.Cuti_belum_terjadi += RowCutiPengajuan.duration
                        End If
                    Next
                End If

                'filter data attend berdasarkan primary karyawan dan cuti
                filter = "[primary karyawan] = " + RowBeginingQty.primary_karyawan.ToString + _
                "AND [primary cuti] = " + RowBeginingQty.primary_leave.ToString
                TblAttend.DefaultView.RowFilter = filter
                If TblAttend.DefaultView.Count > 0 Then
                    For i As Integer = 1 To TblAttend.DefaultView.Count
                        rowAttend = TblAttend.DefaultView(i - 1).Row
                        RowBeginingQty.Cuti_yang_terjadi += rowAttend.is_cuti
                    Next
                End If

                RowBeginingQty.Balance = RowBeginingQty.Begining_Qty - RowBeginingQty.Cuti_belum_terjadi - RowBeginingQty.Cuti_yang_terjadi
                'Balance = Qty (berdasarkan primary) - Qty Deduction + Qty Addition - Cuti yang Terjadi - Cuti pengajuan yg blm terjadi(status open/approved) 

                'update balance
                DtaBeginingQty.UpdateQuery_SetBalance(RowBeginingQty.Balance, RowBeginingQty.primary_leave, RowBeginingQty.primary_karyawan)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaBeginingQty.Dispose()
            TblBeginingQty.Dispose()
            DtaCutiPengajuan.Dispose()
            TblCutiPengajuan.Dispose()
            DtaAttend.Dispose()
            TblAttend.Dispose()
            DtaMaxDate.Dispose()
            TblMaxDate.Dispose()
            ObjCd = Nothing
        End Try
    End Sub

    Enum EnumAddDed
        add
        ded
    End Enum

    Public Function GetQtyLeave(ByVal vPrimaryLeave As Double) As Decimal
        Dim DTAITem As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim Dummy As New Nullable(Of Decimal)
        Try
            DTAITem.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTAITem.ScalarQuery_GetLeaveQty(vPrimaryLeave)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAITem.Dispose()
        End Try
    End Function


    Public Function ApakahEditSah(ByVal vPrimaryCuti As Double, ByVal vPrimaryKaryawan As Double, _
    ByVal vQtyAddDgv As Double, ByVal vQtyDedDgv As Double, ByVal vQtyLeaveTxt As Double, _
    ByVal vEnumAddDed As EnumAddDed) As Boolean
        Dim dta As New DseMasterCutiTableAdapters.DataTable1TableAdapter
        Dim tbl As New DseMasterCuti.DataTable1DataTable
        Dim row As DseMasterCuti.DataTable1Row
        Dim QtyBalanaceDB As Double = 0
        Dim QtyAddDB As Double = 0
        Dim QtyDedDB As Double = 0
        Dim QtyLeaveDB As Double = 0
        Dim BalanceBaru As Double = 0
        Dim filter As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.FillBy_Primary(tbl, vPrimaryCuti)

            'dapetin nilai yang ada di database

            filter = "[primary karyawan] = " + vPrimaryKaryawan.ToString
            tbl.DefaultView.RowFilter = filter
            'harus nya cuma dapat 1 baris
            If tbl.DefaultView.Count = 1 Then
                row = tbl.DefaultView(0).Row
                With row
                    .Qty_Add_DGV = vQtyAddDgv
                    .Qty_Ded_DGV = vQtyDedDgv
                    .Qty_Balance_New = .balance + (.Qty_Add_DGV - .leave_addition) - (.Qty_Ded_DGV - .leave_deduction) + (vQtyLeaveTxt - .leave_qty)
                    If .Qty_Balance_New < 0 Then
                        If vEnumAddDed = EnumAddDed.add Then
                            MsgBox("Minimum addition qty is " + Me.ObjCd.DigitGrouping((vQtyAddDgv - .Qty_Balance_New), 1).ToString, MsgBoxStyle.Information, "Information")
                        Else
                            MsgBox("Maximum deduction qty is " + Me.ObjCd.DigitGrouping((vQtyDedDgv + .Qty_Balance_New), 1).ToString, MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End With
            ElseIf tbl.DefaultView.Count = 0 Then
                If CDbl(vQtyLeaveTxt) + CDbl(vQtyAddDgv) - CDbl(vQtyDedDgv) < 0 Then
                    If vEnumAddDed = EnumAddDed.add Then
                        MsgBox("Minimum addition qty is " + Me.ObjCd.DigitGrouping(Math.Abs(CDbl(vQtyLeaveTxt) - CDbl(vQtyDedDgv)), 1).ToString, MsgBoxStyle.Information, "Information")
                    Else
                        MsgBox("Maximum deduction qty is " + Me.ObjCd.DigitGrouping((CDbl(vQtyLeaveTxt) + CDbl(vQtyAddDgv)), 1).ToString, MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            ElseIf tbl.DefaultView.Count > 1 Then
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function ApakahLeaveEditSah(ByVal vPrimaryCuti As Double, ByVal vfrm As FrmMasterCuti, ByVal vDgv As DataGridView) As Boolean
        Dim dta As New DseMasterCutiTableAdapters.DataTable1TableAdapter
        Dim tbl As New DseMasterCuti.DataTable1DataTable
        Dim row As DseMasterCuti.DataTable1Row
        Dim QtyBalanaceDB As Double = 0
        Dim QtyAddDB As Double = 0
        Dim QtyDedDB As Double = 0
        Dim QtyLeaveDB As Double = 0
        Dim BalanceBaru As Double = 0
        Dim rowdgv As DataGridViewRow
        Dim filter As String
        Dim ListMinus As New List(Of String)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.FillBy_Primary(tbl, vPrimaryCuti)

            ListMinus.Clear()
            For Each rowdgv In vDgv.Rows
                If rowdgv.Cells(vfrm.ColPrimary.Index).Value <> -1 Then
                    filter = "[primary karyawan] = " + rowdgv.Cells(vfrm.ColPrimary.Index).Value.ToString
                    tbl.DefaultView.RowFilter = filter
                    'harus nya cuma dapat 1 baris
                    If tbl.DefaultView.Count = 1 Then
                        row = tbl.DefaultView(0).Row
                        With row
                            .Qty_Add_DGV = CDbl(rowdgv.Cells(vfrm.ColLeaveAddition.Index).Value)
                            .Qty_Ded_DGV = CDbl(rowdgv.Cells(vfrm.ColLeaveDeduction.Index).Value)
                            .Qty_Balance_New = .balance + (.Qty_Add_DGV - .leave_addition) - (.Qty_Ded_DGV - .leave_deduction) + _
                            (CDbl(vfrm.TxtLeaveQty.Text) - .leave_qty)
                            If .Qty_Balance_New < 0 Then
                                ListMinus.Add(.KodeKaryawan)
                            End If
                        End With
                    ElseIf tbl.DefaultView.Count = 0 Then
                        If CDbl(vfrm.TxtLeaveQty.Text) + CDbl(rowdgv.Cells(vfrm.ColLeaveAddition.Index).Value) - CDbl(rowdgv.Cells(vfrm.ColLeaveDeduction.Index).Value) < 0 Then
                            ListMinus.Add(rowdgv.Cells(vfrm.ColCode.Index).Value + " - " + rowdgv.Cells(vfrm.ColName.Index).Value)
                        End If
                    ElseIf tbl.DefaultView.Count > 1 Then
                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            'jika tidk ada yang minus 
            If ListMinus.Count > 0 Then
                MsgBox("Cannot edit leave qty data, the following employee leave balance qty is negative: " + vbNewLine + _
                Me.GetListBalanceMinus(ListMinus), MsgBoxStyle.Information, "Information")
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetListBalanceMinus(ByVal vList As List(Of String)) As String
        Dim RetVal As String = ""
        Dim i As Integer = 0
        Try
            For i = 1 To vList.Count
                RetVal = RetVal + (i).ToString + ". " + vList(i - 1).ToString + "." + vbNewLine
                'i += 1
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    ''@rosadinnn
    ''@20150908 1348
    Public Function IsAlreadyDefaultLeave(ByVal vCurrentPrimary As Double, ByVal vPeriod As Date) As Boolean
        Dim Dta As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim count_ As Nullable(Of Long)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            count_ = Dta.ScalarQuery_GetDefaultLeaveCount(vCurrentPrimary, vPeriod)
            If count_.HasValue = True Then
                If count_.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function IsAlreadyDefaultLeave(ByVal vPeriod As Date) As Boolean
        Dim Dta As New DseMasterCutiTableAdapters.tblmaster_cutiTableAdapter
        Dim count_ As Nullable(Of Long)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            count_ = Dta.ScalarQuery_GetDefaultLeaveCount2(vPeriod)
            If count_.HasValue = True Then
                If count_.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
        End Try
    End Function
End Class