Option Explicit On
Option Strict Off

Public Class ClsTrans_SchDownTime

    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianScheduling
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Dim objItem As New ClsTrans_SchDownTime_Item
    Dim objCd As New Core_Module.ClsCekData
    Dim objNumbering As New Core_Module.ClsMasterTransactionNumbering

    Public Enum EnumSearchBy As Integer
        DOCUMENT_NO
        DATE_
        REASON
        REMARKS
        WORK_UNIT_CODE
        WORK_UNIT_DESC
    End Enum

    Private mPrimary As Double
    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Private mDocumentNo As String
    Public Property DocumentNo() As String
        Get
            Return mDocumentNo
        End Get
        Set(ByVal value As String)
            mDocumentNo = value
        End Set
    End Property

    Private mUrutan As Double
    Public Property Urutan() As Double
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Double)
            mUrutan = value
        End Set
    End Property

    Private mDate As Date
    Public Property Date_() As Date
        Get
            Return mDate
        End Get
        Set(ByVal value As Date)
            mDate = value
        End Set
    End Property

    Private mReason As String
    Public Property Reason() As String
        Get
            Return mReason
        End Get
        Set(ByVal value As String)
            mReason = value
        End Set
    End Property

    Private mRemarks As String
    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Private mCreatedBy As String
    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Private mCreatedDate As Date
    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Private mModifiedBy As String
    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Private mModifiedDate As Date
    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Private mTimestamp As Date
    Public Property Timestamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property

    Public Function ApakahDataBenar(ByVal vPrimary As Double, ByVal vDgv As DataGridView, ByVal f As FrmTrans_SchDownTime, ByVal vBaru As Boolean) As Boolean
        Try
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mDocumentNo) = True Then
                    If Me.objNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime) = True Then
                        Me.mDocumentNo = Me.GetNoBaru(Me.mDate)
                    Else
                        MsgBox("Document No: " & Me.mDocumentNo & " already exist.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                Dim oldDocNo As String = Me.GetdocumentNo(mPrimary)
                If oldDocNo <> Me.mDocumentNo Then
                    If Me.ApakahSudahAda(Me.mDocumentNo) = True Then
                        If Me.objNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime) = True Then
                            Me.mDocumentNo = Me.GetNoBaru(Me.mDate)
                        Else
                            MsgBox("Document No: " & Me.mDocumentNo & " already exist.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If

                Dim oldTimeStamp As Date = Me.GetTimestamp(Me.mPrimary)
                If mTimestamp <> oldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.CariDetail(Me.mPrimary, vDgv, f)
                    End If
                    Return False
                End If
            End If
            If Me.ApakahDataItemBenar(vPrimary, vDgv, f) = False Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataItemBenar(ByVal vPrimary As Double, ByVal vDgv As DataGridView, ByVal f As FrmTrans_SchDownTime) As Boolean
        Dim listConflictTime As String = ""
        Try
            vDgv.AllowUserToAddRows = False
            For Each r As DataGridViewRow In vDgv.Rows
                If r.Cells(f.colWU_Primary.Index).Value <> "-1" And String.IsNullOrEmpty(r.Cells(f.colWU_Code.Index).Value) = False Then
                    listConflictTime += Me.getListConflictSchedule(r.Index + 1, vPrimary, r.Cells(f.colWU_Primary.Index).Value, r.Cells(f.colWU_Code.Index).Value, _
                            CDate(r.Cells(f.colWU_StartDate.Index).Value & " " & r.Cells(f.colWU_StartTime.Index).Value.ToString.Replace(".", ":") & ":00"), _
                            CDate(r.Cells(f.colWU_EndDate.Index).Value & " " & r.Cells(f.colWU_EndTime.Index).Value.ToString.Replace(".", ":") & ":00"))
                End If
            Next
            vDgv.AllowUserToAddRows = True
            If listConflictTime <> "" Then
                MsgBox("Conflict schedule: " & listConflictTime, MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function getListConflictSchedule(ByVal vRowDgv As Integer, ByVal vPrimaryTransExclude As Double, ByVal vPrimaryWorkUnit As Double, ByVal vWorkUnitCode As String, ByVal vStartTime As Date, ByVal vEndTime As Date) As String
        Dim d As New DseTrans_SchDownTimeTableAdapters.DataTable2TableAdapter
        Dim t As New DseTrans_SchDownTime.DataTable2DataTable
        Dim r As DseTrans_SchDownTime.DataTable2Row
        Dim val As String = ""
        Dim otherStartTime, otherEndTime As Date
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            d.Fill(t, vPrimaryTransExclude, vPrimaryWorkUnit)
            For Each r In t
                otherStartTime = r.start_time
                otherEndTime = r.end_time
                If (vStartTime < otherStartTime And vEndTime >= otherStartTime) Or _
                        (vStartTime = otherStartTime) Or _
                        (vStartTime > otherStartTime And vStartTime <= otherEndTime) Then
                    val += vbNewLine & "Work Unit " & vWorkUnitCode & ", row " & vRowDgv.ToString & " with item in transaction no : " & r.No_Transaksi & " " & Format(r.start_time, "dd MMM yyyy HH.mm") & " to " & Format(r.end_time, "dd MMM yyyy HH.mm")
                End If
            Next
            Return val
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vNoTrans As String) As Boolean
        Dim d As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtimeTableAdapter
        Dim n As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            n = d.ScalarQuery_GetCountByNoTransaksi(vNoTrans)
            If n.HasValue Then
                If n.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, Optional ByVal vShowMsg As Boolean = True) As Boolean
        Dim createdBy As String = ""
        Try
            createdBy = Me.GetCreatedBy(vPrimary)
            If createdBy <> vUserCode Then
                If vShowMsg = True Then
                    MsgBox("Only user " & createdBy & " can edit this data.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, Optional ByVal vShowMsg As Boolean = True) As Boolean
        Dim createdBy As String = ""
        Try
            createdBy = Me.GetCreatedBy(vPrimary)
            If createdBy <> vUserCode Then
                If vShowMsg = True Then
                    MsgBox("Only user " & createdBy & " can delete this data.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim dta As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtimeTableAdapter
        Dim s As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            s = dta.ScalarQuery_GetCreatedBy(vPrimary)
            If String.IsNullOrEmpty(s) = False Then
                Return s
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vDgv As DataGridView, ByVal vFrm As FrmTrans_SchDownTime) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
           Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_ScheduleDownTime_SimpanBaru, vUserCode)

            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()

            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, vDgv, vFrm)
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vDgv As DataGridView, ByVal vFrm As FrmTrans_SchDownTime) As Boolean
        Dim DTA As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtimeTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(Me.mPrimary, vDgv, vFrm, True) = True Then
                Me.Urutan = Me.objNumbering.GetUrutan(Me.mDocumentNo, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime, Me.mDate)
                If DTA.InsertQuery(Me.mDocumentNo, Me.mUrutan, Me.mReason, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now, Me.mDate) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mDocumentNo).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    If objItem.SimpanItem(Me.Primary, vDgv, vFrm) = False Then
                        Return False
                    End If
                    Return True
                Else
                    Return False
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

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vDgv As DataGridView, ByVal vFrm As FrmTrans_SchDownTime) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
             Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_ScheduleDownTime_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode, vDgv, vFrm)

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

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal vDgv As DataGridView, ByVal vFrm As FrmTrans_SchDownTime) As Boolean
        Dim DTA As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtimeTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(Me.mPrimary, vDgv, vFrm, False) = True Then
                Me.Urutan = Me.objNumbering.GetUrutan(Me.mDocumentNo, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime, Me.mDate)
                If DTA.UpdateQuery(Me.mDocumentNo, Me.mUrutan, Me.mReason, Me.mRemarks, vUserCode, Now, Now, Me.mDate, Me.mPrimary) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    objItem.HapusItem(Me.Primary)
                    If objItem.SimpanItem(Me.Primary, vDgv, vFrm) = False Then
                        Return False
                    End If
                    Return True
                Else
                    Return False
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

    Public Function Hapus(ByVal vPrimary As Double, ByVal vUser As String, ByVal vDgv As DataGridView, _
            ByVal i As FrmTrans_SchDownTime, Optional ByVal vShowMessage As Boolean = True) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, vPrimary, _
               Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_ScheduleDownTime_Hapus, vUser)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimary, vUser, vDgv, i, vShowMessage)
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

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUser As String, ByVal vDgv As DataGridView, _
            ByVal i As FrmTrans_SchDownTime, Optional ByVal vShowMessage As Boolean = True) As Boolean
        Dim DTA As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtimeTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUser, vShowMessage) = True Then
                DTA.DeleteQuery(vPrimary)
                Return True
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

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vDgv As DataGridView, ByVal vFrm As FrmTrans_SchDownTime) As Boolean
        Dim DTA As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtimeTableAdapter
        Dim TBL As New DseTrans_SchDownTime.tbltrans_sch_downtimeDataTable
        Dim row As DseTrans_SchDownTime.tbltrans_sch_downtimeRow
        Try
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mUrutan = .urutan
                    Me.mDocumentNo = .No_Transaksi
                    Me.mDate = ._Date
                    Me.mReason = .Reason
                    Me.mRemarks = .remarks
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
                objItem.IsiDgvItem(Me.mPrimary, vDgv, vFrm)
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vDocNo As String) As Double
        Dim DTA As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtimeTableAdapter
        Dim Dummy_ As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy_ = DTA.ScalarQuery_GetPrimary(vDocNo)
            If Dummy_.HasValue = True Then
                Return Dummy_.Value
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

    Public Function GetdocumentNo(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtimeTableAdapter
        Dim Dummy_ As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy_ = DTA.ScalarQuery_GetDocNo(vPrimary)
            If String.IsNullOrEmpty(Dummy_) = False Then
                Return Dummy_
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
        Dim DTA As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtimeTableAdapter
        Dim Dummy_ As New Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy_ = DTA.ScalarQuery_GetTimeStamp(vPrimary)
            If Dummy_.HasValue = True Then
                Return Dummy_.Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Date.MinValue
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, ByVal DGV As DataGridView, ByVal TBL As DseTrans_SchDownTime.tbltrans_sch_downtime1DataTable, _
        ByVal SearchCriteria As ClsTrans_SchDownTime.EnumSearchBy, ByVal tsl As ToolStripLabel, ByVal vDari As Date, ByVal vSampai As Date) As Boolean
        Dim DTA As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtime1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        'Dim row As DseTrans_SchDownTime.tbltrans_sch_downtime1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchBy.DOCUMENT_NO Then
                QueryResult = DTA.FillBy_DocNoLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchBy.DATE_ Then
                QueryResult = DTA.FillBy_Date(TBL, vDari, vSampai)
            ElseIf SearchCriteria = EnumSearchBy.REASON Then
                QueryResult = DTA.FillBy_ReasonLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchBy.REMARKS Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchBy.WORK_UNIT_CODE Then
                QueryResult = DTA.FillBy_WorkUnitCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchBy.WORK_UNIT_DESC Then
                QueryResult = DTA.FillBy_WorkUnitDescLIKE(TBL, Criteria + "%")
            End If 

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal vDocNo As String) As Integer
        Dim DTA As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtimeTableAdapter
        Dim Dummy_ As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy_ = DTA.ScalarQuery_GetCountByNoTransaksi(vDocNo)
            If Dummy_.HasValue = True Then
                Return Dummy_.Value
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

    Public Function GetJml(ByVal vCriteria As String, ByVal vSearchBy As ClsTrans_SchDownTime.EnumSearchBy, _
            ByVal vDari As Date, ByVal vSampai As Date) As Double
        Dim DTA As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtime1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchBy.DOCUMENT_NO Then
                If DTA.ScalarQuery_GetCountByDocNoLIKE(vCriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetCountByDocNoLIKE(vCriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchBy.DATE_ Then
                If DTA.ScalarQuery_GetCountByDate(vDari, vSampai).HasValue Then
                    Return DTA.ScalarQuery_GetCountByDate(vDari, vSampai).Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchBy.REASON Then
                If DTA.ScalarQuery_GetCountByReasonLIKE(vCriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetCountByReasonLIKE(vCriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchBy.REMARKS Then
                If DTA.ScalarQuery_GetCountByRemarksLIKE(vCriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetCountByRemarksLIKE(vCriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchBy.WORK_UNIT_CODE Then
                If DTA.ScalarQuery_GetCountByWorkUnitCodeLIKE(vCriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetCountByWorkUnitCodeLIKE(vCriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchBy.WORK_UNIT_DESC Then
                If DTA.ScalarQuery_GetCountByWorkUnitDescLIKE(vCriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetCountByWorkUnitDescLIKE(vCriteria + "%").Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTrans_SchDownTimeTableAdapters.tbltrans_sch_downtimeTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As Double = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim RetVal As String = ""
        Dim TipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime

        Try
            'cari tanggal sampai dan dari
            Me.objNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, TipeTransaksi, vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_MaxUrutan(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_MaxUrutan(TanggalEnd, TanggalStart).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            'cari nomor baru
            RetVal = Me.objNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, TipeTransaksi)
            While Me.ApakahSudahAda(RetVal) = True
                NoUrut = NoUrut + 1
                RetVal = Me.objNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, TipeTransaksi)
            End While

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function
End Class
