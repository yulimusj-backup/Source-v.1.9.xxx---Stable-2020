Option Explicit On
Option Strict Off

Public Class ClsTrans_SPL

    Private ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Private ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mNoDocument As String = ""
    Private mUrutan As Double = 0
    Private mDate As Date = Now
    Private mJobdesc As String = ""
    Private mRemarks As String = ""
    Private mAuthorized As Integer = 0
    Private mPrimaryAuth As Double = -1
    Private mFromDate As Date = Now
    Private mToDate As Date = Now

    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now
    Public print As Boolean = False



    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property NoDocument() As String
        Get
            Return Me.mNoDocument
        End Get
        Set(ByVal value As String)
            mNoDocument = value
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

    Public Property Tanggal() As Date
        Get
            Return Me.mDate
        End Get
        Set(ByVal value As Date)
            mDate = value
        End Set
    End Property

    Public Property JobDesc() As String
        Get
            Return Me.mJobdesc
        End Get
        Set(ByVal value As String)
            mJobdesc = value
        End Set
    End Property

    Public Property Authorized() As Integer
        Get
            Return Me.mAuthorized
        End Get
        Set(ByVal value As Integer)
            mAuthorized = value
        End Set
    End Property

    Public Property PrimaryAuth() As Double
        Get
            Return Me.mPrimaryAuth
        End Get
        Set(ByVal value As Double)
            mPrimaryAuth = value
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

    Public Property FromDate() As Date
        Get
            Return Me.mFromDate
        End Get
        Set(ByVal value As Date)
            mFromDate = value
        End Set
    End Property

    Public Property ToDate() As Date
        Get
            Return Me.mToDate
        End Get
        Set(ByVal value As Date)
            mToDate = value
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

    Public Function ApakahSudahAda(ByVal vNoDokumen As String) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter
        Dim dummy As New Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            dummy = DTA.ScalarQuery_GetJmlByNoDokumen(vNoDokumen)
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
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryDokumen As Double, _
       ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
       Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(CDec(vPrimaryDokumen))

            If CreatedBy = vSalesStaff Then
                Return True
            Else
                If TampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryDokumen As Double, _
        ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
        Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter
        Dim CreatedBy As String = ""
        Try

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If StrictEditting = True Then
                CreatedBy = DTA.ScalarQuery_GetCreatedBy(CDec(vPrimaryDokumen))
                If CreatedBy <> vSalesStaff Then
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            If GetStatusAuth(vPrimaryDokumen) = True Then
                MsgBox("SPL can't be changed or deleted because it's already authorize", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mNoDocument) = True Then
                    Me.mNoDocument = Me.GetNoBaru(Me.mDate)
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalDokumen As Date) As String
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As Double = 0
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SPL, vTanggalDokumen)

            'cari no urut
            If DTA.ScalarQuery_GetMaxUrutan(TanggalStart, TanggalEnd).HasValue = True Then
                NoUrut = DTA.ScalarQuery_GetMaxUrutan(TanggalStart, TanggalEnd).Value + 1
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            Return Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalDokumen, NoUrut, _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SPL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            'ObjNumbering = Nothing
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryDokumen As Double, _
       ByVal vDGVItemWS As DataGridView, ByVal vDGVItemWU As DataGridView, ByVal vFrm As FrmTrans_SPL) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter
        Dim TBL As New DseTrans_SPL.tbltrans_splDataTable
        Dim row As DseTrans_SPL.tbltrans_splRow
        Dim ObjItem As New ClsTrans_SPLItem
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'DTA.FillBy_PrimaryTransaksi(TBL, CDec(Primary))

            If DTA.FillBy_PrimaryDokumen(TBL, CDec(vPrimaryDokumen)) > 0 Then
                row = TryCast(TBL.Rows(0), DseTrans_SPL.tbltrans_splRow)
                'row = TBL.Rows(0)
                With row
                    Me.mPrimary = vPrimaryDokumen
                    Me.mNoDocument = .document_no
                    Me.mUrutan = .urutan
                    Me.mDate = ._date
                    Me.mJobdesc = .Job_Desc
                    Me.mRemarks = .remarks
                    Me.mAuthorized = .is_authorized
                    Me.mPrimaryAuth = .primary_user_auth
                    Me.mFromDate = .from_date
                    Me.mToDate = .to_date

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With

                ObjItem.IsiDGVItemWS(vPrimaryDokumen, vDGVItemWS, vFrm)
                ObjItem.IsiDGVItemWU(vPrimaryDokumen, vDGVItemWU, vFrm)
                Return True
            Else
                Me.ResetObject(vDGVItemWS)
                Me.ResetObject(vDGVItemWU)
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            'ObjItem = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vDGVItem As DataGridView)
        Try
            mPrimary = -1
            mNoDocument = ""
            mUrutan = 0
            mDate = Now
            mJobdesc = ""
            mRemarks = ""
            mFromDate = Now
            mToDate = Now

            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today
            vDGVItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
              EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
              Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long

        Dim DTA As New DseTrans_SPLTableAdapters.tblCaritrans_splTableAdapter
        Dim RetVal As Nullable(Of Long)


        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByNoDokumen Then
                'Return DTA.ScalarQuery_GetJmlByNoDokumenLIKE(vKriteria + "%")
                RetVal = DTA.ScalarQuery_GetJmlByNoDokumenLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByTanggalTransaksi Then
                'Return DTA.ScalarQuery_GetJmlByTanggalDariSampai(vTanggalSampai, vTanggalDari)
                RetVal = DTA.ScalarQuery_GetJmlByTanggalDariSampai(vTanggalDari, vTanggalSampai)
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByJobdesc Then
                'Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(vKriteria + "%")
                RetVal = DTA.ScalarQuery_GetJmlByJobDescLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                'Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(vKriteria + "%")
                RetVal = DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%")
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

    Public Function IsiDGV(ByVal Criteria As String, _
   ByVal DGV As DataGridView, ByVal TBL As DseTrans_SPL.tblCaritrans_splDataTable, _
   ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
   Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
   Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tblCaritrans_splTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoDokumen Then
                QueryResult = DTA.FillBy_NoDokumenLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TanggalDariSampai(TBL, vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = EnumSearchCriteria.ByJobdesc Then
                QueryResult = DTA.FillBy_JobDescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView

            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Document No"
                DGV.Columns(2).HeaderText = "Date"
                DGV.Columns(3).HeaderText = "Reason"

                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > 10000000 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Ticks / TimeSpan.TicksPerSecond) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, 0, MidpointRounding.AwayFromZero)) + _
                    " miliseconds"
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

    Public Overridable Function Hapus(ByVal vPrimaryDokumen As Double, _
    ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimaryDokumen, StrictSalesStaffDelete, UserCode) = True Then
                If DTA.Delete(CDec(vPrimaryDokumen)) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vNoDokumen As String, ByVal vTipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering, _
    ByVal vtanggaldokumen As Date, ByVal vDGVitemWS As DataGridView, ByVal vDGVitemWU As DataGridView, ByVal vUserCode As String, ByVal vfrmSPL As FrmTrans_SPL, ByVal vTbLog As DseTrans_SPL.tbltrans_spl_item_logDataTable) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter
        Dim ObjItem As New ClsTrans_SPLItem
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True) = True Then
                Me.Urutan = ObjNumbering.GetUrutan(vNoDokumen, vTipeTransaksi, vtanggaldokumen)
                If DTA.Insert(Me.mNoDocument, Me.mUrutan, Me.mDate, Me.mRemarks, Me.mAuthorized, _
                Me.mFromDate, Me.mToDate, vUserCode, Now, vUserCode, Now, Now, Me.mJobdesc, Me.mPrimaryAuth) Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimaryByNoDokumen(Me.mNoDocument).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    '//JANGAN LUPA!
                    ObjItem.SimpanItemWS(vDGVitemWS, mPrimary, vfrmSPL)
                    ObjItem.SimpanItemWU(vDGVitemWU, mPrimary, vfrmSPL, vTbLog)
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoDocument + "' for key 3" Then
                Return Me.SimpanBaru(vNoDokumen, vTipeTransaksi, vtanggaldokumen, vDGVitemWS, vDGVitemWU, vUserCode, vfrmSPL, vTbLog)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function
    Public Overridable Function SimpanPerubahan(ByVal vNoDokumen As String, ByVal vTipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering, _
    ByVal vtanggaldokumen As Date, ByVal vDGVitemWS As DataGridView, ByVal vDGVitemWU As DataGridView, ByVal vUserCode As String, ByVal vfrmSPL As FrmTrans_SPL, ByVal vTbLog As DseTrans_SPL.tbltrans_spl_item_logDataTable) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter
        Dim ObjItem As New ClsTrans_SPLItem
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False) = True Then
                Me.Urutan = ObjNumbering.GetUrutan(vNoDokumen, vTipeTransaksi, vtanggaldokumen)
                If DTA.Update(Me.mNoDocument, Me.mUrutan, Me.mDate, Me.mRemarks, _
                Me.mAuthorized, Me.mFromDate, Me.mToDate, vUserCode, Now, Now, Me.mJobdesc, Me.mPrimaryAuth, CDec(Me.mPrimary), Me.mTimestamp) Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode


                    '//JANGAN LUPA!
                    ObjItem.HapusItemWS(Me.mPrimary)
                    ObjItem.SimpanItemWS(vDGVitemWS, mPrimary, vfrmSPL)

                    ObjItem.HapusItemWU(Me.mPrimary)
                    ObjItem.SimpanItemWU(vDGVitemWU, mPrimary, vfrmSPL, vTbLog)

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

    Public Function GetPrimary(ByVal vNoDokumen As String) As Double
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryByNoDokumen(vNoDokumen).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryByNoDokumen(vNoDokumen).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoDokumen(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoDokumenByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetNoDokumenByPrimary(vPrimary)
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

    Public Function GetTanggal(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetTanggalDokumenByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetTanggalDokumenByPrimary(vPrimary)
            Else
                Return Now
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Now
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCreatedBy(ByVal vPrimaryTransaksi As String) As String
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)) = False Then
                Return DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
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

    Public Function CetakanSPL(ByVal vPrimarySPL As Double, ByVal vCrvSPL As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean
        Dim DS As New DseTrans_SPL
        Dim Dta As New DseTrans_SPLTableAdapters.DataTable4TableAdapter
        Dim Tbl As New DseTrans_SPL.DataTable4DataTable
        Dim Row As DseTrans_SPL.DataTable4Row

        Dim TBLCetak As DseTrans_SPL.DataTable3DataTable = DS.DataTable3
        Dim RowCetak As DseTrans_SPL.DataTable3Row

        Dim Report As New CrpTrans_SPL
        Dim frmOption As New FrmTrans_SPL_Cetak_Option

        Dim parTotal_ As String = ""

        Try
            If frmOption.ShowDialog = DialogResult.OK Then
                print = True
                Dta.Connection.ConnectionString = Core_Module.GetConstring
                If Dta.Fill(Tbl, vPrimarySPL) > 0 Then
                    For Each Row In Tbl
                        If Row.is_early_ot = 1 Then
                            RowCetak = TBLCetak.NewDataTable3Row
                            RowCetak.document_no = Row.document_no
                            RowCetak._date = Row._date
                            RowCetak.from_date = Row.from_date
                            RowCetak.to_date = Row.to_date
                            RowCetak.Reason = Row.Job_Desc
                            RowCetak.Date_Log = Row.Date_log
                            RowCetak.Work_Shift = Row.Work_Shift
                            RowCetak.Work_Unit = Row.work_unit_code
                            RowCetak.Type = "Early OT"
                            RowCetak.Duration = ObjCD.Convert_ToString(Row.early_ot_duration)
                            RowCetak.Delay = ObjCD.Convert_ToString(Row.early_delay_duration)
                            RowCetak.Break = ObjCD.Convert_ToString(Row.early_break_duration)
                            RowCetak.Delay_Before_Break = ObjCD.Convert_ToString(Row.delay_before_early_break)
                            TBLCetak.Rows.Add(RowCetak)
                        End If
                        If Row.is_ot = 1 Then
                            RowCetak = TBLCetak.NewDataTable3Row
                            RowCetak.document_no = Row.document_no
                            RowCetak._date = Row._date
                            RowCetak.from_date = Row.from_date
                            RowCetak.to_date = Row.to_date
                            RowCetak.Reason = Row.Job_Desc
                            RowCetak.Date_Log = Row.Date_log
                            RowCetak.Work_Shift = Row.Work_Shift
                            RowCetak.Work_Unit = Row.work_unit_code
                            RowCetak.Type = "OT"
                            RowCetak.Duration = ObjCD.Convert_ToString(Row.ot_duration)
                            RowCetak.Delay = ObjCD.Convert_ToString(Row.delay_duration)
                            RowCetak.Break = ObjCD.Convert_ToString(Row.break)
                            RowCetak.Delay_Before_Break = ObjCD.Convert_ToString(Row.delay_before_break)
                            TBLCetak.Rows.Add(RowCetak)
                        End If
                    Next
                    parTotal_ = Me.ObjCD.Convert_ToString(Me.CalculateSubTotal(TBLCetak))
                End If
            Else
                print = False
            End If

            For Each RowCetak In TBLCetak
                If RowCetak.Break = "00.00" Then
                    RowCetak.Break = "-"
                End If
                If RowCetak.Duration = "00.00" Then
                    RowCetak.Duration = "-"
                End If
                If RowCetak.Delay = "00.00" Then
                    RowCetak.Delay = "-"
                End If
                If RowCetak.Delay_Before_Break = "00.00" Then
                    RowCetak.Delay_Before_Break = "-"
                End If

            Next

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_ShowSubTotal.ParameterFieldName, frmOption.CheckBox1.Checked)
            Report.SetParameterValue(Report.Parameter_Total.ParameterFieldName, parTotal_)
            vCrvSPL.ReportSource = Report
            vCrvSPL.Zoom(100)
            vCrvSPL.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
        Return True
    End Function

    Private Function CalculateSubTotal(ByRef vTbl As DseTrans_SPL.DataTable3DataTable) As Integer
        Dim row As DseTrans_SPL.DataTable3Row
        Dim tblDummy As New DseTrans_SPL.DataTable5DataTable
        Dim rowDummy As DseTrans_SPL.DataTable5Row
        Dim Total_ As Integer = 0
        Try
            For Each row In vTbl
                tblDummy.DefaultView.RowFilter = "[Date] = #" & row.Date_Log.Month.ToString & "/" & row.Date_Log.Day.ToString & "/" & row.Date_Log.Year.ToString & "#"
                If tblDummy.DefaultView.Count = 0 Then
                    rowDummy = tblDummy.NewDataTable5Row
                    rowDummy._Date = row.Date_Log
                    rowDummy.Sub_Total = row.Duration
                    tblDummy.Rows.Add(rowDummy)
                Else
                    rowDummy = tblDummy.DefaultView(0).Row
                    rowDummy.Sub_Total = Me.ObjCD.Convert_ToString(Me.ObjCD.Convert_ToMinutes(rowDummy.Sub_Total) + Me.ObjCD.Convert_ToMinutes(row.Duration))

                End If
            Next
            For Each row In vTbl
                tblDummy.DefaultView.RowFilter = "[Date] = #" & row.Date_Log.Month.ToString & "/" & row.Date_Log.Day.ToString & "/" & row.Date_Log.Year.ToString & "#"
                If tblDummy.DefaultView.Count > 0 Then
                    rowDummy = tblDummy.DefaultView(0).Row
                    row.Sub_Total_Duration = rowDummy.Sub_Total
                    'Total_ += ObjCD.Convert_ToMinutes(row.Sub_Total_Duration)

                Else
                    MsgBox("Error")
                End If
                row.Date_String = row.Date_Log.Year.ToString & "-" & row.Date_Log.Month.ToString & "-" & row.Date_Log.Day.ToString
            Next
            For Each row In vTbl
                If tblDummy.DefaultView.Count > 0 Then
                    Total_ += ObjCD.Convert_ToMinutes(row.Duration)
                End If
            Next
            ObjCD.Convert_ToString(Total_)
            Return Total_
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
        End Try
    End Function


    Private Function CalculateSubTotalAuthorize(ByRef vTbl As DseTrans_SPL.DataTable3DataTable) As Integer
        Dim row As DseTrans_SPL.DataTable3Row
        Dim tblDummy As New DseTrans_SPL.DataTable5DataTable
        Dim rowDummy As DseTrans_SPL.DataTable5Row
        Try
            For Each row In vTbl
                tblDummy.DefaultView.RowFilter = "[Date] = #" & row.Date_Log.Month.ToString & "/" & row.Date_Log.Day.ToString & "/" & row.Date_Log.Year.ToString & "# AND [Primary SPL] = " & row.Primary_SPL.ToString & ""
                If tblDummy.DefaultView.Count = 0 Then
                    rowDummy = tblDummy.NewDataTable5Row
                    rowDummy._Date = row.Date_Log
                    rowDummy.Sub_Total = row.Duration
                    rowDummy.Primary_SPL = row.Primary_SPL
                    tblDummy.Rows.Add(rowDummy)
                Else
                    rowDummy = tblDummy.DefaultView(0).Row
                    rowDummy.Sub_Total = Me.ObjCD.Convert_ToString(Me.ObjCD.Convert_ToMinutes(rowDummy.Sub_Total) + Me.ObjCD.Convert_ToMinutes(row.Duration))

                End If
            Next
            For Each row In vTbl
                tblDummy.DefaultView.RowFilter = "[Date] = #" & row.Date_Log.Month.ToString & "/" & row.Date_Log.Day.ToString & "/" & row.Date_Log.Year.ToString & "# AND [Primary SPL] = " & row.Primary_SPL.ToString & ""
                If tblDummy.DefaultView.Count > 0 Then
                    rowDummy = tblDummy.DefaultView(0).Row
                    row.Sub_Total_Duration = rowDummy.Sub_Total
                Else
                    MsgBox("Error")
                End If
                row.Date_String = row.Date_Log.Year.ToString & "-" & row.Date_Log.Month.ToString & "-" & row.Date_Log.Day.ToString
            Next
            For Each row In vTbl
                tblDummy.DefaultView.RowFilter = "[Primary SPL] = " & row.Primary_SPL.ToString & ""
                If tblDummy.DefaultView.Count > 0 Then
                    For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                        rowDummy = tblDummy.DefaultView(i).Row
                        row.Total_Per_SPL = ObjCD.Convert_ToString(ObjCD.Convert_ToMinutes(rowDummy.Sub_Total) + ObjCD.Convert_ToMinutes(row.Total_Per_SPL))
                    Next
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
        End Try
    End Function
    Public Function CetakanSPLAuthorize(ByVal vListPrimary As List(Of Double), ByVal vCrvSPL As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean
        Dim DS As New DseTrans_SPL
        Dim Dta As New DseTrans_SPLTableAdapters.DataTable4TableAdapter
        Dim Tbl As New DseTrans_SPL.DataTable4DataTable
        Dim Row As DseTrans_SPL.DataTable4Row

        Dim TBLCetak As DseTrans_SPL.DataTable3DataTable = DS.DataTable3
        Dim RowCetak As DseTrans_SPL.DataTable3Row

        Dim Report As New CrpTrans_SPL_Batch
        Dim frmOption As New FrmTrans_SPL_Cetak_Option

        Dim parTotal_ As String = ""
        Dim lastCode As String = ""
        Try
            If frmOption.ShowDialog = DialogResult.OK Then
                print = True
                Dta.Connection.ConnectionString = Core_Module.GetConstring
                Dta.ClearBeforeFill = False
                For Each primary As Double In vListPrimary
                    Dta.Fill(Tbl, primary)
                Next
                If Tbl.Count > 0 Then
                    For Each Row In Tbl
                        If Row.is_early_ot = 1 Then
                            RowCetak = TBLCetak.NewDataTable3Row
                            RowCetak.Primary_SPL = Row.Primary_SPL
                            RowCetak.document_no = Row.document_no
                            RowCetak._date = Row._date
                            RowCetak.from_date = Row.from_date
                            RowCetak.to_date = Row.to_date
                            RowCetak.Reason = Row.Job_Desc
                            RowCetak.Date_Log = Row.Date_log
                            RowCetak.Work_Shift = Row.Work_Shift
                            RowCetak.Work_Unit = Row.work_unit_code
                            RowCetak.Type = "Early OT"
                            RowCetak.Duration = ObjCD.Convert_ToString(Row.early_ot_duration)
                            RowCetak.Delay = ObjCD.Convert_ToString(Row.early_delay_duration)
                            RowCetak.Break = ObjCD.Convert_ToString(Row.early_break_duration)
                            RowCetak.Delay_Before_Break = ObjCD.Convert_ToString(Row.delay_before_early_break)
                            TBLCetak.Rows.Add(RowCetak)
                        End If
                        If Row.is_ot = 1 Then
                            RowCetak = TBLCetak.NewDataTable3Row
                            RowCetak.Primary_SPL = Row.Primary_SPL
                            RowCetak.document_no = Row.document_no
                            RowCetak._date = Row._date
                            RowCetak.from_date = Row.from_date
                            RowCetak.to_date = Row.to_date
                            RowCetak.Reason = Row.Job_Desc
                            RowCetak.Date_Log = Row.Date_log
                            RowCetak.Work_Shift = Row.Work_Shift
                            RowCetak.Work_Unit = Row.work_unit_code
                            RowCetak.Type = "OT"
                            RowCetak.Duration = ObjCD.Convert_ToString(Row.ot_duration)
                            RowCetak.Delay = ObjCD.Convert_ToString(Row.delay_duration)
                            RowCetak.Break = ObjCD.Convert_ToString(Row.break)
                            RowCetak.Delay_Before_Break = ObjCD.Convert_ToString(Row.delay_before_break)
                            TBLCetak.Rows.Add(RowCetak)
                        End If
                    Next
                    Me.ObjCD.Convert_ToString(Me.CalculateSubTotalAuthorize(TBLCetak))
                End If
            Else
                print = False
            End If

            For Each RowCetak In TBLCetak
                If RowCetak.Break = "00.00" Then
                    RowCetak.Break = "-"
                End If
                If RowCetak.Duration = "00.00" Then
                    RowCetak.Duration = "-"
                End If
                If RowCetak.Delay = "00.00" Then
                    RowCetak.Delay = "-"
                End If
                If RowCetak.Delay_Before_Break = "00.00" Then
                    RowCetak.Delay_Before_Break = "-"
                End If
                lastCode = RowCetak.document_no
            Next

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_LastCode.ParameterFieldName, lastCode)
            Report.SetParameterValue(Report.Parameter_ShowSubTotal.ParameterFieldName, frmOption.CheckBox1.Checked)
            vCrvSPL.ReportSource = Report
            vCrvSPL.Zoom(100)
            vCrvSPL.Show()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
        Return True
    End Function

    Public Function ApplyToWorkUnitGroup(ByVal vDGV As DataGridView, ByVal vFrm As FrmTrans_SPL) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tblmaster_workunitTableAdapter
        Dim TBL As New DseTrans_SPL.tblmaster_workunitDataTable
        Dim rowItem As DseTrans_SPL.tblmaster_workunitRow
        Dim rowDGV As DataGridViewRow
        Dim ij As New Core_Module.FrmMasterWorkUnitGroup_Cari

        Try
            ij.BackColorAll = vFrm.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryWorkUnitGroup = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                vDGV.AllowUserToAddRows = False
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                If DTA.FillBy_PrimaryWorkUnit(TBL, ij.PrimaryWorkUnitGroup) > 0 Then

                    For Each rowItem In TBL

                        If vFrm.GetIndexItemWU(rowItem.primary, -1) = -1 Then
                            vDGV.Rows.Add()
                            rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                            rowDGV.Cells(vFrm.colWU_No.Index).Value = vDGV.Rows.Count
                            rowDGV.Cells(vFrm.colWU_Code.Index).Value = rowItem.work_unit_code
                            rowDGV.Cells(vFrm.colWU_Desc.Index).Value = rowItem.description
                            rowDGV.Cells(vFrm.colWU_Primary.Index).Value = rowItem.primary
                        End If
                    Next
                End If
            End If
            vDGV.AllowUserToAddRows = True
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try

    End Function
    Public Function ApplyToAllWokUnit(ByVal vDGV As DataGridView, ByVal vFrm As FrmTrans_SPL) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tblmaster_workunitTableAdapter
        Dim TBL As New DseTrans_SPL.tblmaster_workunitDataTable
        Dim rowItem As DseTrans_SPL.tblmaster_workunitRow
        Dim rowDGV As DataGridViewRow
        Try
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGV.Rows.Clear()
            vDGV.AllowUserToAddRows = False
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(TBL) > 0 Then
                For Each rowItem In TBL
                    vDGV.Rows.Add()
                    rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.colWU_No.Index).Value = ""
                        .Cells(vFrm.colWU_Code.Index).Value = rowItem.work_unit_code
                        .Cells(vFrm.colWU_Desc.Index).Value = rowItem.description
                        .Cells(vFrm.colWU_Primary.Index).Value = rowItem.primary
                    End With
                Next
                Me.ObjCD.SetNumber(vDGV)
            End If
            vDGV.AllowUserToAddRows = True
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub CreateLog(ByVal vDGV As DataGridView, ByVal vFrm As FrmTrans_SPL)
        Dim dtaWS As New DseTrans_SPLTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim tblWS As New DseTrans_SPL.tbltrans_mo_workshift_allocateDataTable
        Dim rowWS As DseTrans_SPL.tbltrans_mo_workshift_allocateRow
        Try
            vFrm.mTblLog.Clear()
            dtaWS.Connection.ConnectionString = Core_Module.GetConstring
            dtaWS.Fill(tblWS, vFrm.DtpFrom.Value.Date, vFrm.DtpTo.Value.Date)
            For Each dgvr As DataGridViewRow In vDGV.Rows
                If dgvr.Cells(vFrm.colWU_Primary.Index).Value <> "-1" Then
                    For i As Integer = 0 To DateDiff(DateInterval.Day, vFrm.DtpFrom.Value.Date, vFrm.DtpTo.Value.Date)
                        tblWS.DefaultView.RowFilter = "[date] = #" & vFrm.DtpFrom.Value.AddDays(i).Month & "/" & vFrm.DtpFrom.Value.AddDays(i).Day & "/" & vFrm.DtpFrom.Value.AddDays(i).Year & "# and [primary work unit] = " & dgvr.Cells(vFrm.colWU_Primary.Index).Value.ToString
                        If tblWS.DefaultView.Count > 0 Then
                            For u As Integer = 0 To tblWS.DefaultView.Count - 1
                                rowWS = tblWS.DefaultView(u).Row
                                CreateLogTbl(vFrm.mTblLog, rowWS.primary_workshift, rowWS.primary_work_unit, vFrm.DtpFrom.Value.AddDays(i).Date, vFrm.dgvWorkshift, vFrm)
                            Next
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateLogTbl(ByVal vTbl As DseTrans_SPL.tbltrans_spl_item_logDataTable, _
            ByVal vPrimaryWS As Double, ByVal vPrimaryWU As Double, ByVal vDate As Date, ByVal vDGV As DataGridView, ByVal vFrm As FrmTrans_SPL)

        Try
            For Each r As DataGridViewRow In vDGV.Rows
                If r.Cells(vFrm.colPrimaryWorkshift.Index).Value <> "-1" Then
                    If r.Cells(vFrm.colPrimaryWorkshift.Index).Value = vPrimaryWS Then
                        vFrm.mRowLog = vFrm.mTblLog.Newtbltrans_spl_item_logRow
                        With vFrm.mRowLog
                            .primary = -1
                            .primary_spl = vFrm.txtPrimary.Text
                            ._date = vDate
                            .primary_wokunit = vPrimaryWU
                            .primary_workshift = vPrimaryWS
                            If r.Cells(vFrm.colEarlyOT.Index).Value = True Then
                                .is_early_ot = 1
                            Else
                                .is_early_ot = 0
                            End If
                            .early_ot_duration = ObjCD.Convert_ToMinutes(r.Cells(vFrm.colEarlyOTDuration.Index).Value)
                            .early_delay_duration = ObjCD.Convert_ToMinutes(r.Cells(vFrm.colEarlyDelayDuration.Index).Value)
                            .early_break_duration = ObjCD.Convert_ToMinutes(r.Cells(vFrm.ColEarlyBreak.Index).Value)
                            .delay_before_early_break = ObjCD.Convert_ToMinutes(r.Cells(vFrm.ColDelayBeforeEarlyBreak.Index).Value)

                            If r.Cells(vFrm.colOT.Index).Value = True Then
                                .is_ot = 1
                            Else
                                .is_ot = 0
                            End If
                            .ot_duration = ObjCD.Convert_ToMinutes(r.Cells(vFrm.colOTDuration.Index).Value)
                            .delay_duration = ObjCD.Convert_ToMinutes(r.Cells(vFrm.colDelayDuration.Index).Value)
                            .break = ObjCD.Convert_ToMinutes(r.Cells(vFrm.ColBreak.Index).Value)
                            .delay_before_break = ObjCD.Convert_ToMinutes(r.Cells(vFrm.ColDelayBeforeBreak.Index).Value)

                        End With
                        vFrm.mTblLog.Rows.Add(vFrm.mRowLog)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahSudahDiallocate(ByVal vDGV As DataGridView, ByVal vFrm As FrmTrans_SPL) As Boolean
        Dim dtaWS As New DseTrans_SPLTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim tblWS As New DseTrans_SPL.tbltrans_mo_workshift_allocateDataTable
        'Dim rowWS As DseTrans_SPL.tbltrans_mo_workshift_allocateRow
        Dim listWU As New List(Of String)
        Try
            dtaWS.Connection.ConnectionString = Core_Module.GetConstring
            dtaWS.Fill(tblWS, vFrm.DtpFrom.Value.Date, vFrm.DtpTo.Value.Date)

            For Each rd As DataGridViewRow In vDGV.Rows
                If rd.Cells(vFrm.colWU_Primary.Index).Value <> "-1" Then
                    For i As Integer = 0 To DateDiff(DateInterval.Day, vFrm.DtpFrom.Value.Date, vFrm.DtpTo.Value.Date)
                        tblWS.DefaultView.RowFilter = "[date] = #" & vFrm.DtpFrom.Value.AddDays(i).Month & "/" & vFrm.DtpFrom.Value.AddDays(i).Day & "/" & vFrm.DtpFrom.Value.AddDays(i).Year & "# and [primary work unit] = " & rd.Cells(vFrm.colWU_Primary.Index).Value.ToString
                        If tblWS.DefaultView.Count = 0 Then

                            If listWU.Contains(rd.Cells(vFrm.colWU_Code.Index).Value) = False Then
                                listWU.Add(rd.Cells(vFrm.colWU_Code.Index).Value)
                            End If
                        End If
                    Next
                End If
            Next

            If listWU.Count > 0 Then
                Dim msg As String = ""
                For Each s As String In listWU
                    msg = msg & s & ", "
                Next
                msg = msg.Substring(0, msg.Length - 2) & "."
                MsgBox("These work unit should be workshift allocate first: " & msg, MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahSPLSudahdibuat(ByVal vDGV As DataGridView, ByVal vFrm As FrmTrans_SPL) As Boolean
        Dim dtaLOG As New DseTrans_SPLTableAdapters.tbltrans_spl_item_logTableAdapter
        Dim tblLOG As New DseTrans_SPL.tbltrans_spl_item_logDataTable

        Dim listWU As New List(Of String)
        Dim listTanggal As New List(Of String)
        Try
            dtaLOG.Connection.ConnectionString = Core_Module.GetConstring
            dtaLOG.Fill(tblLOG)

            For Each rd As DataGridViewRow In vDGV.Rows
                If rd.Cells(vFrm.colWU_Primary.Index).Value <> "-1" Then
                    For i As Integer = 0 To DateDiff(DateInterval.Day, vFrm.DtpFrom.Value.Date, vFrm.DtpTo.Value.Date)
                        tblLOG.DefaultView.RowFilter = "[date] = #" & vFrm.DtpFrom.Value.AddDays(i).Month & "/" & vFrm.DtpFrom.Value.AddDays(i).Day & "/" & vFrm.DtpFrom.Value.AddDays(i).Year & "# and [primary wokunit] = " & rd.Cells(vFrm.colWU_Primary.Index).Value.ToString
                        If tblLOG.DefaultView.Count > 0 Then
                            If listWU.Contains(rd.Cells(vFrm.colWU_Code.Index).Value) = False Then
                                Dim tgl As Date = vFrm.DtpFrom.Value.AddDays(i)
                                listWU.Add(rd.Cells(vFrm.colWU_Code.Index).Value & "," & Format(tgl, "dd MMM yyyy"))

                            End If
                        End If
                    Next
                End If
            Next

            If listWU.Count > 0 Then
                Dim msg, wu, _date As String
                Dim hasil(1) As String
                msg = ""
                For Each list As String In listWU
                    hasil = Split(list, ",")
                    wu = hasil(0)
                    _date = hasil(1)
                    msg += "SPL Workunit " & wu & " has already been created on date : " & _date & vbNewLine
                Next
                MsgBox(msg, MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

        Return True
    End Function

    Public Function apakahSudahBundling(ByVal vDGV As DataGridView, ByVal vFRM As FrmTrans_SPL) As Boolean
        Dim dtaBDL As New DseTrans_SPLTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim tblBDL As New DseTrans_SPL.tbltransaksi_stockbundlingDataTable
        Try
            dtaBDL.Connection.ConnectionString = Core_Module.GetConstring
            dtaBDL.Fill(tblBDL)

            For Each rd As DataGridViewRow In vDGV.Rows
                If rd.Cells(vFRM.colWU_Primary.Index).Value <> "-1" Then
                    For i As Integer = 0 To DateDiff(DateInterval.Day, vFRM.DtpFrom.Value.Date, vFRM.DtpTo.Value.Date)
                        tblBDL.DefaultView.RowFilter = "[primary work unit] = " & rd.Cells(vFRM.colWU_Primary.Index).Value.ToString & " AND [end time actual] >= #" & vFRM.DtpFrom.Value.AddDays(i).Month & "/" & vFRM.DtpFrom.Value.AddDays(i).Day & "/" & vFRM.DtpFrom.Value.AddDays(i).Year & " 23:59:59# AND [start time actual] <= #" & vFRM.DtpFrom.Value.AddDays(i).Month & "/" & vFRM.DtpFrom.Value.AddDays(i).Day & "/" & vFRM.DtpFrom.Value.AddDays(i).Year & "# "
                        If tblBDL.DefaultView.Count > 0 Then
                            Return False
                        End If
                    Next
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetStatusAuth(ByVal vPrimaryTransaksi As String) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_splTableAdapter
        Dim dummy As New Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetStatsAtuhByPrimary(vPrimaryTransaksi)) = False Then
                Return DTA.ScalarQuery_GetStatsAtuhByPrimary(vPrimaryTransaksi)
            Else
                Return ""
            End If
            dummy = DTA.ScalarQuery_GetStatsAtuhByPrimary(vPrimaryTransaksi)
            If dummy.HasValue Then
                If dummy.Value = 1 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
