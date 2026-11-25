Option Explicit On
Option Strict Off

Public Class ClsTransaksiJournalQuickAdd
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mCode As String = ""
    Private mTipeQuickAdd As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.AdjustmenJournal
    Private mRemarks As String = ""
    Private mDesc As String = ""

    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now

    Public Property Desc() As String
        Get
            Return Me.mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
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

    Public Property Code() As String
        Get
            Return mCode
        End Get
        Set(ByVal value As String)
            mCode = value
        End Set
    End Property

    Public Property TipeQuickAdd() As Integer
        Get
            Return mTipeQuickAdd
        End Get
        Set(ByVal value As Integer)
            mTipeQuickAdd = value
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

    Public Property Timestamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
          ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
          Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Try
            '2. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If CreatedBy <> vSalesStaff Then
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
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

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
        ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
        Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Dim CreatedBy As String = ""
        Try

            '4. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If CreatedBy = vSalesStaff Then
                Return True
            Else
                If TampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim dta As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Try
            If String.IsNullOrEmpty(dta.ScalarQuery_GetDescByPrimary(vPrimary)) = False Then
                Return dta.ScalarQuery_GetDescByPrimary(vPrimary)
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

    Public Function GetDesc_ByCode(ByVal vCode As String) As String
        Dim dta As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Try
            If String.IsNullOrEmpty(dta.ScalarQuery_GetDescByCode(vCode)) = False Then
                Return dta.ScalarQuery_GetDescByCode(vCode)
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

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim dta As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Try
            If String.IsNullOrEmpty(dta.ScalarQuery_GetCodeByPrimary(vPrimary)) = False Then
                Return dta.ScalarQuery_GetCodeByPrimary(vPrimary)
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

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByCode(vCode).HasValue Then
                If DTA.ScalarQuery_GetJmlByCode(vCode).Value > 0 Then
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

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mCode) = True Then
                    MsgBox("Duplication in reference number '" + Me.mCode + "', please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                'no po tidak boleh duplikat
                Dim NoLama As String = Me.GetCode(Me.mPrimary)
                If Me.Code <> NoLama Then
                    If Me.ApakahSudahAda(Me.mCode) = True Then
                        MsgBox("Duplication in reference number '" + Me.mCode + "', please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'return value
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria, _
    ByVal TipeJournal As Integer) As Double
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tblCariTrans_Jurnal_Quick_AddTableAdapter
        Try
            If vSearchBy = EnumSearchCriteria.ByReferenceCode Then
                Return DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%", TipeJournal).Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%", TipeJournal).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
   ByVal DGV As DataGridView, ByVal TBL As DseTransaksiJurnalQuickAdd.tblCariTrans_Jurnal_Quick_AddDataTable, _
   ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal TipeJournal As Integer) As Boolean
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tblCariTrans_Jurnal_Quick_AddTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByReferenceCode Then
                QueryResult = DTA.FillBy_RefCodeTipe0(TBL, Criteria + "%", TipeJournal)
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_RemarksTipe0(TBL, Criteria + "%", TipeJournal)
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Reference number"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Remarks"
                DGV.Columns(3).Visible = False
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vDGVItem As DataGridView, _
       ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_quickAdd, ByVal TipeJournal As Integer) As Boolean
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Dim ObjItem As New ClsTransaksiJurnalQuickAdd_Item
        Try
            Me.mTipeQuickAdd = TipeJournal
            If Me.ApakahDataBenar(True) = True Then
                If DTA.Insert(Me.Code, Me.mTipeQuickAdd, Me.Remarks, vUserCode, Now, _
                 vUserCode, Now, Now, Me.mDesc) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimaryByCode(Me.mCode).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem(vDGVItem, Me.mPrimary, Frm)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mCode + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru(vDGVItem, vUserCode, Frm, TipeJournal)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_quickAdd, ByVal TipeJournal As Integer) As Boolean
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Dim ObjItem As New ClsTransaksiJurnalQuickAdd_Item
        Try
            Me.mTipeQuickAdd = TipeJournal
            If Me.ApakahDataBenar(False) = True Then
                If DTA.Update(Me.Code, Me.mTipeQuickAdd, Me.Remarks, vUserCode, Now, Now, _
                Me.mDesc, Me.Primary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)

                    'simpan data item baru
                    ObjItem.SimpanItem(vDGVItem, Me.mPrimary, frm)

                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mCode + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanPerubahan(vDGVItem, vUserCode, frm, TipeJournal)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal usercode As String) As Boolean
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, usercode) = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mCode = ""
            mTipeQuickAdd = ClsTransaksiJournal.EnumTipeJournal.AdjustmenJournal
            mRemarks = ""
            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function CariDetail(ByVal vPrimaryMaster As Double, _
    ByVal vDGVItem As DataGridView, ByVal frm As FrmTransaksiJournal_quickAdd, _
    ByVal tipe_journal As Integer) As Boolean
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Dim TBL As New DseTransaksiJurnalQuickAdd.tbltrans_jurnal_quick_addDataTable
        Dim row As DseTransaksiJurnalQuickAdd.tbltrans_jurnal_quick_addRow
        Dim ObjItem As New ClsTransaksiJurnalQuickAdd_Item
        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryMaster, tipe_journal) > 0 Then
                'salin data header
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = vPrimaryMaster
                    Me.mCode = .Code
                    Me.mTipeQuickAdd = .Tipe
                    Me.mRemarks = .Remarks
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mDesc = .Desc
                End With
                'salin data item ada 2 item 
                'CodeNo
                ObjItem.IsiDGVItem(vDGVItem, vPrimaryMaster, frm)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDGVItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function GetPrimaryRefCode(ByVal vRefCode As String) As Double
        Dim dta As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Try
            If dta.ScalarQuery_GetPrimaryByCode(vRefCode).HasValue Then
                Return dta.ScalarQuery_GetPrimaryByCode(vRefCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetRefCode(ByVal Vprimary As Double) As String
        Dim dta As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Try
            If String.IsNullOrEmpty(dta.ScalarQuery_GetCodeByPrimary(Vprimary)) = False Then
                Return dta.ScalarQuery_GetCodeByPrimary(Vprimary)
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

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim dta As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Try
            If String.IsNullOrEmpty(dta.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                Return dta.ScalarQuery_GetCreatedBy(vPrimary)
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

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim Dta As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Try
            If Dta.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return Dta.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function IsiDGVQuickAdd(ByVal vTipeJurnal As ClsTransaksiJournal.EnumTipeJournal, _
    ByVal vDgv As DataGridView, ByVal vTotalReceivement As Double, ByVal vKodeQuickAdd As String, _
    ByVal vIndexKodeAkun As Integer, _
    ByVal vIndex_Desc As Integer, ByVal vIndexPrimaryAkun As Integer, ByVal vIndexDebet As Integer, _
    ByVal vIndexKredit As Integer, ByVal vIndexRemarks As Integer, ByVal ObjCd As Core_Module.ClsCekData, _
    ByVal DefaultPrec As Integer, ByVal vConfirmFirst As Boolean, ByVal vIndexPrimaryDepartement As Integer, _
    ByVal vIndexPrimaryJob As Integer, ByVal TextQuickAdd As String) As Boolean
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseTransaksiJurnalQuickAdd.DataTable1DataTable
        Dim row As DseTransaksiJurnalQuickAdd.DataTable1Row
        Dim rowdgv As DataGridViewRow
        Dim MasukkanJumlah As Boolean = False
        Dim Hasil As MsgBoxResult = MsgBoxResult.Yes
      
        Try
            vDgv.AllowUserToAddRows = False
            If vDgv.Rows.Count > 0 Then
                If vConfirmFirst = True Then
                    Hasil = MsgBox("Insert journal using quick add?", MsgBoxStyle.YesNo, "Confirm")
                Else
                    Hasil = MsgBoxResult.Yes
                End If
            End If

            If Hasil = MsgBoxResult.Yes Then
                vDgv.Rows.Clear()
                DTA.Fill(TBL, vKodeQuickAdd, vTipeJurnal)
                If TBL.Rows.Count = 2 Then
                    MasukkanJumlah = True
                ElseIf vTotalReceivement = 0 And TBL.Rows.Count > 2 Then
                    MasukkanJumlah = True
                End If

                For Each row In TBL.Rows
                    vDgv.Rows.Add()
                    rowdgv = vDgv.Rows(vDgv.Rows.Count - 1)
                    With rowdgv
                        .Cells(vIndex_Desc).Value = row.Description
                        .Cells(vIndexDebet).Value = ObjCd.DigitGrouping(0, DefaultPrec)
                        .Cells(vIndexKredit).Value = ObjCd.DigitGrouping(0, DefaultPrec)
                        .Cells(vIndexKodeAkun).Value = row.Account_Code
                        .Cells(vIndexPrimaryAkun).Value = row.Primary_Akun
                        .Cells(vIndexRemarks).Value = ""
                        .Cells(vIndexPrimaryDepartement).Value = -1
                        .Cells(vIndexPrimaryJob).Value = -1

                        If MasukkanJumlah = True Then
                            If vTotalReceivement = 0 Then
                                If row.Saldo_Normal = 0 Then
                                    .Cells(vIndexDebet).Value = ObjCd.DigitGrouping(1, DefaultPrec)
                                Else
                                    .Cells(vIndexKredit).Value = ObjCd.DigitGrouping(1, DefaultPrec)
                                End If
                            Else
                                If vTotalReceivement > 0 Then
                                    If row.Saldo_Normal = 0 Then
                                        .Cells(vIndexDebet).Value = ObjCd.DigitGrouping(vTotalReceivement, DefaultPrec)
                                    Else
                                        .Cells(vIndexKredit).Value = ObjCd.DigitGrouping(vTotalReceivement, DefaultPrec)
                                    End If
                                Else
                                    If row.Saldo_Normal = 0 Then
                                        .Cells(vIndexKredit).Value = ObjCd.DigitGrouping(Math.Abs(vTotalReceivement), DefaultPrec)
                                    Else
                                        .Cells(vIndexDebet).Value = ObjCd.DigitGrouping(Math.Abs(vTotalReceivement), DefaultPrec)
                                    End If
                                End If
                            End If
                        Else
                            If row.Saldo_Normal = 0 Then
                                .Cells(vIndexDebet).Value = ObjCd.DigitGrouping(1, DefaultPrec)
                            Else
                                .Cells(vIndexKredit).Value = ObjCd.DigitGrouping(1, DefaultPrec)
                            End If
                        End If
                    End With
                Next
                vDgv.AllowUserToAddRows = True
                Return True
            Else
                vDgv.AllowUserToAddRows = False
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub IsiCboQuickAdd(ByVal vTipeJurnal As ClsTransaksiJournal.EnumTipeJournal, ByVal Cbo As ComboBox)
        Dim DTA As New DseTransaksiJurnalQuickAddTableAdapters.tbltrans_jurnal_quick_addTableAdapter
        Dim TBL As New DseTransaksiJurnalQuickAdd.tbltrans_jurnal_quick_addDataTable
        Dim row As DseTransaksiJurnalQuickAdd.tbltrans_jurnal_quick_addRow

        Try
            Cbo.Items.Clear()
            DTA.FillBy_TipeJurnal(TBL, vTipeJurnal)
            Cbo.Items.Add("")
            For Each row In TBL.Rows
                Cbo.Items.Add(row.Code)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub
End Class
