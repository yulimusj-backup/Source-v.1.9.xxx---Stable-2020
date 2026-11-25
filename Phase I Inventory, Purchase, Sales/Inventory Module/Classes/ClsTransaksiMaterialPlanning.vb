Option Explicit On
Option Strict Off

Public Class ClsTransaksiMaterialPlanning
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mNoUrut As Double = -1
    Private mNoTransaksi As String = ""
    Private mTanggalTransaksi As Date = Now
    Private mRemarks As String = ""
    Private mRefDate As Date = #1/1/1950#
    Private mIsClosed As Integer = 0

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Enum EnumActionType As Integer
        BUY_CREATE_PO
        BUY_ADDTO_PO
        PRODUCE_CREATE_MO
        PRODUCE_ADDTO_MO
    End Enum

    Public Enum EnumTipeExclusion As Integer
        Exclude
        Include
    End Enum

    Public Enum EnumTipeTransaksiExclusion As Integer
        MaterialPlanning
        ShippingForecast
    End Enum

    ''rosaidn-r8
    Public Property IsClosed() As Integer
        Get
            Return Me.mIsClosed
        End Get
        Set(ByVal value As Integer)
            Me.mIsClosed = value
        End Set
    End Property

    Public Property RefDate() As Date
        Get
            Return Me.mRefDate
        End Get
        Set(ByVal value As Date)
            mRefDate = value
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

    Public Property TanggalTransaksi() As Date
        Get
            Return Me.mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return Me.mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property NoUrut() As Double
        Get
            Return Me.mNoUrut
        End Get
        Set(ByVal value As Double)
            mNoUrut = value
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

    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
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

    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
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

    Public Function GetTanggalMP(ByVal vPrimaryMP As Double) As Date
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDate(vPrimaryMP).HasValue = True Then
                Return DTA.ScalarQuery_GetDate(vPrimaryMP).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
       ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
       Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTglSampaiDari(vTanggalSampai, vTanggalDari).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE(Criteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoBaru1(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As Double = 0
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cari tanggal sampai dan dari
            Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP, vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetNoUrutMax(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetNoUrutMax(TanggalEnd, TanggalStart).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            Return Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable, _
    ByVal vDT_ItemGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable, _
    ByVal vDT_ItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable, _
    ByVal vDT_ItemPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, _
    ByVal vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, _
    ByVal vDT_ItemSettingLocation As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, _
    ByVal vDT_ItemQty As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable, _
    ByVal vDT_ItemMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable, _
    ByVal vDT_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable, _
    ByVal vDT_ItemFollowUp As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupDataTable, _
    ByVal vDT_ItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable, _
    ByVal vDT_ItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim TBL As New DseTransaksiMaterialPlanning2.tbltrans_material_planningDataTable
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planningRow
        Dim ObjItem As New ClsTransaksiMaterialPlanning_item

        Try
            'fil basis data
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mNoTransaksi = .doc_no
                    Me.mNoUrut = .urutan
                    Me.mPrimary = .primary
                    Me.mRemarks = .remarks
                    Me.mTanggalTransaksi = ._date
                    Me.mRefDate = .ref_date
                    Me.mIsClosed = .is_closed
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With

                'salin data item 
                ObjItem.IsiDGV(Me.mPrimary, vDT, vDT_ItemGIT, vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, vDT_ItemFollowUp, vDT_ItemSOBDelivery, _
                vDT_ItemPrOBDelivery)

                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDT)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiMaterialPlanning2.tbltrans_material_planning1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_DocNo(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TglSampaiDari(TBL, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_Remarks(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns(2).DefaultCellStyle = ij
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
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryMP As Double, _
    ByVal vUserCode As String, Optional ByVal vStrictEdit As Boolean = True, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim CreatedBy As String = ""
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '0. jika strict edit maka cek dulu siapa yang membuat
            If vStrictEdit = True Then
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimaryMP)) = True Then
                    CreatedBy = ""
                Else
                    CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryMP)
                End If
                If CreatedBy = vUserCode Then
                Else
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_MaterialPlanningDashboard_MP_OverrideEdit") = False Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            If Me.IsMPClosed(vPrimaryMP) Then
                If TampilkanPesan = True Then
                    MsgBox("This MP is closed, you can't edit this data.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryMP As Double, ByVal vUserCode As String, _
    Optional ByVal vStrictEdit As Boolean = True, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'Suatu transaksi do tidak bisa diedit jika:
            '0. jika strict edit maka harus dicek siapa yang buat
            If vStrictEdit = True Then
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimaryMP)) = True Then
                    CreatedBy = ""
                Else
                    CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryMP)
                End If
                If CreatedBy = vUserCode Then
                Else
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_MaterialPlanningDashboard_MP_OverrideDelete") = False Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            If Me.GetAuthLevelNeeded(vPrimaryMP) > 0 Then
                If Me.GetAuthLevelCurrent(vPrimaryMP) = Me.GetAuthLevelNeeded(vPrimaryMP) Then
                    MsgBox("This MP is already authorized, you can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, _
    ByVal vDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable, _
    ByVal vPrimaryMP As Double) As Boolean
        Try
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimary(Me.mNoTransaksi) <> Me.mPrimary Then
                            Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                        End If
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimary(Me.mNoTransaksi) <> Me.mPrimary Then
                            MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            '2. Jika tidak ada item
            If vDT.Rows.Count = 0 Then
                MsgBox("No item detected, process cancelled.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoMP As String) As Double
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_getPrimary(vNoMP).HasValue = True Then
                Return DTA.ScalarQuery_getPrimary(vNoMP).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable, ByVal vUserCode As String, _
    ByVal ij As FrmTrans_MaterialPlanningDashboard, _
    ByVal vDT_ItemGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable, _
    ByVal vDT_ItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable, _
    ByVal vDT_ItemPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, _
    ByVal vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, _
    ByVal vDT_ItemSettingLocation As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, _
    ByVal vDT_ItemQty As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable, _
    ByVal vDT_ItemMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable, _
    ByVal vDT_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable, _
    ByVal vDT_ItemFollowUp As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupDataTable, _
    ByVal vDT_ItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable, _
    ByVal vDT_ItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            Core_Module.SetConstring_DTAAntrian(Me.DTAInOUT)
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiMP_SimpanBaru, vUserCode)
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
                            retval = Me.SimpanBaru1(vDT, vUserCode, ij, vDT_ItemGIT, vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                            vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, vDT_ItemFollowUp, vDT_ItemSOBDelivery, _
                            vDT_ItemPrOBDelivery)

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

    Public Overridable Function SimpanBaru1(ByVal vDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable, ByVal vUserCode As String, _
    ByVal ij As FrmTrans_MaterialPlanningDashboard, _
    ByVal vDT_ItemGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable, _
    ByVal vDT_ItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable, _
    ByVal vDT_ItemPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, _
    ByVal vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, _
    ByVal vDT_ItemSettingLocation As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, _
    ByVal vDT_ItemQty As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable, _
    ByVal vDT_ItemMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable, _
    ByVal vDT_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable, _
    ByVal vDT_ItemFollowUp As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupDataTable, _
    ByVal vDT_ItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable, _
    ByVal vDT_ItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim ObjItem As New ClsTransaksiMaterialPlanning_item
        Dim Urutan As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, vDT, Me.mPrimary) = True Then
                'insert query
                Urutan = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, _
                Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP, Me.mTanggalTransaksi)

                Dim NeedToAuth As Integer = 0
                Dim objUser As New Core_Module.ClsUser
                If objUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_MaterialPlanningDashboard_MP_AuthorizeRequired") Then
                    NeedToAuth = 3
                End If
                If DTA.InsertQuery(Me.mNoTransaksi, Urutan, Me.mTanggalTransaksi, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now, Me.mRefDate, NeedToAuth) > 0 Then

                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_getPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = Urutan
                    Me.mTimestamp = DTA.ScalarQuery_gettimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item baru
                    ObjItem.Simpan(Me.mPrimary, vDT, vDT_ItemGIT, vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                    vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, vDT_ItemFollowUp, vDT_ItemSOBDelivery, _
                    vDT_ItemPrOBDelivery)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    Return Me.SimpanBaru1(vDT, vUserCode, ij, vDT_ItemGIT, vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                    vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, vDT_ItemFollowUp, vDT_ItemSOBDelivery, _
                    vDT_ItemPrOBDelivery)
                Else
                    MsgBox("Document Number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable, ByVal vUserCode As String, _
    ByVal ij As FrmTrans_MaterialPlanningDashboard, _
    ByVal vDT_ItemGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable, _
    ByVal vDT_ItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable, _
    ByVal vDT_ItemPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, _
    ByVal vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, _
    ByVal vDT_ItemSettingLocation As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, _
    ByVal vDT_ItemQty As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable, _
    ByVal vDT_ItemMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable, _
    ByVal vDT_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable, _
    ByVal vDT_ItemFollowUp As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupDataTable, _
    ByVal vDT_ItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable, _
    ByVal vDT_ItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            'ambil no antrian
            Core_Module.SetConstring_DTAAntrian(Me.DTAInOUT)
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiMP_SimpanPerubahan, vUserCode)

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
                            retval = Me.SimpanPerubahan1(vDT, vUserCode, ij, vDT_ItemGIT, vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                            vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, vDT_ItemFollowUp, vDT_ItemSOBDelivery, _
                            vDT_ItemPrOBDelivery)

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

    Public Overridable Function SimpanPerubahan1(ByVal vDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable, ByVal vUserCode As String, _
    ByVal ij As FrmTrans_MaterialPlanningDashboard, _
    ByVal vDT_ItemGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable, _
    ByVal vDT_ItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable, _
    ByVal vDT_ItemPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, _
    ByVal vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, _
    ByVal vDT_ItemSettingLocation As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, _
    ByVal vDT_ItemQty As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable, _
    ByVal vDT_ItemMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable, _
    ByVal vDT_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable, _
    ByVal vDT_ItemFollowUp As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupDataTable, _
    ByVal vDT_ItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable, _
    ByVal vDT_ItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim ObjItem As New ClsTransaksiMaterialPlanning_item
        Dim Urutan As Double = -1
        Dim ObjEmail As New E_mail_Module.ClsTransEmail

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False, vDT, Me.mPrimary) = True Then
                Urutan = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, _
                Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP, Me.mTanggalTransaksi)

                If DTA.Update(Me.mNoTransaksi, Urutan, Me.mTanggalTransaksi, Me.mRemarks, vUserCode, Now, Now, Me.mRefDate, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = Urutan
                    Me.mTimestamp = DTA.ScalarQuery_gettimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(Me.mPrimary, vDT, vDT_ItemGIT, vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                    vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, vDT_ItemFollowUp, vDT_ItemSOBDelivery, _
                    vDT_ItemPrOBDelivery)

                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    Return Me.SimpanPerubahan1(vDT, vUserCode, ij, vDT_ItemGIT, vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                    vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, vDT_ItemFollowUp, vDT_ItemSOBDelivery, _
                    vDT_ItemPrOBDelivery)
                Else
                    MsgBox("Document number duplication, please change document number", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjEmail = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal UserCode As String, _
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
            Core_Module.SetConstring_DTAAntrian(Me.DTAInOUT)
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiMP_Hapus, UserCode)

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
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, UserCode, TampilkanPesan)
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

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, ByVal UserCode As String, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimaryTransaksi, UserCode, TampilkanPesan) = True Then
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

    Public Sub ResetObject(ByVal vDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable)
        Try
            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now
            Me.mNoTransaksi = ""
            Me.mNoUrut = -1
            Me.mPrimary = -1
            Me.mRemarks = ""
            Me.mTanggalTransaksi = Now
            Me.mTimestamp = Now
            Me.mRefDate = #1/1/1950#

            vDT.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetTimestamp(ByVal vPrimaryMP As Double) As Date
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_gettimestamp(vPrimaryMP).HasValue = True Then
                Return DTA.ScalarQuery_gettimestamp(vPrimaryMP).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDocumentNo(ByVal vPrimaryMP As Double) As String
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDocNo(vPrimaryMP)) = False Then
                Return DTA.ScalarQuery_GetDocNo(vPrimaryMP)
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

    ' Public Function ExplainHistoricalPrice(ByVal vPrimaryProduct As Double, ByVal vDGV As DataGridView, _
    'ByVal ObjCD As Core_Module.ClsCekData, ByVal vRefDate As Date, ByVal vFrm As FrmTransaksiPO_HistoricalPrice) As Double
    '     Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.DataTable5TableAdapter
    '     Dim row As DseTransaksiMaterialPlanning2.DataTable5Row
    '     Dim rowDGV As DataGridViewRow
    '     Dim TBL As New DseTransaksiMaterialPlanning2.DataTable5DataTable

    '     Try
    '         Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '         DTA.Connection.ConnectionString = Core_Module.GetConstring

    '         If vRefDate = #1/1/1950# Then
    '             DTA.Fill(TBL, vPrimaryProduct)
    '         Else
    '             DTA.FillBy_RefDate(TBL, vPrimaryProduct, vRefDate)
    '         End If

    '         vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '         vDGV.Rows.Clear()

    '         For Each row In TBL.Rows
    '             vDGV.Rows.Add()
    '             rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)

    '             With row
    '                 rowDGV.Cells(vFrm.ColCurrency.Index).Value = .Currency_Code
    '                 rowDGV.Cells(vFrm.ColDate.Index).Value = .Tanggal_Transaksi
    '                 rowDGV.Cells(vFrm.ColDiscount.Index).Value = ObjCD.DigitGrouping(.Discount, 2)
    '                 rowDGV.Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Nett_Price, .Default_Precision)
    '                 rowDGV.Cells(vFrm.ColOrderQty.Index).Value = ObjCD.DigitGrouping(.PO_Qty, .Decimal_Product)
    '                 rowDGV.Cells(vFrm.ColPONumber.Index).Value = .PO_Number
    '                 rowDGV.Cells(vFrm.ColSupplier.Index).Value = .Supplier_Name
    '                 rowDGV.Cells(vFrm.ColUnit.Index).Value = .Unit_Code
    '                 rowDGV.Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.Unit_Price, .Default_Precision)
    '             End With
    '         Next

    '         vFrm.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
    '         vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '     Catch ex As Exception
    '         MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '     Finally
    '         DTA.Dispose()
    '         TBL.Dispose()
    '     End Try
    ' End Function

    Public Sub Cetak(ByVal vPrimaryTrans As Double, ByVal vRequiredOnly As Boolean, _
    ByVal vOrderPointFilter As Date, ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal DS As DseTransaksiMaterialPlanning2, ByVal vReport As CrpTrans_MaterialPlanning, ByVal vGrouping As String, ByVal vGroupingFilter As String, _
    ByVal vSOBOnly As Boolean, ByVal vPrOBOnly As Boolean, ByVal vFollowUpOnly As Boolean, ByVal vShowCodeNo As Boolean, ByVal vShowDesc As Boolean, _
    ByVal vGrouping2 As FrmTrans_MaterialPlanningDashboard_Print.EnumGroupingCetakan, ByVal vReport2 As CrpTrans_MaterialPlanning2, ByVal vShowLegend As Boolean, _
    ByVal vGoods As Boolean, ByVal vService As Boolean)
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.TBlCetakTableAdapter
        Dim row, row1 As DseTransaksiMaterialPlanning2.TBlCetakRow
        Dim TBLDummy As New DseTransaksiMaterialPlanning2.TBlCetakDataTable
        Dim i As Double = 0
        Dim vDate As Date = Today
        Dim vNoTrans As String = ""
        Dim Lanjut As Boolean
        Dim TipeProductGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim TipeProductService As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'atur sort
            If vGoods = False Then
                'service only
                TipeProductGoods = Core_Module.clsMasterProduct.EnumTipeProduct.Services
            ElseIf vService = False Then
                'goods only
                TipeProductService = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
            End If

            If vShowCodeNo = True Then
                'berdasarkan code no
                DTA.Fill(TBLDummy, vPrimaryTrans, TipeProductGoods, TipeProductService)
            ElseIf vShowCodeNo = False Then
                DTA.FillBy_DescSort(TBLDummy, vPrimaryTrans, TipeProductGoods, TipeProductService)
            End If

            i = 0
            If vOrderPointFilter <> #1/1/1950# Then
                While i < TBLDummy.Rows.Count
                    row = TBLDummy.Rows(i)
                    If row.order_point <= vOrderPointFilter Then
                        i += 1
                    Else
                        TBLDummy.Rows.Remove(row)
                    End If
                End While
            End If
            i = 0
            If vRequiredOnly = True Then
                While i < TBLDummy.Rows.Count
                    row = TBLDummy.Rows(i)
                    If row.qty_required > 0 Then
                        i += 1
                    Else
                        TBLDummy.Rows.Remove(row)
                    End If
                End While
            End If

            If vGroupingFilter <> "" Then
                i = 0
                While i < TBLDummy.Rows.Count
                    row = TBLDummy.Rows(i)

                    If vGrouping = "Category" Then
                        If row.Category_Code <> vGroupingFilter Then
                            TBLDummy.Rows.Remove(row)
                        Else
                            i += 1
                            row.Order_Strategy_Str = "Category: " + row.Category_Desc
                        End If
                    ElseIf vGrouping = "Product Group" Then
                        If row.Product_Group_Code <> vGroupingFilter Then
                            TBLDummy.Rows.Remove(row)
                        Else
                            i += 1
                            row.Order_Strategy_Str = "Group: " + row.Product_Group_Desc
                        End If
                    ElseIf vGrouping = "Brand" Then
                        If row.Brand_Code <> vGroupingFilter Then
                            TBLDummy.Rows.Remove(row)
                        Else
                            i += 1
                            row.Order_Strategy_Str = "Brand: " + row.Brand_Desc
                        End If
                    ElseIf vGrouping = "Sub Brand" Then
                        If row.Sub_Brand_Code <> vGroupingFilter Then
                            TBLDummy.Rows.Remove(row)
                        Else
                            i += 1
                            row.Order_Strategy_Str = "Sub Brand: " + row.Sub_Brand_Desc
                        End If
                    End If
                End While
            End If

            DS.TBlCetak.Rows.Clear()
            For Each row In TBLDummy.Rows
                Lanjut = True
                If vSOBOnly = True Then
                    If row.qty_sob <= 0 Then
                        Lanjut = False
                    End If
                End If
                If Lanjut = True Then
                    If vPrOBOnly = True Then
                        If row.qty_prob <= 0 Then
                            Lanjut = False
                        End If
                    End If
                End If
                If Lanjut = True Then
                    If vFollowUpOnly = True Then
                        If row.modify = 0 Then
                            Lanjut = False
                        End If
                    End If
                End If

                If Lanjut = True Then

                    If row.procurement_strategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order Then
                        row.Order_Strategy_Str = "Build To Order"
                    ElseIf row.procurement_strategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                        row.Order_Strategy_Str = "Build To Stock"
                    Else
                        'kalo ada strategi lain yang digunakan, maka salah
                        MsgBox("Logic error, please contact your vendor.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    'If row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty Then
                    '    row.Order_Strategy_Str = "Maximum Order Qty"
                    'ElseIf row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Minimum_Order_Qty Then
                    '    row.Order_Strategy_Str = "Minimum Order Qty"
                    'ElseIf row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty Then
                    '    row.Order_Strategy_Str = "Optimum Order Qty"
                    'ElseIf row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Max_Qty Then
                    '    row.Order_Strategy_Str = "Maximum Qty"
                    'ElseIf row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty Then
                    '    row.Order_Strategy_Str = "Minimum Qty"
                    'ElseIf row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Optimum_Qty Then
                    '    row.Order_Strategy_Str = "Optimum Qty"
                    'ElseIf row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Safety_Stock Then
                    '    row.Order_Strategy_Str = "Safety Stock Qty"
                    'End If

                    If vShowCodeNo = True And vShowDesc = False Then
                        row.Product_Code = row.Product_Code + "(" + row.Unit_Code + ")"
                    ElseIf vShowCodeNo = True And vShowDesc = True Then
                        row.Product_Code = row.Product_Code + "(" + row.Unit_Code + ")" + vbCrLf + row.Description
                    ElseIf vShowCodeNo = False And vShowDesc = True Then
                        row.Product_Code = row.Description + "(" + row.Unit_Code + ")"
                    End If
                    row.Qty_Free_Stock = row.qty_on_hand + row.qty_pob + row.qty_git - row.qty_sob - row.qty_prob - row.qty_pending_do - row.qty_inquiry

                    row1 = DS.TBlCetak.NewTBlCetakRow
                    For ij As Integer = 0 To DS.TBlCetak.Columns.Count - 1
                        row1.Item(ij) = row.Item(ij)
                    Next
                    DS.TBlCetak.Rows.Add(row1)
                End If
            Next

            If vGrouping2 = FrmTrans_MaterialPlanningDashboard_Print.EnumGroupingCetakan.NoGrouping Then
                vReport.SetDataSource(DS)
                vReport.SetParameterValue(vReport.Parameter_Date.ParameterFieldName, Me.GetTanggalMP(vPrimaryTrans))
                vReport.SetParameterValue(vReport.Parameter_Document_No.ParameterFieldName, Me.GetDocumentNo(vPrimaryTrans))
                If vShowLegend = True Then
                    vReport.SetParameterValue(vReport.Parameter_show_legend.ParameterFieldName, 1)
                Else
                    vReport.SetParameterValue(vReport.Parameter_show_legend.ParameterFieldName, 0)
                End If
                vCrv.ReportSource = vReport
                vCrv.Show()
            Else
                If vGrouping2 = FrmTrans_MaterialPlanningDashboard_Print.EnumGroupingCetakan.GroupByBrand Then
                    For Each row1 In DS.TBlCetak.Rows
                        row1.Group2 = row1.Brand_Code + " - " + row1.Brand_Desc
                    Next
                ElseIf vGrouping2 = FrmTrans_MaterialPlanningDashboard_Print.EnumGroupingCetakan.GroupByCategory Then
                    For Each row1 In DS.TBlCetak.Rows
                        row1.Group2 = row1.Category_Code + " - " + row1.Category_Desc
                    Next
                ElseIf vGrouping2 = FrmTrans_MaterialPlanningDashboard_Print.EnumGroupingCetakan.GroupByProductGroup Then
                    For Each row1 In DS.TBlCetak.Rows
                        row1.Group2 = row1.Product_Group_Code + " - " + row1.Product_Group_Desc
                    Next
                ElseIf vGrouping2 = FrmTrans_MaterialPlanningDashboard_Print.EnumGroupingCetakan.GroupBySubBrand Then
                    For Each row1 In DS.TBlCetak.Rows
                        row1.Group2 = row1.Sub_Brand_Code + " - " + row1.Sub_Brand_Desc
                    Next
                End If
                vReport2.SetDataSource(DS)
                vReport2.SetParameterValue(vReport2.Parameter_Date.ParameterFieldName, Me.GetTanggalMP(vPrimaryTrans))
                If vShowLegend = True Then
                    vReport2.SetParameterValue(vReport2.Parameter_show_legend.ParameterFieldName, 1)
                Else
                    vReport2.SetParameterValue(vReport2.Parameter_show_legend.ParameterFieldName, 0)
                End If
                vReport2.SetParameterValue(vReport2.Parameter_Document_No.ParameterFieldName, Me.GetDocumentNo(vPrimaryTrans))
                vCrv.ReportSource = vReport2
                vCrv.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBLDummy.Dispose()
        End Try
    End Sub

    Public Sub New()
        Core_Module.SetConstring_DTAAntrian(Me.DTAInOUT)
    End Sub

    'rosadin-r8
    Public Function GetAuthLevelNeeded(ByVal vPrimaryMP As Double) As Integer
        Dim dta As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim val As New Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            val = dta.ScalarQuery_GetAuthLvNeeded(vPrimaryMP)
            If val.HasValue Then
                Return val.Value
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
        End Try
    End Function
    Public Function GetAuthLevelCurrent(ByVal vPrimaryMP As Double) As Integer
        Dim dta As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim val As New Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            val = dta.ScalarQuery_GetCurrentAuthLv(vPrimaryMP)
            If val.HasValue Then
                Return val.Value
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function
    Public Function CloseMP(ByVal vPrimaryMP As Double, ByVal vClose As Boolean) As Boolean
        Dim dta As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim Close As Integer = 0
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If vClose Then
                Close = 1
            End If
            If dta.UpdateQuery_IsClosed(Close, vPrimaryMP) > 0 Then
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
    Public Function IsMPClosed(ByVal vPrimaryMP As Double) As Boolean
        Dim dta As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim val As New Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            val = dta.ScalarQuery_GetCloseStatus(vPrimaryMP)
            If val.HasValue Then
                If val.Value = 0 Then
                    Return False
                ElseIf val.Value = 1 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
    Public Function GetCreatedBy(ByVal vprimary As Double) As String
        Dim d As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim val As String = String.Empty
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            val = d.ScalarQuery_GetCreatedBy(vprimary)
            If String.IsNullOrEmpty(val) = False Then
                Return val
            End If
            Return String.Empty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return String.Empty
        End Try
    End Function
    Public Function GetOpenMPCount(ByVal vUserCode As String) As Integer
        Dim d As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim val As String = String.Empty
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            val = d.ScalarQuery_GetOpenMPCount(vUserCode)
            If String.IsNullOrEmpty(val) = False Then
                Return val
            End If
            Return String.Empty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return String.Empty
        End Try
    End Function

    Public Function GetOpenMPDocNo(ByVal vUserCode As String) As String
        Dim dta As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planningTableAdapter
        Dim val As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            val = dta.ScalarQuery_GetOpenMPDocumentNo(vUserCode)
            If Not String.IsNullOrEmpty(val) Then
                Return val
            End If
            Return String.Empty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return String.Empty
        End Try
    End Function
End Class
