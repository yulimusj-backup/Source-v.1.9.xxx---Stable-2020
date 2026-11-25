Option Explicit On
Option Strict Off

Public Class ClsTrans_ManufacturingOrder
    Dim mPrimary As Double = -1
    Dim mPrimaryUser_Auth As Double = -1
    Dim mAuth_Time As Date = #1/1/1950#
    Dim mNoTransaksi As String = ""
    Dim mUrutan As Double = 0
    Dim mTanggalTransaksi As Date = Now
    Public mIsReady_ForScheduling As EnumPilihSchedule
    Dim mRemarks As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Public mStatusWO As EnumPilihStatusWO
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Dim ObjMO_Item As New ClsTrans_ManufacturingOrder_Item

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Public Enum EnumPilihSchedule
        Schedule_tdk_Aktif
        Schedule_Aktif
    End Enum

    Public Enum EnumPilihStatusWO
        Free
        Planned
        Scheduled
        On_Progress
        Finish
    End Enum

    Public Enum PilihFormatNumber
        MO
        JO
        WO
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property Urutan() As Double
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Double)
            mUrutan = value
        End Set
    End Property

    Public Property PrimaryUser_Auth() As Double
        Get
            Return mPrimaryUser_Auth
        End Get
        Set(ByVal value As Double)
            mPrimaryUser_Auth = value
        End Set
    End Property

    Public Property Auth_Time() As Date
        Get
            Return mAuth_Time
        End Get
        Set(ByVal value As Date)
            mAuth_Time = value
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

    Public Property TanggalTransaksi() As Date
        Get
            Return mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property

    Public Property IsReady_ForScheduling() As EnumPilihSchedule
        Get
            Return mIsReady_ForScheduling
        End Get
        Set(ByVal value As EnumPilihSchedule)
            mIsReady_ForScheduling = value
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

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim Retval As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimary(vNoTransaksi)
            If Retval.HasValue = True Then
                If Retval.Value > -1 Then
                    Return Retval.Value
                Else
                    Return -1
                End If
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim Retval As New Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            If Retval.HasValue = True Then
                Return Retval.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetNoTransaksi(vPrimary)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Retval
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vDGVItemJO As DataGridView, ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, ByVal vUserCode As String, ByVal Frm As FrmTrans_ManufacturingOrder, _
    ByVal vTblItemRMAdd As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable, _
    ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable, _
    ByVal vTblItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try

            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.FrmTrans_ManufacturingOrder_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vDGVItemJO, vTblItemBOM, vTblProsesOrder, vTblItemWO, vTblItemRM, vUserCode, Frm, vTblItemRMAdd, vTblItemWIP, vTblItemDelivery)
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

    Private Function SimpanBaru1(ByVal vDGVItemJO As DataGridView, ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, ByVal vUserCode As String, ByVal Frm As FrmTrans_ManufacturingOrder, _
    ByVal vTblItemRMAdd As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable, _
    ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable, _
    ByVal vTblItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryDataTable) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim TipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO
        Dim DTA1 As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_itemTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim rowWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(vDGVItemJO, vTblItemBOM, vTblProsesOrder, vTblItemWO, vTblItemRM, True, Frm, vTblItemRMAdd, vTblItemWIP, vTblItemDelivery, vUserCode) = True Then
                Me.mUrutan = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO, Me.mTanggalTransaksi)

                If DTA.Insert(Me.mNoTransaksi, Me.mUrutan, Me.mTanggalTransaksi, Me.mIsReady_ForScheduling, Me.mRemarks, _
                 vUserCode, Now, vUserCode, Now, Now, Me.mPrimaryUser_Auth, Me.mAuth_Time) > 0 Then
                    Me.mPrimary = Me.GetPrimary(Me.mNoTransaksi)
                    Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'untuk simpan baru, balance mp pada item wo = target produksi
                    For Each rowWO In vTblItemWO
                        rowWO.Qty_Balance_MP = rowWO.Balance_Result  '- rowWO.Qty_WIP
                    Next

                    Me.BetulkanPenomoranJOdanWO(vDGVItemJO, vTblItemWO, Frm)

                    'simpan all item
                    Me.ObjMO_Item.SimpanAllItem(Me.mPrimary, vDGVItemJO, vTblItemBOM, vTblProsesOrder, _
                    vTblItemWO, vTblItemRM, vUserCode, Frm, vTblItemRMAdd, vTblItemWIP, vTblItemDelivery)

                    'update qty target
                    For Each rowDGV In vDGVItemJO.Rows
                        If rowDGV.Cells(Frm.ColItem_CodeNo.Index).Value <> "" Then
                            Me.ObjMO_Item.SetBalanceMP_SimpanBaru(rowDGV.Cells(Frm.ColItem_Identifier.Index).Value)
                            DTA1.UpdateQuery_SetQtyBalance(CDec(rowDGV.Cells(Frm.ColItem_QtyTarget.Index).Value), rowDGV.Cells(Frm.ColItem_Identifier.Index).Value)
                        End If
                    Next

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi, PilihFormatNumber.MO, -1, Frm)
                    Return Me.SimpanBaru1(vDGVItemJO, vTblItemBOM, vTblProsesOrder, vTblItemWO, vTblItemRM, vUserCode, Frm, vTblItemRMAdd, vTblItemWIP, vTblItemDelivery)
                Else
                    MsgBox("MO Number duplication, please change MO number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub BetulkanPenomoranJOdanWO(ByVal vDGVItemJO As DataGridView, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal Frm As FrmTrans_ManufacturingOrder)
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim TblItemWO_Empty As New DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
        Try
            ''betulin no JO
            vDGVItemJO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItemJO.AllowUserToAddRows = False
            vDGVItemJO.Sort(Frm.ColItem_Urutan, System.ComponentModel.ListSortDirection.Ascending)
            If Me.ApakahDataItemJOsudahAda(vDGVItemJO.Item(Frm.ColItem_JONumber.Index, 0).Value) Then
                Dim MaximumUrutanDgv As Double = 0
                For Each rowDgv As DataGridViewRow In vDGVItemJO.Rows
                    rowDgv.Cells(Frm.ColItem_JONumber.Index).Value = Me.GetNoBaru1(Frm.DtpTglTransaksi.Value.Date, PilihFormatNumber.JO, MaximumUrutanDgv, Frm, vDGVItemJO.Rows.IndexOf(rowDgv))
                    MaximumUrutanDgv = rowDgv.Cells(Frm.ColItem_Urutan.Index).Value
                Next
            End If
            vDGVItemJO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            ''betulin no WO
            vTblItemWO.DefaultView.RowFilter = ""
            vTblItemWO.DefaultView.Sort = "[Urutan]"
            rowItemWO = vTblItemWO.DefaultView(0).Row
            If Me.ApakahDataItemWOsudahAda(rowItemWO.WO_Number, TblItemWO_Empty) Then
                Dim MaximumUrutanWO As Double = 0
                For Each DV As Data.DataRowView In vTblItemWO.DefaultView
                    rowItemWO = DV.Row
                    rowItemWO.WO_Number = Me.GetNoBaru1(Frm.DtpTglTransaksi.Value.Date, ClsTrans_ManufacturingOrder.PilihFormatNumber.WO, MaximumUrutanWO, Frm, rowItemWO.Urutan, TblItemWO_Empty, False)
                    MaximumUrutanWO = rowItemWO.Urutan
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function SimpanPerubahan(ByVal vDGVItemJO As DataGridView, ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, ByVal vUserCode As String, ByVal Frm As FrmTrans_ManufacturingOrder, _
    ByVal vTblItemRMAdd As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable, _
    ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable, _
    ByVal vItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try

            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.FrmTrans_ManufacturingOrder_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vDGVItemJO, vTblItemBOM, vTblProsesOrder, vTblItemWO, vTblItemRM, vUserCode, Frm, vTblItemRMAdd, vTblItemWIP, vItemDelivery)
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
            'ObjAntri = Nothing
            SW = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal vDGVItemJO As DataGridView, ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, ByVal vUserCode As String, ByVal Frm As FrmTrans_ManufacturingOrder, _
    ByVal vTblItemRMAdd As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable, _
    ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable, _
    ByVal vItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryDataTable) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim TipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO
        Dim ObjItem As New ClsTrans_ManufacturingOrder_Item
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(vDGVItemJO, vTblItemBOM, vTblProsesOrder, vTblItemWO, vTblItemRM, False, Frm, vTblItemRMAdd, vTblItemWIP, vItemDelivery, vUserCode) = True Then
                Me.Urutan = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, TipeTransaksi, Me.mTanggalTransaksi)

                If DTA.Update(Me.mNoTransaksi, Me.mUrutan, Me.mTanggalTransaksi, Me.mIsReady_ForScheduling, Me.mRemarks, _
                 vUserCode, Now, Now, Me.mPrimaryUser_Auth, Me.mAuth_Time, Me.mPrimary) > 0 Then
                    Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'update balance mp untuk setiap item wo yang sebelumnya sudah ada
                    Me.ObjMO_Item.SetBalanceMP_SimpanPerubahan(Me.mPrimary, vTblItemWO, vTblItemRM, vTblItemWIP)

                    'hapus semua item lalu simpan semua item 
                    Me.ObjMO_Item.UpdateQtySpareMP(Me.mPrimary, vTblItemRM)
                    Me.ObjMO_Item.HapusAllItem(Me.mPrimary)
                    Me.ObjMO_Item.SimpanAllItem(Me.mPrimary, vDGVItemJO, vTblItemBOM, vTblProsesOrder, _
                    vTblItemWO, vTblItemRM, vUserCode, Frm, vTblItemRMAdd, vTblItemWIP, vItemDelivery)


                    'update balance qty item mo
                    For Each rowDGV In vDGVItemJO.Rows
                        If rowDGV.Cells(Frm.ColItem_CodeNo.Index).Value <> "" Then
                            Me.ObjMO_Item.SetMOBalance(rowDGV.Cells(Frm.ColItem_Identifier.Index).Value)
                        End If
                    Next

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi, PilihFormatNumber.MO, -1, Frm)
                    Return Me.SimpanPerubahan1(vDGVItemJO, vTblItemBOM, vTblProsesOrder, vTblItemWO, vTblItemRM, vUserCode, Frm, vTblItemRMAdd, vTblItemWIP, vItemDelivery)
                Else
                    MsgBox("MO Number duplication, please change MO number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, _
    ByVal Frm As FrmTrans_ManufacturingOrder, ByVal vNoTransaksi As String, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try

            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiStockBundling_Hapus, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, vUserCode, Frm, vNoTransaksi, vTblItemWO)
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

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, _
    ByVal Frm As FrmTrans_ManufacturingOrder, ByVal vNoTransaksi As String, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimaryTransaksi, vUserCode, vNoTransaksi, vTblItemWO, True) = True Then
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vDGVItemJO As DataGridView, ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, _
    ByVal vBaru As Boolean, ByVal vFrm As FrmTrans_ManufacturingOrder, _
    ByVal vTblItemRMAdd As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable, _
    ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable, _
    ByVal vDTItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryDataTable, ByVal vUserCode As String) As Boolean
        Try
            'tidak boleh double
            If vBaru = True Then
                If Me.Apakahsudahada(mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi, ClsTrans_ManufacturingOrder.PilihFormatNumber.MO, -1, vFrm)
                    ' MsgBox("Duplication MO Number, please change your MO Number", _
                    'MsgBoxStyle.Information, "Information")
                    ' Return False
                End If
            Else
                Dim productGroupLama As String = Me.GetNoTransaksi(mPrimary)
                If productGroupLama <> mNoTransaksi Then
                    If Me.Apakahsudahada(mNoTransaksi) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi, ClsTrans_ManufacturingOrder.PilihFormatNumber.MO, -1, vFrm)
                        'MsgBox("Duplication MO Number, please change your MO Number", _
                        ' MsgBoxStyle.Information, "Information")
                        'Return False
                    End If
                End If

                'data salah jika nilai timestamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim oldTimeStamp As Date = Me.GetTimestamp(mPrimary)
                If mTimestamp <> oldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.CariDetail(Me.mPrimary, vDGVItemJO, vTblItemBOM, vTblProsesOrder, vTblItemWO, vTblItemRM, vFrm, Me.ObjCD, vTblItemRMAdd, vTblItemWIP, vDTItemDelivery, vUserCode)
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vSalesStaff As String, ByVal vNoTransaksi As String, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        'Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim ListIdentifierWO As New List(Of String)
        Dim CB As String = ""
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cek untuk memastikan bila transaksi belum di hapus oleh user lain
            If Me.Apakahsudahada(vNoTransaksi) = False Then
                If vTampilkanPesan = True Then
                    MsgBox("The manufacturing order: " + vNoTransaksi + " has been deleted by another user, please choose another manufacturing order.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            CB = DTA.ScalarQuery_GetCreatedBy(vPrimary)
            If CB <> vSalesStaff Then
                If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTrans_ManufacturingOrder_HapusTransaksi") = False Then
                    If vTampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Else
                    'do nothing

                End If
            End If

            'jika sudah dijadwalkan
            If Me.IsScheduled(vPrimary) = True Then
                MsgBox("The following manufacturing order data is scheduled" & _
                ", you can’t delete this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh hapus bila mo sudah di buat bundling
            If ApakahSudahAdaBundling(vPrimary, vTblItemWO) = False Then
                MsgBox("The following manufacturing order data is reference by one or several bundling" & _
                ", you can’t delete this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh hapus bila sudah di authorize
            If GetReadyForSchedule(mPrimary) = EnumPilihSchedule.Schedule_Aktif Then
                MsgBox("Please cancel the authorization", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapusItemJO(ByVal vPrimary As Double, ByVal vIdentifierJO As String, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim RowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim ListIdentifierWO As New List(Of String)

        Try
            'jika sudah dijadwalkan, maka tidak bisa hapus jo
            If Me.IsScheduled(vPrimary, vIdentifierJO) = True Then
                Return False
            End If

            'cek bundling, bila wo tertentu yang akan di reorder sudah di bundling maka tidak bisa reorder
            vTblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
            If vTblItemWO.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblItemWO.DefaultView
                    RowItemWO = dv.Row
                    If ListIdentifierWO.Contains(RowItemWO.Identifier) = False Then
                        ListIdentifierWO.Add(RowItemWO.Identifier)
                    End If
                Next

                If ListIdentifierWO.Count > 0 Then
                    For Each IdentifierWO As String In ListIdentifierWO
                        If GetJmlBundling(IdentifierWO, vIdentifierJO) > 0 Then
                            Return False
                        End If
                    Next
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ApakahBolehHapusItemBOM(ByVal vIdentifierBOM As String, ByVal vIdentifierBOM_Parent As String, _
    ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable) As Boolean
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow

        Try

            If vIdentifierBOM <> "-1" Then
                vTblItemBOM.DefaultView.RowFilter = "[Identifier BOM] = '" & vIdentifierBOM.ToString _
                & "' AND [Identifier BOM Parent] = '" & vIdentifierBOM_Parent & "'"
            Else
                vTblItemBOM.DefaultView.RowFilter = "[Identifier BOM Parent] = '" & vIdentifierBOM_Parent & "'"
            End If

            If vTblItemBOM.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblItemBOM.DefaultView
                    rowItemBOM = dv.Row
                    If GetJmlBundling(rowItemBOM.Identifier_WO, rowItemBOM.Identifier_JO) > 0 Then
                        Return False
                    Else
                        If ApakahBolehHapusItemBOM("-1", rowItemBOM.Identifier_BOM, vTblItemBOM) = False Then
                            Return False
                        End If
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function CekChild(ByVal vIdentifierBOM As Double, ByVal vIdentifierBOM_Parent As Double, _
    ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable) As Boolean
        'Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow

        'Try
        '    vTblItemBOM.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO.ToString & "' AND [Primary Product] = " _
        '    & vPrimaryProduct.ToString & " AND [Primary Parent Product] = " & vPrimaryParentProduct
        '    If vTblItemBOM.DefaultView.Count > 0 Then
        '        For Each dv As Data.DataRowView In vTblItemBOM.DefaultView
        '            rowItemBOM = dv.Row
        '            If GetJmlBundling(rowItemBOM.Identifier_WO) > 0 Then
        '                Return False
        '            Else
        '                CekChild(vIdentifierJO, rowItemBOM.Primary_Product_RM, rowItemBOM.Primary_Product, vTblItemBOM)
        '            End If
        '        Next
        '    End If

        '    Return True
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Function

    Public Function ApakahBolehEditItemBOM(ByVal vIdentifierBOM As String, ByVal vIdentifierBOM_Parent As String, _
    ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable) As Boolean
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow

        Try
            If vIdentifierBOM <> "-1" Then
                vTblItemBOM.DefaultView.RowFilter = "[Identifier BOM] = '" & vIdentifierBOM.ToString _
                & "' AND [Identifier BOM Parent] = '" & vIdentifierBOM_Parent & "'"
            Else
                vTblItemBOM.DefaultView.RowFilter = "[Identifier BOM Parent] = '" & vIdentifierBOM_Parent & "'"
            End If

            If vTblItemBOM.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblItemBOM.DefaultView
                    rowItemBOM = dv.Row
                    If GetJmlBundling(rowItemBOM.Identifier_WO, rowItemBOM.Identifier_JO) > 0 Then
                        Return False
                    Else
                        If ApakahBolehHapusItemBOM("-1", rowItemBOM.Identifier_BOM, vTblItemBOM) = False Then
                            Return False
                        End If
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vSalesStaff As String, _
    ByVal vNoTransaksi As String, Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim CB As String = ""
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cek untuk memastikan bila transaksi belum di hapus oleh user lain
            If Me.GetNoTransaksi(vPrimary) = "" Then
                If vTampilkanPesan = True Then
                    MsgBox("The manufacturing order: " + vNoTransaksi + " has been deleted by another user, please choose another manufacturing order.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            CB = DTA.ScalarQuery_GetCreatedBy(vPrimary)
            If CB <> vSalesStaff Then
                If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTrans_ManufacturingOrder_EditTransaksi") = False Then
                    If vTampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                Else
                    'do nothing
                End If
            End If

            'tidak boleh edit jika mo sudah di buat bundling

            'jika sudah authorize hanya plan time dan additional yang boleh di edit

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vDGVItemJO As DataGridView, ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, ByVal vTblitemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, ByVal Frm As FrmTrans_ManufacturingOrder, ByVal ObjCD As Core_Module.ClsCekData, _
    ByVal vTblItemRMAdd As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable, _
    ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable, _
    ByVal vTBlItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryDataTable, _
    ByVal vUserCode As String, Optional ByVal vTblJO_Temp As DseTrans_ManufacturingOrder.JO_TempDataTable = Nothing) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try

            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.FrmTrans_ManufacturingOrder_CariDetail, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.CariDetail1(vPrimary, vDGVItemJO, vTblItemBOM, vTblProsesOrder, _
                            vTblitemWO, vTblItemRM, Frm, ObjCD, vTblItemRMAdd, vTblItemWIP, vTBlItemDelivery, vUserCode, vTblJO_Temp)
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

    Public Function CariDetail1(ByVal vPrimary As Double, ByVal vDGVItemJO As DataGridView, ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, ByVal vTblitemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, ByVal Frm As FrmTrans_ManufacturingOrder, ByVal ObjCD As Core_Module.ClsCekData, _
    ByVal vTblItemRMAdd As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable, _
    ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable, _
    ByVal vTBlItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryDataTable, _
    ByVal vUserCode As String, Optional ByVal vTblJO_Temp As DseTrans_ManufacturingOrder.JO_TempDataTable = Nothing) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim TBL As New DseTrans_ManufacturingOrder.tbltrans_moDataTable
        Dim row As DseTrans_ManufacturingOrder.tbltrans_moRow
        Dim DtaItemBOM As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_bomTableAdapter
        Dim DtaProsesOrder As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_processorderTableAdapter
        Dim DtaItemWO As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim DtaItemRM As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim DtaItemRMAdd As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rm_addTableAdapter
        Dim DTAItemWIP As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wipTableAdapter
        Dim DTAItemDelivery As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_itemdeliveryTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemWIP.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.FillBy_Primary(TBL, CDec(vPrimary))
            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseTrans_ManufacturingOrder.tbltrans_moRow)
                With row
                    Me.mPrimary = row.Primary
                    Me.TanggalTransaksi = row._Date
                    Me.mNoTransaksi = row.MO_Number
                    Me.mUrutan = row.Urutan
                    Me.mAuth_Time = row.Auth_Time
                    Me.mPrimaryUser_Auth = row.Primary_User_Auth
                    Me.mIsReady_ForScheduling = row.Is_Ready_for_Scheduling
                    Me.mRemarks = row.Remarks
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With

                DTAItemDelivery.Connection.ConnectionString = Core_Module.GetConstring
                DTAItemDelivery.FillBy_Primary(vTBlItemDelivery, vPrimary)

                'isi dgv item JO
                Me.ObjMO_Item.IsiDGVItemJO(vPrimary, vDGVItemJO, vTblJO_Temp, Frm, ObjCD, vTBlItemDelivery)

                'isi tbl item BOM
                DtaItemBOM.Connection.ConnectionString = Core_Module.GetConstring
                DtaItemBOM.FillBy_PrimaryMO(vTblItemBOM, vPrimary)

                'isi tbl item Proses Order
                DtaProsesOrder.Connection.ConnectionString = Core_Module.GetConstring
                DtaProsesOrder.FillBy_PrimaryMO(vTblProsesOrder, vPrimary)

                'isi tbl item Work Order
                DtaItemWO.Connection.ConnectionString = Core_Module.GetConstring
                DtaItemWO.FillBy_PrimaryMO(vTblitemWO, vPrimary)

                'isi tbl item Raw Material
                DtaItemRM.Connection.ConnectionString = Core_Module.GetConstring
                DtaItemRM.FillBy_PrimaryMO(vTblItemRM, vPrimary)

                'isi tbl item Raw Material Add
                DtaItemRMAdd.Connection.ConnectionString = Core_Module.GetConstring
                DtaItemRMAdd.Fill(vTblItemRMAdd, vPrimary)

                'isi tbl wip
                DTAItemWIP.Fill(vTblItemWIP, vPrimary)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
            DtaItemBOM.Dispose()
            DtaProsesOrder.Dispose()
            DtaItemWO.Dispose()
            DtaItemRM.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
         ByVal DGV As DataGridView, ByVal TBL As DseTrans_ManufacturingOrder.tbltrans_mo1DataTable, _
         ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
         Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
         Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean

        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDate Then
                QueryResult = DTA.FillBy_TglDariSampai(TBL, vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByWONumber Then
                QueryResult = DTA.FillBy_WONumberLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescriptionLIKE(TBL, Criteria + "%")
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Mo Number"
                DGV.Columns(2).HeaderText = "Date"
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(3).HeaderText = "Authorization"
                DGV.Columns(4).HeaderText = "Remarks"

                DGV.Columns(1).MinimumWidth = 80
                DGV.Columns(2).MinimumWidth = 70
                DGV.Columns(3).MinimumWidth = 80
                DGV.Columns(4).MinimumWidth = 80
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
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
            TBL.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       Core_Module.EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
       Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Double
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo1TableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                RetVal = DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByDate Then
                RetVal = DTA.ScalarQuery_GetJmlByTglDariSampai(vTanggalDari, vTanggalSampai)
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByRemarks Then
                RetVal = DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByWONumber Then
                RetVal = DTA.ScalarQuery_GetJmlByWONumber(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByCode Then
                RetVal = DTA.ScalarQuery_GetJmlByProductLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByDescription Then
                RetVal = DTA.ScalarQuery_GetJmlByDescriptionLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Apakahsudahada(ByVal vNoTransaksi As String) As Boolean
        Dim Retval As Nullable(Of Long)
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJml(vNoTransaksi)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetQtyRef(ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo1TableAdapter
        Dim Retval As New Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetQtyRef(vPrimaryProduct)
            If Retval.HasValue = True Then
                Return Retval.Value
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlBundling(ByVal vIdentifierWO As String, ByVal vIdentifierJO As String) As Integer
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim Retval As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJmlBundlingByIdentifierWO(vIdentifierWO, vIdentifierJO)
            If Retval.HasValue = True Then
                Return Retval.Value
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vIndexFormat As ClsTrans_ManufacturingOrder.PilihFormatNumber, _
    'ByVal Frm As FrmTrans_ManufacturingOrder, ByVal vRowIndex As Integer, ByVal vMaxUrutan As Integer, ByRef vMaxUrutanResult As Integer) As String
    '    Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
    '    Dim TanggalStart, TanggalEnd As Date
    '    Dim NoUrut As Double = 0
    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim RetVal As String = ""
    '    Dim TipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO
    '    Try
    '        If vIndexFormat = PilihFormatNumber.MO Then
    '            TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO
    '        Else
    '            TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.JO
    '        End If

    '        Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, TipeTransaksi, vTanggalTransaksi)

    '        If TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO Then
    '            If DTA.ScalarQuery_MAXUrutan(TanggalStart, TanggalEnd).HasValue = True Then
    '                NoUrut = CStr(DTA.ScalarQuery_MAXUrutan(TanggalStart, TanggalEnd).Value + 1)
    '                If NoUrut <= 0 Then
    '                    NoUrut = 1
    '                End If
    '            Else
    '                NoUrut = 1
    '            End If

    '        ElseIf TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.JO Then
    '            If vRowIndex = 0 Then
    '                If DTA.ScalarQuery_MAXUrutan_Item(TanggalStart, TanggalEnd).HasValue = True Then
    '                    NoUrut = DTA.ScalarQuery_MAXUrutan_Item(TanggalStart, TanggalEnd).Value + 1
    '                End If
    '            Else
    '                NoUrut = vMaxUrutan + 1
    '            End If
    '        End If

    '        vMaxUrutanResult = NoUrut
    '        RetVal = Me.ObjNumbering.GetNomorTransaksiBaru(vTanggalTransaksi, NoUrut, TipeTransaksi)
    '        If TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO Then
    '            While Me.Apakahsudahada(RetVal) = True
    '                NoUrut = NoUrut + 1
    '                RetVal = Me.ObjNumbering.GetNomorTransaksiBaru(vTanggalTransaksi, NoUrut, TipeTransaksi)
    '            End While
    '        ElseIf TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.JO Then
    '            While Me.ApakahDataItemsudahAda(NoUrut, Frm.DgvItem.Item(Frm.ColItem_Urutan.Index, vRowIndex).Value, Frm) = True
    '                'While Me.ApakahDataItemsudahAda(Frm.DgvItem.Item(Frm.ColItem_JONumber.Index, vRowIndex).Value) = True
    '                NoUrut = NoUrut + 1
    '                vMaxUrutanResult = NoUrut
    '                RetVal = Me.ObjNumbering.GetNomorTransaksiBaru(vTanggalTransaksi, NoUrut, TipeTransaksi)
    '                'RetVal = Me.GetNoBaru(vTanggalTransaksi, vIndexFormat, Frm, 0, vMaxUrutan, vMaxUrutanResult)
    '            End While
    '        End If

    '        Return RetVal
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    Finally
    '        DTA.Dispose()
    '        ObjCD = Nothing
    '    End Try
    'End Function

    Public Function GetNoBaru1(ByVal vTanggalTransaksi As Date, ByVal vIndexFormat As ClsTrans_ManufacturingOrder.PilihFormatNumber, _
    ByVal vMaxUrutan As Integer, Optional ByVal Frm As FrmTrans_ManufacturingOrder = Nothing, Optional ByRef vRowIndex As Integer = -1, _
    Optional ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable = Nothing, Optional ByVal vAlwaysRecheck As Boolean = True) As String
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As Double = 0
        Dim RetVal As String = ""
        Dim TipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO
        Dim Max_Urutan_DB As Double = 0

        Try
            If vIndexFormat = PilihFormatNumber.MO Then
                TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO
            ElseIf vIndexFormat = PilihFormatNumber.JO Then
                TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.JO
            ElseIf vIndexFormat = PilihFormatNumber.WO Then
                TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.WO
            End If

            Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, TipeTransaksi, vTanggalTransaksi)

            If TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO Then
                If DTA.ScalarQuery_MAXUrutan(TanggalStart, TanggalEnd).HasValue = True Then
                    NoUrut = CStr(DTA.ScalarQuery_MAXUrutan(TanggalStart, TanggalEnd).Value + 1)
                    If NoUrut <= 0 Then
                        NoUrut = 1
                    End If
                Else
                    NoUrut = 1
                End If
            ElseIf TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.JO Then
                Max_Urutan_DB = DTA.ScalarQuery_MAXUrutan_Item_JO(TanggalStart, TanggalEnd).Value

                If Max_Urutan_DB = 0 And vMaxUrutan = 0 Then
                    NoUrut = 1
                ElseIf Max_Urutan_DB = 0 And vMaxUrutan > 0 Then
                    NoUrut = vMaxUrutan + 1
                ElseIf Max_Urutan_DB > 0 And vMaxUrutan = 0 Then
                    NoUrut = Max_Urutan_DB + 1
                ElseIf Max_Urutan_DB > 0 And vMaxUrutan > 0 Then
                    If vMaxUrutan <= Max_Urutan_DB Then
                        While vMaxUrutan <= Max_Urutan_DB
                            vMaxUrutan += 1
                        End While
                        NoUrut = vMaxUrutan
                    Else
                        NoUrut = vMaxUrutan + 1
                    End If
                End If

                If vRowIndex <> -1 Then
                    Frm.DgvItem.Item(Frm.ColItem_Urutan.Index, vRowIndex).Value = NoUrut
                End If
            ElseIf TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.WO Then
                Max_Urutan_DB = DTA.ScalarQuery_MAXUrutan_Item_WO(TanggalStart, TanggalEnd).Value

                If Max_Urutan_DB = 0 And vMaxUrutan = 0 Then
                    NoUrut = 1
                ElseIf Max_Urutan_DB = 0 And vMaxUrutan > 0 Then
                    NoUrut = vMaxUrutan
                ElseIf Max_Urutan_DB > 0 And vMaxUrutan = 0 Then
                    NoUrut = Max_Urutan_DB + 1
                ElseIf Max_Urutan_DB > 0 And vMaxUrutan > 0 Then
                    If vMaxUrutan <= Max_Urutan_DB Then
                        While vMaxUrutan <= Max_Urutan_DB
                            vMaxUrutan += 1
                        End While
                        NoUrut = vMaxUrutan
                    Else
                        NoUrut = vMaxUrutan + 1
                    End If
                End If

                vRowIndex = NoUrut
            End If

            RetVal = Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, TipeTransaksi)

            If vAlwaysRecheck And TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO Then
                While Me.Apakahsudahada(RetVal) = True
                    NoUrut = NoUrut + 1
                    RetVal = Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, TipeTransaksi)
                End While
            ElseIf vAlwaysRecheck And TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.JO Then
                While Me.ApakahDataItemJOsudahAda(RetVal) = True
                    NoUrut = NoUrut + 1
                    RetVal = Me.GetNoBaru1(vTanggalTransaksi, PilihFormatNumber.JO, NoUrut, Frm)
                End While
            ElseIf vAlwaysRecheck And TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.WO Then
                While Me.ApakahDataItemWOsudahAda(RetVal, vTblItemWO) = True
                    NoUrut = NoUrut + 1
                    RetVal = Me.GetNoBaru1(vTanggalTransaksi, PilihFormatNumber.WO, NoUrut, Frm, vRowIndex, vTblItemWO)
                End While
            End If

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function ApakahDataItemsudahAda(ByVal Urutan As String, ByVal JoNumber As String, ByVal frm As FrmTrans_ManufacturingOrder) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim Retval As New Nullable(Of Long)
        Dim TanggalStart As Date
        Dim TanggalEnd As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.JO, frm.DtpTglTransaksi.Value)
            If TanggalEnd.Month = 12 And TanggalStart.Month = 1 Then
                Retval = DTA.ScalarQuery_GetJmlByUrutan(Urutan)
            Else
                Retval = DTA.ScalarQuery_GetJmlByJONumber(JoNumber)
            End If

            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahDataItemJOsudahAda(ByVal vJoNumber As String) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim Retval As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJmlByJONumber(vJoNumber)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahDataItemWOsudahAda(ByVal vWoNumber As String, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim Retval As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cek wo number yang sama ke datatable temporary
            vTblItemWO.DefaultView.RowFilter = "[WO Number] = '" & vWoNumber & "'"
            If vTblItemWO.DefaultView.Count > 0 Then
                Return True
            End If

            'cek wo number yang sama ke database
            Retval = DTA.ScalarQuery_GetJmlByWONumber(vWoNumber)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTypeAccProduct(ByVal vPrimaryProduct As Double) As Integer
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo1TableAdapter
        Dim Retval As New Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetTypeAccProduct(vPrimaryProduct)
            If Retval.HasValue = True Then
                Return Retval.Value
            Else
                Return -1
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlItemBOM(ByVal vPrimaryProduct As Double) As Integer
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo1TableAdapter
        Dim Retval As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJmlItemBom(vPrimaryProduct)
            If Retval.HasValue = True Then
                Return Retval.Value
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetMaxIdentifierJO_ProsesOrder(ByVal vPrimaryProduct As Double) As String
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.Process_Order_BeforeTableAdapter
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetMaxIdentifierJO(vPrimaryProduct)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetQtyProduce_InTbl(ByVal vPrimaryProduct As Double, ByVal vQtyTarget As Decimal, ByVal vIdentifierJO As String, _
    ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable) As Decimal
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim QtyProduce As Decimal = 0

        Try
            vTblItemBOM.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString & " AND [Identifier JO] = '" & vIdentifierJO & "'"
            If vTblItemBOM.DefaultView.Count > 0 Then
                rowItemBOM = vTblItemBOM.DefaultView(0).Row
                QtyProduce = vQtyTarget / (rowItemBOM.Good_Percentage_Process / 100)
            End If

            Return QtyProduce
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetGoodPercentageOfFG(ByVal vPrimaryProduct As Double, ByVal vIdentifierJO As String, _
            ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable) As Decimal
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Try
            vTblItemBOM.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString & " AND [Identifier JO] = '" & vIdentifierJO & "'"
            If vTblItemBOM.DefaultView.Count > 0 Then
                rowItemBOM = vTblItemBOM.DefaultView(0).Row
                Return rowItemBOM.Good_Percentage_Process
            End If
            Return 100
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 100
        End Try
    End Function

    'Public Function GetQtyTarget_InTbl(ByVal vPrimaryProduct As Double, ByVal vIdentifierJO As String, _
    'ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable) As Decimal
    '    Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
    '    Dim QtyTarget As Decimal = 0

    '    Try
    '        vTblItemRM.DefaultView.RowFilter = "[Primary Product Code RM] = " & vPrimaryProduct.ToString & " AND [Identifier JO] = '" & vIdentifierJO & "'"
    '        If vTblItemRM.DefaultView.Count > 0 Then
    '            rowItemRM = vTblItemRM.DefaultView(0).Row
    '            QtyTarget = rowItemRM.Required_Qty
    '        End If

    '        Return QtyTarget
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return 0
    '    End Try
    'End Function

    Public Function GetJONumber(ByVal vIdentifierJO As String) As String
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item1TableAdapter
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJONumber(vIdentifierJO)
            If String.IsNullOrEmpty(RetVal) = False Then
                Return RetVal
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

    Public Function IsJOFinished(ByVal vIdentifierJO As String) As Boolean
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item1TableAdapter
        Dim RetVal As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_Get_Is_Closed(vIdentifierJO)
            If RetVal.HasValue = True Then
                If RetVal.Value = 0 Then
                    Return False
                Else
                    Return True
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

    Public Function GetEarliestStartTime(ByVal vIdentifierJO As String) As Date
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item1TableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetEarliestStartTime(vIdentifierJO)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJONumber_InDgvJO(ByVal vIdentifierJO As String, ByVal vDgvItemJO As DataGridView, _
    ByVal ij As FrmTrans_ManufacturingOrder) As String
        Dim rowD As DataGridViewRow
        Dim JONumber As String = ""
        Dim LastIdx As Integer = 0

        Try
            If ij.btnOK.Enabled = True Then
                LastIdx = 2
            Else
                LastIdx = 1
            End If

            If vDgvItemJO.Rows.Count > 0 Then
                For i As Integer = 0 To vDgvItemJO.RowCount - LastIdx
                    rowD = vDgvItemJO.Rows(i)
                    If rowD.Cells(ij.ColItem_Identifier.Index).Value = vIdentifierJO Then
                        JONumber = rowD.Cells(ij.ColItem_JONumber.Index).Value
                    End If
                Next
            End If

            Return JONumber
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetIdentifierWO_InTblBOMs(ByVal vIdentifierBOM As String, _
    ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable) As String
        Dim rowItemBom As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim IdentifierWO As String = ""

        Try
            vTblItemBOM.DefaultView.RowFilter = "[Identifier BOM] = '" & vIdentifierBOM & "'"
            If vTblItemBOM.DefaultView.Count > 0 Then
                rowItemBom = vTblItemBOM.DefaultView(0).Row
                IdentifierWO = rowItemBom.Identifier_WO
            End If

            Return IdentifierWO
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetWONumber_InTbl(ByVal vIdentifierWO As String, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vIdentifierJO As String) As String
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim WONumber As String = ""

        Try
            vTblItemWO.DefaultView.RowFilter = "[Identifier] = '" & vIdentifierWO & "' AND [Identifier JO] = '" + vIdentifierJO + "'"
            If vTblItemWO.DefaultView.Count > 0 Then
                rowItemWO = vTblItemWO.DefaultView(0).Row
                WONumber = rowItemWO.WO_Number
            End If

            Return WONumber
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetWONumber(ByVal vIdentifierWO As String, ByVal vIdentifierJO As String) As String
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_woTableAdapter

        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetWONumber_ByIdentifierWOJO(vIdentifierWO, vIdentifierJO)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Retval
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAdaBundling(ByVal vPrimaryMO As Double, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable) As Boolean
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow

        Try
            vTblItemWO.DefaultView.RowFilter = "[Primary MO] = " & vPrimaryMO
            If vTblItemWO.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblItemWO.DefaultView
                    rowItemWO = dv.Row
                    If GetJmlBundling(rowItemWO.Identifier, rowItemWO.Identifier_JO) > 0 Then
                        Return False
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ApakahJOSudahSelesai(ByVal vIdentifierJO As String) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_itemTableAdapter
        Dim IsClosed As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            IsClosed = DTA.ScalarQuery_GetStatusClosed(vIdentifierJO)
            If IsClosed.HasValue = True Then
                If IsClosed.Value = 1 Then
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

    '''cek otorisasi
    Public Function ApakahSudahDiOtorisasi(ByVal vPrimaryMO As Double) As Boolean
        Try
            If Me.GetReadyForSchedule(vPrimaryMO) = EnumPilihSchedule.Schedule_tdk_Aktif Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetReadyForSchedule(ByVal vPrimary As Double) As EnumPilihSchedule
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim Retval As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetReadyForSchedule(vPrimary)
            If Retval.HasValue = True Then
                If Retval.Value = 0 Then
                    Return EnumPilihSchedule.Schedule_tdk_Aktif
                ElseIf Retval.Value = 1 Then
                    Return EnumPilihSchedule.Schedule_Aktif
                Else
                    Return EnumPilihSchedule.Schedule_tdk_Aktif
                End If
            End If
            Return EnumPilihSchedule.Schedule_tdk_Aktif
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumPilihSchedule.Schedule_tdk_Aktif
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Authorize(ByVal vPrimaryAuthorize As Double, _
    ByVal vPrimaryMO As Double, ByVal EnumReadyIsSchedule As Integer, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim DTAWO As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim TimeStamp1 As Date
        Dim ObjMySQL As New Core_Module.ClsMySQLReference

        Try
            TimeStamp1 = ObjMySQL.GetServerTime
            If EnumReadyIsSchedule = ClsTrans_ManufacturingOrder.EnumPilihSchedule.Schedule_Aktif Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTAWO.Connection.ConnectionString = Core_Module.GetConstring

                DTA.UpdateQuery_IsReadyForSchedule(EnumReadyIsSchedule, vPrimaryAuthorize, TimeStamp1, _
                vUserCode + " - " + Format(TimeStamp1, "dd MMM yyy HH.mm"), vPrimaryMO)
                DTAWO.UpdateQuery_StatusByPrimaryMO(EnumPilihStatusWO.Planned, vPrimaryMO)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAWO.Dispose()
        End Try
    End Function

    Public Function CancelAuthorize(ByVal vTimeAutorize As Date, ByVal vPrimaryAuthorize As Double, _
    ByVal vPrimarymo As Double, ByVal EnumReadyIsSchedule As Integer, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim DTAWO As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_woTableAdapter

        Try
            If EnumReadyIsSchedule = ClsTrans_ManufacturingOrder.EnumPilihSchedule.Schedule_tdk_Aktif Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTAWO.Connection.ConnectionString = Core_Module.GetConstring

                DTA.UpdateQuery_IsReadyForSchedule(EnumReadyIsSchedule, vPrimaryAuthorize, vTimeAutorize, "-", vPrimarymo)
                DTAWO.UpdateQuery_StatusByPrimaryMO(EnumPilihStatusWO.Free, vPrimarymo)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAWO.Dispose()
        End Try
    End Function

    Public Function IsiListBoxProses(ByVal vlbo As CheckedListBox, ByVal vIdentifierJO As String, _
    ByVal vTBL As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, _
    ByVal vTbl_2 As DseTrans_ManufacturingOrder.ProsesOrder_FilterDataTable) As Boolean
        Dim Row_2 As DseTrans_ManufacturingOrder.ProsesOrder_FilterRow
        Dim Row As DseTrans_ManufacturingOrder.tbltrans_mo_processorderRow
        Dim ObjProses As New Core_Module.ClsMasterProses
        Dim ListPrimaryProses As New List(Of Double)

        Try
            vTbl_2.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
            For i As Integer = 0 To vTbl_2.DefaultView.Count - 1
                Row_2 = vTbl_2.DefaultView(i).Row
                ListPrimaryProses.Add(Row_2.Primary_Proses)
            Next

            vTBL.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
            If vTBL.DefaultView.Count > 0 Then
                For i As Integer = 0 To vTBL.DefaultView.Count - 1
                    Row = vTBL.DefaultView(i).Row
                    With Row
                        vTbl_2.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO.ToString & "' AND [Primary Proses] = " & .primary_process
                        If vTbl_2.DefaultView.Count = 0 Then
                            Row_2 = vTbl_2.NewProsesOrder_FilterRow
                            Row_2.Primary_Proses = .primary_process
                            Row_2.Proses_Desc = ObjProses.GetDesc(.primary_process)
                            Row_2.Identifier_JO = vIdentifierJO
                            Row_2.Check = 1
                            vTbl_2.Rows.Add(Row_2)
                        End If

                        If ListPrimaryProses.Contains(.primary_process) = True Then
                            ListPrimaryProses.Remove(.primary_process)
                        End If
                    End With
                Next

                If ListPrimaryProses.Count > 0 Then
                    For Each Primary As Double In ListPrimaryProses
                        vTbl_2.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "' AND [Primary Proses] = " & Primary
                        If vTbl_2.DefaultView.Count > 0 Then
                            Row_2 = vTbl_2.DefaultView(0).Row
                            vTbl_2.Rows.Remove(Row_2)
                        End If
                    Next
                End If

                vTbl_2.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
                If vTbl_2.DefaultView.Count > 0 Then
                    vlbo.DataSource = vTbl_2.DefaultView
                    For i As Integer = 0 To vTbl_2.DefaultView.Count - 1
                        Row_2 = vTbl_2.DefaultView(i).Row
                        If Row_2.Check = 1 Then
                            vlbo.SetItemCheckState(i, CheckState.Checked)
                        End If
                    Next
                    vlbo.DisplayMember = "Proses Desc"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjProses = Nothing
        End Try
    End Function

    Public Function GetDate(ByVal vPrimaryMO As Double) As Date
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim Retval As New Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetDate(vPrimaryMO)
            If Retval.HasValue = True Then
                Return Retval.Value
            Else
                Return "1/1/1950"
            End If
            Return "1/1/1950"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "1/1/1950"
        Finally
            DTA.Dispose()
        End Try
    End Function

    'berkaitan update GP Actual RM
    Public Function get_GPActualRM(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim GP_Actual, QtyRefProcess, QtyRM, QtyReject, QtyUsed, QtyActualProduce As Decimal

        Try
            QtyRefProcess = get_qtyRefProcess(videntifier_wo, vprimary_product, vIdentifierJO)
            QtyRM = get_qtyRM(videntifier_wo, vprimary_product, vIdentifierJO)
            QtyReject = get_qtyRejectBundling(videntifier_wo, vprimary_product, vIdentifierJO)
            QtyUsed = get_qtyUsedRM(videntifier_wo, vprimary_product, vIdentifierJO)
            QtyActualProduce = get_qtyActualProduceBundling(videntifier_wo, vIdentifierJO)

            Dim a As Decimal
            If (QtyUsed + QtyReject) = 0 Then
                GP_Actual = 0
            Else
                If QtyRefProcess > 0 Then a = QtyRM / QtyRefProcess Else a = 0
                GP_Actual = ObjCD.DigitGrouping(((a * QtyActualProduce) / (QtyUsed + QtyReject)) * 100, 2)
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_GPActual(GP_Actual, vprimary_product, videntifier_wo, vIdentifierJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function get_qtyRefProcess(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_bomTableAdapter
        Dim retval As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetQtyRefProcessByIdentifierWOnPrimaryProductRM(videntifier_wo, vprimary_product, vIdentifierJO)
            If retval.HasValue = True Then
                Return retval.Value
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_qtyRM(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_bomTableAdapter
        Dim retval As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetQtyRMbyIdentifierWOnPrimaryProductRM(videntifier_wo, vprimary_product, vIdentifierJO)
            If retval.HasValue = True Then
                Return retval.Value
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_qtyRejectBundling(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_ManufacturingOrderTableAdapters.Bundling_RejectTableAdapter
        Dim tbl As New DseTrans_ManufacturingOrder.Bundling_RejectDataTable
        Dim row As DseTrans_ManufacturingOrder.Bundling_RejectRow
        Dim retval As Decimal = 0

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.FillBy_IdentifierWO(tbl, videntifier_wo, vprimary_product, vIdentifierJO)
            If tbl.Count > 0 Then
                For Each row In tbl.Rows
                    retval += row.qty
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
            tbl.Dispose()
        End Try
    End Function

    Public Function get_qtyUsedRM(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim retval As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetUsedQtyByIdentifierWOnPrimaryProductRM(videntifier_wo, vprimary_product, vIdentifierJO)
            If retval.HasValue = True Then
                Return retval.Value
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_qtyActualProduceBundling(ByVal videntifier_wo As String, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim tbl As New DseTrans_StockBundling.tbltransaksi_stockbundlingDataTable
        Dim row As DseTrans_StockBundling.tbltransaksi_stockbundlingRow
        Dim retval As Decimal = 0

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.FillBy_IdentifierWO(tbl, videntifier_wo, vIdentifierJO)
            If tbl.Count > 0 Then
                For Each row In tbl.Rows
                    retval += row.total_produce
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
            tbl.Dispose()
        End Try
    End Function

    Public Function IsiDGVAuthorize(ByVal vPrimarySPL As Double, ByVal vUserCode As String, ByVal vDGVAuthorize0 As DataGridView, ByVal vDGVAuthorize1 As DataGridView, ByVal vDGVAuthorize2 As DataGridView, _
            ByVal vFrm As FrmTrans_ManufacturingOrder_Authorize) As Boolean
        Dim objCd As New Core_Module.ClsCekData
        Dim objUser As New Core_Module.ClsUser
        Dim dta As New DseTrans_ManufacturingOrderTableAdapters.DataTable4TableAdapter
        Dim tbl As New DseTrans_ManufacturingOrder.DataTable4DataTable
        Dim row As DseTrans_ManufacturingOrder.DataTable4Row
        Dim dtaClosed As New DseTrans_ManufacturingOrderTableAdapters.DataTable5TableAdapter
        Dim tblClosed As New DseTrans_ManufacturingOrder.DataTable5DataTable
        Dim rowClosed As DseTrans_ManufacturingOrder.DataTable5Row
        Dim rowDgv As DataGridViewRow
        Dim isReadyForScheduling, isReadyForScheduling2 As Integer
        Dim fromDate As Date = #1/1/1950#
        Dim toDate As Date = #1/1/2200#
        Dim moPrimary, productPrimary As New Nullable(Of Double)
        Dim listOfUser As New List(Of String)
        Try

            ''init
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGVAuthorize0.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVAuthorize1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVAuthorize2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVAuthorize0.Rows.Clear()
            vDGVAuthorize1.Rows.Clear()
            vDGVAuthorize2.Rows.Clear()
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dtaClosed.Connection.ConnectionString = Core_Module.GetConstring

            ''set params
            If vFrm.cboStatus.SelectedIndex = 0 Then
                isReadyForScheduling = 0
                isReadyForScheduling2 = 1
            ElseIf vFrm.cboStatus.SelectedIndex = 1 Then
                isReadyForScheduling = 1
                isReadyForScheduling2 = 1
            ElseIf vFrm.cboStatus.SelectedIndex = 2 Then
                isReadyForScheduling = 0
                isReadyForScheduling2 = 0
            End If
            If vFrm.ChkMODate.Checked Then
                fromDate = vFrm.DtpFrom.Value.Date
                toDate = vFrm.DtpTo.Value.Date
            End If
            If vFrm.TxtPrimaryMO.Text <> "-1" Then
                moPrimary = vFrm.TxtPrimaryMO.Text
            End If
            If vFrm.txtPrimaryCodeNo.Text <> "-1" Then
                productPrimary = vFrm.txtPrimaryCodeNo.Text
            End If

            ''load database
            dta.Fill(tbl, isReadyForScheduling, isReadyForScheduling2, fromDate, toDate, moPrimary, productPrimary)
            ''set AlloDbNull property parameter; mo_primary and product_primary = True

            ''filter user created
            vFrm.DgvUser.AllowUserToAddRows = False
            For Each rowDgv In vFrm.DgvUser.Rows
                If Not String.IsNullOrEmpty(rowDgv.Cells(vFrm.ColCreatedBy.Index).Value) Then
                    If Not listOfUser.Contains(rowDgv.Cells(vFrm.ColCreatedBy.Index).Value) Then
                        listOfUser.Add(rowDgv.Cells(vFrm.ColCreatedBy.Index).Value)
                    End If
                End If
            Next
            vFrm.DgvUser.AllowUserToAddRows = True
            If listOfUser.Count > 0 Then
                For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                    row = tbl.Rows(i)
                    If Not listOfUser.Contains(row.Created_By) Then
                        tbl.Rows.RemoveAt(i)
                    End If
                Next
            End If

            ''filter hanya yang belum close jo yang bisa cancel auth
            If vFrm.TboAuthorize.SelectedIndex = 1 Then
                Dim moPrimaryMin As Double = Double.MaxValue
                Dim moPrimaryMax As Double = Double.MinValue
                For Each row In tbl
                    If row.Primary < moPrimaryMin Then
                        moPrimaryMin = row.Primary
                    End If
                    If row.Primary > moPrimaryMax Then
                        moPrimaryMax = row.Primary
                    End If
                Next
                dtaClosed.Fill(tblClosed, moPrimaryMin, moPrimaryMax)
                For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                    row = tbl.Rows(i)
                    tblClosed.DefaultView.RowFilter = "[Primary MO] = " & row.Primary.ToString
                    If tblClosed.DefaultView.Count > 0 Then
                        tbl.Rows.RemoveAt(i)
                    End If
                Next
            End If

            ''load to datagridview
            For Each row In tbl
                If vFrm.TboAuthorize.SelectedIndex = 0 Then
                    vDGVAuthorize0.Rows.Add()
                    rowDgv = vDGVAuthorize0.Rows(vDGVAuthorize0.Rows.Count - 1)
                    rowDgv.Cells(vFrm.ColDgvAuth_Primary.Index).Value = row.Primary
                    rowDgv.Cells(vFrm.ColDgvAuth_MONo.Index).Value = row.MO_Number
                    rowDgv.Cells(vFrm.ColDgvAuth_Date.Index).Value = Format(row._Date, "dd MMM yyyy")
                ElseIf vFrm.TboAuthorize.SelectedIndex = 1 Then
                    If row.Kode_User = vUserCode Then
                        If row.IsSum_of_PrimaryNull Then ''hanya tampilkan yang belum ada schedule
                            vDGVAuthorize1.Rows.Add()
                            rowDgv = vDGVAuthorize1.Rows(vDGVAuthorize1.Rows.Count - 1)
                            rowDgv.Cells(vFrm.ColDgvUnAuth_Primary.Index).Value = row.Primary
                            rowDgv.Cells(vFrm.ColDgvUnAuth_MONo.Index).Value = row.MO_Number
                            rowDgv.Cells(vFrm.ColDgvUnAuth_Date.Index).Value = Format(row._Date, "dd MMM yyyy")
                        End If
                    End If
                ElseIf vFrm.TboAuthorize.SelectedIndex = 2 Then
                    vDGVAuthorize2.Rows.Add()
                    rowDgv = vDGVAuthorize2.Rows(vDGVAuthorize2.Rows.Count - 1)
                    rowDgv.Cells(vFrm.col3_Primary.Index).Value = row.Primary
                    rowDgv.Cells(vFrm.col3_DocNo.Index).Value = row.MO_Number
                    If row.Is_Ready_for_Scheduling = 0 Then
                        rowDgv.Cells(vFrm.col3_Status.Index).Value = "Un-Authorized"
                    ElseIf row.Is_Ready_for_Scheduling = 1 Then
                        rowDgv.Cells(vFrm.col3_Status.Index).Value = "Authorized by " & row.Kode_User & ""
                    End If
                    rowDgv.Cells(vFrm.col3_Date.Index).Value = Format(row._Date, "dd MMM yyyy")
                End If
            Next

            vDGVAuthorize0.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVAuthorize1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVAuthorize2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub FilterCreatedBy(ByRef vTBL As DseTrans_ManufacturingOrder.tbltrans_moAuthDataTable, ByVal vFrm As FrmTrans_ManufacturingOrder_Authorize)
        Dim vRow As DseTrans_ManufacturingOrder.tbltrans_moAuthRow
        Dim tblMOAuth As New DseTrans_ManufacturingOrder.tbltrans_moAuthDataTable
        Dim rowMoAuth As DseTrans_ManufacturingOrder.tbltrans_moAuthRow
        Dim filterUser As String = ""
        Try
            'pindahkan isi tbl dari DB ke data table
            For Each vRow In vTBL
                rowMoAuth = tblMOAuth.Newtbltrans_moAuthRow
                For i As Integer = 0 To vTBL.Columns.Count - 1
                    rowMoAuth.Item(i) = vRow.Item(i)
                Next
                tblMOAuth.Rows.Add(rowMoAuth)
            Next

            For Each rd As DataGridViewRow In vFrm.DgvUser.Rows
                If rd.Cells(vFrm.ColCreatedBy.Index).Value <> "" Then
                    If filterUser = "" Then
                        filterUser = "[created by] = '" & rd.Cells(vFrm.ColCreatedBy.Index).Value & "'"
                    Else
                        filterUser = filterUser & " OR [created by] = '" & rd.Cells(vFrm.ColCreatedBy.Index).Value & "'"
                    End If
                End If
            Next
            tblMOAuth.DefaultView.RowFilter = filterUser
            vTBL.Rows.Clear()
            For i As Integer = 0 To tblMOAuth.DefaultView.Count - 1
                rowMoAuth = tblMOAuth.DefaultView(i).Row
                vRow = vTBL.Newtbltrans_moAuthRow
                For i2 As Integer = 0 To tblMOAuth.Columns.Count - 1
                    vRow.Item(i2) = rowMoAuth.Item(i2)
                Next
                vTBL.Rows.Add(vRow)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Public Sub Get_argb(ByVal vIdentifierJO As String, ByRef a As Integer, ByRef r As Integer, ByRef g As Integer, ByRef b As Integer)
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item3TableAdapter
        Dim TBL As New DseTrans_ManufacturingOrder.tbltrans_mo_item3DataTable
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_item3Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vIdentifierJO)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                a = row.a
                r = row.r
                g = row.g
                b = row.b
            Else
                a = 0
                r = 0
                g = 0
                b = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub Set_argb(ByVal vIdentifierJO As String, ByVal a As Integer, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer)
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_Set_argb(a, r, g, b, vIdentifierJO)
            DTA.UpdateQuery_set_argb2(a, r, g, b, vIdentifierJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function IsScheduled(ByVal vPrimaryMO As Double) As Boolean
        Dim dta As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim n As New Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            n = dta.ScalarQuery_GetJmlInSch(vPrimaryMO)
            If n.HasValue Then
                If n.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
        End Try
    End Function

    Public Function IsScheduled(ByVal vPrimaryMO As Double, ByVal vIdentifierJO As String) As Boolean
        Dim dta As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim n As New Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            n = dta.ScalarQuery_GetJmlSch(vPrimaryMO, vIdentifierJO)
            If n.HasValue Then
                If n.Value > 0 Then
                    Return True
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
        End Try
    End Function

    Public Function GetScheduledMOCount(ByVal vPrimaryMO As Double) As Integer
        Dim d As New DseTrans_ManufacturingOrder_SetFinishJOTableAdapters.DataTable1TableAdapter
        Dim v As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            v = d.ScalarQuery_GetScheduledMOCount(vPrimaryMO)
            If v.HasValue Then
                Return v.Value
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetClosedJOCount(ByVal vPrimaryMO As Double) As Integer
        Dim d As New DseTrans_ManufacturingOrder_SetFinishJOTableAdapters.DataTable1TableAdapter
        Dim v As New Nullable(Of Long)

        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            v = d.ScalarQuery_GetClosedJOCount(vPrimaryMO)
            If v.HasValue Then
                Return v.Value
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function ApakahBolehCloseJO(ByVal vIdentifierJO As String, ByVal vShowMessage As Boolean) As Boolean
        Dim d As New DseTrans_ManufacturingOrder_SetFinishJOTableAdapters.DataTable1TableAdapter
        Dim v As Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            v = d.ScalarQuery_IsAuthJO(vIdentifierJO)
            If v.HasValue Then
                If v.Value <> 1 Then
                    If vShowMessage Then
                        MsgBox("MO for this JO isn't authorized, you can only close JO with authorized MO.", MsgBoxStyle.Information, "Information.")
                    End If
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally

        End Try
    End Function

    Public Function GetAuthStatus(ByVal vPrimaryMO As Double) As Boolean
        Dim d As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim n As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            n = d.ScalarQuery_GetAuthStatus(vPrimaryMO)
            If n.HasValue Then
                If n.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function GetRawMaterialRequest(ByVal vPrimaryProduct As Double) As Decimal
        Dim DTARM As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_item_rm2TableAdapter
        Dim TBLRM As New DseTransaksiMaterialPlanning2.tbltrans_mo_item_rm2DataTable
        Dim rowRM As DseTransaksiMaterialPlanning2.tbltrans_mo_item_rm2Row
        Dim DTAWIP As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_item_wip2TableAdapter
        Dim TBLWIP As New DseTransaksiMaterialPlanning2.tbltrans_mo_item_wip2DataTable
        Dim rowWIP As DseTransaksiMaterialPlanning2.tbltrans_mo_item_wip2Row
        Dim QtyRM, QtyWIP As Decimal

        Try
            QtyRM = 0
            QtyWIP = 0

            DTARM.Connection.ConnectionString = Core_Module.GetConstring
            DTAWIP.Connection.ConnectionString = Core_Module.GetConstring

            'item rm req + item wip
            DTARM.Fill(TBLRM, vPrimaryProduct)
            QtyRM = 0
            For Each rowRM In TBLRM.Rows
                QtyRM += rowRM.Balance_Qty
            Next

            DTAWIP.Fill(TBLWIP, vPrimaryProduct)
            QtyWIP = 0
            For Each rowWIP In TBLWIP.Rows
                QtyWIP += rowWIP.Qty_Balance
            Next

            Return QtyWIP + QtyRM
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTARM.Dispose()
        End Try
    End Function

    Public Function GetMOB(ByVal vPrimaryProduct As Double) As Decimal
        Dim DTAMOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_item2TableAdapter
        Dim TBLMOB As New DseTransaksiMaterialPlanning2.tbltrans_mo_item2DataTable
        Dim rowMOB As DseTransaksiMaterialPlanning2.tbltrans_mo_item2Row
        Dim QtyMOB As Decimal

        Try
            QtyMOB = 0

            DTAMOB.Connection.ConnectionString = Core_Module.GetConstring

            'item MOB req + item wip
            DTAMOB.Fill(TBLMOB, vPrimaryProduct)
            QtyMOB = 0
            For Each rowMOB In TBLMOB.Rows
                QtyMOB += rowMOB.Balance_Qty
            Next

            Return QtyMOB
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAMOB.Dispose()
        End Try
    End Function

    Public Function GetListOfIdentifierJO(ByVal vPrimaryMO As Double) As List(Of String)
        Dim d As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_itemTableAdapter
        Dim t As New DseTrans_ManufacturingOrder.tbltrans_mo_itemDataTable
        Dim r As DseTrans_ManufacturingOrder.tbltrans_mo_itemRow
        GetListOfIdentifierJO = New List(Of String)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            d.FillBy_PrimaryMO(t, vPrimaryMO)
            For Each r In t
                GetListOfIdentifierJO.Add(r.Identifier)
            Next
            Return GetListOfIdentifierJO
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetListOfIdentifierJO
        End Try
    End Function

    Public Function GetIdentifierJO(ByVal vJONo As String) As String
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_itemTableAdapter
        Dim RetVal As String

        Try
            RetVal = DTA.ScalarQuery_GetIdentifierJO(vJONo)
            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJmlJO(ByVal vJONo As String) As Double
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_itemTableAdapter
        Dim Jml As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Jml = DTA.ScalarQuery_GetJml(vJONo)
            If Jml.HasValue = True Then
                Return Jml.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahWIP(ByVal vIdentifierWO As String, ByVal vIdentifierJO As String, ByVal vPrimaryCodeNo As Double, _
    ByVal vPrimaryMO As Double, ByVal vDefaultRetVal As Boolean) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim TBL As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_IdentifierJO_WO(TBL, vPrimaryMO, vIdentifierJO, vIdentifierWO)

            For Each row In TBL.Rows
                If row.Primary_Product_Code_RM = vPrimaryCodeNo Then
                    If row.is_wip = 1 Then
                        Return True
                    Else
                        Return False
                    End If
                End If
            Next

            Return vDefaultRetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function
End Class
