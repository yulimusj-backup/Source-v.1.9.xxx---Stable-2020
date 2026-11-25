Option Explicit On
Option Strict Off

Imports System.Net
Public Class ClsAntrian
    Inherits System.Windows.Forms.Timer
    Dim ObjMySQL As New ClsMySQLReference
    Dim ObjRnd As New Random
    Dim DTAInOUT As New DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim DTA1 As New DseAntrianTableAdapters.tbltrans_antrianTableAdapter

    'waktu tunggu antrian
    Public mWaitTime As Integer = 60

    Public Function GetWaitTime() As Integer
        Return Me.mWaitTime
    End Function

    Public Property WaitTime() As Integer
        Get
            Return Me.mWaitTime
        End Get
        Set(ByVal value As Integer)
            mWaitTime = value
        End Set
    End Property

    Public Enum EnumTipeTransaksiAsalAntrian As Integer
        NoType
        MasterKompetensi
        MasterProduct
        MasterProses
        MasterJobPosition
        AntrianTunggal
        TransaksiStockInOut_SimpanBaru
        TransaksiStockInOut_SimpanPerubahan
        TransaksiStockInOut_Hapus
        TransaksiDO_SimpanBaru
        TransaksiDO_SimpanPerubahan
        TransaksiDO_Hapus
        TransaksiInvoice_SimpanBaru
        TransaksiInvoice_SimpanPerubahan
        TransaksiInvoice_Hapus
        TransaksiInvoice_Quick_SimpanBaru
        TransaksiInvoice_Quick_SimpanPerubahan
        TransaksiInvoice_Quick_Hapus
        TransaksiJurnal_General_SimpanBaru
        TransaksiJurnal_General_SimpanPerubahan
        TransaksiJurnal_General_Hapus
        TransaksiJurnal_Sales_SimpanBaru
        TransaksiJurnal_Sales_SimpanPerubahan
        TransaksiJurnal_Sales_Hapus
        TransaksiJurnal_Purchase_SimpanBaru
        TransaksiJurnal_Purchase_SimpanPerubahan
        TransaksiJurnal_Purchase_Hapus
        TransaksiJurnal_Asset_Dep_SimpanBaru
        TransaksiJurnal_Asset_Dep_SimpanPerubahan
        TransaksiJurnal_Asset_Acq_SimpanBaru
        TransaksiJurnal_Asset_Acq_SimpanPerubahan
        TransaksiJurnal_Cash_SimpanBaru
        TransaksiJurnal_Cash_SimpanPerubahan
        TransaksiJurnal_EOM_SimpanBaru
        TransaksiJurnal_EOM_SimpanPerubahan
        TransaksiJurnal_EOM_Hapus
        TransaksiPurchaseReturn_SimpanBaru
        TransaksiPurchaseReturn_SimpanPerubahan
        TransaksiPurchaseReturn_Hapus
        TransaksiStockBundling_SimpanBaru
        TransaksiStockBundling_SimpanPerubahan
        TransaksiStockBundling_Hapus
        TransaksiAOP_SimpanBaru
        TransaksiAOP_SimpanPerubahan
        TransaksiAOP_Hapus
        Transaksi_Numbering_SimpanBaru
        Transaksi_Numbering_SimpanPerubahan
        Transaksi_Numbering_Hapus
        TransaksiSalesReturn_SimpanBaru
        TransaksiSalesReturn_SimpanPerubahan
        TransaksiSalesReturn_Hapus
        TransaksiMP_SimpanBaru
        TransaksiMP_SimpanPerubahan
        TransaksiMP_Hapus
        TransaksiSF_SimpanBaru
        TransaksiSF_SimpanPerubahan
        TransaksiSF_Hapus
        MasterUser_TempPass_SimpanBaru
        MasterUser_TempPass_SimpanPerubahan
        MasterCuti_SimpanBaru
        MasterCuti_SimpanPerubahan
        MasterCuti_Hapus
        MasterKaryawan_SimpanBaru
        MasterKaryawan_SimpanPerubahan
        MasterKaryawan_Hapus
        MasterPayrollAccess_SimpanBaru
        MasterPayrollAccess_SimpanPerubahan
        MasterPayrollAccess_Hapus
        MasterSalaryFormula_SimpanBaru
        MasterSalaryFormula_SimpanPerubahan
        MasterSalaryFormula_Hapus
        MasterSalaryDescription_SimpanBaru
        MasterSalaryDescription_SimpanPerubahan
        MasterSalaryDescription_Hapus
        MasterWorkshift_SimpanBaru
        MasterWorkshift_SimpanPerubahan
        MasterWorkshift_Hapus
        Trans_OtherDeduction_SimpanBaru
        Trans_OtherDeduction_SimpanPerubahan
        Trans_OtherDeduction_Hapus
        Trans_Attd_SimpanBaru
        Trans_Attd_SimpanPerubahan
        Trans_Attd_Hapus
        TransaksiLeaveSetting_SimpanBaru
        TransaksiLeaveSetting_SimpanPerubahan
        TransaksiLeaveSetting_Hapus
        TransaksiLeaveApplication_SimpanBaru
        TransaksiLeaveApplication_SimpanPerubahan
        TransaksiLeaveApplication_Hapus
        Trans_Payment_SimpanBaru
        Trans_Payment_SimpanPerubahan
        Trans_Payment_Hapus
        Trans_Tax_Beginning_Balance_SimpanBaru
        Trans_Tax_Beginning_Balance_SimpanPerubahan
        Trans_Tax_Beginning_Balance_Hapus
        Trans_Loan_SimpanBaru
        Trans_Loan_SimpanPerubahan
        Trans_Loan_Hapus
        Trans_Loan_Payment_SimpanBaru
        Trans_Loan_Payment_SimpanPerubahan
        Trans_Loan_Payment_Hapus
        Trans_Salary_SimpanBaru
        Trans_Salary_SimpanPerubahan
        Trans_Salary_Hapus
        Trans_SalesBudget_SimpanBaru
        Trans_SalesBudget_SimpanPerubahan
        Trans_SalesBudget_Hapus
        Trans_Purchase_Request_SimpanBaru
        Trans_Purchase_Request_SimpanPerubahan
        Trans_Purchase_Request_Hapus
        Trans_Purchase_Request_Cancel_SimpanBaru
        Trans_Purchase_Request_Cancel_SimpanPerubahan
        Trans_Purchase_Request_Cancel_Hapus
        Trans_Purchase_Request_PickUp_SimpanBaru
        Trans_Purchase_Request_PickUp_SimpanPerubahan
        Trans_Purchase_Request_PickUp_Hapus
        TransaksiPurchaseOrder_SimpanBaru
        TransaksiPurchaseOrder_SimpanPerubahan
        TransaksiPurchaseOrder_Hapus
        TransaksiPackingList_SimpanBaru
        TransaksiPackingList_SimpanPerubahan
        TransaksiPackingList_Hapus
        CreditDebitNote_SimpanBaru
        CreditDebitNote_SimpanPerubahan
        CreditDebitNote_Hapus
        TransaksiCustom_SimpanBaru
        TransaksiCustom_SimpanPerubahan
        TransaksiCustom_Hapus
        Trans_StockTransfer_SimpanBaru
        Trans_StockTransfer_SimpanPerubahan
        Trans_StockTransfer_Hapus
        Trans_TaxRecalculate_SimpanBaru
        Trans_TaxRecalculate_SimpanPerubahan
        Trans_TaxRecalculate_Hapus
        Trans_ScheduleDownTime_SimpanBaru
        Trans_ScheduleDownTime_SimpanPerubahan
        Trans_ScheduleDownTime_Hapus
        FrmTrans_ManufacturingOrder_PrintJO_SimpanBaru
        FrmTrans_ManufacturingOrder_SimpanBaru
        FrmTrans_ManufacturingOrder_SimpanPerubahan
        FrmTrans_ManufacturingOrder_Hapus
        FrmTrans_ManufacturingOrder_CariDetail
        TransaksiStockInOut_LoadDataTBLQty
        TransaksiDO_LoadDataTBLQty
        TransaksiBarcode_SimpanBaru
        TransaksiBarcode_SimpanPerubahan
        TransaksiBarcode_Hapus
        MasterWarehouse_SimpanBaru
        MasterWarehouse_SimpanPerubahan
        MasterWarehouse_Hapus
        MasterStorageArea_SimpanBaru
        MasterStorageArea_SimpanPerubahan
        MasterStorageArea_Hapus
        MasterProduct_UpdateControlState
    End Enum

    Public Enum EnumTipeAntrian As Integer
        AntrianStockInOut_DO_Invoice
        AntrianMaster
        AntrianTransaksiPOS
        AntrianTransferDataPOS
        AntrianAssetModule
        AntrianTransaksiHRPayroll
        AntrianTunggalSalaryFormula
        AntrianTunggalWorkshiftAllocate
        AntrianPurchaseRequest
        AntrianScheduling
        AntrianTunggalSchedulingOT
        AntrianTransaksiProduksi
        AntrianCustom1
        AntrianCustom2
        AntrianCustom3
        AntrianTransaksiSPLGroup 'trf source infiniti to finalize
    End Enum

    ''' <summary>
    ''' parameter interval check adalah waktu yg dispare untuk pengecekan data.
    ''' data dianggap tidak valid jika sudah
    ''' </summary>
    ''' <remarks></remarks>
    Public IntervalCheck As Double = 5000
    Public IntervalUpdate As Double = 1000

    Private mIdentifier As String = ""
    Private mPrimaryAntrian As Double = -1

    Public Sub GetNoAntrian(ByRef vIdentifier As String, ByVal ObjRnd As Random, _
    ByVal DTAInOUT As DseAntrianTableAdapters.tbltrans_antrianTableAdapter, ByVal vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian, _
    ByRef vPrimary_Antrian As Double, ByVal vPrimaryTransaksiAsal As Double, ByVal vTipeTransaksiAsal As EnumTipeTransaksiAsalAntrian, _
    ByVal vNamaUser As String)
        Dim NamaKomputer As String = ""
        Dim IP_Add As String = ""
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTAInOUT.Connection.ConnectionString = Core_Module.GetConstring

            NamaKomputer = My.Computer.Name
            IP_Add = ObjCD.GetIpAddress_Local
            DTAInOUT.Connection.ConnectionString = Core_Module.GetConstring

            vIdentifier = ObjRnd.Next(1000000)
            'jangan sampai sama dengan identifier yang sudah ada
            While DTAInOUT.ScalarQuery_GetJmlIdentifier(vTipeAntrian, vIdentifier).Value > 0
                vIdentifier = ObjRnd.Next(1000000)
            End While
            'masukan data baru ke antrian
            If System.Net.Dns.GetHostAddresses("").Length > 2 Then
                DTAInOUT.Insert(vTipeAntrian, vIdentifier, vPrimaryTransaksiAsal, vTipeTransaksiAsal, vNamaUser, My.Computer.Name + "-" + _
                System.Net.Dns.GetHostAddresses("")(2).ToString)
            Else
                DTAInOUT.Insert(vTipeAntrian, vIdentifier, vPrimaryTransaksiAsal, vTipeTransaksiAsal, vNamaUser, My.Computer.Name + "- localhost")
            End If

            'cek jangan sampai ada 2 data yang sama
            While DTAInOUT.ScalarQuery_GetJmlIdentifier(vTipeAntrian, vIdentifier).Value > 1
                'bila ada yang sama maka harus masukkan data baru
                DTAInOUT.DeleteQuery(vTipeAntrian, vIdentifier)
                'minta identifier baru
                vIdentifier = ObjRnd.Next(1000000)
                'jangan sampai sama dengan identifier yang sudah ada
                While DTAInOUT.ScalarQuery_GetJmlIdentifier(vTipeAntrian, vIdentifier).Value > 0
                    vIdentifier = ObjRnd.Next(1000000)
                    'masukan data baru ke antrian
                    DTAInOUT.Insert(vTipeAntrian, vIdentifier, vPrimaryTransaksiAsal, vTipeTransaksiAsal, vNamaUser, My.Computer.Name)
                End While
            End While

            'ambil no antrian
            vPrimary_Antrian = DTAInOUT.Scalar_Query_GetPrimary(vIdentifier, vTipeAntrian).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Function StartAntrianTunggal(ByVal vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian, ByVal vNamaUser As String, _
    ByRef VNamaUserLain As String, ByRef vNamaKomputerLain As String) As Boolean
        Dim LoopGo As Boolean = True
        Dim PrimaryMin As Double = -1
        Dim TS_Min As Date = #1/1/1950#
        Dim TS As Date = #1/1/1950#
        Dim DTA As New DseAntrianTableAdapters.tbltrans_antrianTableAdapter
        Dim TBL As New DseAntrian.tbltrans_antrianDataTable
        Dim row As DseAntrian.tbltrans_antrianRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.Interval = Me.IntervalUpdate

            'masukkan data pertama
            Me.GetNoAntrian(Me.mIdentifier, Me.ObjRnd, Me.DTAInOUT, vTipeAntrian, Me.mPrimaryAntrian, -1, EnumTipeTransaksiAsalAntrian.AntrianTunggal, _
            vNamaUser)

            'cek kembali datanya dengan menggunakan loop
            While LoopGo = True
                'Apakah masih data pertama,
                PrimaryMin = Me.GetPrimaryMin(vTipeAntrian)

                'jika masih data pertama, maka return true dan jalankan timer
                If PrimaryMin = Me.mPrimaryAntrian Then
                    LoopGo = False
                    Me.Enabled = True
                    LoopGo = False
                    Return True
                Else
                    'jika bukan data pertama, maka 
                    '0. get timestamp
                    TS_Min = Me.GetTimestamp(PrimaryMin)
                    TS = Me.GetTimestamp(Me.mPrimaryAntrian)

                    '1. cek data pertama, nilai timestampnya > me.interval cek
                    If (TS - TS_Min).TotalMilliseconds > Me.IntervalCheck Then
                        '    1.1. Jika ya, maka hapuskan data pertama, kemudian cek kembali
                        Me.HapusAntrian(PrimaryMin)
                        LoopGo = True
                    Else
                        '    1.2. Jika tidak, maka return false dan hapuskan datanya
                        Me.EndAntrianTunggal()
                        Me.Enabled = False
                        LoopGo = False

                        DTA.FillBy_Primary(TBL, PrimaryMin)

                        If TBL.Rows.Count > 0 Then
                            row = TBL.Rows(0)
                            VNamaUserLain = row.Nama_User
                            vNamaKomputerLain = row.Nama_Komputer
                        Else
                            VNamaUserLain = ""
                            vNamaKomputerLain = ""
                        End If

                        Return False
                    End If
                End If
            End While

            MsgBox("Logic error, please contact your vendor if you see this message")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function HapusAntrian(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.Delete(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub HapusAntrian_HanyaYgLogOff()
        Dim DTA As New DseAntrianTableAdapters.tbltrans_antrianTableAdapter
        Dim TBL As New DseAntrian.tbltrans_antrianDataTable
        Dim row As DseAntrian.tbltrans_antrianRow
        Dim DTA1 As New DseTransaksiMonitorTableAdapters.tbltrans_monitorTableAdapter
        Dim TBL1 As New DseTransaksiMonitor.tbltrans_monitorDataTable
        Dim Timestamp As Date = Now
        Dim ObjMySQL As New Core_Module.ClsMySQLReference
        Dim ListNamaUser As New List(Of String)
        Dim ObjMonitor As New ClsTransaksiMonitor

        Try
            'hapuskan item monitor yang time key > 5 menit
            Timestamp = ObjMySQL.GetServerTime
            DTA1.DeleteQuery(Timestamp - TimeSpan.FromMilliseconds(ObjMonitor.DeleteInterval))

            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'load semua user yang online
            DTA1.FillBy_SortUser(TBL1, CInt(ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License))
            'load antrian
            DTA.FillBy_All(TBL)
            For Each row In TBL.Rows
                TBL1.DefaultView.RowFilter = "[user name] = '" + row.Nama_User + "'"
                If TBL1.DefaultView.Count = 0 Then
                    'hapus antrian atas user yang tidak log on
                    DTA.DeleteQuery_ByNamaUser(row.Nama_User)
                End If
                ListNamaUser.Add(row.Nama_User)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
            DTA.Dispose()
            ObjMySQL = Nothing
            ObjMonitor = Nothing
        End Try
    End Sub

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryMin(ByVal vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian) As Double
        Dim DTA As New DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryMin(vTipeAntrian).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryMin(vTipeAntrian).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function EndAntrianTunggal() As Boolean
        Dim DTA As New DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(Me.mPrimaryAntrian)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub ClsAntrian_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Tick
        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.UpdateQuery_Timestamp(Me.ObjMySQL.GetServerTime, Me.mPrimaryAntrian)
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowQueue(ByVal vDGV As DataGridView, ByVal vFrm As FrmTransaksiMonitor)
        Dim DTA As New DseAntrianTableAdapters.tbltrans_antrianTableAdapter
        Dim TBL As New DseAntrian.tbltrans_antrianDataTable
        Dim row As DseAntrian.tbltrans_antrianRow
        Dim rowDGV As DataGridViewRow

        Try
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGV.Rows.Clear()
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_All(TBL)
            For Each row In TBL.Rows
                vDGV.Rows.Add()
                rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                rowDGV.Cells(vFrm.ColQueue_No.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(vFrm.ColQueue_ComputerID.Index).Value = row.Nama_Komputer
                rowDGV.Cells(vFrm.ColQueue_TransactionCode.Index).Value = row.Tipe_Transaksi.ToString + "-" + row.Tipe_Transaksi_Acuan.ToString
                rowDGV.Cells(vFrm.ColQueue_UserName.Index).Value = row.Nama_User
                rowDGV.Cells(vFrm.ColQueue_Timestamp.Index).Value = Format(row.Timestamp, "dd MMM yyyy - HH:mm:ss")
            Next

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub
End Class
