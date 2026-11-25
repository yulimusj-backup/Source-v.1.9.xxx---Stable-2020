
Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Public Class ClsTrans_Cuti_Approval

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransferDataPOS
    Dim vTipeTransaksiAsal As Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian = Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiLeaveApplication_SimpanBaru
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim vObjAntrian As New Core_Module.ClsAntrian
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

    Public Enum EnumStatus As Integer
        WaitingForApproval
        Approved
        Rejected
    End Enum

    Private Function ApakahBolehEdit(ByVal vPrimary As Double, Optional ByVal vShowMsg As Boolean = True) As Boolean
        Dim dta As New DseTrans_Cuti_ApprovalTableAdapters.tbltrans_attendanceTableAdapter
        Dim tbl As New DseTrans_Cuti_Approval.tbltrans_attendanceDataTable
        Dim row As DseTrans_Cuti_Approval.tbltrans_attendanceRow
        Dim DtaCuti As New DseTrans_Cuti_ApprovalTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim ListStr As String = ""
        Try
            If dta.Fill(tbl, vPrimary) > 0 Then
                Dim i As Integer = 1
                For Each row In tbl
                    ListStr = ListStr + i.ToString + ". " + row.no_transaksi.ToString + " ( " + Format(row.start_date, "dd MMM yyyy") + " - " + Format(row.end_date, "dd MMM yyyy") + " )" + vbNewLine
                    i = i + 1
                Next
                If vShowMsg = True Then
                    MsgBox("Cannot modify " + DtaCuti.ScalarQuery_GetNoTranByPrimary(vPrimary).ToString + " because referring to transaction attendance:" + vbNewLine + ListStr)
                End If
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            tbl.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vStatus As EnumStatus, ByVal vApprovedDate As Date, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vPrimary As Double, ByVal i As FrmTransaksiCuti_Approval, _
    ByVal vIdxDgv As Integer) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            'ambil no antrian
            Me.vObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, Me.vPrimary_Antrian, _
            vPrimary, Me.vTipeTransaksiAsal, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= Me.vObjAntrian.GetWaitTime Then
                    Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, vStatus, vApprovedDate, vObjCD, vPrimary, i, vIdxDgv)
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vStatus As EnumStatus, ByVal vApprovedDate As Date, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vPrimary As Double, ByVal i As FrmTransaksiCuti_Approval, _
    ByVal vIdxDgv As Integer) As Boolean
        Dim DTA As New DseTrans_Cuti_ApprovalTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim dtaPengajuan As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim PrimaryKaryawanLama As Double = -1
        Dim PrimaryCutiLama As Double = -1
        Dim PengajuanCutiDate As Date = Now
        Dim ObjCuti As New Core_Module.ClsMasterCuti
        Try
            If Me.ApakahBolehEdit(vPrimary) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                If vStatus = EnumStatus.WaitingForApproval Then
                    DTA.UpdateQuery_SetStatus(vStatus, "-", #1/1/1950#, vPrimary)
                Else
                    DTA.UpdateQuery_SetStatus(vStatus, vUserCode, vApprovedDate, vPrimary)
                End If
                'Get Primary Karyawan dan Primary Cuti By Primary Pengajuan-nya(--->vPrimary)
                PrimaryKaryawanLama = i.dgvItem.Item(i.ColPrimaryEmployee.Index, vIdxDgv).Value
                PrimaryCutiLama = i.dgvItem.Item(i.ColPrimaryLeaveAplication.Index, vIdxDgv).Value
                PengajuanCutiDate = i.dgvItem.Item(i.ColDate.Index, vIdxDgv).Value
                'Set Balance By Primary Karyawan dan Primary Cuti
                ObjCuti.SetBalance(PrimaryKaryawanLama, PrimaryCutiLama, PengajuanCutiDate)

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            dtaPengajuan.Dispose()
        End Try
    End Function

    'Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vStatus As EnumStatus, ByVal vApprovedDate As Date, _
    ' ByVal vObjCD As Core_Module.ClsCekData, ByVal vPrimary As Double, _
    ' ByVal i As FrmTransaksiCuti_Approval) As Boolean
    '    Dim CheckInterval As Double = 0
    '    Dim TryCounter As Integer = 0
    '    Dim Start As Double = 0
    '    Dim retval As Boolean = True
    '    Dim SW As New Stopwatch

    '    Try
    '        'ambil no antrian
    '        Me.vObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, Me.vPrimary_Antrian, _
    '        vPrimary, Me.vTipeTransaksiAsal, vUserCode)

    '        'coba simpan data baru
    '        Me.Clock_Antri.Reset()
    '        Me.Clock_Antri.Start()
    '        While Start = 0
    '            If SW.Elapsed.TotalSeconds >= Me.vObjAntrian.GetWaitTime Then
    '                Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
    '                MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
    '                Exit Function
    '            Else
    '                'hanya mencoba ke database setelah interval waktu tertentu saja
    '                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
    '                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
    '                        'clock antri matikan
    '                        Me.Clock_Antri.Stop()
    '                        'sudah saatnya melakukan proses simpan baru
    '                        retval = Me.SimpanBaru1(vUserCode, vStatus, vApprovedDate, vObjCD, vPrimary, i)
    '                        'hapus data antri untuk transaksi ini
    '                        If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
    '                            MsgBox("Logic error(1), contact your provider when you see this message", _
    '                            MsgBoxStyle.Information, "Information")
    '                        End If
    '                        'nilai balikan
    '                        Return retval
    '                    Else
    '                        'kontrol interval pengecekan
    '                        TryCounter += 1
    '                        CheckInterval = ((TryCounter \ 5) + 1) * 500
    '                        Me.Clock_Antri.Reset()
    '                        Me.Clock_Antri.Start()
    '                        Threading.Thread.Sleep(1000)
    '                    End If
    '                End If
    '            End If
    '        End While
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        SW = Nothing
    '    End Try
    'End Function

    'Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vStatus As EnumStatus, ByVal vApprovedDate As Date, _
    'ByVal vObjCD As Core_Module.ClsCekData, ByVal vPrimary As Double, _
    'ByVal i As FrmTransaksiCuti_Approval) As Boolean
    '    Dim DTA As New DseTrans_Cuti_ApprovalTableAdapters.tbltrans_cuti_pengajuanTableAdapter
    '    Dim dtaPengajuan As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
    '    Dim PrimaryKaryawanLama As Double = -1
    '    Dim PrimaryCutiLama As Double = -1
    '    Dim ObjCuti As New Core_Module.ClsMasterCuti
    '    Try
    '        If Me.ApakahBolehEdit(vPrimary) = True Then
    '            DTA.Connection.ConnectionString = Core_Module.GetConstring
    '            If vStatus = EnumStatus.WaitingForApproval Then
    '                DTA.UpdateQuery_SetStatus(vStatus, "-", #1/1/1950#, vPrimary)
    '            Else
    '                DTA.UpdateQuery_SetStatus(vStatus, vUserCode, vApprovedDate, vPrimary)
    '            End If
    '            'Get Primary Karyawan dan Primary Cuti By Primary Pengajuan-nya(--->vPrimary)
    '            PrimaryKaryawanLama = dtaPengajuan.ScalarQuery_GetPrimaryKaryawanByPrimary(i.TxtPrimary.Text)
    '            PrimaryCutiLama = dtaPengajuan.ScalarQuery_GetPrimaryCutiByPrimary(i.TxtPrimary.Text)
    '            'Set Balance By Primary Karyawan dan Primary Cuti
    '            ObjCuti.SetBalance(i.TxtPrimaryEmployee.Text, CDbl(i.cboLeave.SelectedValue), i.dtpStart.Value)

    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        dtaPengajuan.Dispose()
    '    End Try
    'End Function

    Public Function GetStatusNotes(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_Cuti_ApprovalTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetStatusNotes(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetStatusNotes(vPrimary)
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

    Public Sub LoadData(ByVal Frm As FrmTransaksiCuti_Approval, ByVal vStatus As EnumStatus, _
    ByRef vTBL As DseTrans_Cuti_Approval.tbltrans_cuti_pengajuanDataTable)
        Frm.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        Frm.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        Dim DTA As New DseTrans_Cuti_ApprovalTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim row As DseTrans_Cuti_Approval.tbltrans_cuti_pengajuanRow
        Dim rowD As DataGridViewRow
        Dim CboDvg As DataGridViewComboBoxCell

        Try
            vTBL.Clear()
            Frm.dgvItem.Rows.Clear()
            'TODO: This line of code loads data into the 'DseTrans_Cuti_Approval.tbltrans_cuti_pengajuan' table. You can move, or remove it, as needed.
            DTA.FillBy_Status_Approval(vTBL, vStatus)
            If Frm.dgvItem.ColumnCount > 0 Then
                Frm.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"

                For Each row In vTBL
                    Frm.dgvItem.Rows.Add()
                    rowD = Frm.dgvItem.Rows(Frm.dgvItem.RowCount - 1)
                    With row
                        rowD.Cells(Frm.ColDocumentNo.Index).Value = .no_transaksi
                        rowD.Cells(Frm.ColDate.Index).Value = .tanggal
                        rowD.Cells(Frm.ColEmployeeCode.Index).Value = .Kode_Karyawan
                        rowD.Cells(Frm.ColEmployeeName.Index).Value = .Full_Name
                        rowD.Cells(Frm.ColPrimaryLeaveAplication.Index).Value = .primary_cuti
                        rowD.Cells(Frm.ColPrimaryEmployee.Index).Value = .primary_karyawan
                        rowD.Cells(Frm.ColEnumStatus.Index).Value = .status
                        CboDvg = rowD.Cells(Frm.ColStatus.Index)
                        rowD.Cells(Frm.ColStatus.Index).Value = CboDvg.Items(.status)
                        rowD.Cells(Frm.ColPrimary.Index).Value = .primary
                    End With
                Next

                If Frm.dgvItem.RowCount > 0 Then
                    Frm.dgvItem.ClearSelection()
                    Frm.dgvItem.Rows(0).Selected = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Frm.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            DTA.Dispose()
        End Try
    End Sub

    Public Function UpdateStatusNotes(ByVal vStatusNotes As String, ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTrans_Cuti_ApprovalTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_StatusNotes(vStatusNotes, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetJmlByPrimaryEmployeeNDateItemAttendance(ByVal vPrimaryEmployee As Double, _
    ByVal vEndDate As Date, ByVal vStartDate As Date) As Double
        Dim DTA As New DseTrans_Cuti_ApprovalTableAdapters.tbltrans_attendance_itemTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlAttendanceByEmployeeNDate(vPrimaryEmployee, vEndDate, vStartDate)
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

    Public Function ApakahDataBenar(ByVal vDGVItem As DataGridView, ByVal i As FrmTransaksiCuti_Approval) As Boolean
        Dim RowDGV As DataGridViewRow
        Dim Ketemu As Boolean = False
        Try
            If vDGVItem.Rows.Count > 0 Then
                For Each RowDGV In vDGVItem.Rows
                    If RowDGV.Cells(i.ColModifiy.Index).Value = True Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    MsgBox("Please checked modify data", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            'cek apakah sudah ada attendance jika sudah maka status terakhir tidak boleh dirubah
            'MsgBox("Tolong direpair")
            If Me.GetJmlByPrimaryEmployeeNDateItemAttendance(i.TxtPrimaryEmployee.Text, i.DtpEnd.Value.Date, i.dtpStart.Value.Date) > 0 Then
                MsgBox("You can not change the approval status of leave because no attendance data for the date in question.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
