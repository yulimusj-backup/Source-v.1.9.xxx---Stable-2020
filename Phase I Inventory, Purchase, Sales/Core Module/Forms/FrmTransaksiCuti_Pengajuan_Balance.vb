Option Explicit On
Option Strict Off
Imports System.Windows.Forms

Public Class FrmTransaksiCuti_Pengajuan_Balance
    Public mPrimaryEmployee As Double = -1
    Public mPrimaryLeave As Double = -1
    Public mPeriod As Date = #1/1/1950#
    Public ObjCd As ClsCekData

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Pengajuan_Balance_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.OlahDataLeave(Me.mPrimaryEmployee, Me.mPrimaryLeave)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub OlahDataLeave(ByVal vPrimaryEmployee As Double, ByVal vPrimaryLeave As Double)
        Dim DtaHeader As New DseMasterCuti_SetBalanceTableAdapters.tblmaster_cuti_beginning1_headerTableAdapter
        Dim TblHeader As New DseMasterCuti_SetBalance.tblmaster_cuti_beginning1_headerDataTable
        Dim RowHeader As DseMasterCuti_SetBalance.tblmaster_cuti_beginning1_headerRow
        Dim DtaMaxDate As New DseMasterCuti_SetBalanceTableAdapters.tblMAX_Date1TableAdapter
        Dim TblMaxDate As New DseMasterCuti_SetBalance.tblMAX_Date1DataTable
        Dim RowMaxdate As DseMasterCuti_SetBalance.tblMAX_Date1Row
        Dim DtaCutiPengajuan As New DseMasterCuti_SetBalanceTableAdapters.tbltrans_cuti_pengajuan1TableAdapter
        Dim TblCutiPengajuan As New DseMasterCuti_SetBalance.tbltrans_cuti_pengajuan1DataTable
        Dim RowCutiPengajuan As DseMasterCuti_SetBalance.tbltrans_cuti_pengajuan1Row
        Dim DtaAttend As New DseMasterCuti_SetBalanceTableAdapters.TblTrans_Attendance1TableAdapter
        Dim TblAttend As New DseMasterCuti_SetBalance.TblTrans_Attendance1DataTable
        Dim rowAttend As DseMasterCuti_SetBalance.TblTrans_Attendance1Row
        Dim TblItem As New DseMasterCuti_SetBalance.tblmaster_cuti_beginning1_itemDataTable
        Dim RowItem As DseMasterCuti_SetBalance.tblmaster_cuti_beginning1_itemRow
        Dim RowDgv As DataGridViewRow
        Dim Qty As Double = 0
        Try
            'Koding untuk mendapatkan data Leave  
            If vPrimaryEmployee = -1 Or vPrimaryLeave = -1 Then
                Exit Sub
            End If
            'header
            DtaHeader.Fill(TblHeader, Me.mPrimaryLeave, Me.mPrimaryEmployee)
            If TblHeader.Rows.Count < 1 Then
                Exit Sub
            End If
            RowHeader = TblHeader.Rows(0)
            If RowHeader.leave_qty = CInt(RowHeader.leave_qty) Then
                Me.TxtQtyLeave.Text = Me.ObjCd.DigitGrouping(RowHeader.leave_qty, 0)
            Else
                Me.TxtQtyLeave.Text = Me.ObjCd.DigitGrouping(RowHeader.leave_qty, 1)
            End If
            If RowHeader.leave_addition = CInt(RowHeader.leave_addition) Then
                Me.TxtLeaveAddition.Text = Me.ObjCd.DigitGrouping(RowHeader.leave_deduction, 0)
            Else
                Me.TxtLeaveAddition.Text = Me.ObjCd.DigitGrouping(RowHeader.leave_addition, 1)
            End If
            If RowHeader.leave_deduction = CInt(RowHeader.leave_deduction) Then
                Me.TxtLeaveDeduction.Text = Me.ObjCd.DigitGrouping(RowHeader.leave_deduction, 0)
            Else
                Me.TxtLeaveDeduction.Text = Me.ObjCd.DigitGrouping(RowHeader.leave_deduction, 1)
            End If
            If RowHeader.Begining_Qty = CInt(RowHeader.Begining_Qty) Then
                Me.TxtTotalBeginning.Text = Me.ObjCd.DigitGrouping(RowHeader.Begining_Qty, 0)
            Else
                Me.TxtTotalBeginning.Text = Me.ObjCd.DigitGrouping(RowHeader.Begining_Qty, 1)
            End If

            'item
            DtaMaxDate.Fill(TblMaxDate, Me.ObjCd.GetLastDateOfTheYear(Me.mPeriod), Me.mPrimaryEmployee)
            DtaCutiPengajuan.Fill(TblCutiPengajuan, Me.ObjCd.GetLastDateOfTheYear(Me.mPeriod), Me.ObjCd.GetFisrtDateOfTheYear(Me.mPeriod), _
            Me.mPrimaryLeave, Me.mPrimaryEmployee)
            DtaAttend.Fill(TblAttend, Me.ObjCd.GetLastDateOfTheYear(Me.mPeriod), Me.ObjCd.GetFisrtDateOfTheYear(Me.mPeriod), _
            Me.mPrimaryLeave, Me.mPrimaryEmployee)
            'kosongkan tabel item
            TblItem.Rows.Clear()
            'cuti yang belum terjadi
            For Each RowCutiPengajuan In TblCutiPengajuan.Rows
                If TblMaxDate.Rows.Count > 0 Then
                    RowMaxdate = TblMaxDate.Rows(0)
                    RowCutiPengajuan._date = RowMaxdate.MAX_Date
                End If
                'jika date > maxdate
                If RowCutiPengajuan._date > RowCutiPengajuan.TglMAX Then
                    RowItem = TblItem.NewRow
                    RowItem.No_Document = RowCutiPengajuan.No_Document
                    RowItem.Tanggal_Transaksi = RowCutiPengajuan._date
                    RowItem.Qty_Used = RowCutiPengajuan.duration
                    RowItem.Primary_Transaksi_Rujukan = RowCutiPengajuan.primary_transaksi
                    RowItem.Tipe_Transaksi_Rujukan = 0
                    TblItem.Addtblmaster_cuti_beginning1_itemRow(RowItem)
                End If
            Next
            'cuti yang sudah terjadi
            For Each rowAttend In TblAttend.Rows
                If TblMaxDate.Rows.Count > 0 Then
                    RowMaxdate = TblMaxDate.Rows(0)
                    rowAttend._date = RowMaxdate.MAX_Date
                End If
                'jika date <= maxdate
                If rowAttend._date <= rowAttend.TglMAX Then
                    RowItem = TblItem.NewRow
                    RowItem.No_Document = rowAttend.No_Document
                    RowItem.Tanggal_Transaksi = rowAttend._date
                    RowItem.Qty_Used = rowAttend.is_cuti
                    RowItem.Primary_Transaksi_Rujukan = rowAttend.primary_transaksi
                    RowItem.Tipe_Transaksi_Rujukan = 0
                    TblItem.Addtblmaster_cuti_beginning1_itemRow(RowItem)
                End If
            Next
            'sort
            Qty = 0
            For Each RowItem In TblItem.Rows
                Me.dgvExplain.Rows.Add()
                RowDgv = Me.dgvExplain.Rows(Me.dgvExplain.Rows.Count - 1)
                RowDgv.Cells(Me.ColNo.Index).Value = Me.dgvExplain.Rows.Count
                RowDgv.Cells(Me.ColNoDocument.Index).Value = RowItem.No_Document
                RowDgv.Cells(Me.ColTanggalTransaksi.Index).Value = Format(RowItem.Tanggal_Transaksi, "dd MMM yyyy")
                If RowItem.Qty_Used = CInt(RowItem.Qty_Used) Then
                    RowDgv.Cells(Me.ColQtyUsed.Index).Value = Me.ObjCd.DigitGrouping(RowItem.Qty_Used, 0)
                Else
                    RowDgv.Cells(Me.ColQtyUsed.Index).Value = Me.ObjCd.DigitGrouping(RowItem.Qty_Used, 1)
                End If
                Qty += RowItem.Qty_Used
                If (CDbl(Me.TxtTotalBeginning.Text) - Qty) = CInt(CDbl(Me.TxtTotalBeginning.Text) - Qty) Then
                    RowDgv.Cells(Me.ColBalance.Index).Value = Me.ObjCd.DigitGrouping(CDbl(Me.TxtTotalBeginning.Text) - Qty, 0)
                Else
                    RowDgv.Cells(Me.ColBalance.Index).Value = Me.ObjCd.DigitGrouping(CDbl(Me.TxtTotalBeginning.Text) - Qty, 1)
                End If
                RowDgv.Cells(Me.ColPrimaryTransaksiRujukan.Index).Value = RowItem.Primary_Transaksi_Rujukan
                RowDgv.Cells(Me.ColTipeTransaksiRujukan.Index).Value = RowItem.Tipe_Transaksi_Rujukan
            Next

            If Qty = CInt(Qty) Then
                Me.TxtTotalUsed.Text = Me.ObjCd.DigitGrouping(Qty, 0)
            Else
                Me.TxtTotalUsed.Text = Me.ObjCd.DigitGrouping(Qty, 1)
            End If

            If Me.dgvExplain.RowCount > 0 Then
                Me.TxtTotalBalance.Text = Me.ObjCd.DigitGrouping(Me.dgvExplain.Item(Me.ColBalance.Index, Me.dgvExplain.RowCount - 1).Value, 0)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaHeader.Dispose()
            TblHeader.Dispose()
            DtaAttend.Dispose()
            TblAttend.Dispose()
            DtaCutiPengajuan.Dispose()
            TblCutiPengajuan.Dispose()
            DtaMaxDate.Dispose()
            TblAttend.Dispose()
        End Try
    End Sub
End Class