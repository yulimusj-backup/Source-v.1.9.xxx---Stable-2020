Option Explicit On
Option Strict Off
Imports System.Windows.Forms

Public Class FrmMasterCuti_Balance
    Public mPrimaryEmployee As Double = -1
    Public mPrimaryLeave As Double = -1
    Public mPeriod As Date = #1/1/1950#
    Public ObjCd As ClsCekData
    Public FrmMasterLeaveRef As FrmMasterCuti
    Public FrmMasterLeaveRef_vIndexRef As Integer = 0

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Pengajuan_Balance_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.chkDetail.Checked = True Then
                Me.chkDetail.Checked = False
            Else
                Me.OlahDataLeave_FrmMasterLeave()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub OlahDataLeave_FrmMasterLeave()
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
        Dim RowDgv, RowDgv1 As DataGridViewRow
        Dim Qty As Double = 0
        Dim ijk As Double = 0
        Dim NextInd As Double = 0
        Dim DV As DataRowView

        Try
            Me.dgvExplain.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvExplain.Rows.Clear()

            'Koding untuk mendapatkan data Leave  
            If Me.mPrimaryEmployee = -1 Then
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'header
            RowDgv = Me.FrmMasterLeaveRef.dgvEmployee.Rows(Me.FrmMasterLeaveRef_vIndexRef)
            Me.TxtQtyLeave.Text = Me.ObjCd.DigitGrouping(Me.FrmMasterLeaveRef.TxtLeaveQty.Text, 1)
            Me.TxtLeaveAddition.Text = Me.ObjCd.DigitGrouping(RowDgv.Cells(Me.FrmMasterLeaveRef.ColLeaveAddition.Index).Value, 1)
            Me.TxtLeaveDeduction.Text = Me.ObjCd.DigitGrouping(RowDgv.Cells(Me.FrmMasterLeaveRef.ColLeaveDeduction.Index).Value, 1)
            Me.TxtTotalBeginning.Text = Me.ObjCd.DigitGrouping(CDbl(Me.TxtQtyLeave.Text) + CDbl(Me.TxtLeaveAddition.Text) - CDbl(Me.TxtLeaveDeduction.Text), 1)

            'item
            DtaMaxDate.Fill(TblMaxDate, Me.ObjCd.GetLastDateOfTheYear(Me.mPeriod), Me.mPrimaryEmployee)
            DtaCutiPengajuan.Fill(TblCutiPengajuan, Me.ObjCd.GetLastDateOfTheYear(Me.mPeriod), Me.ObjCd.GetFisrtDateOfTheYear(Me.mPeriod), _
            Me.mPrimaryLeave, Me.mPrimaryEmployee)
            DtaAttend.Fill(TblAttend, Me.ObjCd.GetLastDateOfTheYear(Me.mPeriod), Me.ObjCd.GetFisrtDateOfTheYear(Me.mPeriod), _
            Me.mPrimaryLeave, Me.mPrimaryEmployee)

            'kosongkan tabel item
            TblItem.Rows.Clear()

            '1. cuti yang belum terjadi
            For Each RowCutiPengajuan In TblCutiPengajuan.Rows
                If TblMaxDate.Rows.Count > 0 Then
                    RowMaxdate = TblMaxDate.Rows(0)
                    RowCutiPengajuan.TglMAX = RowMaxdate.MAX_Date
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

            '2. cuti yang sudah terjadi
            For Each rowAttend In TblAttend.Rows
                If TblMaxDate.Rows.Count > 0 Then
                    RowMaxdate = TblMaxDate.Rows(0)
                    rowAttend.TglMAX = RowMaxdate.MAX_Date
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
            TblItem.DefaultView.Sort = "[Tanggal Transaksi], [No Document]"
            TblItem.DefaultView.RowFilter = ""
            For Each DV In TblItem.DefaultView
                RowItem = DV.Row
                Me.dgvExplain.Rows.Add()
                RowDgv = Me.dgvExplain.Rows(Me.dgvExplain.Rows.Count - 1)
                RowDgv.Cells(Me.ColNo.Index).Value = Me.dgvExplain.Rows.Count
                RowDgv.Cells(Me.ColNoDocument.Index).Value = RowItem.No_Document
                RowDgv.Cells(Me.ColTanggalTransaksi.Index).Value = Format(RowItem.Tanggal_Transaksi, "dd MMM yyyy")
                RowDgv.Cells(Me.ColQtyUsed.Index).Value = Me.ObjCd.DigitGrouping(RowItem.Qty_Used, 1)
                Qty += RowItem.Qty_Used
                RowDgv.Cells(Me.ColBalance.Index).Value = Me.ObjCd.DigitGrouping(CDbl(Me.TxtTotalBeginning.Text) - Qty, 1)
                RowDgv.Cells(Me.ColPrimaryTransaksiRujukan.Index).Value = RowItem.Primary_Transaksi_Rujukan
                RowDgv.Cells(Me.ColTipeTransaksiRujukan.Index).Value = RowItem.Tipe_Transaksi_Rujukan
            Next
            Me.TxtTotalUsed.Text = Me.ObjCd.DigitGrouping(Qty, 1)

            If Me.dgvExplain.RowCount > 0 Then
                Me.TxtTotalBalance.Text = Me.ObjCd.DigitGrouping(Me.dgvExplain.Item(Me.ColBalance.Index, Me.dgvExplain.RowCount - 1).Value, 1)
            Else
                Me.TxtTotalBalance.Text = Me.TxtTotalBeginning.Text
            End If

            'gabungkan data baris untuk beberapa data yang berasal dari sumber yang sama
            ijk = 0
            If Me.chkDetail.Checked = False Then
                While ijk <= Me.dgvExplain.Rows.Count - 2
                    RowDgv = Me.dgvExplain.Rows(ijk)
                    NextInd = ijk + 1

                    If CDbl(RowDgv.Cells(Me.ColQtyUsed.Index).Value) = 0 Then
                        Me.dgvExplain.Rows.Remove(RowDgv)
                    Else
                        While NextInd <= Me.dgvExplain.Rows.Count - 1
                            RowDgv1 = Me.dgvExplain.Rows(NextInd)
                            If RowDgv.Cells(Me.ColPrimaryTransaksiRujukan.Index).Value = RowDgv1.Cells(Me.ColPrimaryTransaksiRujukan.Index).Value And _
                            RowDgv.Cells(Me.ColTipeTransaksiRujukan.Index).Value = RowDgv1.Cells(Me.ColTipeTransaksiRujukan.Index).Value And _
                            RowDgv.Cells(Me.ColNoDocument.Index).Value = RowDgv1.Cells(Me.ColNoDocument.Index).Value Then
                                RowDgv.Cells(Me.ColQtyUsed.Index).Value = Me.ObjCd.DigitGrouping(CDbl(RowDgv.Cells(Me.ColQtyUsed.Index).Value) + _
                                CDbl(RowDgv1.Cells(Me.ColQtyUsed.Index).Value), 1)
                                RowDgv.Cells(Me.ColBalance.Index).Value = Me.ObjCd.DigitGrouping(CDbl(RowDgv.Cells(Me.ColBalance.Index).Value) - _
                                CDbl(RowDgv1.Cells(Me.ColQtyUsed.Index).Value), 1)

                                If IsNothing(RowDgv.Cells(Me.ColTanggalTransaksi_End.Index).Value) = True Then
                                    RowDgv.Cells(Me.ColTanggalTransaksi_End.Index).Value = RowDgv1.Cells(Me.ColTanggalTransaksi.Index).Value
                                Else
                                    If CDate(RowDgv.Cells(Me.ColTanggalTransaksi_End.Index).Value) < CDate(RowDgv1.Cells(Me.ColTanggalTransaksi.Index).Value) Then
                                        RowDgv.Cells(Me.ColTanggalTransaksi_End.Index).Value = RowDgv1.Cells(Me.ColTanggalTransaksi.Index).Value
                                    End If
                                End If

                                Me.dgvExplain.Rows.Remove(RowDgv1)
                            Else
                                NextInd += 1
                            End If
                        End While

                        ijk += 1
                    End If
                End While
                Me.ColTanggalTransaksi.HeaderText = "Start Date"
                Me.ColTanggalTransaksi_End.HeaderText = "End Date"
                Me.ColTanggalTransaksi_End.Visible = True
            Else
                While ijk <= Me.dgvExplain.Rows.Count - 1
                    RowDgv = Me.dgvExplain.Rows(ijk)

                    If CDbl(RowDgv.Cells(Me.ColQtyUsed.Index).Value) = 0 Then
                        Me.dgvExplain.Rows.Remove(RowDgv)
                    Else
                        ijk += 1
                    End If
                End While
                Me.ColTanggalTransaksi.HeaderText = "Date"
                Me.ColTanggalTransaksi_End.Visible = False
            End If

            'isikan data untuk end date yg sama dengan begin date dan hanya ada 1 baris saja
            For Each RowDgv In Me.dgvExplain.Rows
                If IsNothing(RowDgv.Cells(Me.ColTanggalTransaksi_End.Index).Value) = True Then
                    RowDgv.Cells(Me.ColTanggalTransaksi_End.Index).Value = RowDgv.Cells(Me.ColTanggalTransaksi.Index).Value
                Else
                    If String.IsNullOrEmpty(RowDgv.Cells(Me.ColTanggalTransaksi_End.Index).Value.ToString) = True Then
                        RowDgv.Cells(Me.ColTanggalTransaksi_End.Index).Value = RowDgv.Cells(Me.ColTanggalTransaksi.Index).Value
                    End If
                End If
            Next

            Me.ObjCd.SetNumber(Me.dgvExplain)
            Me.dgvExplain.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaAttend.Dispose()
            TblAttend.Dispose()
            DtaCutiPengajuan.Dispose()
            TblCutiPengajuan.Dispose()
            DtaMaxDate.Dispose()
            TblAttend.Dispose()
        End Try
    End Sub

    Private Sub chkDetail_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDetail.CheckedChanged
        Try
            Me.OlahDataLeave_FrmMasterLeave()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class