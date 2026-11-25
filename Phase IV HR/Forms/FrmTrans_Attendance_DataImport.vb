Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms
Public Class FrmTrans_Attendance_DataImport
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjQuick As New ClsTrans_Attendance_quick
    Dim ObjAtt As New ClsTrans_Attendance
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub BtnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShow.Click
        Try
            Me.ObjQuick.Show(Me.DtpPeriod_Attd.Value, Me.DtpTo_Attd.Value, Me)

            ObjSettingUser.SetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_Quick_Type_Filter, Me.CboTypeDateFilter.SelectedIndex, Me.UserCode)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttendanceData_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVAttendanceData.CellContentClick
        Dim ObjKary As New Core_Module.ClsMasterKaryawan
        Dim TblOT_Order As New DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable
        Dim frmAtt As New FrmTrans_Attendance

        Try
            If e.ColumnIndex = Me.ColEdit.Index Then
                If Me.ObjAtt.GetJmlTransAtt(Me.DGVAttendanceData.Item(Me.ColPrimaryTransaksi.Index, e.RowIndex).Value) > 0 Then
                    'Parameter yang menunjang pengecekan Apakah Boleh Edit
                    Me.ObjAtt.Primary = Me.DGVAttendanceData.Item(Me.ColPrimaryTransaksi.Index, e.RowIndex).Value
                    Me.ObjAtt.PrimaryEmp = ObjKary.GetPrimaryByCode(Me.DGVAttendanceData.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value)
                    Me.ObjAtt.StrDate = Me.DGVAttendanceData.Item(Me.ColStartDate.Index, e.RowIndex).Value
                    Me.ObjAtt.EndDate = Me.DGVAttendanceData.Item(Me.ColEndDate.Index, e.RowIndex).Value
                    '------------------------------------------------------
                    If Me.ObjAtt.ApakahBolehEdit(Me.DGVAttendanceData.Item(Me.ColPrimaryTransaksi.Index, e.RowIndex).Value, _
                    Me.UserCode, False) = True Then
                        Dim frm As New FrmTrans_Attendance()
                        frm.ObjAtt.CariDetail(Me.DGVAttendanceData.Item(Me.ColPrimaryTransaksi.Index, e.RowIndex).Value, frm, TblOT_Order, False)
                        frm.SalinKeField()
                        frm.DariQuick = True
                        frm.Baru = False
                        frm.UserCode = Me.UserCode
                        frm.BackColorAll = Me.BackColorAll
                        frm.ShowDialog()
                        frm.Dispose()
                    End If
                Else
                    MsgBox("This transaction is already deleted.", MsgBoxStyle.Information, "Information")
                    Me.BtnShow.PerformClick()
                End If
            ElseIf e.ColumnIndex = Me.ColDocumentNo.Index Then
                Dim frm As New FrmTrans_Attendance()
                frm.BackColorAll = Me.BackColorAll
                frm.UserCode = Me.UserCode
                frm.ShowAtt_FromQuick = True
                frm.PrimaryAtt_FromQuick = Me.DGVAttendanceData.Item(Me.ColPrimaryTransaksi.Index, e.RowIndex).Value
                frm.ShowDialog()
                frm.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblChecked_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblChecked.LinkClicked
        Dim row As DataGridViewRow
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each row In DGVAttendanceData.Rows
                row.Cells(Me.ColChkPrint.Index).Value = True
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LblUnchecked_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUnchecked.LinkClicked
        Dim row As DataGridViewRow
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each row In DGVAttendanceData.Rows
                row.Cells(Me.ColChkPrint.Index).Value = False
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LblCheckedDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblCheckedDelete.LinkClicked
        Dim row As DataGridViewRow
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each row In DGVAttendanceData.Rows
                row.Cells(Me.ColChkDelete.Index).Value = True
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LblUncheckedDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUncheckedDelete.LinkClicked
        Dim row As DataGridViewRow
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each row In DGVAttendanceData.Rows
                row.Cells(Me.ColChkDelete.Index).Value = False
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNew.Click
        Try
            Me.Text = "Quick Attendance"
            Dim frm As New FrmTrans_Attendance_Quick2
            frm.User = Me.UserCode
            frm.BackColorAll = Me.BackColorAll
            frm.dtpPeriod.Value = Me.DtpPeriod_Attd.Value
            frm.DtpTo.Value = Me.DtpTo_Attd.Value
            frm.ShowDialog()

            'add by ahdar
            If Me.CboTypeDateFilter.SelectedIndex = 0 Then
                Me.DtpPeriod_Attd.Value = frm.DtpTanggalTransaksi.Value
                Me.DtpTo_Attd.Value = frm.DtpTanggalTransaksi.Value
            Else
                Me.DtpPeriod_Attd.Value = frm.dtpPeriod.Value
                Me.DtpTo_Attd.Value = frm.DtpTo.Value
            End If

            Me.BtnShow.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_DataImport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CboTypeDateFilter.Items.Add("Transaction Date")
            Me.CboTypeDateFilter.Items.Add("Period")
            Me.CboTypeDateFilter.SelectedIndex = 0

            Me.DGVAttendanceData.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Dim UserSetting As String = ObjSettingUser.GetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_Quick_Type_Filter, Me.UserCode)
            If UserSetting = "1" Then
                Me.CboTypeDateFilter.SelectedIndex = 1
            Else
                Me.CboTypeDateFilter.SelectedIndex = 0
            End If

            Me.ColEmployeeCode.Visible = True
            Me.ColDocumentNo.Frozen = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahBolehPrint() As Boolean
        Dim RowDGv As DataGridViewRow
        Dim Ketemu As Boolean = False
        Try
            'Buat Batasan Apakah Boleh Print, sama dengan batasan quick salary
            If Me.DtpPeriod_Attd.Value.Date > DtpTo_Attd.Value Then
                MsgBox("Period to cannot less than period from.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.DGVAttendanceData.Rows.Count <= 0 Then
                MsgBox("Please input data item.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.DGVAttendanceData.Rows.Count > 0 Then
                For Each RowDGv In Me.DGVAttendanceData.Rows
                    If RowDGv.Cells(Me.ColChkPrint.Index).Value = True Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    MsgBox("Please checked coloumn print.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CekApakahBolehHapus(Optional ByVal vShowMessage As Boolean = True) As Boolean
        Dim RowDGV As DataGridViewRow
        Dim Ketemu As Boolean = False
        Try
            If Me.DGVAttendanceData.Rows.Count > 0 Then
                For Each RowDGV In Me.DGVAttendanceData.Rows
                    If RowDGV.Cells(Me.ColChkDelete.Index).Value = True Then
                        Ketemu = True
                        Exit For
                    End If
                Next
            End If
            If Ketemu = False Then
                If vShowMessage = True Then
                    MsgBox("Please choose at least 1 data to delete.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click
        Dim FrmError As New FrmTrans_Gaji_MsgBox
        Dim TBLMsgBox As New DseTrans_Attendance.TBLMessageBoxDataTable
        Dim ListPrimaryAttn As New List(Of Double)
        Dim PrimaryAttn As Double = -1
        Dim Hasil As MsgBoxResult
        Dim ObjHapusAtt As New ClsTrans_Attendance
        Dim TblOT_Order As New DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable
        Dim frmAtt As New FrmTrans_Attendance

        Try
            If Me.CekApakahBolehHapus() = True Then
                Me.ObjQuick.ApakahBolehHapus_Quick(ListPrimaryAttn, TBLMsgBox, True, Me.UserCode, Me)
                If TBLMsgBox.Rows.Count > 0 Then
                    TBLMsgBox.DefaultView.Sort = "[Employee Code], [Employee Name]"
                    FrmError.dgvItem.DataSource = TBLMsgBox.DefaultView
                    FrmError.dgvItem.Columns(0).Visible = False
                    FrmError.LblProgress.Text = "Total Rows = " + TBLMsgBox.Rows.Count.ToString
                    FrmError.Text = "The following employee list can't be deleted"
                    FrmError.ShowDialog()
                End If
            End If
            If Me.CekApakahBolehHapus(False) = True Then
                If ListPrimaryAttn.Count > 0 Then
                    Hasil = MsgBox("Delete data?", MsgBoxStyle.YesNo, "Confirm")
                    If Hasil = MsgBoxResult.Yes Then
                        For Each PrimaryAttn In ListPrimaryAttn
                            'delete data
                            ObjHapusAtt.Primary = PrimaryAttn
                            ObjHapusAtt.StrDate = ObjHapusAtt.GetStartPeriod(PrimaryAttn)
                            ObjHapusAtt.EndDate = ObjHapusAtt.GetEndPeriod(PrimaryAttn)
                            ObjHapusAtt.PrimaryEmp = ObjHapusAtt.GetPrimaryEmployeeAttendance(PrimaryAttn)
                            'ObjAtt.IsiTblSpl(ObjHapusAtt.PrimaryEmp, ObjHapusAtt.StrDate.Date, ObjHapusAtt.EndDate.Date, TblOT_Order)
                            ObjHapusAtt.Hapus(PrimaryAttn, Me.UserCode, frmAtt, TblOT_Order, True)
                        Next
                        MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        Me.BtnShow.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLMsgBox.Dispose()
            FrmError.Dispose()
            ListPrimaryAttn = Nothing
            ObjHapusAtt = Nothing
        End Try
    End Sub

    Private Sub BtnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Dim FrmCetak As New FrmTrans_Attendance_Cetak
        Dim FrmOpt As New FrmTrans_Attendance_Cetak_Opt
        Dim RowDgv As DataGridViewRow
        Dim RowSetting As DseTrans_Attendance_Cetak_IGI.TblSettingRow

        Try
            If Me.ApakahBolehPrint() = True Then
                FrmCetak.mShowQuickAttendance = True
                FrmCetak.mTblSetting.Rows.Clear()
                If Me.DGVAttendanceData.Rows.Count > 0 Then
                    For Each RowDgv In Me.DGVAttendanceData.Rows
                        If RowDgv.Cells(Me.ColChkPrint.Index).Value = True Then
                            RowSetting = FrmCetak.mTblSetting.NewTblSettingRow
                            RowSetting.primary_transaksi = RowDgv.Cells(Me.ColPrimaryTransaksi.Index).Value
                            RowSetting.no_transaksi = RowDgv.Cells(Me.ColNo.Index).Value
                            RowSetting.end_date = RowDgv.Cells(Me.ColEndDate.Index).Value
                            RowSetting.start_date = RowDgv.Cells(Me.ColStartDate.Index).Value
                            RowSetting.Employee_Code = RowDgv.Cells(Me.ColEmployeeCode.Index).Value
                            RowSetting.index = RowDgv.Index
                            FrmCetak.mTblSetting.Rows.Add(RowSetting)
                        End If
                    Next
                    FrmCetak.mIndexActual = 0
                    FrmCetak.mIndexMin = 0
                    FrmCetak.mIndexMax = FrmCetak.mTblSetting.Rows.Count - 1
                End If
                'Showdialog
                FrmOpt.ShowDialog()
                FrmCetak.FormatCetak = FrmOpt.CboFormat.SelectedIndex
                FrmCetak.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCetak.Dispose()
            FrmOpt.Dispose()
        End Try
    End Sub

    'Private Sub BtnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
    '    Dim FrmCetak As New FrmTrans_Attendance_Cetak
    '    Dim RowDgv As DataGridViewRow
    '    Dim RowSetting As DseTrans_Attendance_Cetak_IGI.TblSettingRow
    '    Try
    '        If Me.ApakahBolehPrint() = True Then
    '            FrmCetak.mShowQuickAttendance = True
    '            FrmCetak.mTblSetting.Rows.Clear()
    '            If Me.DGVAttendanceData.Rows.Count > 0 Then
    '                For Each RowDgv In Me.DGVAttendanceData.Rows
    '                    If RowDgv.Cells(Me.ColChkPrint.Index).Value = True Then
    '                        RowSetting = FrmCetak.mTblSetting.NewTblSettingRow
    '                        RowSetting.primary_transaksi = RowDgv.Cells(Me.ColPrimaryTransaksi.Index).Value
    '                        RowSetting.no_transaksi = RowDgv.Cells(Me.ColNo.Index).Value
    '                        RowSetting.end_date = RowDgv.Cells(Me.ColEndDate.Index).Value
    '                        RowSetting.start_date = RowDgv.Cells(Me.ColStartDate.Index).Value
    '                        RowSetting.Employee_Code = RowDgv.Cells(Me.ColEmployeeCode.Index).Value
    '                        RowSetting.index = RowDgv.Index
    '                        FrmCetak.mTblSetting.Rows.Add(RowSetting)
    '                    End If
    '                Next
    '                FrmCetak.mIndexActual = 0
    '                FrmCetak.mIndexMin = 0
    '                FrmCetak.mIndexMax = FrmCetak.mTblSetting.Rows.Count - 1
    '            End If
    '            'Showdialog
    '            FrmCetak.ShowDialog()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        FrmCetak.Dispose()
    '    End Try
    'End Sub

    Private Sub DGVAttendanceData_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVAttendanceData.Sorted
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            ObjCD.SetNumber(Me.DGVAttendanceData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Private Sub llbEmail_CekAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbEmail_CekAll.LinkClicked
        Dim row As DataGridViewRow
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each row In DGVAttendanceData.Rows
                row.Cells(Me.ColChkEmail.Index).Value = True
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub llbEmail_UncekAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbEmail_UncekAll.LinkClicked
        Dim row As DataGridViewRow
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each row In DGVAttendanceData.Rows
                row.Cells(Me.ColChkEmail.Index).Value = False
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DGVAttendanceData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DGVAttendanceData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnSendEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSendEmail.Click
        Dim ij As New FrmTrans_Attendance_SendEmail
        Dim ListPrimary As New List(Of Double)
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DGVAttendanceData.Rows
                If rowDGV.Cells(Me.ColChkEmail.Index).Value = Boolean.TrueString Then
                    ListPrimary.Add(rowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value)
                End If
            Next

            If ListPrimary.Count > 0 Then
                ij.UserCode = Me.UserCode
                ij.BackColorAll = Me.BackColorAll
                ij.ListPrimaryAttendance = ListPrimary
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub lllSetDataSorting_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lllSetDataSorting.LinkClicked
        Dim i As New FrmTransaksiQuick_Sorting

        Try
            i.UserCode = Me.UserCode
            i.BackColorAll = Me.BackColorAll
            i.TipeTransaksi = ClsTransaksiQuick_Sorting.EnumTipeTransaksi.QuickAttendance
            i.Text = "Quick Attendance Data Sorting"
            i.ShowDialog()
            If i.ShowUlang = True Then
                Me.BtnShow.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class