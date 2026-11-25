Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class FrmTransaksiCuti_Approval
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjApp As New ClsTrans_Cuti_Approval
    Dim ObjCuti As New Core_Module.ClsMasterCuti_Item
    Dim Status As ClsTrans_Cuti_Approval.EnumStatus
    Dim Buffer As String
    Dim Buffer_Status As Integer = 0
    Dim Buffer_date As Date = #1/1/1950#
    Dim TBL As New DseTrans_Cuti_Approval.tbltrans_cuti_pengajuanDataTable

    Dim mPrimaryCuti As Double = -1

    Public ij As New FrmTransaksiCuti_Pengajuan_Balance

    Public Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.GroupBox1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.GroupBox2.Controls, True, True, True)
            Me.TBL.Rows.Clear()
            'Me.cboLeave.DataSource = Nothing
            'Me.cboLeave.Items.Clear()
            Me.TxtPrimary.Text = "-1"
            Me.DtpDate.Value = Now
            Me.TxtPrimaryEmployee.Text = "-1"
            'Me.TxtPrimaryLeave.Text = "-1"
            Me.dtpStart.Value = Now
            Me.DtpEnd.Value = Now
            Me.ChkHDStart.Checked = False
            Me.ChkHDEnd.Checked = False
            Me.TxtDuration.Text = "0.0"
            Me.TxtBalance.Text = "0.0"
            Me.DtpApprovedBy.Value = Now
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KendalikanButton(ByVal vEnumStatus As ClsTrans_Cuti_Approval.EnumStatus)
        Try
            Me.BtnApprove.Enabled = False
            Me.BtnReject.Enabled = False
            Me.BtnWait.Enabled = False
            If vEnumStatus = ClsTrans_Cuti_Approval.EnumStatus.WaitingForApproval Then
                Me.BtnApprove.Enabled = True
                Me.BtnReject.Enabled = True
            End If
            If vEnumStatus = ClsTrans_Cuti_Approval.EnumStatus.Approved Then
                Me.BtnWait.Enabled = True
                Me.BtnReject.Enabled = True
            End If
            If vEnumStatus = ClsTrans_Cuti_Approval.EnumStatus.Rejected Then
                Me.BtnApprove.Enabled = True
                Me.BtnWait.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SettingDefaultValue()
        Try
            Me.RboWaitingForApproval.Checked = True
            Me.KendalikanButton(ClsTrans_Cuti_Approval.EnumStatus.WaitingForApproval)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Approval_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            If Me.RboApproved.Checked = True Or Me.RboRejected.Checked = True Or Me.RboWaitingForApproval.Checked = True Then
                Me.LoadData()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Approval_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Status = ClsTrans_Cuti_Approval.EnumStatus.WaitingForApproval
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.SettingDefaultValue()
            Me.ObjApp.LoadData(Me, ClsTrans_Cuti_Approval.EnumStatus.WaitingForApproval, Me.TBL)
            'Me.ObjApp.LoadData(TBL, Me.dgvItem, ClsTrans_Cuti_Approval.EnumStatus.WaitingForApproval, Me)
            Me.Text = "Leave Approval"
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboWaitingForApproval_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboWaitingForApproval.Enter
        Try
            Me.Buffer_Status = Me.Status
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboWaitingForApproval_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboWaitingForApproval.CheckedChanged
        Try
            If Me.RboWaitingForApproval.Checked = True Then
                Me.Status = ClsTrans_Cuti_Approval.EnumStatus.WaitingForApproval
                If Me.Buffer_Status <> Me.Status Then
                    Me.LoadData()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboApproved_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboApproved.Enter
        Try
            Me.Buffer_Status = Me.Status
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadData()
       
        Try
            If Me.RboApproved.Checked = True Then
                Me.ObjApp.LoadData(Me, ClsTrans_Cuti_Approval.EnumStatus.Approved, Me.TBL)
                Me.KendalikanButton(ClsTrans_Cuti_Approval.EnumStatus.Approved)
            ElseIf Me.RboRejected.Checked = True Then
                Me.ObjApp.LoadData(Me, ClsTrans_Cuti_Approval.EnumStatus.Rejected, Me.TBL)
                Me.KendalikanButton(ClsTrans_Cuti_Approval.EnumStatus.Rejected)
            Else
                Me.ObjApp.LoadData(Me, ClsTrans_Cuti_Approval.EnumStatus.WaitingForApproval, Me.TBL)
                Me.KendalikanButton(ClsTrans_Cuti_Approval.EnumStatus.WaitingForApproval)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboApproved_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboApproved.CheckedChanged
        Try
            If Me.RboApproved.Checked = True Then
                Me.Status = ClsTrans_Cuti_Approval.EnumStatus.Approved
                If Me.Buffer_Status <> Me.Status Then
                    Me.LoadData()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboRejected_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboRejected.Enter
        Try
            Me.Buffer_Status = Me.Status
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboRejected_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboRejected.CheckedChanged
        Try
            If Me.RboRejected.Checked = True Then
                Me.Status = ClsTrans_Cuti_Approval.EnumStatus.Rejected
                If Me.Buffer_Status <> Me.Status Then
                    Me.LoadData()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColModifiy.Index Then
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
        Dim Filter As String = ""
        Try
            If Me.dgvItem.Rows.Count > 0 Then
                If e.RowIndex = -1 Then
                    Exit Sub
                End If

                If e.ColumnIndex = Me.ColModifiy.Index Then
                    Dim RowDGV As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
                    If RowDGV.Cells(Me.ColModifiy.Index).Value = False Then
                        RowDGV.Cells(Me.ColModifiy.Index).Value = True
                    Else
                        RowDGV.Cells(Me.ColModifiy.Index).Value = False
                    End If
                End If

                If e.ColumnIndex = Me.ColStatusNotes.Index Then
                    Dim ij As New FrmTransaksiCuti_Approval_StatusNotes_1
                    If e.ColumnIndex = Me.ColStatusNotes.Index Then
                        If e.RowIndex = -1 Then
                            Exit Sub
                        End If
                        ij.mPrimaryTrans = Me.dgvItem(Me.ColPrimary.Index, e.RowIndex).Value
                        ij.FrmUtama = Me
                        ij.BackColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        Me.ObjApp.LoadData(Me, Me.Status, Me.TBL)
                        'Me.ObjApp.LoadData(Me.TBL, Me.dgvItem, Me.Status, Me)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnApprove.Click
        Dim RowDGV As DataGridViewRow
        Dim Primary As Double
        Dim ListPrimary As New List(Of Double)
        Dim Hasil As MsgBoxResult
        Dim StrCutiNotUpd As String = ""
        Dim CountSave As Integer = 0

        Try
            If Me.dgvItem.Rows.Count > 0 Then
                If Me.ObjApp.ApakahDataBenar(Me.dgvItem, Me) = True Then
                    Hasil = MsgBox("Approve ?", MsgBoxStyle.YesNo, "Confirm")
                    If Hasil = MsgBoxResult.Yes Then
                        Dim ij As New Core_Module.DialogPassword
                        ij.txtUserName.Text = Me.UserCode
                        ij.ShowDialog()
                        If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                            For Each RowDGV In Me.dgvItem.Rows
                                If RowDGV.Cells(Me.ColModifiy.Index).Value = True Then
                                    If ListPrimary.Contains(RowDGV.Cells(Me.ColPrimary.Index).Value) = False Then
                                        ListPrimary.Add(RowDGV.Cells(Me.ColPrimary.Index).Value)
                                        If Me.ObjApp.SimpanBaru(Me.UserCode, ClsTrans_Cuti_Approval.EnumStatus.Approved, Now, Me.ObjCD, _
                                           RowDGV.Cells(Me.ColPrimary.Index).Value, Me, RowDGV.Index) = False Then
                                            If StrCutiNotUpd = "" Then
                                                StrCutiNotUpd = "Cannot modify leave data because referring to transaction attendance:"
                                                StrCutiNotUpd &= vbCrLf & RowDGV.Cells(Me.ColDocumentNo.Index).Value
                                            Else
                                                StrCutiNotUpd &= vbCrLf & RowDGV.Cells(Me.ColDocumentNo.Index).Value
                                            End If
                                        Else
                                            CountSave += 1
                                        End If
                                    End If
                                End If
                            Next
                            If ListPrimary.Count > 0 Then
                                If StrCutiNotUpd <> "" Then
                                    MsgBox(StrCutiNotUpd, MsgBoxStyle.Information, "Information")
                                End If
                                If CountSave > 0 Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.KosongkanField()
                                    Me.ObjApp.LoadData(Me, Me.Status, Me.TBL)
                                    'Me.ObjApp.LoadData(Me.TBL, Me.dgvItem, Me.Status, Me)
                                Else
                                    MsgBox("No data stored", MsgBoxStyle.Information, "Information")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub BtnReject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnReject.Click
        Dim RowDGV As DataGridViewRow
        Dim Primary As Double
        Dim ListPrimary As New List(Of Double)
        Dim Hasil As MsgBoxResult
        Dim StrCutiNotUpd As String = ""
        Dim CountSave As Integer = 0

        Try
            If Me.dgvItem.Rows.Count > 0 Then
                If Me.ObjApp.ApakahDataBenar(Me.dgvItem, Me) = True Then
                    Hasil = MsgBox("Reject ?", MsgBoxStyle.YesNo, "Confirm")
                    If Hasil = MsgBoxResult.Yes Then
                        Dim ij As New Core_Module.DialogPassword
                        ij.txtUserName.Text = Me.UserCode
                        ij.ShowDialog()
                        If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                            For Each RowDGV In Me.dgvItem.Rows
                                If RowDGV.Cells(Me.ColModifiy.Index).Value = True Then
                                    If ListPrimary.Contains(RowDGV.Cells(Me.ColPrimary.Index).Value) = False Then
                                        ListPrimary.Add(RowDGV.Cells(Me.ColPrimary.Index).Value)
                                        If Me.ObjApp.SimpanBaru(Me.UserCode, ClsTrans_Cuti_Approval.EnumStatus.Rejected, Now, Me.ObjCD, _
                                           RowDGV.Cells(Me.ColPrimary.Index).Value, Me, RowDGV.Index) = False Then
                                            If StrCutiNotUpd = "" Then
                                                StrCutiNotUpd = "Cannot modify leave data because referring to transaction attendance:"
                                                StrCutiNotUpd &= vbCrLf & RowDGV.Cells(Me.ColDocumentNo.Index).Value
                                            Else
                                                StrCutiNotUpd &= vbCrLf & RowDGV.Cells(Me.ColDocumentNo.Index).Value
                                            End If
                                        Else
                                            CountSave += 1
                                        End If
                                    End If
                                End If
                            Next

                            If ListPrimary.Count > 0 Then
                                If StrCutiNotUpd <> "" Then
                                    MsgBox(StrCutiNotUpd, MsgBoxStyle.Information, "Information")
                                End If
                                If CountSave > 0 Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.KosongkanField()
                                    Me.ObjApp.LoadData(Me, Me.Status, Me.TBL)
                                    'Me.ObjApp.LoadData(Me.TBL, Me.dgvItem, Me.Status, Me)
                                Else
                                    MsgBox("No data stored", MsgBoxStyle.Information, "Information")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnWait_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnWait.Click
        Dim RowDGV As DataGridViewRow
        Dim Primary As Double
        Dim ListPrimary As New List(Of Double)
        Dim Hasil As MsgBoxResult
        Dim StrCutiNotUpd As String = ""
        Dim CountSave As Integer = 0

        Try
            If Me.dgvItem.Rows.Count > 0 Then
                If Me.ObjApp.ApakahDataBenar(Me.dgvItem, Me) = True Then
                    Hasil = MsgBox("Wait for approval ?", MsgBoxStyle.YesNo, "Confirm")
                    If Hasil = MsgBoxResult.Yes Then
                        Dim ij As New Core_Module.DialogPassword
                        ij.txtUserName.Text = Me.UserCode
                        ij.ShowDialog()
                        If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                            For Each RowDGV In Me.dgvItem.Rows
                                If RowDGV.Cells(Me.ColModifiy.Index).Value = True Then
                                    If ListPrimary.Contains(RowDGV.Cells(Me.ColPrimary.Index).Value) = False Then
                                        ListPrimary.Add(RowDGV.Cells(Me.ColPrimary.Index).Value)
                                        If Me.ObjApp.SimpanBaru(Me.UserCode, ClsTrans_Cuti_Approval.EnumStatus.WaitingForApproval, Now, Me.ObjCD, _
                                           RowDGV.Cells(Me.ColPrimary.Index).Value, Me, RowDGV.Index) = False Then
                                            If StrCutiNotUpd = "" Then
                                                StrCutiNotUpd = "Cannot modify leave data because referring to transaction attendance:"
                                                StrCutiNotUpd &= vbCrLf & RowDGV.Cells(Me.ColDocumentNo.Index).Value
                                            Else
                                                StrCutiNotUpd &= vbCrLf & RowDGV.Cells(Me.ColDocumentNo.Index).Value
                                            End If
                                        Else
                                            CountSave += 1
                                        End If
                                    End If
                                End If
                            Next
                            If ListPrimary.Count > 0 Then
                                If StrCutiNotUpd <> "" Then
                                    MsgBox(StrCutiNotUpd, MsgBoxStyle.Information, "Information")
                                End If
                                If CountSave > 0 Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.KosongkanField()
                                    Me.ObjApp.LoadData(Me, Me.Status, Me.TBL)
                                    'Me.ObjApp.LoadData(Me.TBL, Me.dgvItem, Me.Status, Me)
                                Else
                                    MsgBox("No data stored", MsgBoxStyle.Information, "Information")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'by ahdar budiman
    Private Sub LblDuration_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDuration.LinkClicked
        Dim ij As New FrmTransaksiCuti_Pengajuan_ExplainDuration
        Dim RowDGV As DataGridViewRow
        Dim DTA As New DseTrans_Cuti_ApprovalTableAdapters.DataTable1TableAdapter
        Dim TBLDGV As New DseTrans_Cuti_Approval.DataTable1DataTable
        Dim Row As DseTrans_Cuti_Approval.DataTable1Row
        Dim ObjCuti_Item As New ClsTrans_Cuti_Pengajuan_Item
        Try
            If Me.TxtPrimary.Text <> "" And Me.TxtPrimary.Text <> "-1" Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                ij.dgvExplain.Rows.Clear()
                DTA.FillBy_Primary(TBLDGV, Me.TxtPrimary.Text)
                If TBLDGV.Rows.Count > 0 Then
                    For Each Row In TBLDGV.Rows
                        ij.dgvExplain.Rows.Add()
                        RowDGV = ij.dgvExplain.Rows(ij.dgvExplain.Rows.Count - 1)
                        RowDGV.Cells(ij.ColPrimaryWorkShift.Index).Value = Row.primary_workshift
                        RowDGV.Cells(ij.ColWorkShift.Index).Value = Row.Code_Workshift
                        If Row.status = ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Leave Then
                            RowDGV.Cells(ij.ColDuration.Index).Value = Me.ObjCD.DigitGrouping(Row.duration, 1)
                        Else
                            RowDGV.Cells(ij.ColDuration.Index).Value = Me.ObjCD.DigitGrouping(0, 1)
                        End If
                        RowDGV.Cells(ij.ColDate.Index).Value = Format(Row._date, "dd MMM yyyy")
                        RowDGV.Cells(ij.ColStatus.Index).Value = ObjCuti_Item.GetString_ByEnumStatusExplain(Row.status)
                    Next
                End If
                ij.dgvExplain.ReadOnly = True
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBLDGV.Dispose()
            ij.Dispose()
        End Try
    End Sub

    'Private Sub LblDuration_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDuration.LinkClicked
    '    Dim ij As New FrmTransaksiCuti_Pengajuan_ExplainDuration
    '    Dim RowDGV As DataGridViewRow
    '    Dim DTA As New DseTrans_Cuti_ApprovalTableAdapters.DataTable1TableAdapter
    '    Dim TBLDGV As New DseTrans_Cuti_Approval.DataTable1DataTable
    '    Dim Row As DseTrans_Cuti_Approval.DataTable1Row

    '    Try
    '        If Me.TxtPrimary.Text <> "" And Me.TxtPrimary.Text <> "-1" Then
    '            DTA.Connection.ConnectionString = Core_Module.GetConstring
    '            ij.dgvExplain.Rows.Clear()
    '            DTA.FillBy_Primary(TBLDGV, Me.TxtPrimary.Text)
    '            If TBLDGV.Rows.Count > 0 Then
    '                For Each Row In TBLDGV.Rows
    '                    ij.dgvExplain.Rows.Add()
    '                    RowDGV = ij.dgvExplain.Rows(ij.dgvExplain.Rows.Count - 1)
    '                    RowDGV.Cells(ij.ColPrimaryWorkShift.Index).Value = Row.primary_workshift
    '                    RowDGV.Cells(ij.ColWorkShift.Index).Value = Row.Code_Workshift
    '                    RowDGV.Cells(ij.ColDuration.Index).Value = Me.ObjCD.DigitGrouping(Row.duration, 1)
    '                    RowDGV.Cells(ij.ColDate.Index).Value = Format(Row._date, "dd MMM yyyy")
    '                Next
    '            End If
    '            ij.ShowDialog()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        TBLDGV.Dispose()
    '        ij.Dispose()
    '    End Try
    'End Sub

    Private Sub LblBalance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblBalance.LinkClicked
        'Dim DummyStartDate As Date
        Dim Dummy As String = ""
        Dim JmlAttendance As Integer = 0

        Try
            If Me.TxtPrimary.Text <> "" And Me.TxtPrimary.Text <> "-1" Then
                If Me.TxtReason.Enabled = False Then
                    If Me.TxtPrimary.Text <> "-1" Then
                        Me.ij.mPrimaryEmployee = Me.TxtPrimaryEmployee.Text
                        Me.ij.mPrimaryLeave = Me.TxtPrimaryCuti.Text
                        Me.ij.mPeriod = Me.DtpDate.Value
                        'Me.ij.mBaru = Me.Baru
                        'Me.ij.mFrm = Me
                        Me.ij.ObjCd = Me.ObjCD
                        Me.ij.Balance = 0
                        Me.ij.mPrimaryTransaksiPengajuan = Me.TxtPrimary.Text

                        Me.ij.ShowDialog()
                    End If
                Else
                    If Me.TxtPrimaryCuti.Text <> "-1" And Me.TxtPrimaryEmployee.Text <> "-1" Then
                        Me.ij.mPrimaryEmployee = Me.TxtPrimaryEmployee.Text
                        Me.ij.mPrimaryLeave = Me.TxtPrimaryCuti.Text
                        Me.ij.mPeriod = Me.DtpDate.Value
                        'Me.ij.mBaru = Me.Baru
                        'Me.ij.mFrm = Me
                        Me.ij.ObjCd = Me.ObjCD
                        Me.ij.Balance = 0
                        Me.ij.mPrimaryTransaksiPengajuan = Me.TxtPrimary.Text

                        Me.ij.ShowDialog()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpStart_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStart.Enter
        Try
            Me.Buffer_date = Me.dtpStart.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpStart_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStart.Leave
        Try
            If Me.dtpStart.Value.Year <> Me.DtpEnd.Value.Year Then
                MsgBox("Start date and End date must have same year", MsgBoxStyle.Information, "Information")
                Me.dtpStart.Value = Me.Buffer_date
                Me.dtpStart.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpEnd_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpEnd.Enter
        Try
            Me.Buffer_date = Me.DtpEnd.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpEnd_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpEnd.Leave
        Try
            If Me.dtpStart.Value.Year <> Me.DtpEnd.Value.Year Then
                MsgBox("Start date and End date must have same year", MsgBoxStyle.Information, "Information")
                Me.DtpEnd.Value = Me.Buffer_date
                Me.DtpEnd.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub dgvItem_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dgvItem.CellFormatting
        'Dim CboDvg As DataGridViewComboBoxCell
        Try
            'If e.ColumnIndex = Me.ColStatus.Index Then
            'CboDvg = Me.dgvItem.Item(Me.ColStatus.Index, e.RowIndex)
            'Me.dgvItem.Item(Me.ColStatus.Index, e.RowIndex).Value = CboDvg.Items(Me.dgvItem.Item(Me.ColEnumStatus.Index, e.RowIndex).Value)
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.dgvItem.RowCount > 0 Then
                If Me.dgvItem.SelectedRows.Count > 0 Then
                    Me.IsiFormByDgvSelection(Me.dgvItem.Item(Me.ColPrimary.Index, Me.dgvItem.SelectedRows(0).Index).Value)
                End If
            Else
                Me.KosongkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiFormByDgvSelection(ByVal vPrimaryTransCuti As Double)
        Dim row As DseTrans_Cuti_Approval.tbltrans_cuti_pengajuanRow

        Try
            Me.TBL.DefaultView.RowFilter = "[primary] = " & vPrimaryTransCuti
            If Me.TBL.DefaultView.Count > 0 Then
                row = Me.TBL.DefaultView(0).Row
                With row
                    Me.TxtPrimary.Text = .primary
                    Me.TxtCode.Text = .no_transaksi
                    Me.DtpDate.Value = .tanggal
                    Me.TxtPrimaryEmployee.Text = .primary_karyawan
                    Me.TxtEmployeeCode.Text = .Kode_Karyawan
                    Me.TxtEmployeeName.Text = .Full_Name
                    Me.TxtPosition.Text = .Kode_Job_Position
                    Me.TxtPrimaryCuti.Text = .primary_cuti
                    Me.TextBox1.Text = .CodeCuti
                    Me.dtpStart.Value = .start_date
                    Me.DtpEnd.Value = .end_date
                    If .half_day_start = 1 Then
                        Me.ChkHDStart.Checked = True
                    Else
                        Me.ChkHDStart.Checked = False
                    End If
                    If .half_day_end = 1 Then
                        Me.ChkHDEnd.Checked = True
                    Else
                        Me.ChkHDEnd.Checked = False
                    End If
                    Me.TxtDuration.Text = Me.ObjCD.DigitGrouping(.duration, 1)
                    Me.TxtBalance.Text = Me.ObjCD.DigitGrouping(.balance, 1)
                    Me.TxtEnumStatus.Text = .status
                    If Me.TxtEnumStatus.Text = "0" Then
                        Me.TxtStatus.Text = "Waiting for approval"
                    ElseIf Me.TxtEnumStatus.Text = "1" Then
                        Me.TxtStatus.Text = "Approved"
                    ElseIf Me.TxtEnumStatus.Text = "2" Then
                        Me.TxtStatus.Text = "Rejected"
                    End If
                    Me.TxtApprovedBy.Text = .approved_by
                    Me.DtpApprovedBy.Value = .approved_date
                    Me.TxtReason.Text = .reason
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class