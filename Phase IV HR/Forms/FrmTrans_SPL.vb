Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms, System.Threading
Public Class FrmTrans_SPL
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim Obj As New ClsTrans_SPL
    Dim ObjItem As New ClsTrans_SPL_Item
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjLicense As New ClsLicense
    Dim objEmp As New Core_Module.ClsMasterKaryawan
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjSettingsUser As New Core_Module.ClsTransaksiSettingUser

    Dim Buffer As String = ""
    Dim EraseRowIndex As Integer = 0
    Dim BufferTransDate As Date = #1/1/1950#
    Dim BufferDate As Date = #1/1/1950#
    Dim ApakahDateFrom As Boolean = False
    Dim ApaShiftKey_InDtp As Boolean = False
    Public FromAtt As Boolean = False

    Dim FrmStatus As New FrmTrans_SPLApprovalNotes

    Public TblEmp As New DseTrans_SPL.tbltrans_spl_overtime_order_item_employeeDataTable
    Public TblEmpShift As New DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftDataTable
    Public TblSchSetup As New DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupDataTable
    Public TblSchDet As New DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailDataTable
    Public TblApp As New DseTrans_SPL.tbltrans_spl_overtime_order_item_approvalDataTable
    Public TblAppOT As New DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otDataTable

    Enum EnumSchSetup As Integer
        OT_Enum
        Delay_Enum
        Break_Enum
        DelayBefBreak_Enum
        HolidayOT_Enum
        HolidayDelay_Enum
        HolidayBreak_Enum
        HolidayDelayBefBreak_Enum
    End Enum

    Enum EnumOvertimeEndEdit As Integer
        Start_Enum
        Break_Enum
        End_Enum
    End Enum

    Private Sub SetField_EditScheduleEmp()
        Try
            Me.dgvSchEmp.ReadOnly = False
            Me.dgvSchEmp.AllowUserToDeleteRows = True
            Me.btnOK.Enabled = True
            Me.btnCancel.Enabled = True

            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            Me.TabControl1.SelectedIndex = 1
            Me.TabControl2.SelectedIndex = 2
            If Me.dgvSchEmp.Rows.Count > 0 Then
                Me.dgvSchEmp.Rows(0).Selected = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.txtPrimary.BackColor, False)
            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, False, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            Me.dgvEmployee.AllowUserToAddRows = True
            'Me.dgvSchSetup.AllowUserToAddRows = True
            'Me.dgvSchDetail.AllowUserToAddRows = True
            'Me.dgvApproval.AllowUserToAddRows = True

            Me.dgvEmployee.AllowUserToDeleteRows = True
            Me.dgvSchDetail.AllowUserToDeleteRows = True
            Me.dgvSchEmp.AllowUserToDeleteRows = True

            Me.dgvEmployee.ReadOnly = False
            Me.dgvSchSetup.ReadOnly = False
            Me.dgvSchEmp.ReadOnly = False
            Me.dgvSchDetail.ReadOnly = False
            Me.dgvApproval.ReadOnly = False

            If Me.baru = True Then
                Me.TxtTotalDuration.Text = "00.00"
                Me.TxtStatus.Text = "Waiting For Approval"

                Me.TxtAppliedBy.Text = ObjSettingUser.GetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_AppliedBy, Me.UserCode)
                Me.TxtAppliedPosition.Text = ObjSettingUser.GetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_AppliedPosition, Me.UserCode)
            End If

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.llbSendApprovalRequest.Enabled = False
            Me.llbBudget.Enabled = False
            Me.TxtReason.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, False, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)

            Me.dgvEmployee.AllowUserToAddRows = False
            Me.dgvSchSetup.AllowUserToAddRows = False
            Me.dgvSchDetail.AllowUserToAddRows = False
            Me.dgvSchEmp.AllowUserToAddRows = False
            Me.dgvApproval.AllowUserToAddRows = False

            Me.dgvEmployee.AllowUserToDeleteRows = False
            Me.dgvSchDetail.AllowUserToDeleteRows = False
            Me.dgvSchEmp.AllowUserToDeleteRows = False

            Me.dgvEmployee.ReadOnly = True
            Me.dgvSchSetup.ReadOnly = True
            Me.dgvSchDetail.ReadOnly = True
            Me.dgvSchEmp.ReadOnly = True
            Me.dgvApproval.ReadOnly = True

            Me.CancelButton = Me.btnExit
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)

            If Me.Obj.Status = ClsTrans_SPL.EnumStatus.Rejected Then
                Me.llbBudget.Enabled = False
            Else
                Me.llbBudget.Enabled = True
            End If
            Me.llbSendApprovalRequest.Enabled = True
            Me.btnFind.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Core_Module.ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            Core_Module.ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            Core_Module.ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)

            Me.llbBudget.Text = "Budget"

            Me.txtPrimary.Text = "-1"
            Me.dgvEmployee.Rows.Clear()
            Me.dgvSchSetup.Rows.Clear()
            Me.dgvSchDetail.Rows.Clear()
            Me.dgvSchEmp.Rows.Clear()
            Me.dgvApproval.Rows.Clear()

            'clear table yang di keep
            Me.TblEmp = New DseTrans_SPL.tbltrans_spl_overtime_order_item_employeeDataTable
            Me.TblEmpShift = New DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftDataTable
            Me.TblSchSetup = New DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupDataTable
            Me.TblSchDet = New DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailDataTable
            Me.TblApp = New DseTrans_SPL.tbltrans_spl_overtime_order_item_approvalDataTable
            Me.TblAppOT = New DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otDataTable
            Me.FrmStatus.TxtStatusNotes.Text = ""

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim DocNoBuffer As String = ""

        Try
            'tidak boleh kosong
            KoleksiNamaData.Add("Reason")
            KoleksiTextBox.Add(Me.TxtReason)
            KoleksiNamaData.Add("Applied By")
            KoleksiTextBox.Add(Me.TxtAppliedBy)
            KoleksiNamaData.Add("Applied By Position")
            KoleksiTextBox.Add(Me.TxtAppliedPosition)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'tidak boleh double
            If Me.baru = True Then
                If Me.Obj.ApakahSudahAda(Me.txtDocumentNo.Text) = True Then
                    DocNoBuffer = Me.txtDocumentNo.Text
                    Me.txtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.txtPrimary.Text)
                End If
            Else
                Dim CodeLama As String = Me.Obj.GetNoTransaksiByPrimary(Me.txtPrimary.Text)
                If CodeLama <> Me.txtDocumentNo.Text Then
                    If Me.Obj.ApakahSudahAda(Me.txtDocumentNo.Text) = True Then
                        DocNoBuffer = Me.txtDocumentNo.Text
                        Me.txtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.txtPrimary.Text)
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses(edit)
                Dim OldTimeStamp As Date = Me.Obj.GetTimeStampByPrimary(Me.txtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.txtPrimary.Text, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'minimal input satu baris employee
            If Me.dgvEmployee.RowCount - 1 = 0 Then
                If String.IsNullOrEmpty(Me.dgvEmployee.Item(Me.colEmpl_Code.Index, 0).Value) = True Then
                    Me.TabControl1.SelectedIndex = 0
                    MsgBox("Please input item employee list", MsgBoxStyle.Information, "Information")
                    Me.dgvEmployee.ClearSelection()
                    Me.dgvEmployee.Rows(0).Selected = True

                    Return False
                End If
            End If

            'minimal harus ada 1 baris 
            If Me.TblSchSetup.Rows.Count = 0 Then
                Me.TabControl1.SelectedIndex = 1
                Me.TabControl2.SelectedIndex = 0
                MsgBox("Please input item schedule setup", MsgBoxStyle.Information, "Information")

                Return False
            Else
                Dim rowSchSetup As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupRow
                Dim rowD As DataGridViewRow
                Dim RowIdx As Integer = -1
                Dim ColIdx As Integer = -1
                Dim ApakahReturnFalse As Boolean = False

                Me.TblSchSetup.DefaultView.RowFilter = ""
                Me.TblSchSetup.DefaultView.Sort = "[Shift Code]"
                For Count As Integer = 0 To Me.TblSchSetup.DefaultView.Count - 1
                    rowSchSetup = Me.TblSchSetup.DefaultView(Count).Row
                    RowIdx = Count
                    If rowSchSetup.Is_Early_OT = 0 And rowSchSetup.Is_OT = 0 And rowSchSetup.Is_Holiday_OT = 0 Then
                        ApakahReturnFalse = True
                        Exit For
                    End If

                    If rowSchSetup.Is_Early_OT = 1 And rowSchSetup.Early_Duration = 0 Then
                        ColIdx = Me.ColSetup_EarlyOT.Index
                        ApakahReturnFalse = True
                        Exit For
                    ElseIf rowSchSetup.Is_OT = 1 And rowSchSetup.OT_Duration = 0 Then
                        ColIdx = Me.ColSetup_OT.Index
                        ApakahReturnFalse = True
                        Exit For
                    ElseIf rowSchSetup.Is_Holiday_OT = 1 And rowSchSetup.Holiday_OT_Duration = 0 Then
                        ColIdx = Me.ColSetup_HolOT.Index
                        ApakahReturnFalse = True
                        Exit For
                    ElseIf rowSchSetup.Early_Break > 0 And rowSchSetup.Early_Delay_Before_Break = 0 Then
                        ColIdx = Me.ColSetup_HolOTDelayBeforeBreak.Index
                        ApakahReturnFalse = True
                        Exit For
                    ElseIf rowSchSetup.OT_Break > 0 And rowSchSetup.OT_Delay_Before_Break = 0 Then
                        ColIdx = Me.ColSetup_OTDelayBeforeBreak.Index
                        ApakahReturnFalse = True
                        Exit For
                    ElseIf rowSchSetup.Holiday_OT_Break > 0 And rowSchSetup.Holiday_Delay_Before_Break = 0 Then
                        ColIdx = Me.ColSetup_HolOTDelayBeforeBreak.Index
                        ApakahReturnFalse = True
                        Exit For
                    End If
                Next

                If ApakahReturnFalse = True Then
                    'MsgBox("Overtime order can only input one schedule", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 1
                    Me.TabControl2.SelectedIndex = 0
                    rowD = Me.dgvSchSetup.Rows(RowIdx)
                    Me.dgvSchSetup.ClearSelection()

                    If ColIdx <> -1 Then
                        MsgBox("Please input positive hours at " & Me.dgvSchSetup.Columns(ColIdx).HeaderText, MsgBoxStyle.Information, "Information")
                        If rowD.Cells(ColIdx).Visible = True Then
                            rowD.Cells(ColIdx).Selected = True
                        Else
                            rowD.Selected = True
                        End If
                    Else
                        MsgBox("Please use overtime at workshift " & Me.dgvSchSetup.Item(Me.ColSetup_ShiftCode.Index, RowIdx).Value, MsgBoxStyle.Information, "Information")
                        rowD.Selected = True
                    End If

                    Return False
                End If
            End If

            If Me.TblSchDet.Rows.Count = 0 Then
                Me.TabControl1.SelectedIndex = 1
                Me.TabControl2.SelectedIndex = 1
                MsgBox("Please input item schedule detail", MsgBoxStyle.Information, "Information")
                Me.dgvSchDetail.ClearSelection()
                Return False
            End If

            If Me.TblApp.Rows.Count = 0 Then
                Me.TabControl1.SelectedIndex = 2
                MsgBox("Please input item approval", MsgBoxStyle.Information, "Information")
                Me.dgvApproval.ClearSelection()
                Return False
            Else
                Dim TotalCeklis As Double = 0
                Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approvalRow
                For Each row In Me.TblApp.Rows
                    If row.Req_Approval = 1 Then
                        TotalCeklis += 1
                    End If
                Next

                If TotalCeklis = 0 Then
                    Me.TabControl1.SelectedIndex = 2
                    MsgBox("Please checklist req approval minimal 1 user", MsgBoxStyle.Information, "Information")
                    Me.dgvApproval.Rows(0).Selected = True
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .NoTransaksi = Me.txtDocumentNo.Text
                .Date_ = Me.dtpDate.Value.Date
                .Reason = Me.TxtReason.Text
                .AppliedBy = Me.TxtAppliedBy.Text
                .AppliedByPosition = Me.TxtAppliedPosition.Text
                '.PrimaryApprovedBy = Me.txtPrimary.Text 'diupdate dari form approval overtime
                '.ApprovedTime = Me.txtDocumentNo.Text 'diupdate dari form approval overtime
                .DateFrom_ = Me.DtpFrom.Value.Date
                .DateTo_ = Me.DtpTo.Value.Date
                .TotalDuration = Me.ObjCD.Convert_ToMinutes(Me.TxtTotalDuration.Text)
                '.StatusNotes = Me.TxtStatus.Text
                '.Status = Me.txtDocumentNo.Text 'diupdate dari form approval overtime
                .Remarks = Me.txtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.txtPrimary.Text = .Primary
                Me.txtDocumentNo.Text = .NoTransaksi
                Me.dtpDate.Value = .Date_
                Me.TxtReason.Text = .Reason
                Me.TxtAppliedBy.Text = .AppliedBy
                Me.TxtAppliedPosition.Text = .AppliedByPosition
                '.PrimaryApprovedBy = Me.txtPrimary.Text 'diupdate dari form approval overtime
                '.ApprovedTime = Me.txtDocumentNo.Text 'diupdate dari form approval overtime
                Me.DtpFrom.Value = .DateFrom_
                Me.DtpTo.Value = .DateTo_
                Me.TxtTotalDuration.Text = Me.ObjCD.Convert_ToString(.TotalDuration)
                If .Status = ClsTrans_SPL.EnumStatus.WaitingForApproval Then
                    Me.TxtStatus.Text = "Waiting For Approval"
                ElseIf .Status = ClsTrans_SPL.EnumStatus.Approved Then
                    Me.TxtStatus.Text = "Approved By " & Me.ObjUser.GetUserCode(.PrimaryApprovedBy) & " - " & Format(.ApprovedTime, "dd-MMM-yyyy") & " " & .ApprovedTime.ToShortTimeString
                Else
                    Me.TxtStatus.Text = "Rejected By " & Me.ObjUser.GetUserCode(.PrimaryApprovedBy) & " - " & Format(.ApprovedTime, "dd-MMM-yyyy") & " " & .ApprovedTime.ToShortTimeString
                End If
                Me.FrmStatus.TxtStatusNotes.Text = .StatusNotes
                '.Status = Me.txtDocumentNo.Text 'diupdate dari form approval overtime
                Me.txtRemarks.Text = .Remarks

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                If Me.txtPrimary.Text <> -1 Then
                    If Me.TabControl1.SelectedIndex = 1 Then
                        If Me.TabControl2.SelectedIndex = 0 Then
                            Me.ObjItem.IsiDGVItemSchSetup(Me.dgvSchSetup, Me.txtPrimary.Text, Me, Me.TblSchSetup)
                        ElseIf Me.TabControl2.SelectedIndex = 1 Then
                            Me.ObjItem.IsiDGVItemSchDetail(Me.dgvSchDetail, Me.txtPrimary.Text, Me, Me.TblSchDet)
                        Else
                            Me.ObjItem.IsiDGVItemSchEmp(Me.dgvSchEmp, Me.txtPrimary.Text, Me, Me.TblAppOT)
                        End If
                    ElseIf Me.TabControl1.SelectedIndex = 2 Then
                        Me.ObjItem.IsiDGVItemApproval(Me.dgvApproval, Me.txtPrimary.Text, Me, Me.TblApp)
                    End If
                End If

                'isi list primary shift
                'Me.IsiList_PrimaryShift()

                Me.GetBudgetSpl(False, True)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SPL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            'Me.ColSchEmp_IsHoliday.Visible = True
            'Me.llbBudget.Text = "Over Budget"
            Me.colEmpl_TotalHours.Visible = False

            Me.LlbShift.Visible = False
            Me.LlbAppOT.Visible = False
            Me.Llbsetup.Visible = False
            Me.LlbDetail.Visible = False
            Me.LlbApp.Visible = False

            Me.ColSchEmp_NormalWH.Visible = True
            'Me.ColSchEmp_NormalWH.Visible = False
            'Me.ColSchEmp_StartEarly.Visible = False
            'Me.ColSchEmp_BreakEarly.Visible = False
            'Me.ColSchEmp_EndEarly.Visible = False
            'Me.ColSchEmp_StartOT.Visible = False
            'Me.ColSchEmp_BreakOT.Visible = False
            'Me.ColSchEmp_EndOT.Visible = False
            Me.ColSchEmp_OverrideOT.Visible = True

            'Me.llbColSetting.Visible = False
            'Me.objDepartement = New ClsMasterDepartement
            Me.baru = True
            Me.NonAktifkanDetail()

            If FromAtt = True Then
                Core_Module.ModCekData.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            End If

            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSchDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSchSetup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSchEmp.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvApproval.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.dgvEmployee.AllowUserToAddRows = False
            Me.dgvSchSetup.AllowUserToDeleteRows = False
            Me.dgvSchDetail.AllowUserToDeleteRows = False
            Me.dgvApproval.AllowUserToDeleteRows = False
            Me.dgvEmployee.MultiSelect = False
            Me.dgvSchSetup.MultiSelect = False
            Me.dgvSchDetail.MultiSelect = False
            Me.dgvSchEmp.MultiSelect = False
            Me.dgvApproval.MultiSelect = False

            'MsgBox("form ini bisa disetting tergantung kebijakan klient. apakah memungkina menimpa nilai overtime pada form attendance. apakah wajib menggunakan spl. apakah nilai yang digunakan yang terkacil antara actual vs spl ketika menghitung attendance")
            Me.dtpDate.Left = Me.SplitContainer1.Panel1.Width - Me.dtpDate.Width
            Me.txtDocumentNo.Width = Me.dtpDate.Left - 5 - Me.txtDocumentNo.Left
            Me.TxtReason.Width = Me.SplitContainer1.Panel1.Width - Me.TxtReason.Left
            Me.TxtReason.Height = Me.TxtAppliedBy.Top - 5 - Me.TxtReason.Top
            Me.TxtAppliedBy.Width = Me.SplitContainer1.Panel1.Width - Me.TxtAppliedBy.Left
            Me.TxtAppliedPosition.Width = Me.SplitContainer1.Panel1.Width - Me.TxtAppliedPosition.Left

            Me.DtpTo.Width = Me.SplitContainer1.Panel2.Width - Me.DtpTo.Left
            Me.TxtTotalDuration.Width = Me.SplitContainer1.Panel2.Width - Me.TxtTotalDuration.Left
            Me.TxtStatus.Width = Me.SplitContainer1.Panel2.Width - 130
            Me.txtRemarks.Width = Me.SplitContainer1.Panel2.Width - 130
            Me.llbNotes.Left = Me.TxtStatus.Right + 5
            Me.llbBudget.Left = Me.txtRemarks.Right + 5

            Me.SettingKolom()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SPL_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            If Me.txtRemarks.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            If Me.Obj.SimpanBaru(Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim hasilApproval As MsgBoxResult
        Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanBaru(Me.UserCode, Me) = True Then
                            ObjSettingUser.SetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_AppliedBy, Me.TxtAppliedBy.Text, Me.UserCode)
                            ObjSettingUser.SetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_AppliedPosition, Me.TxtAppliedPosition.Text, Me.UserCode)

                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.Obj.CariDetail(Me.Obj.Primary, Me)
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()

                            If Me.ApaSendNotifEmail = True Then
                                hasilApproval = MsgBox("Send approval request ?", MsgBoxStyle.YesNo, "Confirm")
                                If hasilApproval = MsgBoxResult.Yes Then
                                    Me.SendApprovalReq()
                                End If
                            End If
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me) Then
                            ObjSettingUser.SetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_AppliedBy, Me.TxtAppliedBy.Text, Me.UserCode)
                            ObjSettingUser.SetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_AppliedPosition, Me.TxtAppliedPosition.Text, Me.UserCode)

                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.Obj.CariDetail(Me.Obj.Primary, Me)
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.txtDocumentNo.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongkanField()
                    Else
                        Me.NonAktifkanDetail()
                        Me.Obj.CariDetail(Me.txtPrimary.Text, Me)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.NonAktifkanDetail()
                Me.KosongkanField()
            End If

            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTrans_SPLCari

        Try
            hj.BackColorAll = Me.BackColorAll
            hj.Frm = Me
            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                If Me.Obj.CariDetail(hj.PrimaryTransaksi, Me) = True Then
                    Me.baru = False
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                    Me.NonAktifkanDetail()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_SPL_New") = False Then
                MsgBox("You don't have permission to perform this action", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.Obj.ResetObject()
            Me.KosongkanField()
            Me.AktifkanDetail()

            Me.txtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.txtPrimary.Text)

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim TutupAksesDetailOT As Boolean = True

        Try
            If Me.txtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDocumentNo.Text <> "" Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_SPL_EditDetailOT") = True Then
                    TutupAksesDetailOT = False
                End If

                If Me.Obj.ApakahBolehedit(Me.UserCode, Me.txtPrimary.Text, TutupAksesDetailOT) = True Then
                    Me.baru = False
                    Me.Obj.CariDetail(Me.txtPrimary.Text, Me)
                    Me.AktifkanDetail()
                    Me.txtDocumentNo.Focus()
                Else
                    If TutupAksesDetailOT = False Then
                        MsgBox("You can edit schedule employee only.", MsgBoxStyle.Information, "Information")
                        Me.SetField_EditScheduleEmp()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.txtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDocumentNo.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.UserCode, Me.txtPrimary.Text) = True Then
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.txtPrimary.Text, Me.UserCode, Me) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub llbColSetting_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbColSetting.LinkClicked
        Dim ij As New FrmTrans_SPL_ColumnsSettings

        Try
            ij.UserCode = Me.UserCode
            ij.ShowDialog()
            Me.SettingKolom()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub SettingKolom()
        Try
            'OT
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowEarlyOT, Me.UserCode) = "1" Then
                Me.ColSetup_IsEarlyOT.Visible = True
                Me.ColSetup_EarlyOT.Visible = True
                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowEarlyDelay, Me.UserCode) = "1" Then
                    Me.ColSetup_EarlyDelay.Visible = True
                Else
                    Me.ColSetup_EarlyDelay.Visible = False
                End If
                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowEarlyBreak, Me.UserCode) = "1" Then
                    Me.ColSetup_EarlyBreak.Visible = True
                    Me.ColSetup_EarlyDelayBeforeBreak.Visible = True
                Else
                    Me.ColSetup_EarlyBreak.Visible = False
                    Me.ColSetup_EarlyDelayBeforeBreak.Visible = False
                End If
            Else
                Me.ColSetup_IsEarlyOT.Visible = False
                Me.ColSetup_EarlyOT.Visible = False
                Me.ColSetup_EarlyDelay.Visible = False
                Me.ColSetup_EarlyBreak.Visible = False
                Me.ColSetup_EarlyDelayBeforeBreak.Visible = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowOT, Me.UserCode) = "1" Then
                Me.ColSetup_IsOT.Visible = True
                Me.ColSetup_OT.Visible = True
                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowDelay, Me.UserCode) = "1" Then
                    Me.ColSetup_OTDelay.Visible = True
                Else
                    Me.ColSetup_OTDelay.Visible = False
                End If
                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowBreak, Me.UserCode) = "1" Then
                    Me.ColSetup_OTBreak.Visible = True
                    Me.ColSetup_OTDelayBeforeBreak.Visible = True
                Else
                    Me.ColSetup_OTBreak.Visible = False
                    Me.ColSetup_OTDelayBeforeBreak.Visible = False
                End If
            Else
                Me.ColSetup_IsOT.Visible = False
                Me.ColSetup_OT.Visible = False
                Me.ColSetup_OTDelay.Visible = False
                Me.ColSetup_OTBreak.Visible = False
                Me.ColSetup_OTDelayBeforeBreak.Visible = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowHolidayOT, Me.UserCode) = "1" Then
                Me.ColSetup_IsHolOT.Visible = True
                Me.ColSetup_HolOT.Visible = True
                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowHolidayDelay, Me.UserCode) = "1" Then
                    Me.ColSetup_HolOTDelay.Visible = True
                Else
                    Me.ColSetup_HolOTDelay.Visible = False
                End If
                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_OvertimeOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_OvertimeOrder_ShowHolidayBreak, Me.UserCode) = "1" Then
                    Me.ColSetup_HolOTBreak.Visible = True
                    Me.ColSetup_HolOTDelayBeforeBreak.Visible = True
                Else
                    Me.ColSetup_HolOTBreak.Visible = False
                    Me.ColSetup_HolOTDelayBeforeBreak.Visible = False
                End If
            Else
                Me.ColSetup_IsHolOT.Visible = False
                Me.ColSetup_HolOT.Visible = False
                Me.ColSetup_HolOTDelay.Visible = False
                Me.ColSetup_HolOTBreak.Visible = False
                Me.ColSetup_HolOTDelayBeforeBreak.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReason_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReason.GotFocus, TxtAppliedBy.GotFocus, TxtAppliedPosition.GotFocus, txtRemarks.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadonly.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReason_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReason.Leave
        Me.TxtReason.Text = Me.TxtReason.Text.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
    End Sub

    Private Sub txtRemarks_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.Leave
        Me.txtRemarks.Text = Me.txtRemarks.Text.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
    End Sub

    Private Sub TxtReason_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReason.LostFocus, TxtAppliedBy.LostFocus, TxtAppliedPosition.LostFocus, txtRemarks.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadonly.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimary.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSendApprovalRequest_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSendApprovalRequest.LinkClicked
        Dim hasil As MsgBoxResult
        Dim MsgString As String = ""

        Try
            If Me.txtPrimary.Text = -1 Or String.IsNullOrEmpty(Me.txtPrimary.Text) = True Then
                Exit Sub
            End If

            If Me.Obj.Status = ClsTrans_SPL.EnumStatus.Approved Then
                MsgString = "This transaction has been approved by " & Me.ObjUser.GetUserCode(Me.Obj.PrimaryApprovedBy)
            ElseIf Me.Obj.Status = ClsTrans_SPL.EnumStatus.Rejected Then
                MsgString = "This transaction has been rejected by " & Me.ObjUser.GetUserCode(Me.Obj.PrimaryApprovedBy)
            End If

            If MsgString <> "" Then
                hasil = MsgBox(MsgString & vbCrLf & "Send approval request ?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If

            Me.SendApprovalReq()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApaSendNotifEmail() As Boolean
        Dim retval As Boolean = False

        Try
            Me.TblEmp.DefaultView.RowFilter = "[notification email ot] = 1"
            If Me.TblEmp.DefaultView.Count > 0 Then
                retval = True
            End If
            Me.TblEmp.DefaultView.RowFilter = ""

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return retval
        End Try
    End Function

    Private Sub SendApprovalReq()
        Dim rowApp As DseTrans_SPL.tbltrans_spl_overtime_order_item_approvalRow
        Dim ij As New FrmTrans_SPL_SendEmail
        Dim ListPrimary As New List(Of Double)

        Try
            For Each rowApp In Me.TblApp
                If rowApp.Req_Approval = 1 Then
                    ListPrimary.Add(rowApp.Primary_User)
                End If
            Next

            If ListPrimary.Count > 0 Then
                ij.UserCode = Me.UserCode
                ij.OvertimeOrderNo = Me.txtDocumentNo.Text
                ij.BackColorAll = Me.BackColorAll
                ij.ListPrimaryUserApp = ListPrimary
                ij.TblApp = Me.TblApp
                ij.ShowDialog()
                If ij.ApakahSendEmail = True Then
                   Me.Obj.CariDetail(Me.txtPrimary.Text, Me)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub LlbAdvance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbAdvance.LinkClicked
        If Me.btnOK.Enabled = True Then
            Me.ContextMenuStrip1.Show(Me.LlbAdvance, 0, 20)
        End If
    End Sub

    Private Sub llbNotes_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbNotes.LinkClicked
        Try
            If Me.txtPrimary.Text <> "-1" Then
                Me.FrmStatus.TxtStatusNotes.Dock = DockStyle.Fill
                Me.FrmStatus.btnOK.SendToBack()
                Me.FrmStatus.btnCancel.SendToBack()
                Me.FrmStatus.btnOK.TabStop = False
                Me.FrmStatus.btnCancel.TabStop = False
                Me.FrmStatus.MaximizeBox = False
                Me.FrmStatus.MinimizeBox = False
                Me.FrmStatus.TxtStatusNotes.ReadOnly = True
                Me.FrmStatus.BackColorAll = Me.txtReadonly.BackColor
                Me.FrmStatus.mPrimaryTrans = Me.txtPrimary.Text
                Me.FrmStatus.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbBudget_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbBudget.LinkClicked
        Try
            Me.GetBudgetSpl(True, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetBudgetSpl(ByVal vShowFormBudget As Boolean, ByVal vUpdateLlbText As Boolean)
        Dim i As New FrmTrans_SPLApprovalBudget
        Dim i2 As New FrmTrans_SPLApproval
        Dim dta As New DseTrans_SPLTableAdapters.tbltrans_spl_approvalTableAdapter
        Dim tbl As New DseTrans_SPL.tbltrans_spl_approvalDataTable
        Dim TblSplApp As New DseTrans_SPL.tbltrans_spl_approvalDataTable
        Dim TblSplUsed As New DseTrans_SPL.ot_order_approvalDataTable
        Dim TblSplAppDep2 As New DseTrans_SPL.ot_order_approval_2DataTable
        Dim DateFilterFrom As Date = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpFrom.Value.Date)
        Dim DateFilterTo As Date = Me.ObjCD.GetLasDateOfTheMonth(Me.DtpTo.Value.Date)
        Dim ListPrimaryDept As New List(Of Double)
        Dim ListDept As New List(Of String)
        Dim StatusSPL As Integer = 0

        Try
            StatusSPL = Me.Obj.GetStatus(Me.txtPrimary.Text)
            If StatusSPL = ClsTrans_SPL.EnumStatus.Rejected Then
                If vShowFormBudget = True Then
                    MsgBox("Overtime budget rejected, not include budget.", MsgBoxStyle.Information, "Information")
                    Me.llbBudget.Enabled = False
                    Exit Sub
                End If
            End If

            If Me.txtPrimary.Text = -1 Or String.IsNullOrEmpty(Me.txtPrimary.Text) = True Then
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.ClearBeforeFill = False
            dta.Fill_SplGroupByDate(TblSplApp, ClsTrans_SPL.EnumStatus.WaitingForApproval, DateFilterFrom.Date, DateFilterTo.Date)
            dta.Fill_SplGroupByDate(TblSplApp, ClsTrans_SPL.EnumStatus.Approved, DateFilterFrom.Date, DateFilterTo.Date)
            dta.Fill_Date(tbl, ClsTrans_SPL.EnumStatus.WaitingForApproval, DateFilterTo.Date, DateFilterFrom.Date)
            dta.Fill_Date(tbl, ClsTrans_SPL.EnumStatus.Approved, DateFilterTo.Date, DateFilterFrom.Date)

            i2.IsiTblBudgetPerDepartment(tbl, TblSplApp, TblSplUsed, TblSplAppDep2, DateFilterFrom.Date, DateFilterTo.Date, False)

            If vShowFormBudget = True Then
                Me.TblEmp.DefaultView.RowFilter = Nothing
                Me.TblEmp.DefaultView.Sort = "[Dept Code]"
                Dim rowEmp As DseTrans_SPL.tbltrans_spl_overtime_order_item_employeeRow
                For Each dv As DataRowView In Me.TblEmp.DefaultView
                    rowEmp = dv.Row
                    If ListPrimaryDept.Contains(rowEmp.Primary_Deptartment) = False Then
                        ListPrimaryDept.Add(rowEmp.Primary_Deptartment)
                    End If
                Next
                Me.TblEmp.DefaultView.RowFilter = Nothing
                Me.TblEmp.DefaultView.Sort = Nothing
                'For Each rowD As DataGridViewRow In Me.dgvEmployee.Rows
                '    If ListPrimaryDept.Contains(rowD.Cells(Me.colEmpl_PrimaryDept.Index).Value) = False Then
                '        ListPrimaryDept.Add(rowD.Cells(Me.colEmpl_PrimaryDept.Index).Value)
                '    End If
                'Next

                i.PrimarySPL = CDbl(Me.txtPrimary.Text)
                i.TblSplAppDep2 = TblSplAppDep2
                i.TblSplUsed = TblSplUsed
                i.ListPrimaryDept = ListPrimaryDept
                i.PeriodFrom = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpFrom.Value.Date)
                i.PeriodTo = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpTo.Value.Date)
                i.BackColorAll = Me.BackColorAll
                i.FromOvertimeOrder = True
                i.ShowDialog()
            End If

            If vUpdateLlbText = True Then
                Me.llbBudget.Text = "Budget"
                If i2.ApakahOverBudget(CDbl(Me.txtPrimary.Text), TblSplAppDep2) = True And StatusSPL <> ClsTrans_SPL.EnumStatus.Rejected Then
                    Me.llbBudget.Text = "Over Budget"
                End If
            End If
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub dtpDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.Enter
        Try
            Me.BufferTransDate = Me.dtpDate.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.Leave
        Try
            If Me.txtRemarks.Enabled = True Then
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpFrom_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpFrom.Enter
        Try
            Me.ApakahDateFrom = True
            Me.BufferDate = Me.DtpFrom.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpFrom_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles DtpFrom.PreviewKeyDown
        Try
            If (e.Modifiers = Keys.Shift And e.KeyCode = Keys.Tab) Then
                ApaShiftKey_InDtp = True
            Else
                ApaShiftKey_InDtp = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpFrom_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpFrom.Leave
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.ApaShiftKey_InDtp = True Then
                    Me.TxtAppliedPosition.Focus()
                Else
                    Me.DtpTo.Focus()
                End If
                If Me.BufferDate <> Me.DtpFrom.Value.Date Then
                    Me.ChangeDtpFromTo()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpTo_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpTo.Enter
        Try
            Me.ApakahDateFrom = False
            Me.BufferDate = Me.DtpTo.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpTo_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles DtpTo.PreviewKeyDown
        Try
            If (e.Modifiers = Keys.Shift And e.KeyCode = Keys.Tab) Then
                ApaShiftKey_InDtp = True
            Else
                ApaShiftKey_InDtp = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpTo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpTo.Leave
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.ApaShiftKey_InDtp = True Then
                    Me.DtpFrom.Focus()
                Else
                    Me.llbNotes.Focus()
                End If
                If Me.BufferDate <> Me.DtpTo.Value.Date Then
                    Me.ChangeDtpFromTo()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChangeDtpFromTo()
        'table adapter = tbltrans_spl_item_employee
        Dim Dta1 As New DseTrans_SPLTableAdapters.DataTable1TableAdapter
        Dim Tbl1 As New DseTrans_SPL.DataTable1DataTable
        Dim row1 As DseTrans_SPL.DataTable1Row
        'table adapter = tbltrans_attendance
        Dim Dta2 As New DseTrans_SPLTableAdapters.tbltrans_attendanceTableAdapter
        Dim Tbl2 As New DseTrans_SPL.tbltrans_attendanceDataTable

        Dim rowD As DataGridViewRow
        Dim ListMsg_1 As String = ""
        Dim ListMsg_2 As String = ""

        Try
            If Me.DtpFrom.Value.Date > Me.DtpTo.Value.Date Then
                MsgBox("Date from value can't greater then date to value", MsgBoxStyle.Information, "Information")
                If Me.ApakahDateFrom = True Then
                    Me.DtpFrom.Value = Me.BufferDate
                    'Me.DtpFrom.Focus()
                Else
                    Me.DtpTo.Value = Me.BufferDate
                    'Me.DtpTo.Focus()
                End If

                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor

            Dta1.Connection.ConnectionString = Core_Module.GetConstring
            Dta2.Connection.ConnectionString = Core_Module.GetConstring
            Dta1.Fill(Tbl1, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, ClsTrans_SPL.EnumStatus.Rejected) 'yang bukan reject
            Dta2.Fill(Tbl2, Me.DtpTo.Value.Date, Me.DtpFrom.Value.Date)

            For i As Integer = 0 To Me.dgvEmployee.RowCount - 2
                rowD = Me.dgvEmployee.Rows(i)
                If Me.Obj.ApaSudahAdaPadaTransaksiLain(Me.txtPrimary.Text, rowD.Cells(Me.colEmpl_Primary.Index).Value, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, Tbl1) = True Then
                    If ListMsg_2 = "" Then
                        ListMsg_2 &= "Employee : " & rowD.Cells(Me.colEmpl_Name.Index).Value
                    Else
                        ListMsg_2 &= ", " & rowD.Cells(Me.colEmpl_Name.Index).Value
                    End If
                End If
                If Me.Obj.ApaSudahAdaAttendance(Me.txtPrimary.Text, rowD.Cells(Me.colEmpl_Primary.Index).Value, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, Tbl2) = True Then
                    If ListMsg_1 = "" Then
                        ListMsg_1 &= "Employee : " & rowD.Cells(Me.colEmpl_Name.Index).Value
                    Else
                        ListMsg_1 &= ", " & rowD.Cells(Me.colEmpl_Name.Index).Value
                    End If
                End If
            Next

            If ListMsg_1 <> "" Then
                MsgBox("Employee data is created at attendance transaction." & vbCrLf & ListMsg_1, MsgBoxStyle.Information, "Information")
                GoTo ToBuffer
            End If
            If ListMsg_2 <> "" Then
                MsgBox("Employee data is created at other transaction." & vbCrLf & ListMsg_2, MsgBoxStyle.Information, "Information")
                GoTo ToBuffer
            End If

            If Me.dgvEmployee.RowCount <= 1 Then
                Exit Sub
            End If

            If Me.CekSudahAllocate(Nothing, -1, True, True) = True Then
                For i As Integer = 0 To Me.dgvEmployee.RowCount - 2
                    rowD = Me.dgvEmployee.Rows(i)
                    Me.IsiTbl_EmpShift(rowD.Cells(Me.colEmpl_Primary.Index).Value)
                Next

                Me.IsiTbl_SchSetup() 'IsiTbl_SchDetail di panggil dalam fungsi ini

                If Me.TabControl1.SelectedIndex = 1 Then
                    If Me.TabControl2.SelectedIndex = 0 Then
                        Me.ObjItem.IsiDGVItemSchSetup(Me.dgvSchSetup, Me.txtPrimary.Text, Me, Me.TblSchSetup)
                    ElseIf Me.TabControl2.SelectedIndex = 1 Then
                        Me.ObjItem.IsiDGVItemSchDetail(Me.dgvSchDetail, Me.txtPrimary.Text, Me, Me.TblSchDet)
                    Else
                        Me.ObjItem.IsiDGVItemSchEmp(Me.dgvSchEmp, Me.txtPrimary.Text, Me, Me.TblAppOT)
                    End If
                ElseIf Me.TabControl1.SelectedIndex = 2 Then
                    Me.ObjItem.IsiDGVItemApproval(Me.dgvApproval, Me.txtPrimary.Text, Me, Me.TblApp)
                End If
            Else
ToBuffer:
                If Me.ApakahDateFrom = True Then
                    Me.DtpFrom.Value = Me.BufferDate
                    Me.DtpFrom.Focus()
                Else
                    Me.DtpTo.Value = Me.BufferDate
                    Me.DtpTo.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.baru = True Then
                Me.txtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.txtPrimary.Text)
            Else
                Dim TransDate As Date = Me.Obj.GetTanggalTransaksi(Me.txtPrimary.Text)
                If Me.dtpDate.Value.Month = TransDate.Month And _
                Me.dtpDate.Value.Year = TransDate.Year Then
                    Me.txtDocumentNo.Text = Me.Obj.GetNoTransaksiByPrimary(Me.txtPrimary.Text)
                Else
                    If Me.dtpDate.Value.Month <> Me.Obj.Date_.Month Or _
                    Me.dtpDate.Value.Year <> Me.Obj.Date_.Year Then
                        Me.txtDocumentNo.Text = Me.Obj.GetNoBaru(Me.dtpDate.Value, Me.txtPrimary.Text)
                    Else
                        Me.txtDocumentNo.Text = Me.Obj.GetNoTransaksiByPrimary(Me.txtPrimary.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.dgvEmployee.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvEmployee)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            'Threading.Thread.Sleep(1)
            Me.ChangeDtpFromTo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            'If ApakahDateFrom = True Then
            '    If Me.BufferDate <> Me.DtpFrom.Value.Date Then
            '        Me.ChangeDtpFromTo()
            '    End If
            'Else
            '    If Me.BufferDate <> Me.DtpTo.Value.Date Then
            '        Me.ChangeDtpFromTo()
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            Me.dgvSchDetail.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvSchDetail)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
        Try
            Me.dgvApproval.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvApproval)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker6_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker6.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker6_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted
        Try
            Me.dgvSchEmp.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvSchEmp)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvEmployee.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colEmpl_Code.Index Or e.ColumnIndex = Me.colEmpl_Remarks.Index Then
                'If e.ColumnIndex = Me.colEmpl_Code.Index Then
                '    If Me.dgvEmployee.Item(Me.colEmpl_Primary.Index, e.RowIndex).Value <> -1 Then
                '        e.Cancel = True
                '        Exit Sub
                '    End If
                'End If
                Me.Buffer = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemEmployee(ByVal vPrimaryEmp As Double, ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1

        Try
            For ij As Integer = 0 To Me.dgvEmployee.Rows.Count - 2
                RowDGV = Me.dgvEmployee.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.colEmpl_Primary.Index).Value = vPrimaryEmp Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvEmployee_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEmployee.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvEmployee.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim PrimaryAcuan As Double = -1
        Dim EmpCode As String = ""
        Dim EmpNIK As String = ""
        Dim EmpName As String = ""
        Dim PositionPrimary As Double = -1
        Dim PositionCode As String = ""
        Dim PositionDesc As String = ""
        Dim DeptPrimary As Double = -1
        Dim DeptCode As String = ""
        Dim DeptDesc As String = ""
        Dim ListPrimaryAppGroup As New List(Of Double)
        Dim PrimaryEmpAwal As Double = -1

        Try
            If IsNothing(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                    Exit Sub
                End If

                If e.ColumnIndex = Me.colEmpl_Code.Index Then
                    PrimaryEmpAwal = Me.dgvEmployee.Item(Me.colEmpl_Primary.Index, e.RowIndex).Value

                    'FUNGSI COL EMPLOYEE CODE
                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.objEmp.ApakahSudahAda(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimaryAcuan = Me.objEmp.GetPrimaryByCode(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                        index = Me.GetIndexItemEmployee(PrimaryAcuan, e.RowIndex)
                        If index < 0 Then
                            'jika employee ditambahkan maka isi dgv sch setup, dgv sch detail, dgv approval
                            If Me.GetEmployeeData(PrimaryAcuan, EmpCode, EmpNIK, EmpName, PositionPrimary, PositionCode, PositionDesc, DeptPrimary, DeptCode, DeptDesc, ListPrimaryAppGroup) = False Then
                                GoTo DeleteRows
                            End If
                            Me.dgvEmployee.Item(Me.colEmpl_Primary.Index, e.RowIndex).Value = PrimaryAcuan
                            Me.dgvEmployee.Item(Me.colEmpl_PrimaryPos.Index, e.RowIndex).Value = PositionPrimary
                            Me.dgvEmployee.Item(Me.colEmpl_PrimaryDept.Index, e.RowIndex).Value = DeptPrimary
                            Me.dgvEmployee.Item(Me.colEmpl_Name.Index, e.RowIndex).Value = EmpName
                            Me.dgvEmployee.Item(Me.colEmpl_NIK.Index, e.RowIndex).Value = EmpNIK
                            Me.dgvEmployee.Item(Me.colEmpl_JobPosition.Index, e.RowIndex).Value = PositionCode & " - " & PositionDesc
                            Me.dgvEmployee.Item(Me.colEmpl_Department.Index, e.RowIndex).Value = DeptCode & " - " & DeptDesc
                        Else
                            If e.RowIndex <> index Then
                                MsgBox("Duplication employee code, please input another code.", MsgBoxStyle.Information, "Information")
                                GoTo DeleteRows
                            End If
                        End If
                    Else
                        Dim ij As New FrmMasterKaryawanCari
                        ij.txtCari.Text = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.BackColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        If ij.PrimaryEmploye <> -1 Then
                            index = Me.GetIndexItemEmployee(ij.PrimaryEmploye, e.RowIndex)
                            If index < 0 Then
                                'jika employee ditambahkan maka isi dgv sch setup, dgv sch detail, dgv approval
                                If Me.GetEmployeeData(ij.PrimaryEmploye, EmpCode, EmpNIK, EmpName, PositionPrimary, PositionCode, PositionDesc, DeptPrimary, DeptCode, DeptDesc, ListPrimaryAppGroup) = False Then
                                    GoTo DeleteRows
                                End If
                                Me.dgvEmployee.Item(Me.colEmpl_Primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                Me.dgvEmployee.Item(Me.colEmpl_PrimaryPos.Index, e.RowIndex).Value = PositionPrimary
                                Me.dgvEmployee.Item(Me.colEmpl_PrimaryDept.Index, e.RowIndex).Value = DeptPrimary
                                Me.dgvEmployee.Item(Me.colEmpl_Code.Index, e.RowIndex).Value = EmpCode
                                Me.dgvEmployee.Item(Me.colEmpl_Name.Index, e.RowIndex).Value = EmpName
                                Me.dgvEmployee.Item(Me.colEmpl_NIK.Index, e.RowIndex).Value = EmpNIK
                                Me.dgvEmployee.Item(Me.colEmpl_JobPosition.Index, e.RowIndex).Value = PositionCode & " - " & PositionDesc
                                Me.dgvEmployee.Item(Me.colEmpl_Department.Index, e.RowIndex).Value = DeptCode & " - " & DeptDesc
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication employee code, please input another code.", MsgBoxStyle.Information, "Information")
                                    GoTo DeleteRows
                                Else
                                    'jika employee ditambahkan maka isi dgv sch setup, dgv sch detail, dgv approval
                                    If Me.GetEmployeeData(ij.PrimaryEmploye, EmpCode, EmpNIK, EmpName, PositionPrimary, PositionCode, PositionDesc, DeptPrimary, DeptCode, DeptDesc, ListPrimaryAppGroup) = False Then
                                        GoTo DeleteRows
                                    End If
                                    Me.dgvEmployee.Item(Me.colEmpl_Primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    Me.dgvEmployee.Item(Me.colEmpl_PrimaryPos.Index, e.RowIndex).Value = PositionPrimary
                                    Me.dgvEmployee.Item(Me.colEmpl_PrimaryDept.Index, e.RowIndex).Value = DeptPrimary
                                    Me.dgvEmployee.Item(Me.colEmpl_Code.Index, e.RowIndex).Value = EmpCode
                                    Me.dgvEmployee.Item(Me.colEmpl_Name.Index, e.RowIndex).Value = EmpName
                                    Me.dgvEmployee.Item(Me.colEmpl_NIK.Index, e.RowIndex).Value = EmpNIK
                                    Me.dgvEmployee.Item(Me.colEmpl_JobPosition.Index, e.RowIndex).Value = PositionCode & " - " & PositionDesc
                                    Me.dgvEmployee.Item(Me.colEmpl_Department.Index, e.RowIndex).Value = DeptCode & " - " & DeptDesc
                                End If
                            End If
                        ElseIf ij.PrimaryEmploye = -1 Then
                            GoTo DeleteRows
                        End If

                        ij = Nothing
                    End If

                    If Me.dgvEmployee.Item(Me.colEmpl_Primary.Index, e.RowIndex).Value <> PrimaryEmpAwal Then
                        Me.HapusTbl_EmpShift(PrimaryEmpAwal, True)
                        Me.HapusTbl_Approval(row)
                    End If

                    Me.IsiTbl_EmpShift(Me.dgvEmployee.Item(Me.colEmpl_Primary.Index, e.RowIndex).Value)
                    Me.IsiTbl_Approval(ListPrimaryAppGroup)
                    Me.IsiTbl_SchSetup() 'IsiTbl_SchDetail di panggil dalam fungsi ini
                ElseIf e.ColumnIndex = Me.colEmpl_Remarks.Index Then
                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
                End If
            Else
DeleteRows:
                If String.IsNullOrEmpty(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
                If String.IsNullOrEmpty(Me.Buffer) = True Then
                    If Me.dgvEmployee.Rows.Count - 1 <> e.RowIndex And e.ColumnIndex <> Me.colEmpl_Remarks.Index Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker2.RunWorkerAsync()
                    End If
                Else
                    If e.ColumnIndex <> Me.colEmpl_Remarks.Index Then
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Else
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvEmployee.RowsAdded
        Dim row As DataGridViewRow = Me.dgvEmployee.Rows(Me.dgvEmployee.RowCount - 1)

        Try
            row.Cells(Me.colEmpl_No.Index).Value = Me.dgvEmployee.RowCount
            row.Cells(Me.colEmpl_Code.Index).Value = ""
            row.Cells(Me.colEmpl_NIK.Index).Value = ""
            row.Cells(Me.colEmpl_Name.Index).Value = ""
            row.Cells(Me.colEmpl_JobPosition.Index).Value = ""
            row.Cells(Me.colEmpl_Department.Index).Value = ""
            row.Cells(Me.colEmpl_Remarks.Index).Value = ""
            row.Cells(Me.colEmpl_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvEmployee.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvEmployee, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvEmployee.UserDeletedRow
        Try
            'jika employee dihapus maka hapus juga dgv sch setup, dgv sch detail, dgv approval yang terkait dengan employee tsb
            If Me.dgvEmployee.RowCount - 1 = 0 Then
                'fungsi ini jika dgv nya habis maka hapus tbl approval karena jika user mempunyai full all employee tidak terhapus di fungsi hapus approval
                Me.TblApp.Rows.Clear()
            End If

            Me.UpdateTotalDuration()
            Me.ObjCD.SetNumber(Me.dgvEmployee, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvEmployee.UserDeletingRow
        Dim row As DataGridViewRow = Me.dgvEmployee.Rows(e.Row.Index)

        Try
            'jika employee dihapus maka cek dgv sch setup, dgv sch detail, dgv approval sudah digunakan pada transaksi attendance
            'dapatkan primary emp, primary shift, primary user
            'hapus tbl employee shift ada di dalam fungsi ini (hapus tbl sch setup, hapus tbl sch detail)
            Me.HapusTbl_EmpShift(Me.dgvEmployee.Item(Me.colEmpl_Primary.Index, e.Row.Index).Value, True)
            Me.HapusTbl_SchEmp_1(Me.dgvEmployee.Item(Me.colEmpl_Primary.Index, e.Row.Index).Value)
            'hapus tbl approval
            Me.HapusTbl_Approval(row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchSetup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSchSetup.CellBeginEdit
        Dim DurationVal As Double = 0

        Try
            If e.ColumnIndex <> Me.ColSetup_No.Index And e.ColumnIndex <> Me.ColSetup_ShiftCode.Index And _
            e.ColumnIndex <> Me.ColSetup_ShiftDesc.Index Then
                'is early
                If e.ColumnIndex = Me.ColSetup_EarlyOT.Index Or e.ColumnIndex = Me.ColSetup_EarlyDelay.Index Or _
                e.ColumnIndex = Me.ColSetup_EarlyBreak.Index Or e.ColumnIndex = Me.ColSetup_EarlyDelayBeforeBreak.Index Then
                    If Me.dgvSchSetup.Item(Me.ColSetup_IsEarlyOT.Index, e.RowIndex).Value = Boolean.FalseString Then
                        e.Cancel = True
                        Exit Sub
                    Else
                        If e.ColumnIndex = Me.ColSetup_EarlyDelayBeforeBreak.Index Then
                            DurationVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_EarlyBreak.Index, e.RowIndex).Value)
                            If DurationVal <= 0 Then
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                'is ot
                If e.ColumnIndex = Me.ColSetup_OT.Index Or e.ColumnIndex = Me.ColSetup_OTDelay.Index Or _
                e.ColumnIndex = Me.ColSetup_OTBreak.Index Or e.ColumnIndex = Me.ColSetup_OTDelayBeforeBreak.Index Then
                    If Me.dgvSchSetup.Item(Me.ColSetup_IsOT.Index, e.RowIndex).Value = Boolean.FalseString Then
                        e.Cancel = True
                        Exit Sub
                    Else
                        If e.ColumnIndex = Me.ColSetup_OTDelayBeforeBreak.Index Then
                            DurationVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_OTBreak.Index, e.RowIndex).Value)
                            If DurationVal <= 0 Then
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                'is holiday ot
                If e.ColumnIndex = Me.ColSetup_HolOT.Index Or e.ColumnIndex = Me.ColSetup_HolOTDelay.Index Or _
                e.ColumnIndex = Me.ColSetup_HolOTBreak.Index Or e.ColumnIndex = Me.ColSetup_HolOTDelayBeforeBreak.Index Then
                    If Me.dgvSchSetup.Item(Me.ColSetup_IsHolOT.Index, e.RowIndex).Value = Boolean.FalseString Then
                        e.Cancel = True
                        Exit Sub
                    Else
                        If e.ColumnIndex = Me.ColSetup_HolOTDelayBeforeBreak.Index Then
                            DurationVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_HolOTBreak.Index, e.RowIndex).Value)
                            If DurationVal <= 0 Then
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If
                    End If
                End If

                Me.Buffer = Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchSetup_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSchSetup.CellContentClick
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchSetup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSchSetup.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvSchSetup.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim PrimaryAcuan As Double = -1
        Dim duration As Double
        Dim BreakDuration As Double
        Dim delaybeforebreak As Double
        Dim MaksHours_Str As String = ""

        Try
            If IsNothing(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                    Exit Sub
                End If

                'is true/false ########################################################################################
                If e.ColumnIndex = Me.ColSetup_IsEarlyOT.Index Then
                    If Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString Then
                        Me.dgvSchSetup.Item(Me.ColSetup_EarlyOT.Index, e.RowIndex).Value = "00.00"
                        Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelay.Index, e.RowIndex).Value = "00.00"
                        Me.dgvSchSetup.Item(Me.ColSetup_EarlyBreak.Index, e.RowIndex).Value = "00.00"
                        Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelayBeforeBreak.Index, e.RowIndex).Value = "00.00"
                    End If
                ElseIf e.ColumnIndex = Me.ColSetup_IsOT.Index Then
                    If Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString Then
                        Me.dgvSchSetup.Item(Me.ColSetup_OT.Index, e.RowIndex).Value = "00.00"
                        Me.dgvSchSetup.Item(Me.ColSetup_OTDelay.Index, e.RowIndex).Value = "00.00"
                        Me.dgvSchSetup.Item(Me.ColSetup_OTBreak.Index, e.RowIndex).Value = "00.00"
                        Me.dgvSchSetup.Item(Me.ColSetup_OTDelayBeforeBreak.Index, e.RowIndex).Value = "00.00"
                    End If
                ElseIf e.ColumnIndex = Me.ColSetup_IsHolOT.Index Then
                    If Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString Then
                        Me.dgvSchSetup.Item(Me.ColSetup_HolOT.Index, e.RowIndex).Value = "00.00"
                        Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelay.Index, e.RowIndex).Value = "00.00"
                        Me.dgvSchSetup.Item(Me.ColSetup_HolOTBreak.Index, e.RowIndex).Value = "00.00"
                        Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelayBeforeBreak.Index, e.RowIndex).Value = "00.00"
                    End If

                    'duration early ########################################################################################
                ElseIf e.ColumnIndex = Me.ColSetup_EarlyOT.Index Then
                    If ObjCD.Is_Time_String(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                        MsgBox("Early OT in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    duration = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_EarlyOT.Index, e.RowIndex).Value)
                    BreakDuration = 0 'Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_EarlyBreak.Index, e.RowIndex).Value)
                    delaybeforebreak = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelayBeforeBreak.Index, e.RowIndex).Value)
                    If duration <= (BreakDuration + delaybeforebreak) And duration = 0 And delaybeforebreak <> 0 Then
                        MsgBox("Duration not smallest than " & Me.ObjCD.Convert_ToString(BreakDuration + delaybeforebreak + 1) & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.CekScheduleSetupHours(Me.dgvSchSetup.Item(Me.ColSetup_EarlyOT.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelay.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_EarlyBreak.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelayBeforeBreak.Index, e.RowIndex).Value, MaksHours_Str, EnumSchSetup.OT_Enum) = False Then
                        MsgBox("Early OT in not more than " & MaksHours_Str & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColSetup_EarlyDelay.Index Then
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                        MsgBox("Early delay in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If ObjCD.Is_Time_String(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.CekScheduleSetupHours(Me.dgvSchSetup.Item(Me.ColSetup_EarlyOT.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelay.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_EarlyBreak.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelayBeforeBreak.Index, e.RowIndex).Value, MaksHours_Str, EnumSchSetup.Delay_Enum) = False Then
                        MsgBox("Early delay in not more than " & MaksHours_Str & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColSetup_EarlyBreak.Index Then
                    If ObjCD.Is_Time_String(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                        MsgBox("Early break in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = 0 Then
                        Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelayBeforeBreak.Index, e.RowIndex).Value = "00.00"
                    End If

                    If Me.CekScheduleSetupHours(Me.dgvSchSetup.Item(Me.ColSetup_EarlyOT.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelay.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_EarlyBreak.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelayBeforeBreak.Index, e.RowIndex).Value, MaksHours_Str, EnumSchSetup.Break_Enum) = False Then
                        MsgBox("Early break in not more than " & MaksHours_Str & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColSetup_EarlyDelayBeforeBreak.Index Then
                    If ObjCD.Is_Time_String(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                        MsgBox("Early delay before break in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    duration = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_EarlyOT.Index, e.RowIndex).Value)
                    BreakDuration = 0 'Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_EarlyBreak.Index, e.RowIndex).Value)
                    delaybeforebreak = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelayBeforeBreak.Index, e.RowIndex).Value)
                    If delaybeforebreak >= (duration - BreakDuration) Then
                        If duration = 0 Then
                            MsgBox("Please insert duration first in early OT", MsgBoxStyle.Information, "Information")
                        Else
                            MsgBox("Early delay before break not more than " & Me.ObjCD.Convert_ToString(duration - BreakDuration - 1) & " minutes", MsgBoxStyle.Information, "Information")
                        End If
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.CekScheduleSetupHours(Me.dgvSchSetup.Item(Me.ColSetup_EarlyOT.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelay.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_EarlyBreak.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_EarlyDelayBeforeBreak.Index, e.RowIndex).Value, MaksHours_Str, EnumSchSetup.DelayBefBreak_Enum) = False Then
                        MsgBox("Early delay before break in not more than " & MaksHours_Str & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'duration ot ########################################################################################
                ElseIf e.ColumnIndex = Me.ColSetup_OT.Index Then
                    If ObjCD.Is_Time_String(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                        MsgBox("OT in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    duration = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_OT.Index, e.RowIndex).Value)
                    BreakDuration = 0 'Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_OTBreak.Index, e.RowIndex).Value)
                    delaybeforebreak = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_OTDelayBeforeBreak.Index, e.RowIndex).Value)
                    If duration <= (BreakDuration + delaybeforebreak) And duration = 0 And delaybeforebreak <> 0 Then
                        MsgBox("Duration not smallest than " & Me.ObjCD.Convert_ToString(BreakDuration + delaybeforebreak + 1) & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.CekScheduleSetupHours(Me.dgvSchSetup.Item(Me.ColSetup_OT.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_OTDelay.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_OTBreak.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_OTDelayBeforeBreak.Index, e.RowIndex).Value, MaksHours_Str, EnumSchSetup.OT_Enum) = False Then
                        MsgBox("OT in not more than " & MaksHours_Str & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColSetup_OTDelay.Index Then
                    If ObjCD.Is_Time_String(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                        MsgBox("OT delay in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.CekScheduleSetupHours(Me.dgvSchSetup.Item(Me.ColSetup_OT.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_OTDelay.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_OTBreak.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_OTDelayBeforeBreak.Index, e.RowIndex).Value, MaksHours_Str, EnumSchSetup.Delay_Enum) = False Then
                        MsgBox("OT delay in not more than " & MaksHours_Str & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColSetup_OTBreak.Index Then
                    If ObjCD.Is_Time_String(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                        MsgBox("OT break in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = 0 Then
                        Me.dgvSchSetup.Item(Me.ColSetup_OTDelayBeforeBreak.Index, e.RowIndex).Value = "00.00"
                    End If

                    If Me.CekScheduleSetupHours(Me.dgvSchSetup.Item(Me.ColSetup_OT.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_OTDelay.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_OTBreak.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_OTDelayBeforeBreak.Index, e.RowIndex).Value, MaksHours_Str, EnumSchSetup.Break_Enum) = False Then
                        MsgBox("OT break in not more than " & MaksHours_Str & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColSetup_OTDelayBeforeBreak.Index Then
                    If ObjCD.Is_Time_String(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                        MsgBox("OT delay before break in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    duration = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_OT.Index, e.RowIndex).Value)
                    BreakDuration = 0 'Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_OTBreak.Index, e.RowIndex).Value)
                    delaybeforebreak = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_OTDelayBeforeBreak.Index, e.RowIndex).Value)
                    If delaybeforebreak >= (duration - BreakDuration) Then
                        If duration = 0 Then
                            MsgBox("Please insert duration first in OT", MsgBoxStyle.Information, "Information")
                        Else
                            MsgBox("Delay before break not more than " & Me.ObjCD.Convert_ToString(duration - BreakDuration - 1) & " minutes", MsgBoxStyle.Information, "Information")
                        End If
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.CekScheduleSetupHours(Me.dgvSchSetup.Item(Me.ColSetup_OT.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_OTDelay.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_OTBreak.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_OTDelayBeforeBreak.Index, e.RowIndex).Value, MaksHours_Str, EnumSchSetup.DelayBefBreak_Enum) = False Then
                        MsgBox("OT delay before break in not more than " & MaksHours_Str & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'duration holiday ot ########################################################################################
                ElseIf e.ColumnIndex = Me.ColSetup_HolOT.Index Then
                    If ObjCD.Is_Time_String(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                        MsgBox("Holiday OT in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    duration = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_HolOT.Index, e.RowIndex).Value)
                    BreakDuration = 0 'Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_HolOTBreak.Index, e.RowIndex).Value)
                    delaybeforebreak = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelayBeforeBreak.Index, e.RowIndex).Value)
                    If duration <= (BreakDuration + delaybeforebreak) And duration = 0 And delaybeforebreak <> 0 Then
                        MsgBox("Duration not smallest than " & Me.ObjCD.Convert_ToString(BreakDuration + delaybeforebreak + 1) & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.CekScheduleSetupHours(Me.dgvSchSetup.Item(Me.ColSetup_HolOT.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelay.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_HolOTBreak.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelayBeforeBreak.Index, e.RowIndex).Value, MaksHours_Str, EnumSchSetup.HolidayOT_Enum) = False Then
                        MsgBox("Holiday OT in not more than " & MaksHours_Str & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColSetup_HolOTDelay.Index Then
                    If ObjCD.Is_Time_String(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                        MsgBox("Holiday delay in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.CekScheduleSetupHours(Me.dgvSchSetup.Item(Me.ColSetup_HolOT.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelay.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_HolOTBreak.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelayBeforeBreak.Index, e.RowIndex).Value, MaksHours_Str, EnumSchSetup.HolidayDelay_Enum) = False Then
                        MsgBox("Holiday delay in not more than " & MaksHours_Str & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColSetup_HolOTBreak.Index Then
                    If ObjCD.Is_Time_String(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                        MsgBox("Holiday OT break in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = 0 Then
                        Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelayBeforeBreak.Index, e.RowIndex).Value = "00.00"
                    End If

                    If Me.CekScheduleSetupHours(Me.dgvSchSetup.Item(Me.ColSetup_HolOT.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelay.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_HolOTBreak.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelayBeforeBreak.Index, e.RowIndex).Value, MaksHours_Str, EnumSchSetup.HolidayBreak_Enum) = False Then
                        MsgBox("Holiday break in not more than " & MaksHours_Str & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColSetup_HolOTDelayBeforeBreak.Index Then
                    If ObjCD.Is_Time_String(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                        MsgBox("Holiday OT delay before break in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    duration = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_HolOT.Index, e.RowIndex).Value)
                    BreakDuration = 0 'Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_HolOTBreak.Index, e.RowIndex).Value)
                    delaybeforebreak = Me.ObjCD.Convert_ToMinutes(Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelayBeforeBreak.Index, e.RowIndex).Value)
                    If delaybeforebreak >= (duration - BreakDuration) Then
                        If duration = 0 Then
                            MsgBox("Please insert duration first in holiday OT", MsgBoxStyle.Information, "Information")
                        Else
                            MsgBox("Holiday delay before break not more than " & Me.ObjCD.Convert_ToString(duration - BreakDuration - 1) & " minutes", MsgBoxStyle.Information, "Information")
                        End If
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.CekScheduleSetupHours(Me.dgvSchSetup.Item(Me.ColSetup_HolOT.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelay.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_HolOTBreak.Index, e.RowIndex).Value, _
                    Me.dgvSchSetup.Item(Me.ColSetup_HolOTDelayBeforeBreak.Index, e.RowIndex).Value, MaksHours_Str, EnumSchSetup.HolidayDelayBefBreak_Enum) = False Then
                        MsgBox("Holiday OT delay before break in not more than " & MaksHours_Str & " minutes", MsgBoxStyle.Information, "Information")
                        Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColSetup_Remarks.Index Then
                    Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
                End If

                Me.UpdTbl_SchSetup(row)
            Else
                If e.ColumnIndex <> Me.ColSetup_Remarks.Index Then
                    Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Else
                    Me.dgvSchSetup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekScheduleSetupHours(ByVal vOT As String, ByVal vDelay As String, ByVal vBreak As String, ByVal vDelayBefBreak As String, _
    ByRef vMaksHours As String, ByVal vEnumSetup As EnumSchSetup) As Boolean
        'PASTIKAN JIKA OT + DELAY + BREAK + DELAY BEFORE BREAK TIDAK SAMPAI MELEBIHI JAM KERJA HARI BERIKUT ATAU SELUMNYA (24.00)
        Dim OTVal As Double = Me.ObjCD.Convert_ToMinutes(vOT)
        Dim DelayVal As Double = Me.ObjCD.Convert_ToMinutes(vDelay)
        Dim BreakVal As Double = Me.ObjCD.Convert_ToMinutes(vBreak)
        Dim DelayBefBreakVal As Double = Me.ObjCD.Convert_ToMinutes(vDelayBefBreak)
        Dim TotalHours As Double = OTVal + DelayVal + BreakVal + DelayBefBreakVal
        Dim MaksHours As Double = 0
        Dim MaksHoursDays As Double = 0

        Try
            If vEnumSetup = EnumSchSetup.OT_Enum Then
                MaksHoursDays = (16 * 60)
                MaksHours = MaksHoursDays - (DelayVal + BreakVal + DelayBefBreakVal)
            ElseIf vEnumSetup = EnumSchSetup.Delay_Enum Then
                MaksHoursDays = (16 * 60)
                MaksHours = MaksHoursDays - (OTVal + BreakVal + DelayBefBreakVal)
            ElseIf vEnumSetup = EnumSchSetup.Break_Enum Then
                MaksHoursDays = (16 * 60)
                MaksHours = MaksHoursDays - (OTVal + DelayVal + DelayBefBreakVal)
            ElseIf vEnumSetup = EnumSchSetup.DelayBefBreak_Enum Then
                MaksHoursDays = (16 * 60)
                MaksHours = MaksHoursDays - (OTVal + DelayVal + BreakVal)
            ElseIf vEnumSetup = EnumSchSetup.HolidayOT_Enum Then
                MaksHoursDays = (24 * 60)
                MaksHours = MaksHoursDays - (DelayVal + BreakVal + DelayBefBreakVal)
            ElseIf vEnumSetup = EnumSchSetup.HolidayDelay_Enum Then
                MaksHoursDays = (24 * 60)
                MaksHours = MaksHoursDays - (OTVal + BreakVal + DelayBefBreakVal)
            ElseIf vEnumSetup = EnumSchSetup.HolidayBreak_Enum Then
                MaksHoursDays = (24 * 60)
                MaksHours = MaksHoursDays - (OTVal + DelayVal + DelayBefBreakVal)
            ElseIf vEnumSetup = EnumSchSetup.HolidayDelayBefBreak_Enum Then
                MaksHoursDays = (24 * 60)
                MaksHours = MaksHoursDays - (OTVal + DelayVal + BreakVal)
            End If

            If TotalHours > MaksHoursDays Then
                vMaksHours = Me.ObjCD.Convert_ToString(MaksHours)
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        End Try
    End Function

    Private Sub dgvSchSetup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSchSetup.DataError

    End Sub

    Private Sub dgvSchSetup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvSchSetup.RowsAdded
        Dim row As DataGridViewRow = Me.dgvSchSetup.Rows(Me.dgvSchSetup.RowCount - 1)

        Try
            row.Cells(Me.ColSetup_No.Index).Value = Me.dgvSchSetup.Rows.Count
            row.Cells(Me.ColSetup_ShiftCode.Index).Value = ""
            row.Cells(Me.ColSetup_ShiftDesc.Index).Value = ""
            row.Cells(Me.ColSetup_IsEarlyOT.Index).Value = Boolean.FalseString
            row.Cells(Me.ColSetup_EarlyOT.Index).Value = "00.00"
            row.Cells(Me.ColSetup_EarlyDelay.Index).Value = "00.00"
            row.Cells(Me.ColSetup_EarlyBreak.Index).Value = "00.00"
            row.Cells(Me.ColSetup_EarlyDelayBeforeBreak.Index).Value = "00.00"
            row.Cells(Me.ColSetup_IsOT.Index).Value = Boolean.FalseString
            row.Cells(Me.ColSetup_OT.Index).Value = "00.00"
            row.Cells(Me.ColSetup_OTDelay.Index).Value = "00.00"
            row.Cells(Me.ColSetup_OTBreak.Index).Value = "00.00"
            row.Cells(Me.ColSetup_OTDelayBeforeBreak.Index).Value = "00.00"
            row.Cells(Me.ColSetup_IsHolOT.Index).Value = Boolean.FalseString
            row.Cells(Me.ColSetup_HolOT.Index).Value = "00.00"
            row.Cells(Me.ColSetup_HolOTDelay.Index).Value = "00.00"
            row.Cells(Me.ColSetup_HolOTBreak.Index).Value = "00.00"
            row.Cells(Me.ColSetup_HolOTDelayBeforeBreak.Index).Value = "00.00"
            row.Cells(Me.ColSetup_Remarks.Index).Value = ""
            row.Cells(Me.ColSetup_PrimaryWorkshift.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchSetup_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSchSetup.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvSchSetup, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchSetup_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSchSetup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvSchSetup, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchSetup_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvSchSetup.UserDeletingRow

    End Sub

    Private Sub dgvSchDetail_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSchDetail.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColSch_Start.Index Or e.ColumnIndex = Me.ColSch_Break.Index Or _
            e.ColumnIndex = Me.ColSch_End.Index Then
                If Me.dgvSchDetail.Item(Me.ColOverride.Index, e.RowIndex).Value = True Then
                    Me.Buffer = Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.ColOverride.Index Then
                Me.Buffer = Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value
            ElseIf e.ColumnIndex = Me.ColSch_Remark.Index Then
                Me.Buffer = Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchDetail_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSchDetail.CellContentClick

    End Sub

    Private Sub dgvSchDetail_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSchDetail.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvSchDetail.Rows(e.RowIndex)
        Dim ContainStrip As New List(Of String)
        Dim StartTimeVal As Integer = 0
        Dim BreakOutTimeVal As Integer = 0
        Dim BreakInTimeVal As Integer = 0
        Dim EndTimeVal As Integer = 0
        Dim StartWHVal As Integer = 0
        Dim EndWHVal As Integer = 0
        Dim NormalWH As String = ""
        Dim StartWH As String = ""
        Dim BreakOut As String = ""
        Dim BreakIn As String = ""
        Dim EndWH As String = ""
        Dim TotalDurDbl As Double = 0
        Dim ApakahHitungDuration As Boolean = True

        Try
            If IsNothing(Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                    Exit Sub
                End If

                If e.ColumnIndex = Me.ColSch_Start.Index Or e.ColumnIndex = Me.ColSch_Break.Index Or e.ColumnIndex = Me.ColSch_End.Index Or e.ColumnIndex = Me.ColOverride.Index Then
                    'GET VALUE NORMAL WORK HOURS
                    NormalWH = Me.dgvSchDetail.Item(Me.ColSch_NormalWH.Index, e.RowIndex).Value
                    StartWH = NormalWH.Remove(NormalWH.LastIndexOf("-"), NormalWH.Length - NormalWH.LastIndexOf("-"))
                    EndWH = NormalWH.Remove(0, NormalWH.LastIndexOf("-") + 1)

                    'CEK BATASAN FORMAT JAM & TIDAK BOLEH LEBIH DARI 24 JAM
                    If Me.CekFormatHour(Me.dgvSchDetail.Item(Me.ColSch_Start.Index, e.RowIndex).Value, _
                    Me.dgvSchDetail.Item(Me.ColSch_Break.Index, e.RowIndex).Value, _
                    Me.dgvSchDetail.Item(Me.ColSch_End.Index, e.RowIndex).Value, BreakOut, BreakIn) = False Then
                        'beri nila buffer
                        GoTo SkipFunc_1
                    End If

                    'CONVERT VALUE STRING TO MINUTE
                    StartWHVal = Me.ObjCD.Convert_ToMinutes(StartWH)
                    EndWHVal = Me.ObjCD.Convert_ToMinutes(EndWH)
                    StartTimeVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchDetail.Item(Me.ColSch_Start.Index, e.RowIndex).Value)
                    BreakOutTimeVal = Me.ObjCD.Convert_ToMinutes(BreakOut)
                    BreakInTimeVal = Me.ObjCD.Convert_ToMinutes(BreakIn)
                    EndTimeVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchDetail.Item(Me.ColSch_End.Index, e.RowIndex).Value)

                    'CEK BATASAN BERDASARKAN TIPE OVERTIME
                    Dim TipeOT_Enum As ClsTrans_SPL.EnumTipeOT = Me.GetTypeOT_Enum(Me.dgvSchDetail.Item(Me.ColSch_OTtype.Index, e.RowIndex).Value)
                    Dim UseBreak As Boolean = False
                    If String.IsNullOrEmpty(Me.dgvSchDetail.Item(Me.ColSch_Break.Index, e.RowIndex).Value) = False Then
                        UseBreak = True
                    End If

                    Dim EnumEndEdit As EnumOvertimeEndEdit
                    If e.ColumnIndex = Me.ColSch_Start.Index Then
                        EnumEndEdit = EnumOvertimeEndEdit.Start_Enum
                    ElseIf e.ColumnIndex = Me.ColSch_Break.Index Then
                        EnumEndEdit = EnumOvertimeEndEdit.Break_Enum
                    ElseIf e.ColumnIndex = Me.ColSch_End.Index Then
                        EnumEndEdit = EnumOvertimeEndEdit.End_Enum
                    End If

                    If Me.CekBatasanHitungHours(TipeOT_Enum, StartWHVal, EndWHVal, StartTimeVal, BreakOutTimeVal, BreakInTimeVal, EndTimeVal, UseBreak, EnumEndEdit) = False Then
                        'beri nila buffer
                        GoTo SkipFunc_1
                    End If

                    If UseBreak = False And BreakOutTimeVal = 0 And BreakInTimeVal = 0 Then
                        Me.dgvSchDetail.Item(Me.ColSch_Break.Index, e.RowIndex).Value = ""
                    Else
                        Me.dgvSchDetail.Item(Me.ColSch_Break.Index, e.RowIndex).Value = Me.ObjCD.Convert_ToString(BreakOutTimeVal) & " - " & Me.ObjCD.Convert_ToString(BreakInTimeVal)
                    End If

                    Dim SkipFunc As Boolean = False
                    If SkipFunc = True Then
SkipFunc_1:
                        If e.ColumnIndex = Me.ColSch_Start.Index Then
                            Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        ElseIf e.ColumnIndex = Me.ColSch_Break.Index Then
                            Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        ElseIf e.ColumnIndex = Me.ColSch_End.Index Then
                            Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If

                        Exit Sub
                    End If

                    If e.ColumnIndex = Me.ColOverride.Index Then
                        If Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value = False Then
                            Me.UpdDgv_SchDetailUnOverride(row)
                            ApakahHitungDuration = False
                        End If
                    ElseIf e.ColumnIndex = Me.ColSch_Remark.Index Then
                        Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
                    End If

                    If ApakahHitungDuration = True Then
                        If BreakOutTimeVal = 0 And BreakInTimeVal = 0 Then
                            'jika tidak menggunakan break
                            If StartTimeVal > EndTimeVal Then
                                StartTimeVal = StartTimeVal - (24 * 60)
                            End If
                        Else
                            'jika menggunakan break
                            If BreakInTimeVal > EndTimeVal Then
                                BreakInTimeVal = BreakInTimeVal - (24 * 60)
                            End If
                            If BreakOutTimeVal > BreakInTimeVal Then
                                BreakOutTimeVal = BreakOutTimeVal - (24 * 60)
                            End If
                            If StartTimeVal > BreakOutTimeVal Then
                                StartTimeVal = StartTimeVal - (24 * 60)
                            End If
                        End If

                        TotalDurDbl = (EndTimeVal - BreakInTimeVal) + (BreakOutTimeVal - StartTimeVal)
                        Me.dgvSchDetail.Item(Me.ColSch_Duration.Index, e.RowIndex).Value = Me.ObjCD.Convert_ToString(TotalDurDbl)
                    End If
                End If

                Me.UpdTbl_SchDetail(row)
            Else
                If e.ColumnIndex = Me.ColSch_Break.Index Then
                    StartTimeVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchDetail.Item(Me.ColSch_Start.Index, e.RowIndex).Value)
                    BreakOutTimeVal = 0
                    BreakInTimeVal = 0
                    EndTimeVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchDetail.Item(Me.ColSch_End.Index, e.RowIndex).Value)

                    'jika tidak menggunakan break
                    If StartTimeVal > EndTimeVal Then
                        StartTimeVal = StartTimeVal - (24 * 60)
                    End If

                    TotalDurDbl = (EndTimeVal - BreakInTimeVal) + (BreakOutTimeVal - StartTimeVal)
                    Me.dgvSchDetail.Item(Me.ColSch_Duration.Index, e.RowIndex).Value = Me.ObjCD.Convert_ToString(TotalDurDbl)

                    Me.UpdTbl_SchDetail(row)

                    Exit Sub
                End If

                If e.ColumnIndex <> Me.ColSch_Remark.Index Then
                    Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Else
                    Me.dgvSchDetail.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchDetail_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSchDetail.DataError

    End Sub

    Private Sub dgvSchDetail_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvSchDetail.RowsAdded
        Dim row As DataGridViewRow = Me.dgvSchDetail.Rows(Me.dgvSchDetail.RowCount - 1)

        Try
            row.Cells(Me.ColSch_No.Index).Value = Me.dgvSchDetail.RowCount
            row.Cells(Me.ColSch_Date.Index).Value = ""
            row.Cells(Me.ColSch_DayOfWeek.Index).Value = ""
            row.Cells(Me.ColSch_Shift.Index).Value = ""
            row.Cells(Me.ColSch_NormalWH.Index).Value = ""
            row.Cells(Me.ColSch_OTtype.Index).Value = ""
            row.Cells(Me.ColSch_Start.Index).Value = "00.00 - 00.00"
            row.Cells(Me.ColSch_Break.Index).Value = "00.00"
            row.Cells(Me.ColSch_End.Index).Value = "00.00"
            row.Cells(Me.ColSch_Duration.Index).Value = "00.00"
            row.Cells(Me.ColOverride.Index).Value = Boolean.FalseString
            row.Cells(Me.ColSch_Remark.Index).Value = ""
            row.Cells(Me.ColSch_PrimaryShift.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchDetail_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSchDetail.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvSchDetail, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchDetail_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSchDetail.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvSchDetail, True)
            Me.UpdateTotalDuration()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchDetail_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvSchDetail.UserDeletingRow
        Try
            Me.HapusTbl_SchDetail_1(Me.dgvSchDetail.Item(Me.ColSch_PrimaryShift.Index, e.Row.Index).Value, _
            CDate(Me.dgvSchDetail.Item(Me.ColSch_Date.Index, e.Row.Index).Value), Me.GetTypeOT_Enum(Me.dgvSchDetail.Item(Me.ColSch_OTtype.Index, e.Row.Index).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchEmp_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSchEmp.CellBeginEdit
        Try
            If Me.Obj.ApaSudahDiBuatAttendance(Me.dgvSchEmp.Item(Me.ColSchEmp_PrimaryEmp.Index, e.RowIndex).Value, CDate(Me.dgvSchEmp.Item(Me.ColSchEmp_Date.Index, e.RowIndex).Value)) = True Then
                MsgBox("You can't edit employee in " & Me.dgvSchEmp.Item(Me.ColSchEmp_Date.Index, e.RowIndex).Value & ". Because employee has been created at attendance transaction.", MsgBoxStyle.Information, "Information")
                Me.RefreshDgvSchEmp = False
                e.Cancel = True
                Exit Sub
            End If

            If e.ColumnIndex = Me.ColSchEmp_StartEarly.Index Or e.ColumnIndex = Me.ColSchEmp_BreakEarly.Index Or _
            e.ColumnIndex = Me.ColSchEmp_EndEarly.Index Then
                If Me.dgvSchEmp.Item(Me.ColSchEmp_OverrideEarlyOT.Index, e.RowIndex).Value = True Then
                    Me.Buffer = Me.dgvSchEmp.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.ColSchEmp_StartOT.Index Or e.ColumnIndex = Me.ColSchEmp_BreakOT.Index Or _
            e.ColumnIndex = Me.ColSchEmp_EndOT.Index Then
                If Me.dgvSchEmp.Item(Me.ColSchEmp_OverrideOT.Index, e.RowIndex).Value = True Then
                    Me.Buffer = Me.dgvSchEmp.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.ColSchEmp_OverrideEarlyOT.Index Then
                If Me.dgvSchEmp.Item(Me.ColSchEmp_IsHoliday.Index, e.RowIndex).Value = "1" Then
                    MsgBox("You can't override early overtime, because holiday.", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                Else
                    Me.Buffer = Me.dgvSchEmp.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            ElseIf e.ColumnIndex = Me.ColSchEmp_OverrideOT.Index Then
                Me.Buffer = Me.dgvSchEmp.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchEmp_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSchEmp.CellContentClick

    End Sub

    Private Sub dgvSchEmp_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSchEmp.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvSchEmp.Rows(e.RowIndex)
        Dim StartTimeVal As Integer = 0
        Dim BreakOutTimeVal As Integer = 0
        Dim BreakInTimeVal As Integer = 0
        Dim EndTimeVal As Integer = 0
        Dim TotalDurDbl As Double = 0

        Try
            If IsNothing(Me.dgvSchEmp.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If Me.dgvSchEmp.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                    Exit Sub
                End If

                If e.ColumnIndex = Me.ColSchEmp_StartEarly.Index Or e.ColumnIndex = Me.ColSchEmp_BreakEarly.Index Or e.ColumnIndex = Me.ColSchEmp_EndEarly.Index _
                Or e.ColumnIndex = Me.ColSchEmp_EndEarly.Index Or e.ColumnIndex = Me.ColSchEmp_OverrideEarlyOT.Index Then
                    Me.DgvSchEmp_EndEditEarly(e.ColumnIndex, e.RowIndex)
                ElseIf e.ColumnIndex = Me.ColSchEmp_StartOT.Index Or e.ColumnIndex = Me.ColSchEmp_BreakOT.Index Or e.ColumnIndex = Me.ColSchEmp_EndOT.Index _
                Or e.ColumnIndex = Me.ColSchEmp_EndOT.Index Or e.ColumnIndex = Me.ColSchEmp_OverrideOT.Index Then
                    Me.DgvSchEmp_EndEditOT(e.ColumnIndex, e.RowIndex)
                End If

                Me.UpdTbl_SchEmp(row)
            Else
                If e.ColumnIndex = Me.ColSchEmp_BreakEarly.Index Then
                    StartTimeVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchEmp.Item(Me.ColSchEmp_StartEarly.Index, e.RowIndex).Value)
                    BreakOutTimeVal = 0
                    BreakInTimeVal = 0
                    EndTimeVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchEmp.Item(Me.ColSchEmp_EndEarly.Index, e.RowIndex).Value)

                    TotalDurDbl = (EndTimeVal - BreakInTimeVal) + (BreakOutTimeVal - StartTimeVal)
                    Me.dgvSchEmp.Item(Me.ColSchEmp_EarlyOTDuration.Index, e.RowIndex).Value = Me.ObjCD.Convert_ToString(TotalDurDbl)

                    Me.UpdTbl_SchEmp(row)

                    Exit Sub
                ElseIf e.ColumnIndex = Me.ColSchEmp_BreakOT.Index Then
                    StartTimeVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchEmp.Item(Me.ColSchEmp_StartOT.Index, e.RowIndex).Value)
                    BreakOutTimeVal = 0
                    BreakInTimeVal = 0
                    EndTimeVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchEmp.Item(Me.ColSchEmp_EndOT.Index, e.RowIndex).Value)

                    'jika tidak menggunakan break
                    If StartTimeVal > EndTimeVal Then
                        StartTimeVal = StartTimeVal - (24 * 60)
                    End If

                    TotalDurDbl = (EndTimeVal - BreakInTimeVal) + (BreakOutTimeVal - StartTimeVal)
                    Me.dgvSchEmp.Item(Me.ColSchEmp_OTDuration.Index, e.RowIndex).Value = Me.ObjCD.Convert_ToString(TotalDurDbl)

                    Me.UpdTbl_SchEmp(row)

                    Exit Sub
                End If

                Me.dgvSchEmp.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSchEmp_EndEditEarly(ByVal vColIdx As Integer, ByVal vRowIdx As Integer)
        Dim row As DataGridViewRow = Me.dgvSchEmp.Rows(vRowIdx)
        Dim ContainStrip As New List(Of String)
        Dim StartTimeVal As Integer = 0
        Dim BreakOutTimeVal As Integer = 0
        Dim BreakInTimeVal As Integer = 0
        Dim EndTimeVal As Integer = 0
        Dim StartWHVal As Integer = 0
        Dim EndWHVal As Integer = 0
        Dim NormalWH As String = ""
        Dim StartWH As String = ""
        Dim BreakOut As String = ""
        Dim BreakIn As String = ""
        Dim EndWH As String = ""
        Dim TotalDurDbl As Double = 0
        Dim ApakahHitungDuration As Boolean = True

        Try
            NormalWH = Me.dgvSchEmp.Item(Me.ColSchEmp_NormalWH.Index, vRowIdx).Value
            StartWH = NormalWH.Remove(NormalWH.LastIndexOf("-"), NormalWH.Length - NormalWH.LastIndexOf("-"))
            EndWH = NormalWH.Remove(0, NormalWH.LastIndexOf("-") + 1)

            'CEK BATASAN FORMAT JAM & TIDAK BOLEH LEBIH DARI 24 JAM
            If Me.CekFormatHour(Me.dgvSchEmp.Item(Me.ColSchEmp_StartEarly.Index, vRowIdx).Value, _
            Me.dgvSchEmp.Item(Me.ColSchEmp_BreakEarly.Index, vRowIdx).Value, _
            Me.dgvSchEmp.Item(Me.ColSchEmp_EndEarly.Index, vRowIdx).Value, BreakOut, BreakIn) = False Then
                'beri nila buffer
                GoTo SkipFunc_1
            End If

            StartWHVal = Me.ObjCD.Convert_ToMinutes(StartWH)
            EndWHVal = Me.ObjCD.Convert_ToMinutes(EndWH)
            StartTimeVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchEmp.Item(Me.ColSchEmp_StartEarly.Index, vRowIdx).Value)
            BreakOutTimeVal = Me.ObjCD.Convert_ToMinutes(BreakOut)
            BreakInTimeVal = Me.ObjCD.Convert_ToMinutes(BreakIn)
            EndTimeVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchEmp.Item(Me.ColSchEmp_EndEarly.Index, vRowIdx).Value)

            'CEK BATASAN BERDASARKAN TIPE OVERTIME
            Dim TipeOT_Enum As ClsTrans_SPL.EnumTipeOT = ClsTrans_SPL.EnumTipeOT.Early_Overtime
            Dim UseBreak As Boolean = False
            If String.IsNullOrEmpty(Me.dgvSchEmp.Item(Me.ColSchEmp_BreakEarly.Index, vRowIdx).Value) = False Then
                UseBreak = True
            End If
            'jika hari libur maka tidak bisa menggunakan early ot
            If Me.dgvSchEmp.Item(Me.ColSchEmp_IsHoliday.Index, vRowIdx).Value = "1" Then
                Exit Sub
            End If

            Dim EnumEndEdit As EnumOvertimeEndEdit
            If vColIdx = Me.ColSchEmp_StartEarly.Index Then
                EnumEndEdit = EnumOvertimeEndEdit.Start_Enum
            ElseIf vColIdx = Me.ColSchEmp_BreakEarly.Index Then
                EnumEndEdit = EnumOvertimeEndEdit.Break_Enum
            ElseIf vColIdx = Me.ColSchEmp_EndEarly.Index Then
                EnumEndEdit = EnumOvertimeEndEdit.End_Enum
            End If

            If Me.CekBatasanHitungHours(TipeOT_Enum, StartWHVal, EndWHVal, StartTimeVal, BreakOutTimeVal, BreakInTimeVal, EndTimeVal, UseBreak, EnumEndEdit) = False Then
                'beri nila buffer
                GoTo SkipFunc_1
            End If

            If UseBreak = False And BreakOutTimeVal = 0 And BreakInTimeVal = 0 Then
                Me.dgvSchEmp.Item(Me.ColSchEmp_BreakEarly.Index, vRowIdx).Value = ""
            Else
                Me.dgvSchEmp.Item(Me.ColSchEmp_BreakEarly.Index, vRowIdx).Value = Me.ObjCD.Convert_ToString(BreakOutTimeVal) & " - " & Me.ObjCD.Convert_ToString(BreakInTimeVal)
            End If

            Dim SkipFunc As Boolean = False
            If SkipFunc = True Then
SkipFunc_1:
                If vColIdx = Me.ColSchEmp_StartEarly.Index Then
                    Me.dgvSchEmp.Item(Me.ColSchEmp_StartEarly.Index, vRowIdx).Value = Me.Buffer
                ElseIf vColIdx = Me.ColSchEmp_BreakEarly.Index Then
                    Me.dgvSchEmp.Item(Me.ColSchEmp_BreakEarly.Index, vRowIdx).Value = Me.Buffer
                ElseIf vColIdx = Me.ColSchEmp_EndEarly.Index Then
                    Me.dgvSchEmp.Item(Me.ColSchEmp_EndEarly.Index, vRowIdx).Value = Me.Buffer
                End If

                Exit Sub
            End If

            If vColIdx = Me.ColSchEmp_OverrideEarlyOT.Index Then
                If Me.dgvSchEmp.Item(vColIdx, vRowIdx).Value = False Then
                    Me.UpdDgv_SchEmpUnOverride(vColIdx, row)
                    ApakahHitungDuration = False
                End If
            End If

            If ApakahHitungDuration = True Then
                TotalDurDbl = (EndTimeVal - BreakInTimeVal) + (BreakOutTimeVal - StartTimeVal)
                Me.dgvSchEmp.Item(Me.ColSchEmp_EarlyOTDuration.Index, vRowIdx).Value = Me.ObjCD.Convert_ToString(TotalDurDbl)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSchEmp_EndEditOT(ByVal vColIdx As Integer, ByVal vRowIdx As Integer)
        Dim row As DataGridViewRow = Me.dgvSchEmp.Rows(vRowIdx)
        Dim ContainStrip As New List(Of String)
        Dim StartTimeVal As Integer = 0
        Dim BreakOutTimeVal As Integer = 0
        Dim BreakInTimeVal As Integer = 0
        Dim EndTimeVal As Integer = 0
        Dim StartWHVal As Integer = 0
        Dim EndWHVal As Integer = 0
        Dim NormalWH As String = ""
        Dim StartWH As String = ""
        Dim BreakOut As String = ""
        Dim BreakIn As String = ""
        Dim EndWH As String = ""
        Dim TotalDurDbl As Double = 0
        Dim ApakahHitungDuration As Boolean = True

        Try
            NormalWH = Me.dgvSchEmp.Item(Me.ColSchEmp_NormalWH.Index, vRowIdx).Value
            StartWH = NormalWH.Remove(NormalWH.LastIndexOf("-"), NormalWH.Length - NormalWH.LastIndexOf("-"))
            EndWH = NormalWH.Remove(0, NormalWH.LastIndexOf("-") + 1)

            'CEK BATASAN FORMAT JAM & TIDAK BOLEH LEBIH DARI 24 JAM
            If Me.CekFormatHour(Me.dgvSchEmp.Item(Me.ColSchEmp_StartOT.Index, vRowIdx).Value, _
            Me.dgvSchEmp.Item(Me.ColSchEmp_BreakOT.Index, vRowIdx).Value, _
            Me.dgvSchEmp.Item(Me.ColSchEmp_EndOT.Index, vRowIdx).Value, BreakOut, BreakIn) = False Then
                'beri nila buffer
                GoTo SkipFunc_1
            End If

            StartWHVal = Me.ObjCD.Convert_ToMinutes(StartWH)
            EndWHVal = Me.ObjCD.Convert_ToMinutes(EndWH)
            StartTimeVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchEmp.Item(Me.ColSchEmp_StartOT.Index, vRowIdx).Value)
            BreakOutTimeVal = Me.ObjCD.Convert_ToMinutes(BreakOut)
            BreakInTimeVal = Me.ObjCD.Convert_ToMinutes(BreakIn)
            EndTimeVal = Me.ObjCD.Convert_ToMinutes(Me.dgvSchEmp.Item(Me.ColSchEmp_EndOT.Index, vRowIdx).Value)

            'CEK BATASAN BERDASARKAN TIPE OVERTIME
            Dim TipeOT_Enum As ClsTrans_SPL.EnumTipeOT = ClsTrans_SPL.EnumTipeOT.Overtime
            Dim UseBreak As Boolean = False
            If String.IsNullOrEmpty(Me.dgvSchEmp.Item(Me.ColSchEmp_BreakOT.Index, vRowIdx).Value) = False Then
                UseBreak = True
            End If
            If Me.dgvSchEmp.Item(Me.ColSchEmp_IsHoliday.Index, vRowIdx).Value = "1" Then
                TipeOT_Enum = ClsTrans_SPL.EnumTipeOT.Holiday_Overtime
            End If

            Dim EnumEndEdit As EnumOvertimeEndEdit
            If vColIdx = Me.ColSchEmp_StartOT.Index Then
                EnumEndEdit = EnumOvertimeEndEdit.Start_Enum
            ElseIf vColIdx = Me.ColSchEmp_BreakOT.Index Then
                EnumEndEdit = EnumOvertimeEndEdit.Break_Enum
            ElseIf vColIdx = Me.ColSchEmp_EndOT.Index Then
                EnumEndEdit = EnumOvertimeEndEdit.End_Enum
            End If

            If Me.CekBatasanHitungHours(TipeOT_Enum, StartWHVal, EndWHVal, StartTimeVal, BreakOutTimeVal, BreakInTimeVal, EndTimeVal, UseBreak, EnumEndEdit) = False Then
                'beri nila buffer
                GoTo SkipFunc_1
            End If

            If UseBreak = False And BreakOutTimeVal = 0 And BreakInTimeVal = 0 Then
                Me.dgvSchEmp.Item(Me.ColSchEmp_BreakOT.Index, vRowIdx).Value = ""
            Else
                Me.dgvSchEmp.Item(Me.ColSchEmp_BreakOT.Index, vRowIdx).Value = Me.ObjCD.Convert_ToString(BreakOutTimeVal) & " - " & Me.ObjCD.Convert_ToString(BreakInTimeVal)
            End If

            Dim SkipFunc As Boolean = False
            If SkipFunc = True Then
SkipFunc_1:
                If vColIdx = Me.ColSchEmp_StartOT.Index Then
                    Me.dgvSchEmp.Item(Me.ColSchEmp_StartOT.Index, vRowIdx).Value = Me.Buffer
                ElseIf vColIdx = Me.ColSchEmp_BreakOT.Index Then
                    Me.dgvSchEmp.Item(Me.ColSchEmp_BreakOT.Index, vRowIdx).Value = Me.Buffer
                ElseIf vColIdx = Me.ColSchEmp_EndOT.Index Then
                    Me.dgvSchEmp.Item(Me.ColSchEmp_EndOT.Index, vRowIdx).Value = Me.Buffer
                End If

                Exit Sub
            End If

            If vColIdx = Me.ColSchEmp_OverrideOT.Index Then
                If Me.dgvSchEmp.Item(vColIdx, vRowIdx).Value = False Then
                    Me.UpdDgv_SchEmpUnOverride(vColIdx, row)
                    ApakahHitungDuration = False
                End If
            End If

            If ApakahHitungDuration = True Then
                If BreakOutTimeVal = 0 And BreakInTimeVal = 0 Then
                    'jika tidak menggunakan break
                    If StartTimeVal > EndTimeVal Then
                        StartTimeVal = StartTimeVal - (24 * 60)
                    End If
                Else
                    'jika menggunakan break
                    If BreakInTimeVal > EndTimeVal Then
                        BreakInTimeVal = BreakInTimeVal - (24 * 60)
                    End If
                    If BreakOutTimeVal > BreakInTimeVal Then
                        BreakOutTimeVal = BreakOutTimeVal - (24 * 60)
                    End If
                    If StartTimeVal > BreakOutTimeVal Then
                        StartTimeVal = StartTimeVal - (24 * 60)
                    End If
                End If

                TotalDurDbl = (EndTimeVal - BreakInTimeVal) + (BreakOutTimeVal - StartTimeVal)
                Me.dgvSchEmp.Item(Me.ColSchEmp_OTDuration.Index, vRowIdx).Value = Me.ObjCD.Convert_ToString(TotalDurDbl)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdDgv_SchEmpUnOverride(ByVal vColIdx As Integer, ByVal vrowD As DataGridViewRow)
        Dim rowAppOT As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow
        Dim rowSchDetail As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailRow
        Dim PrimaryEmp As Double = vrowD.Cells(Me.ColSchEmp_PrimaryEmp.Index).Value
        Dim DateFilter As Date = CDate(vrowD.Cells(Me.ColSchEmp_Date.Index).Value)

        Try
            Me.TblAppOT.DefaultView.RowFilter = "[Primary Employee] = " & PrimaryEmp & " AND [Date] = #" & DateFilter & "#"
            If Me.TblAppOT.DefaultView.Count = 0 Then
                MsgBox("Row schedule employee not found", MsgBoxStyle.Information, "Information")
            Else
                rowAppOT = Me.TblAppOT.DefaultView(0).Row
                Me.TblSchDet.DefaultView.RowFilter = "[Primary Shift] = " & rowAppOT.Primary_Shift & " AND [Date] = #" & DateFilter & "#"
                If TblSchDet.DefaultView.Count = 0 Then
                    MsgBox("Row schedule detail not found", MsgBoxStyle.Information, "Information")
                Else
                    For Each dv As DataRowView In Me.TblSchDet.DefaultView
                        rowSchDetail = dv.Row
                        If rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Early_Overtime Then
                            If vColIdx = Me.ColSchEmp_OverrideEarlyOT.Index Then
                                rowAppOT.Start_Time_Early = rowSchDetail.Start
                                rowAppOT.End_Time_Early = rowSchDetail._End
                                If rowSchDetail.Break_Out <> 0 And rowSchDetail.Break_In <> 0 Then
                                    rowAppOT.Break_Out_Time_Early = rowSchDetail.Break_Out
                                    rowAppOT.Break_In_Time_Early = rowSchDetail.Break_In
                                Else
                                    rowAppOT.Break_Out_Time_Early = 0
                                    rowAppOT.Break_In_Time_Early = 0
                                End If
                                rowAppOT.OT_Early_Approved = rowSchDetail.Duration
                            End If

                            'untuk ot ##############################################################################
                            'untuk holiday ot ##############################################################################
                        ElseIf rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Overtime Or _
                        rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Holiday_Overtime Then
                            If vColIdx = Me.ColSchEmp_OverrideOT.Index Then
                                'OVERTIME - hitung start, break, end, duration 
                                'HOLIDAY OVERTIME - hitung start, break, end, duration
                                rowAppOT.Start_Time = rowSchDetail.Start
                                rowAppOT.End_Time = rowSchDetail._End
                                If rowSchDetail.Break_Out <> 0 And rowSchDetail.Break_In <> 0 Then
                                    rowAppOT.Break_Out_Time = rowSchDetail.Break_Out
                                    rowAppOT.Break_In_Time = rowSchDetail.Break_In
                                Else
                                    rowAppOT.Break_Out_Time = 0
                                    rowAppOT.Break_In_Time = 0
                                End If
                                rowAppOT.OT_Aprroved = rowSchDetail.Duration
                            End If
                        End If
                    Next

                    'update dgvschemployee
                    If vColIdx = Me.ColSchEmp_OverrideEarlyOT.Index Then
                        vrowD.Cells(Me.ColSchEmp_StartEarly.Index).Value = Me.ObjCD.Convert_ToString(rowAppOT.Start_Time_Early)
                        If rowAppOT.Break_Out_Time_Early = 0 Or rowAppOT.Break_In_Time_Early = 0 Then
                            vrowD.Cells(Me.ColSchEmp_BreakEarly.Index).Value = ""
                        Else
                            vrowD.Cells(Me.ColSchEmp_BreakEarly.Index).Value = Me.ObjCD.Convert_ToString(rowAppOT.Break_Out_Time_Early) & " - " & Me.ObjCD.Convert_ToString(rowAppOT.Break_In_Time_Early)
                        End If
                        vrowD.Cells(Me.ColSchEmp_EndEarly.Index).Value = Me.ObjCD.Convert_ToString(rowAppOT.End_Time_Early)
                        vrowD.Cells(Me.ColSchEmp_EarlyOTDuration.Index).Value = Me.ObjCD.Convert_ToString(rowAppOT.OT_Early_Approved)
                        vrowD.Cells(Me.ColSchEmp_OverrideEarlyOT.Index).Value = False
                    ElseIf vColIdx = Me.ColSchEmp_OverrideOT.Index Then
                        vrowD.Cells(Me.ColSchEmp_StartOT.Index).Value = Me.ObjCD.Convert_ToString(rowAppOT.Start_Time)
                        If rowAppOT.Break_Out_Time = 0 Or rowAppOT.Break_In_Time = 0 Then
                            vrowD.Cells(Me.ColSchEmp_BreakOT.Index).Value = ""
                        Else
                            vrowD.Cells(Me.ColSchEmp_BreakOT.Index).Value = Me.ObjCD.Convert_ToString(rowAppOT.Break_Out_Time) & " - " & Me.ObjCD.Convert_ToString(rowAppOT.Break_In_Time)
                        End If
                        vrowD.Cells(Me.ColSchEmp_EndOT.Index).Value = Me.ObjCD.Convert_ToString(rowAppOT.End_Time)
                        vrowD.Cells(Me.ColSchEmp_OTDuration.Index).Value = Me.ObjCD.Convert_ToString(rowAppOT.OT_Aprroved)
                        vrowD.Cells(Me.ColSchEmp_OverrideOT.Index).Value = False
                    End If
                End If
            End If

            'Me.UpdateTotalDuration()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchEmp_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSchEmp.DataError

    End Sub

    Private Sub dgvSchEmp_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvSchEmp.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvSchEmp.Rows(Me.dgvSchEmp.Rows.Count - 1)
            rowDGV.Cells(Me.ColSchEmp_No.Index).Value = Me.dgvSchEmp.Rows.Count
            rowDGV.Cells(Me.ColSchEmp_DayOfWeek.Index).Value = ""
            rowDGV.Cells(Me.ColSchEmp_Date.Index).Value = ""
            rowDGV.Cells(Me.ColSchEmp_Shift.Index).Value = ""
            rowDGV.Cells(Me.ColSchEmp_NormalWH.Index).Value = "00.00"
            rowDGV.Cells(Me.ColSchEmp_StartEarly.Index).Value = "00.00"
            rowDGV.Cells(Me.ColSchEmp_BreakEarly.Index).Value = "00.00"
            rowDGV.Cells(Me.ColSchEmp_EndEarly.Index).Value = "00.00"
            rowDGV.Cells(Me.ColSchEmp_EarlyOTDuration.Index).Value = "00.00"
            rowDGV.Cells(Me.ColSchEmp_OverrideEarlyOT.Index).Value = Boolean.FalseString
            rowDGV.Cells(Me.ColSchEmp_StartOT.Index).Value = "00.00"
            rowDGV.Cells(Me.ColSchEmp_BreakOT.Index).Value = "00.00"
            rowDGV.Cells(Me.ColSchEmp_EndOT.Index).Value = "00.00"
            rowDGV.Cells(Me.ColSchEmp_OTDuration.Index).Value = "00.00"
            rowDGV.Cells(Me.ColSchEmp_OverrideOT.Index).Value = Boolean.FalseString
            rowDGV.Cells(Me.ColSchEmp_PrimaryShift.Index).Value = -1
            rowDGV.Cells(Me.ColSchEmp_PrimaryEmp.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchEmp_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSchEmp.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvSchEmp, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSchEmp_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSchEmp.UserDeletedRow
        Try
            If Me.RefreshDgvSchEmp = True Then
                Me.ObjCD.SetNumber(Me.dgvSchEmp, True)
            End If
            Me.UpdateTotalDuration()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim RefreshDgvSchEmp As Boolean = True
    Private Sub dgvSchEmp_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvSchEmp.UserDeletingRow
        Me.RefreshDgvSchEmp = True

        Try
            If Me.Obj.ApaSudahDiBuatAttendance(Me.dgvSchEmp.Item(Me.ColSchEmp_PrimaryEmp.Index, e.Row.Index).Value, CDate(Me.dgvSchEmp.Item(Me.ColSchEmp_Date.Index, e.Row.Index).Value)) = True Then
                MsgBox("You can't remove employee in " & Me.dgvSchEmp.Item(Me.ColSchEmp_Date.Index, e.Row.Index).Value & ". Because employee has been created at attendance transaction.", MsgBoxStyle.Information, "Information")
                Me.RefreshDgvSchEmp = False
                e.Cancel = True
                Exit Sub
            End If

            Me.HapusTbl_SchEmp_2(Me.dgvSchEmp.Item(Me.ColSchEmp_PrimaryShift.Index, e.Row.Index).Value, _
            Me.dgvSchEmp.Item(Me.ColSchEmp_PrimaryEmp.Index, e.Row.Index).Value, _
            CDate(Me.dgvSchEmp.Item(Me.ColSchEmp_Date.Index, e.Row.Index).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvApproval_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvApproval.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColUser_RequestApproval.Index Or e.ColumnIndex = Me.ColUser_Remarks.Index Or _
            e.ColumnIndex = Me.ColUser_Approval.Index Then
                Me.Buffer = Me.dgvApproval.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvApproval_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvApproval.CellContentClick
        Dim i As New FrmTrans_SPLApprovalScope

        Try
            If e.ColumnIndex = Me.ColUser_Approval.Index Then
                i.PrimaryUser = Me.dgvApproval.Item(Me.ColUser_Primary.Index, e.RowIndex).Value
                i.UserApproval = "(" & Me.dgvApproval.Item(Me.ColUser_Code.Index, e.RowIndex).Value & " - " & Me.dgvApproval.Item(Me.ColUser_Name.Index, e.RowIndex).Value & ")"
                i.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub dgvApproval_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvApproval.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvApproval.Rows(e.RowIndex)

        Try
            If IsNothing(Me.dgvApproval.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                Me.UpdTbl_Approval(row)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvApproval_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvApproval.RowsAdded
        Dim row As DataGridViewRow = Me.dgvApproval.Rows(Me.dgvApproval.RowCount - 1)

        Try
            row.Cells(Me.ColUser_No.Index).Value = Me.dgvApproval.Rows.Count
            row.Cells(Me.ColUser_RequestApproval.Index).Value = Boolean.FalseString
            row.Cells(Me.ColUser_Code.Index).Value = ""
            row.Cells(Me.ColUser_Name.Index).Value = ""
            row.Cells(Me.ColUser_Email.Index).Value = ""
            row.Cells(Me.Coluser_LastDelivery.Index).Value = ""
            row.Cells(Me.ColUser_Remarks.Index).Value = ""
            row.Cells(Me.ColUser_Approval.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvApproval_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvApproval.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvApproval, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvApproval_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvApproval.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvApproval, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvApproval_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvApproval.UserDeletingRow

    End Sub

    Public Sub HapusList_PrimaryShift(ByVal vPrimaryShift As Double)
        Dim Dta As New DseTrans_SPLTableAdapters.tblmaster_workshiftTableAdapter
        Dim row As DseTrans_SPL.tblmaster_workshiftRow
        Dim HapusShift As Boolean = True

        Try
            'If Me.dgvEmployee.RowCount > 1 Then
            '    For Each rowD As DataGridViewRow In Me.dgvEmployee.Rows
            '        'cari masih ada employee yang masih menggunakan workshift yang dimaksud
            '        If rowD.Cells(Me.colEmpl_PrimaryShift.Index).Value = vPrimaryShift Then
            '            HapusShift = False
            '        End If
            '    Next
            'Else
            '    Me.TblShit.Clear()
            '    Exit Sub
            'End If

            'If HapusShift = True Then
            '    Me.TblShit.DefaultView.RowFilter = "[primary] = " & vPrimaryShift
            '    If Me.TblShit.DefaultView.Count > 0 Then
            '        row = Me.TblShit.DefaultView(0).Row
            '        Me.TblShit.Rows.Remove(row)
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetEmployeeData(ByVal vPrimaryEmp As Double, ByRef vEmpCode As String, ByRef vEmpNIK As String, _
    ByRef vEmpName As String, ByRef vPrimaryPosition As Double, ByRef vPositionCode As String, ByRef vPositionDesc As String, _
    ByRef vPrimaryDept As Double, ByRef vDeptCode As String, ByRef vDeptDesc As String, ByRef vListPrimaryAuth As List(Of Double), _
    Optional ByVal vCekBatasan As Boolean = True) As Boolean
        Dim Dta As New DseTrans_SPLTableAdapters.get_employeeTableAdapter
        Dim Tbl As New DseTrans_SPL.get_employeeDataTable
        Dim row As DseTrans_SPL.get_employeeRow

        'table adapter = tbltrans_spl_item_employee
        Dim Dta1 As New DseTrans_SPLTableAdapters.DataTable1TableAdapter
        Dim Tbl1 As New DseTrans_SPL.DataTable1DataTable
        Dim row1 As DseTrans_SPL.DataTable1Row
        'table adapter = tbltrans_attendance
        Dim Dta2 As New DseTrans_SPLTableAdapters.tbltrans_attendanceTableAdapter
        Dim Tbl2 As New DseTrans_SPL.tbltrans_attendanceDataTable

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta1.Connection.ConnectionString = Core_Module.GetConstring
            Dta2.Connection.ConnectionString = Core_Module.GetConstring

            If Dta.FillBy_PrimaryEmployee(Tbl, vPrimaryEmp) > 0 Then
                row = Tbl.Rows(0)
                With row
                    If .Aktif = 1 Then
                        MsgBox("Employee not active", MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    vPrimaryEmp = .Primary
                    vEmpCode = .Emp_Code
                    vEmpNIK = .Emp_NIK
                    vEmpName = .Emp_Name
                    vPrimaryPosition = .Primary_Position
                    vPositionCode = .Position_Code
                    vPositionDesc = .Position_Desc
                    vPrimaryDept = .Primary_Dept
                    vDeptCode = .Dept_Code
                    vDeptDesc = .Dept_Desc

                    If vCekBatasan = True Then
                        'pastikan tidak double dalam satu period per employee
                        Dta1.FillBy_AndPrimaryEmp(Tbl1, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, ClsTrans_SPL.EnumStatus.Rejected, vPrimaryEmp) 'yang bukan reject)
                        If Me.Obj.ApaSudahAdaPadaTransaksiLain(Me.txtPrimary.Text, vPrimaryEmp, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, Tbl1) = True Then
                            MsgBox("Employee data is created at other transaction.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If

                        'pastikan belum di buat attendance
                        Dta2.FillBy_AndPrimaryEmp(Tbl2, Me.DtpTo.Value.Date, Me.DtpFrom.Value.Date, vPrimaryEmp)
                        If Me.Obj.ApaSudahAdaAttendance(Me.txtPrimary.Text, vPrimaryEmp, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, Tbl2) = True Then
                            MsgBox("Employee data is created at attendance transaction.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If

                        If Me.CekSudahAllocate(Tbl, .Primary, False, True) = False Then
                            Return False
                        End If
                        If Me.CekApaMemilikiApprovalGroup(Tbl, vListPrimaryAuth, True) = False Then
                            vListPrimaryAuth.Clear()
                            Return False
                        End If
                    End If
                End With

                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function CekSudahAllocate(ByRef vTbl As DseTrans_SPL.get_employeeDataTable, _
    ByVal vPrimaryEmp As Double, ByVal vCekFromDgv As Boolean, ByVal vShowMsgBox As Boolean) As Boolean
        Dim Dta As New DseTrans_SPLTableAdapters.allocate_employeeTableAdapter
        Dim Tbl As New DseTrans_SPL.allocate_employeeDataTable
        Dim row As DseTrans_SPL.allocate_employeeRow
        Dim vrow As DseTrans_SPL.get_employeeRow
        Dim rowD As DataGridViewRow
        Dim IdxDgvStop As Integer = -1

        Dim DtpFromFilter As Date = CDate("1/1/2200")
        Dim DtpToFilter As Date = CDate("1/1/1950")
        Dim ApaExitSub As Boolean = False
        Dim ListEmployee As String = ""

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If vPrimaryEmp = -1 Then
                Dta.Fill_All(Tbl, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date)
            Else
                Dta.Fill(Tbl, vPrimaryEmp, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date)
            End If

            If Tbl.Count > 0 Then
                If vCekFromDgv = True Then
                    For i As Integer = 0 To Me.dgvEmployee.RowCount - 2
                        DtpFromFilter = CDate("1/1/2200")
                        DtpToFilter = CDate("1/1/1950")
                        rowD = Me.dgvEmployee.Rows(i)
                        Tbl.DefaultView.RowFilter = "[Primary] = " & rowD.Cells(Me.colEmpl_Primary.Index).Value
                        For Each dv As DataRowView In Tbl.DefaultView
                            row = dv.Row
                            If row.start_date < DtpFromFilter Then
                                DtpFromFilter = row.start_date
                            End If
                            If row.end_date > DtpToFilter Then
                                DtpToFilter = row.end_date
                            End If
                        Next
                        If DtpFromFilter > Me.DtpFrom.Value.Date Or DtpToFilter < Me.DtpTo.Value.Date Then
                            ApaExitSub = True

                            If ListEmployee = "" Then
                                ListEmployee = rowD.Cells(Me.colEmpl_Name.Index).Value & " (" & rowD.Cells(Me.colEmpl_Code.Index).Value & ")"
                            Else
                                ListEmployee &= ", " & rowD.Cells(Me.colEmpl_Name.Index).Value & " (" & rowD.Cells(Me.colEmpl_Code.Index).Value & ")"
                            End If
                        End If
                    Next
                Else
RemoveRowsEmp:
                    For i As Integer = vTbl.Count - 1 To 0 Step -1
                        vrow = vTbl.Rows(i)

                        DtpFromFilter = CDate("1/1/2200")
                        DtpToFilter = CDate("1/1/1950")
                        Tbl.DefaultView.RowFilter = "[Primary] = " & vrow.Primary
                        For Each dv As DataRowView In Tbl.DefaultView
                            row = dv.Row
                            If row.start_date < DtpFromFilter Then
                                DtpFromFilter = row.start_date
                            End If
                            If row.end_date > DtpToFilter Then
                                DtpToFilter = row.end_date
                            End If
                        Next
                        If DtpFromFilter > Me.DtpFrom.Value.Date Or DtpToFilter < Me.DtpTo.Value.Date Then
                            ApaExitSub = True

                            If ListEmployee = "" Then
                                ListEmployee = vrow.Emp_Name & " (" & vrow.Emp_Code & ")"
                            Else
                                ListEmployee &= ", " & vrow.Emp_Name & " (" & vrow.Emp_Code & ")"
                            End If

                            vTbl.Rows.Remove(vrow)
                        End If
                    Next
                End If
            Else
                ApaExitSub = True
                GoTo RemoveRowsEmp
            End If

            If ApaExitSub = True Then
                If vShowMsgBox = True Then
                    MsgBox("Please allocate workshift first for employee" & vbCrLf & ListEmployee, MsgBoxStyle.Information, "Information")
                End If
                Me.TabControl1.SelectedIndex = 0

                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiTbl_EmpShift(ByVal vPrimaryEmp As Double) As Boolean
        Dim Dta As New DseTrans_SPLTableAdapters.allocate_employeeTableAdapter
        Dim Tbl As New DseTrans_SPL.allocate_employeeDataTable
        Dim row As DseTrans_SPL.allocate_employeeRow
        Dim rowEmpShift As DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftRow

        Dim DtpFromFilter As Date = CDate("1/1/2200")
        Dim DtpToFilter As Date = CDate("1/1/1950")
        Dim ApaExitSub As Boolean = False

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(Tbl, vPrimaryEmp, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date)

            Me.HapusTbl_EmpShift(vPrimaryEmp, False)

            'isi tbl employee shift nya
            For Each row In Tbl
                Me.TblEmpShift.DefaultView.RowFilter = "[Date From] = #" & row.start_date & _
                "# AND [Primary Employee] = " & vPrimaryEmp & " AND [Primary Shift] = " & row.Primary_Workshift
                If Me.TblEmpShift.DefaultView.Count > 0 Then
                    For Each dv As DataRowView In Me.TblEmpShift.DefaultView
                        rowEmpShift = dv.Row
                        If row.start_date > Me.DtpFrom.Value.Date Then
                            rowEmpShift.Date_From = row.start_date
                        Else
                            rowEmpShift.Date_From = Me.DtpFrom.Value.Date
                        End If
                        If row.end_date < Me.DtpTo.Value.Date Then
                            rowEmpShift.Date_To = row.end_date
                        Else
                            rowEmpShift.Date_To = Me.DtpTo.Value.Date
                        End If
                    Next
                Else
                    rowEmpShift = Me.TblEmpShift.Newtbltrans_spl_overtime_order_item_employee_shiftRow
                    rowEmpShift.Primary = -1
                    rowEmpShift.Primary_Transaksi = Me.txtPrimary.Text
                    rowEmpShift.Primary_Employee = vPrimaryEmp
                    rowEmpShift.Emp_Code = row.Emp_Code
                    rowEmpShift.Emp_Name = row.Emp_Name
                    rowEmpShift.Primary_Shift = row.Primary_Workshift
                    rowEmpShift.Shift_Code = row.Shift_Code
                    rowEmpShift.Shift_Desc = row.Shift_Desc
                    If row.start_date > Me.DtpFrom.Value.Date Then
                        rowEmpShift.Date_From = row.start_date
                    Else
                        rowEmpShift.Date_From = Me.DtpFrom.Value.Date
                    End If
                    If row.end_date < Me.DtpTo.Value.Date Then
                        rowEmpShift.Date_To = row.end_date
                    Else
                        rowEmpShift.Date_To = Me.DtpTo.Value.Date
                    End If
                    Me.TblEmpShift.Rows.Add(rowEmpShift)
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub IsiTbl_SchSetup() 'ByVal vPrimaryShift As Double)
        Dim Dta As New DseTrans_SPLTableAdapters.tblmaster_workshiftTableAdapter
        Dim Tbl As New DseTrans_SPL.tblmaster_workshiftDataTable
        Dim row As DseTrans_SPL.tblmaster_workshiftRow
        Dim TblSchSet_Dum As New DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupDataTable
        Dim rowSchSet, rowSchSet_Dum As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupRow
        Dim rowEmpShift As DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftRow
        Dim ListPrimaryShift As New List(Of Double)

        Try
            TblSchSet_Dum = Me.TblSchSetup.Copy
            Me.TblSchSetup.Clear()

            Dta.Connection.ConnectionString = Core_Module.GetConstring
            For Each rowEmpShift In Me.TblEmpShift
                TblSchSet_Dum.DefaultView.RowFilter = "[Primary Shift] = " & rowEmpShift.Primary_Shift
                If TblSchSet_Dum.DefaultView.Count > 0 Then
                    rowSchSet_Dum = TblSchSet_Dum.DefaultView(0).Row
                    rowSchSet = Me.TblSchSetup.Newtbltrans_spl_overtime_order_item_sch_setupRow()
                    For i As Integer = 0 To Me.TblSchSetup.Columns.Count - 1
                        rowSchSet.Item(i) = rowSchSet_Dum.Item(i)
                    Next
                    Me.TblSchSetup.Rows.Add(rowSchSet)

                    TblSchSet_Dum.Rows.Remove(rowSchSet_Dum)
                Else
                    Me.TblSchSetup.DefaultView.RowFilter = "[Primary Shift] = " & rowEmpShift.Primary_Shift
                    If Me.TblSchSetup.DefaultView.Count = 0 Then
                        Dta.Fill(Tbl, ClsMasterWorkShift.EnumTipeWorkshift.HR, rowEmpShift.Primary_Shift)
                        If Tbl.Count > 0 Then
                            row = Tbl.Rows(0)
                            rowSchSet = Me.TblSchSetup.Newtbltrans_spl_overtime_order_item_sch_setupRow
                            rowSchSet.Primary_Transksi = Me.txtPrimary.Text
                            rowSchSet.Primary_Shift = row.primary
                            rowSchSet.Shift_Code = row.Shift_Code
                            rowSchSet.Shift_Desc = row.Shift_Desc
                            rowSchSet.Use_OT_Order = row.use_ot_order
                            rowSchSet.Notif_Email_OT_Order = row.notification_email_ot
                            Me.TblSchSetup.Rows.Add(rowSchSet)
                        End If
                    End If
                End If

                If ListPrimaryShift.Contains(rowEmpShift.Primary_Shift) = False Then
                    ListPrimaryShift.Add(rowEmpShift.Primary_Shift)
                End If
            Next

            For Each PrimaryShift As Double In ListPrimaryShift
                Me.IsiTbl_SchDetail(PrimaryShift)
            Next

            If TblSchSet_Dum.Count > 0 Then
                For Each rowSchSet_Dum In TblSchSet_Dum
                    Me.HapusTbl_SchDetail(rowSchSet_Dum.Primary_Shift)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiTbl_SchDetail(ByVal vPrimaryShift As Double)
        Dim Dta As New DseTrans_SPLTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim Tbl As New DseTrans_SPL.tblmaster_workshift_itemDataTable
        Dim row As DseTrans_SPL.tblmaster_workshift_itemRow
        Dim Dta1 As New DseTrans_SPLTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim Tbl1 As New DseTrans_SPL.tblmaster_workshift_itemholidayDataTable
        Dim row1 As DseTrans_SPL.tblmaster_workshift_itemholidayRow
        Dim rowEmpShift As DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftRow
        Dim rowSchSet As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupRow
        Dim TblSchDet_Dum As New DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailDataTable
        Dim rowSchDetail, rowSchDetail_Dum As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailRow
        Dim DateFilter As Date = Now
        Dim DayOfWeek As Integer = 0

        Try
            Me.TblSchDet.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift
            TblSchDet_Dum.Merge(Me.TblSchDet.DefaultView.ToTable)
            While Me.TblSchDet.DefaultView.Count > 0
                rowSchDetail = Me.TblSchDet.DefaultView(0).Row
                Me.TblSchDet.Rows.Remove(rowSchDetail)
            End While

            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta1.Connection.ConnectionString = Core_Module.GetConstring

            Dta.Fill(Tbl, vPrimaryShift, ClsMasterWorkShift.EnumTipeWorkshift.HR)
            Dta1.Fill(Tbl1, vPrimaryShift, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date)

            '1. cari ke tbl sch setup berdasarkan primary work shift
            Me.TblSchSetup.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift
            If TblSchSetup.DefaultView.Count > 0 Then
                rowSchSet = Me.TblSchSetup.DefaultView(0).Row
                '2 cari duration ot nya berdasarkan primary work shift ke tbl employee shift
                Me.TblEmpShift.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift
                Me.TblEmpShift.DefaultView.Sort = "[Date From] ASC"
                For Each dv As DataRowView In Me.TblEmpShift.DefaultView
                    rowEmpShift = dv.Row
                    DateFilter = rowEmpShift.Date_From
                    While DateFilter <= rowEmpShift.Date_To
                        If DateFilter.DayOfWeek = System.DayOfWeek.Sunday Then
                            DayOfWeek = 0
                        Else
                            DayOfWeek = DateFilter.DayOfWeek
                        End If

                        'untuk early ot ##############################################################################
                        TblSchDet.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift & " AND [Date] = #" & _
                        DateFilter & "# AND [OT Type] = " & ClsTrans_SPL.EnumTipeOT.Early_Overtime
                        If TblSchDet.DefaultView.Count = 0 Then
                            If rowSchSet.Is_Early_OT = 1 And rowSchSet.Early_Duration > 0 Then
                                rowSchDetail = Me.TblSchDet.Newtbltrans_spl_overtime_order_item_sch_detailRow
                                rowSchDetail.Primary_Transaksi = Me.txtPrimary.Text
                                rowSchDetail._Date = DateFilter
                                rowSchDetail.Primary_Shift = vPrimaryShift
                                rowSchDetail.Shift_Code = rowSchSet.Shift_Code
                                rowSchDetail.Shift_Desc = rowSchSet.Shift_Desc

                                'jika early hanya berlaku di weekday saja 
                                Tbl.DefaultView.RowFilter = "[day of week] = " & DayOfWeek
                                If Tbl.DefaultView.Count > 0 Then
                                    row = Tbl.DefaultView(0).Row
                                    If row.tipe_workdays = ClsMasterWorkShift.TipeWorkdays.Weekday And row.tipe_overtime <> ClsMasterWorkShift.TipeOvertime.No_Overtime Then
                                        'pastikan tidak termasuk ke holiday list 
                                        Tbl1.DefaultView.RowFilter = "[date] = #" & DateFilter & "#"
                                        If Tbl1.DefaultView.Count = 0 Then
                                            rowSchDetail.Start_Time = row.schedule_in
                                            rowSchDetail.End_Time = row.schedule_out
                                            rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Early_Overtime

                                            'EARLY OVERTIME - hitung start, break, end, duration
                                            rowSchDetail._End = row.schedule_in - rowSchSet.Early_Delay
                                            If rowSchSet.Early_Break <> 0 And rowSchSet.Early_Delay_Before_Break <> 0 Then
                                                rowSchDetail.Start = rowSchDetail._End - rowSchSet.Early_Duration - rowSchSet.Early_Break
                                                rowSchDetail.Break_Out = rowSchDetail.Start + rowSchSet.Early_Delay_Before_Break
                                                rowSchDetail.Break_In = rowSchDetail.Break_Out + rowSchSet.Early_Break
                                            Else
                                                rowSchDetail.Start = rowSchDetail._End - rowSchSet.Early_Duration
                                            End If

                                            'jika fungsi override digunakan maka tidak diubah, melainkan gunakan hitungan sebelumnya
                                            TblSchDet_Dum.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift & " AND [Date] = #" & _
                                            DateFilter & "# AND [OT Type] = " & ClsTrans_SPL.EnumTipeOT.Early_Overtime
                                            If TblSchDet_Dum.DefaultView.Count > 0 Then
                                                rowSchDetail_Dum = TblSchDet_Dum.DefaultView(0).Row
                                                If rowSchDetail_Dum.Override = 1 Then
                                                    rowSchDetail.Start = rowSchDetail_Dum.Start
                                                    rowSchDetail.Break_Out = rowSchDetail_Dum.Break_Out
                                                    rowSchDetail.Break_In = rowSchDetail_Dum.Break_In
                                                    rowSchDetail._End = rowSchDetail_Dum._End
                                                End If
                                                rowSchDetail.Override = rowSchDetail_Dum.Override
                                                rowSchDetail.Remarks = rowSchDetail_Dum.Remarks
                                            End If
                                            rowSchDetail.Duration = rowSchDetail._End - rowSchDetail.Start - (rowSchDetail.Break_In - rowSchDetail.Break_Out)
                                            '#FUNGSI ADD ROWS KE SCH DETAIL (EARLY OT)
                                            Me.TblSchDet.Rows.Add(rowSchDetail)
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        'untuk ot ##############################################################################
                        If (rowSchSet.Is_OT = 1 And rowSchSet.OT_Duration > 0) Or (rowSchSet.Is_Holiday_OT = 1 And rowSchSet.Holiday_OT_Duration > 0) Then
                            rowSchDetail = Me.TblSchDet.Newtbltrans_spl_overtime_order_item_sch_detailRow
                            rowSchDetail.Primary_Transaksi = Me.txtPrimary.Text
                            rowSchDetail._Date = DateFilter
                            rowSchDetail.Primary_Shift = vPrimaryShift
                            rowSchDetail.Shift_Code = rowSchSet.Shift_Code
                            rowSchDetail.Shift_Desc = rowSchSet.Shift_Desc

                            'cek apakah hari libur
                            Tbl1.DefaultView.RowFilter = "[date] = #" & DateFilter & "#"
                            If Tbl1.DefaultView.Count > 0 Then
                                'cek dulu apakah memperhitungkan ot holiday
                                If rowSchSet.Is_Holiday_OT = 0 And rowSchSet.Holiday_OT_Duration = 0 Then
                                    GoTo skipAddRows
                                End If

                                row1 = Tbl1.DefaultView(0).Row
                                rowSchDetail.Start_Time = row1.sch_in
                                rowSchDetail.End_Time = row1.sch_out
                                If row1.tipe_overtime = ClsMasterWorkShift.TipeOvertime.Weekday Then
                                    rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Overtime
                                Else
                                    rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Holiday_Overtime
                                End If

                                'HOLIDAY OVERTIME (HOLIDAY LIST)- hitung start, break, end, duration
                                rowSchDetail.Start = row1.sch_in + rowSchSet.Holiday_OT_Delay
                                If rowSchDetail.Start >= (24 * 60) Then
                                    rowSchDetail.Start = row1.sch_in + rowSchSet.Holiday_OT_Delay - (24 * 60)
                                End If
                                If rowSchSet.Holiday_OT_Break <> 0 And rowSchSet.Holiday_Delay_Before_Break <> 0 Then
                                    rowSchDetail.Break_Out = rowSchDetail.Start + rowSchSet.Holiday_Delay_Before_Break
                                    If rowSchDetail.Break_Out >= (24 * 60) Then
                                        rowSchDetail.Break_Out = rowSchDetail.Start + rowSchSet.Holiday_Delay_Before_Break - (24 * 60)
                                    End If
                                    rowSchDetail.Break_In = rowSchDetail.Break_Out + rowSchSet.Holiday_OT_Break
                                    If rowSchDetail.Break_In >= (24 * 60) Then
                                        rowSchDetail.Break_In = rowSchDetail.Break_Out + rowSchSet.Holiday_OT_Break - (24 * 60)
                                    End If
                                    rowSchDetail._End = rowSchDetail.Break_In + (rowSchSet.Holiday_OT_Duration - rowSchSet.Holiday_Delay_Before_Break)
                                    If rowSchDetail._End >= (24 * 60) Then
                                        rowSchDetail._End = rowSchDetail.Break_In + (rowSchSet.Holiday_OT_Duration - rowSchSet.Holiday_Delay_Before_Break) - (24 * 60)
                                    End If
                                Else
                                    rowSchDetail._End = rowSchDetail.Start + rowSchSet.Holiday_OT_Duration
                                    If rowSchDetail._End >= (24 * 60) Then
                                        rowSchDetail._End = rowSchDetail.Start + rowSchSet.Holiday_OT_Duration - (24 * 60)
                                    End If
                                End If
                            Else
                                Tbl.DefaultView.RowFilter = "[day of week] = " & DayOfWeek
                                If Tbl.DefaultView.Count > 0 Then
                                    row = Tbl.DefaultView(0).Row
                                    rowSchDetail.Start_Time = row.schedule_in
                                    rowSchDetail.End_Time = row.schedule_out

                                    If row.tipe_workdays = ClsMasterWorkShift.TipeWorkdays.Weekday Then
                                        If rowSchSet.Is_OT <> 1 And rowSchSet.OT_Duration <= 0 Then
                                            GoTo skipAddRows
                                        End If

                                        rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Overtime

                                        'OVERTIME - hitung start, break, end, duration
                                        rowSchDetail.Start = row.schedule_out + rowSchSet.OT_Delay
                                        If rowSchDetail.Start >= (24 * 60) Then
                                            rowSchDetail.Start = row.schedule_out + rowSchSet.OT_Delay - (24 * 60)
                                        End If
                                        If rowSchSet.OT_Break <> 0 And rowSchSet.OT_Delay_Before_Break <> 0 Then
                                            rowSchDetail.Break_Out = rowSchDetail.Start + rowSchSet.OT_Delay_Before_Break
                                            If rowSchDetail.Break_Out >= (24 * 60) Then
                                                rowSchDetail.Break_Out = rowSchDetail.Start + rowSchSet.OT_Delay_Before_Break - (24 * 60)
                                            End If
                                            rowSchDetail.Break_In = rowSchDetail.Break_Out + rowSchSet.OT_Break
                                            If rowSchDetail.Break_In >= (24 * 60) Then
                                                rowSchDetail.Break_In = rowSchDetail.Break_Out + rowSchSet.OT_Break - (24 * 60)
                                            End If
                                            rowSchDetail._End = rowSchDetail.Break_In + (rowSchSet.OT_Duration - rowSchSet.OT_Delay_Before_Break)
                                            If rowSchDetail._End >= (24 * 60) Then
                                                rowSchDetail._End = rowSchDetail.Break_In + (rowSchSet.OT_Duration - rowSchSet.OT_Delay_Before_Break) - (24 * 60)
                                            End If
                                        Else
                                            rowSchDetail._End = rowSchDetail.Start + rowSchSet.OT_Duration
                                            If rowSchDetail._End >= (24 * 60) Then
                                                rowSchDetail._End = rowSchDetail.Start + rowSchSet.OT_Duration - (24 * 60)
                                            End If
                                        End If
                                    ElseIf row.tipe_workdays = ClsMasterWorkShift.TipeWorkdays.Holiday Then
                                        If rowSchSet.Is_Holiday_OT <> 1 And rowSchSet.Holiday_OT_Duration <= 0 Then
                                            GoTo skipAddRows
                                        End If

                                        rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Holiday_Overtime

                                        'HOLIDAY OVERTIME - hitung start, break, end, duration
                                        rowSchDetail.Start = row.schedule_in + rowSchSet.Holiday_OT_Delay
                                        If rowSchDetail.Start >= (24 * 60) Then
                                            rowSchDetail.Start = row.schedule_in + rowSchSet.Holiday_OT_Delay - (24 * 60)
                                        End If
                                        If rowSchSet.Holiday_OT_Break <> 0 And rowSchSet.Holiday_Delay_Before_Break <> 0 Then
                                            rowSchDetail.Break_Out = rowSchDetail.Start + rowSchSet.Holiday_Delay_Before_Break
                                            If rowSchDetail.Break_Out >= (24 * 60) Then
                                                rowSchDetail.Break_Out = rowSchDetail.Start + rowSchSet.Holiday_Delay_Before_Break - (24 * 60)
                                            End If
                                            rowSchDetail.Break_In = rowSchDetail.Break_Out + rowSchSet.Holiday_OT_Break
                                            If rowSchDetail.Break_In >= (24 * 60) Then
                                                rowSchDetail.Break_In = rowSchDetail.Break_Out + rowSchSet.Holiday_OT_Break - (24 * 60)
                                            End If
                                            rowSchDetail._End = rowSchDetail.Break_In + (rowSchSet.Holiday_OT_Duration - rowSchSet.Holiday_Delay_Before_Break)
                                            If rowSchDetail._End >= (24 * 60) Then
                                                rowSchDetail._End = rowSchDetail.Break_In + (rowSchSet.Holiday_OT_Duration - rowSchSet.Holiday_Delay_Before_Break) - (24 * 60)
                                            End If
                                        Else
                                            rowSchDetail._End = rowSchDetail.Start + rowSchSet.Holiday_OT_Duration
                                            If rowSchDetail._End >= (24 * 60) Then
                                                rowSchDetail._End = rowSchDetail.Start + rowSchSet.Holiday_OT_Duration - (24 * 60)
                                            End If
                                        End If
                                    Else
                                        GoTo skipAddRows
                                    End If
                                End If
                            End If

                            TblSchDet.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift & " AND [Date] = #" & _
                            DateFilter & "# AND [OT Type] = " & rowSchDetail.OT_Type
                            If TblSchDet.DefaultView.Count = 0 Then
                                TblSchDet_Dum.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift & " AND [Date] = #" & _
                                                            DateFilter & "# AND [OT Type] = " & rowSchDetail.OT_Type
                                If TblSchDet_Dum.DefaultView.Count > 0 Then
                                    rowSchDetail_Dum = TblSchDet_Dum.DefaultView(0).Row
                                    If rowSchDetail_Dum.Override = 1 Then
                                        rowSchDetail.Start = rowSchDetail_Dum.Start
                                        rowSchDetail.Break_Out = rowSchDetail_Dum.Break_Out
                                        rowSchDetail.Break_In = rowSchDetail_Dum.Break_In
                                        rowSchDetail._End = rowSchDetail_Dum._End
                                    End If
                                    rowSchDetail.Override = rowSchDetail_Dum.Override
                                    rowSchDetail.Remarks = rowSchDetail_Dum.Remarks
                                End If

                                Dim HitungStart As Double = rowSchDetail.Start
                                Dim HitungBreakOut As Double = rowSchDetail.Break_Out
                                Dim HitungBreakIn As Double = rowSchDetail.Break_In
                                Dim HitungEnd As Double = rowSchDetail._End
                                If HitungBreakOut = 0 And HitungBreakIn = 0 Then
                                    If HitungEnd < HitungStart Then
                                        HitungEnd = HitungEnd + (24 * 60)
                                    End If
                                Else
                                    If HitungBreakOut < HitungStart Then
                                        HitungBreakOut = HitungBreakOut + (24 * 60)
                                    End If
                                    If HitungBreakIn < HitungBreakOut Then
                                        HitungBreakIn = HitungBreakIn + (24 * 60)
                                    End If
                                    If HitungEnd < HitungBreakIn Then
                                        HitungEnd = HitungEnd + (24 * 60)
                                    End If
                                End If
                                rowSchDetail.Duration = HitungEnd - HitungStart - (HitungBreakIn - HitungBreakOut)
                                '#FUNGSI ADD ROWS KE SCH DETAIL (OT & HOLIDAY OT)
                                Me.TblSchDet.Rows.Add(rowSchDetail)
                            End If
                        End If
skipAddRows:
                        DateFilter = DateFilter.AddDays(1)
                    End While
                Next
            End If

            Me.IsiTbl_SchEmp(vPrimaryShift)
            'Me.UpdateTotalDuration()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiTbl_SchEmp(ByVal vPrimaryShift As Double)
        Dim Dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approved_otTableAdapter
        Dim TblDum As New DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otDataTable
        Dim row, rowDum As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow
        Dim TblOT_2 As New DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_2DataTable
        Dim rowEmpShift As DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftRow
        Dim rowSchDetail As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailRow
        Dim DateFilter As Date = Now
        Dim StartWH As Double = 0
        Dim EndWH As Double = 0
        Dim StartTimeEarlyOT As Double = 0
        Dim BreakOutTimeEarlyOT As Double = 0
        Dim BreakInTimeEarlyOT As Double = 0
        Dim EndTimeEarlyOT As Double = 0
        Dim HitungDurationEarlyOT As Double = 0
        Dim StartTimeOT As Double = 0
        Dim BreakOutTimeOT As Double = 0
        Dim BreakInTimeOT As Double = 0
        Dim EndTimeOT As Double = 0
        Dim HitungDurationOT As Double = 0
        Dim ApaAddEarlyOT As Boolean = False
        Dim ApaAddOT As Boolean = False
        Dim ApaHariLibur As Boolean = False

        Try
            Me.TblAppOT.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift
            TblDum.Merge(Me.TblAppOT.DefaultView.ToTable)
            While Me.TblAppOT.DefaultView.Count > 0
                row = Me.TblAppOT.DefaultView(0).Row
                Me.TblAppOT.Rows.Remove(row)
            End While

            Me.IsiTbl_SchEmp_Filter(Me.txtPrimary.Text, vPrimaryShift, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, TblOT_2)

            Dta.Connection.ConnectionString = Core_Module.GetConstring
            'Dta.Fill(TblDum, vPrimaryTransaksi)
            Me.TblEmpShift.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift
            Me.TblEmpShift.DefaultView.Sort = "[Primary Employee] ASC, [Date From] ASC"
            For ii As Integer = 0 To Me.TblEmpShift.DefaultView.Count - 1
                rowEmpShift = Me.TblEmpShift.DefaultView(ii).Row
                DateFilter = rowEmpShift.Date_From
                While DateFilter <= rowEmpShift.Date_To
                    Me.TblSchDet.DefaultView.RowFilter = "[Date] = #" & DateFilter & "# AND [Primary Shift] = " & rowEmpShift.Primary_Shift '& " AND [] = " & rowEmpShift.Primary_Employee
                    Me.TblSchDet.DefaultView.Sort = "[OT Type] ASC"
                    If Me.TblSchDet.DefaultView.Count > 0 Then
                        StartWH = 0
                        EndWH = 0
                        StartTimeEarlyOT = 0
                        BreakOutTimeEarlyOT = 0
                        BreakInTimeEarlyOT = 0
                        EndTimeEarlyOT = 0
                        HitungDurationEarlyOT = 0
                        StartTimeOT = 0
                        BreakOutTimeOT = 0
                        BreakInTimeOT = 0
                        EndTimeOT = 0
                        HitungDurationOT = 0

                        ApaAddEarlyOT = False
                        ApaAddOT = False

                        For Each dv As DataRowView In Me.TblSchDet.DefaultView
                            rowSchDetail = dv.Row
                            ApaHariLibur = False

                            'cek kembali apa employee pada tanggal, shift ybs sudah ada pada transaksi lain, jika sudah ada jangan input.
                            If Me.IsiTbl_SchEmp_ApaSudahAda(rowEmpShift.Primary_Employee, DateFilter, rowSchDetail.OT_Type, TblOT_2) = False Then
                                StartWH = rowSchDetail.Start_Time
                                EndWH = rowSchDetail.End_Time
                                If rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Early_Overtime Then
                                    StartTimeEarlyOT = rowSchDetail.Start
                                    BreakOutTimeEarlyOT = rowSchDetail.Break_Out
                                    BreakInTimeEarlyOT = rowSchDetail.Break_In
                                    EndTimeEarlyOT = rowSchDetail._End
                                    HitungDurationEarlyOT += rowSchDetail.Duration
                                    ApaAddEarlyOT = True
                                Else
                                    StartTimeOT = rowSchDetail.Start
                                    BreakOutTimeOT = rowSchDetail.Break_Out
                                    BreakInTimeOT = rowSchDetail.Break_In
                                    EndTimeOT = rowSchDetail._End
                                    HitungDurationOT += rowSchDetail.Duration
                                    ApaAddOT = True

                                    If rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Holiday_Overtime Then
                                        ApaHariLibur = True
                                    End If
                                End If
                            End If
                        Next

                        If ApaAddEarlyOT = True Or ApaAddOT = True Then
                            row = Me.TblAppOT.Newtbltrans_spl_overtime_order_item_approved_otRow
                            row.Primary_Transaksi = Me.txtPrimary.Text
                            row.Primary_Employee = rowEmpShift.Primary_Employee
                            row.Emp_Code = rowEmpShift.Emp_Code
                            row.Emp_Name = rowEmpShift.Emp_Name
                            row.Primary_Shift = rowEmpShift.Primary_Shift
                            row.Shift_Code = rowEmpShift.Shift_Code
                            row.Shift_Desc = rowEmpShift.Shift_Desc
                            row._Date = DateFilter
                            row.Start_WH = StartWH
                            row.End_WH = EndWH

                            If ApaAddEarlyOT = True Then
                                row.Start_Time_Early = StartTimeEarlyOT
                                row.Break_Out_Time_Early = BreakOutTimeEarlyOT
                                row.Break_In_Time_Early = BreakInTimeEarlyOT
                                row.End_Time_Early = EndTimeEarlyOT
                                row.OT_Early_Approved = HitungDurationEarlyOT
                            End If

                            If ApaAddOT = True Then
                                row.Start_Time = StartTimeOT
                                row.Break_Out_Time = BreakOutTimeOT
                                row.Break_In_Time = BreakInTimeOT
                                row.End_Time = EndTimeOT
                                row.OT_Aprroved = HitungDurationOT
                            End If

                            If ApaHariLibur = True Then
                                row.is_holiday = 1
                            Else
                                row.is_holiday = 0
                            End If

                            'jika fungsi override digunakan maka tidak diubah, melainkan gunakan hitungan sebelumnya
                            TblDum.DefaultView.RowFilter = "[Primary Employee] = " & row.Primary_Employee.ToString & " AND [Date] = #" & row._Date.Date & "#"
                            If TblDum.DefaultView.Count > 0 Then
                                rowDum = TblDum.DefaultView(0).Row
                                If rowDum.Override_Early_OT = 1 Then
                                    row.Start_Time_Early = rowDum.Start_Time_Early
                                    row.Break_Out_Time_Early = rowDum.Break_Out_Time_Early
                                    row.Break_In_Time_Early = rowDum.Break_In_Time_Early
                                    row.End_Time_Early = rowDum.End_Time_Early
                                    row.Override_Early_OT = rowDum.Override_Early_OT
                                    row.OT_Early_Approved = rowDum.OT_Early_Approved
                                End If
                                If rowDum.Override_OT = 1 Then
                                    row.Start_Time = rowDum.Start_Time
                                    row.Break_Out_Time = rowDum.Break_Out_Time
                                    row.Break_In_Time = rowDum.Break_In_Time
                                    row.End_Time = rowDum.End_Time
                                    row.Override_OT = rowDum.Override_OT
                                    row.OT_Aprroved = rowDum.OT_Aprroved
                                End If
                            End If

                            Me.TblAppOT.Rows.Add(row)
                        End If
                    End If

                    DateFilter = DateFilter.AddDays(1)
                End While
            Next

            Me.UpdateTotalDuration()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiTbl_SchEmp_Filter(ByVal vPrimarySPL As Double, ByVal vPrimaryShift As Double, ByVal vDateFrom As Date, ByVal vDateTo As Date, _
    ByRef Tbl As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_2DataTable)
        Dim Dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approved_ot_2TableAdapter
        Dim TblDum As New DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_2DataTable
        Dim rowEmpShift As DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftRow
        Dim ListPrimaryEmp As New List(Of Double)
        Dim Filter As String = ""

        Try
            Tbl.Clear()

            For Each rowEmpShift In Me.TblEmpShift
                If ListPrimaryEmp.Contains(rowEmpShift.Primary_Employee) = False Then
                    ListPrimaryEmp.Add(rowEmpShift.Primary_Employee)
                    If Filter = "" Then
                        Filter = "[Primary Employee] = " & rowEmpShift.Primary_Employee
                    Else
                        Filter &= " OR [Primary Employee] = " & rowEmpShift.Primary_Employee
                    End If
                End If
            Next

            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.FillBy_DateFromTo(TblDum, vPrimarySPL, vDateFrom.Date, vDateTo.Date, vPrimaryShift, ClsTrans_SPL.EnumStatus.Rejected)
            TblDum.DefaultView.RowFilter = Filter
            Tbl.Merge(TblDum.DefaultView.ToTable)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function IsiTbl_SchEmp_ApaSudahAda(ByVal vPrimaryEmp As Double, ByVal vDate As Date, ByVal vTipeOT As ClsTrans_SPL.EnumTipeOT, _
    ByVal Tbl As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_2DataTable) As Boolean
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_2Row
        Dim Retval As Boolean = False

        Try
            Tbl.DefaultView.RowFilter = "[Primary Employee] = " & vPrimaryEmp.ToString & " AND [Date] = #" & vDate.Date & "#"
            If Tbl.DefaultView.Count > 0 Then
                row = Tbl.DefaultView(0).Row
                Return True
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Retval
        End Try
    End Function

    Public Function CekApaMemilikiApprovalGroup(ByRef vTbl As DseTrans_SPL.get_employeeDataTable, _
    ByRef vListPrimaryAuth As List(Of Double), ByVal vShowMsgBox As Boolean) As Boolean
        Dim Dta As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_groupTableAdapter
        Dim Tbl As New DseTrans_SPL_Setting.tbltrans_spl_approval_groupDataTable
        Dim row As DseTrans_SPL_Setting.tbltrans_spl_approval_groupRow
        Dim Dta1 As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_item_scopeTableAdapter
        Dim Tbl1 As New DseTrans_SPL_Setting.tbltrans_spl_approval_group_item_scopeDataTable
        Dim row1 As DseTrans_SPL_Setting.tbltrans_spl_approval_group_item_scopeRow
        Dim rowEmp As DseTrans_SPL.get_employeeRow
        Dim Ketemu As Boolean = True
        Dim Retval As Boolean = True
        Dim ListEmployee As String = ""
        Dim ApakahAllEmployee As Boolean = False
        Dim ListPrimaryEmp As New List(Of Double)

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta1.Connection.ConnectionString = Core_Module.GetConstring
            Dta.FillBy_allEmp(Tbl, 1)
            Dta1.Fill_all(Tbl1)

            If Tbl.Count > 0 Then
                'jika user memiliki akses penuh all employee
                For Each row In Tbl
                    If vListPrimaryAuth.Contains(row.Primary) = False Then
                        vListPrimaryAuth.Add(row.Primary)
                    End If
                Next
                ApakahAllEmployee = True
            End If

            Dim Filter As String = ""
            For Each row1 In Tbl1
                With row1
                    If .Primary_Department <> -1 And .Primary_Position <> -1 Then
                        Filter = "[Primary Dept] = " & .Primary_Department & " AND [Primary Position] = " & .Primary_Position
                    ElseIf .Primary_Department <> -1 And .Primary_Position = -1 Then
                        Filter = "[Primary Dept] = " & .Primary_Department
                    ElseIf .Primary_Department = -1 And .Primary_Position <> -1 Then
                        Filter = "[Primary Position] = " & .Primary_Position
                    ElseIf .Primary_Employee <> -1 And .Primary_Department = -1 And .Primary_Position = -1 Then
                        Filter = "[Primary] = " & .Primary_Employee
                    End If

                    vTbl.DefaultView.RowFilter = Filter
                    If vTbl.DefaultView.Count > 0 Then
                        If vListPrimaryAuth.Contains(.Primary_Transaksi) = False Then
                            vListPrimaryAuth.Add(row1.Primary_Transaksi)
                        End If

                        For Each dv As DataRowView In vTbl.DefaultView
                            rowEmp = dv.Row
                            If ListPrimaryEmp.Contains(rowEmp.Primary) = False Then
                                ListPrimaryEmp.Add(rowEmp.Primary)
                            End If
                        Next
                    End If
                End With
            Next

            If ApakahAllEmployee = False Then
                vTbl.DefaultView.RowFilter = ""
                vTbl.DefaultView.Sort = "[Emp_Code] DESC"
                For i As Integer = vTbl.DefaultView.Count - 1 To 0 Step -1
                    rowEmp = vTbl.DefaultView(i).Row
                    If ListPrimaryEmp.Contains(rowEmp.Primary) = False Then
                        If ListEmployee = "" Then
                            ListEmployee = rowEmp.Emp_Name & " (" & rowEmp.Emp_Code & ")"
                        Else
                            ListEmployee &= ", " & rowEmp.Emp_Name & " (" & rowEmp.Emp_Code & ")"
                        End If

                        vTbl.Rows.Remove(rowEmp)
                    End If
                Next
            End If

            If ListEmployee <> "" Then
                If vShowMsgBox = True Then
                    MsgBox("Please create approval group for employee" & vbCrLf & ListEmployee, MsgBoxStyle.Information, "Information")
                End If

                Retval = False
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub IsiTbl_Approval(ByVal vListPrimaryAuth As List(Of Double))
        Dim Dta2 As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_item_userTableAdapter
        Dim Tbl2 As New DseTrans_SPL_Setting.tbltrans_spl_approval_group_item_userDataTable
        Dim row2 As DseTrans_SPL_Setting.tbltrans_spl_approval_group_item_userRow
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approvalRow

        Try
            Dta2.Connection.ConnectionString = Core_Module.GetConstring
            Dta2.Fill_all(Tbl2)

            For Each PrimaryAuth As Double In vListPrimaryAuth
                Tbl2.DefaultView.RowFilter = "[Primary Transaksi] = " & PrimaryAuth
                If Tbl2.DefaultView.Count > 0 Then
                    For Each dv As DataRowView In Tbl2.DefaultView
                        row2 = dv.Row
                        Me.TblApp.DefaultView.RowFilter = "[Primary User] = " & row2.Primary_User
                        If Me.TblApp.DefaultView.Count = 0 Then
                            row = Me.TblApp.Newtbltrans_spl_overtime_order_item_approvalRow
                            row.Primary_Transaksi = Me.txtPrimary.Text
                            row.Primary_User = row2.Primary_User
                            row.Email_Address = row2.Email_Address
                            row.Kode_User = row2.Kode_User
                            row.Nama_User = row2.Nama_User
                            Me.TblApp.Rows.Add(row)
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Public Sub IsiTbl_ApprovalOT(ByVal vPrimaryShift As Double)
    '    Dim Dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approved_otTableAdapter
    '    Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow

    '    Try
    '        Dta.Connection.ConnectionString = Core_Module.GetConstring

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Public Sub UpdTbl_SchSetup(ByVal vrowD As DataGridViewRow)
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupRow
        Dim PrimaryShift As Double = CDbl(vrowD.Cells(Me.ColSetup_PrimaryWorkshift.Index).Value)

        Try
            Me.TblSchSetup.DefaultView.RowFilter = "[Primary Shift] = " & PrimaryShift
            If Me.TblSchSetup.DefaultView.Count > 0 Then
                row = Me.TblSchSetup.DefaultView(0).Row
                With row
                    If vrowD.Cells(Me.ColSetup_IsEarlyOT.Index).Value = False Then
                        .Is_Early_OT = 0
                    Else
                        .Is_Early_OT = 1
                    End If
                    .Early_Duration = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSetup_EarlyOT.Index).Value)
                    .Early_Delay = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSetup_EarlyDelay.Index).Value)
                    .Early_Break = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSetup_EarlyBreak.Index).Value)
                    .Early_Delay_Before_Break = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSetup_EarlyDelayBeforeBreak.Index).Value)
                    If vrowD.Cells(Me.ColSetup_IsOT.Index).Value = False Then
                        .Is_OT = 0
                    Else
                        .Is_OT = 1
                    End If
                    .OT_Duration = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSetup_OT.Index).Value)
                    .OT_Delay = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSetup_OTDelay.Index).Value)
                    .OT_Break = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSetup_OTBreak.Index).Value)
                    .OT_Delay_Before_Break = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSetup_OTDelayBeforeBreak.Index).Value)
                    If vrowD.Cells(Me.ColSetup_IsHolOT.Index).Value = False Then
                        .Is_Holiday_OT = 0
                    Else
                        .Is_Holiday_OT = 1
                    End If
                    .Holiday_OT_Duration = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSetup_HolOT.Index).Value)
                    .Holiday_OT_Delay = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSetup_HolOTDelay.Index).Value)
                    .Holiday_OT_Break = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSetup_HolOTBreak.Index).Value)
                    .Holiday_Delay_Before_Break = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSetup_HolOTDelayBeforeBreak.Index).Value)
                    .Remarks = vrowD.Cells(Me.ColSetup_Remarks.Index).Value

                    'sudah di isi saat input employee
                    '.Primary_Transksi = 0
                    '.Shift_Code = 0
                    '.Primary_Shift = 0
                    '.Shift_Desc = 0
                    '.Use_OT_Order = 0
                    '.Notif_Email_OT_Order = 0

                    'karena ada perubahan di schedule setup maka harus update juga di tabel schedule detail
                    Me.IsiTbl_SchDetail(PrimaryShift)
                End With
            Else
                MsgBox("Row schedule setup not found", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdDgv_SchDetailUnOverride(ByVal vrowD As DataGridViewRow)
        Dim rowSchDetail As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailRow
        Dim rowSchSet As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupRow
        Dim PrimaryShift As Double = vrowD.Cells(Me.ColSch_PrimaryShift.Index).Value
        Dim DateFilter As Date = CDate(vrowD.Cells(Me.ColSch_Date.Index).Value)
        Dim OTType As ClsTrans_SPL.EnumTipeOT = Me.GetTypeOT_Enum(vrowD.Cells(Me.ColSch_OTtype.Index).Value)

        Try
            Me.TblSchDet.DefaultView.RowFilter = "[Primary Shift] = " & PrimaryShift & " AND [Date] = #" & DateFilter & "# AND [OT Type] = " & OTType
            If Me.TblSchDet.DefaultView.Count = 0 Then
                MsgBox("Row schedule detail not found", MsgBoxStyle.Information, "Information")
            Else
                rowSchDetail = Me.TblSchDet.DefaultView(0).Row
                Me.TblSchSetup.DefaultView.RowFilter = "[Primary Shift] = " & PrimaryShift
                If TblSchSetup.DefaultView.Count = 0 Then
                    MsgBox("Row schedule detail not found", MsgBoxStyle.Information, "Information")
                Else
                    rowSchSet = Me.TblSchSetup.DefaultView(0).Row
                    'early overtime ##############################################################################
                    If rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Early_Overtime Then
                        rowSchDetail._End = rowSchDetail.Start_Time - rowSchSet.Early_Delay
                        If rowSchSet.Early_Break <> 0 And rowSchSet.Early_Delay_Before_Break <> 0 Then
                            rowSchDetail.Start = rowSchDetail._End - rowSchSet.Early_Duration - rowSchSet.Early_Break
                            rowSchDetail.Break_Out = rowSchDetail.Start + rowSchSet.Early_Delay_Before_Break
                            rowSchDetail.Break_In = rowSchDetail.Break_Out + rowSchSet.Early_Break
                        Else
                            rowSchDetail.Start = rowSchDetail._End - rowSchSet.Early_Duration
                            rowSchDetail.Break_Out = 0
                            rowSchDetail.Break_In = 0
                        End If
                        rowSchDetail.Duration = rowSchSet.Early_Duration  'rowSchDetail._End - rowSchDetail.Start - (rowSchDetail.Break_In - rowSchDetail.Break_Out)

                        'untuk ot ##############################################################################
                    ElseIf rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Overtime Then
                        'OVERTIME - hitung start, break, end, duration
                        rowSchDetail.Start = rowSchDetail.End_Time + rowSchSet.OT_Delay
                        If rowSchDetail.Start >= (24 * 60) Then
                            rowSchDetail.Start = rowSchDetail.End_Time + rowSchSet.OT_Delay - (24 * 60)
                        End If
                        If rowSchSet.OT_Break <> 0 And rowSchSet.OT_Delay_Before_Break <> 0 Then
                            rowSchDetail.Break_Out = rowSchDetail.Start + rowSchSet.OT_Delay_Before_Break
                            If rowSchDetail.Break_Out >= (24 * 60) Then
                                rowSchDetail.Break_Out = rowSchDetail.Start + rowSchSet.OT_Delay_Before_Break - (24 * 60)
                            End If
                            rowSchDetail.Break_In = rowSchDetail.Break_Out + rowSchSet.OT_Break
                            If rowSchDetail.Break_In >= (24 * 60) Then
                                rowSchDetail.Break_In = rowSchDetail.Break_Out + rowSchSet.OT_Break - (24 * 60)
                            End If
                            rowSchDetail._End = rowSchDetail.Break_In + (rowSchSet.OT_Duration - rowSchSet.OT_Delay_Before_Break)
                            If rowSchDetail._End >= (24 * 60) Then
                                rowSchDetail._End = rowSchDetail.Break_In + (rowSchSet.OT_Duration - rowSchSet.OT_Delay_Before_Break) - (24 * 60)
                            End If
                        Else
                            rowSchDetail.Break_Out = 0
                            rowSchDetail.Break_In = 0
                            rowSchDetail._End = rowSchDetail.Start + rowSchSet.OT_Duration
                            If rowSchDetail._End >= (24 * 60) Then
                                rowSchDetail._End = rowSchDetail.Start + rowSchSet.OT_Duration - (24 * 60)
                            End If
                        End If
                        
                        rowSchDetail.Duration = rowSchSet.OT_Duration
                        'untuk holiday ot ##############################################################################
                    ElseIf rowSchDetail.OT_Type = ClsTrans_SPL.EnumTipeOT.Holiday_Overtime Then
                        'HOLIDAY OVERTIME - hitung start, break, end, duration
                        rowSchDetail.Start = rowSchDetail.Start_Time + rowSchSet.Holiday_OT_Delay
                        If rowSchDetail.Start >= (24 * 60) Then
                            rowSchDetail.Start = rowSchDetail.Start_Time + rowSchSet.Holiday_OT_Delay - (24 * 60)
                        End If
                        If rowSchSet.Holiday_OT_Break <> 0 And rowSchSet.Holiday_Delay_Before_Break <> 0 Then
                            rowSchDetail.Break_Out = rowSchDetail.Start + rowSchSet.Holiday_Delay_Before_Break
                            If rowSchDetail.Break_Out >= (24 * 60) Then
                                rowSchDetail.Break_Out = rowSchDetail.Start + rowSchSet.Holiday_Delay_Before_Break - (24 * 60)
                            End If
                            rowSchDetail.Break_In = rowSchDetail.Break_Out + rowSchSet.Holiday_OT_Break
                            If rowSchDetail.Break_In >= (24 * 60) Then
                                rowSchDetail.Break_In = rowSchDetail.Break_Out + rowSchSet.Holiday_OT_Break - (24 * 60)
                            End If
                            rowSchDetail._End = rowSchDetail.Break_In + (rowSchSet.Holiday_OT_Duration - rowSchSet.Holiday_Delay_Before_Break)
                            If rowSchDetail._End >= (24 * 60) Then
                                rowSchDetail._End = rowSchDetail.Break_In + (rowSchSet.Holiday_OT_Duration - rowSchSet.Holiday_Delay_Before_Break) - (24 * 60)
                            End If
                        Else
                            rowSchDetail.Break_Out = 0
                            rowSchDetail.Break_In = 0
                            rowSchDetail._End = rowSchDetail.Start + rowSchSet.Holiday_OT_Duration
                            If rowSchDetail._End >= (24 * 60) Then
                                rowSchDetail._End = rowSchDetail.Start + rowSchSet.Holiday_OT_Duration - (24 * 60)
                            End If
                        End If

                        rowSchDetail.Duration = rowSchSet.Holiday_OT_Duration
                    End If

                    'update dgvschdetail
                    vrowD.Cells(Me.ColSch_Start.Index).Value = Me.ObjCD.Convert_ToString(rowSchDetail.Start)
                    If rowSchDetail.Break_Out = 0 And rowSchDetail.Break_In = 0 Then
                        vrowD.Cells(Me.ColSch_Break.Index).Value = ""
                    Else
                        vrowD.Cells(Me.ColSch_Break.Index).Value = Me.ObjCD.Convert_ToString(rowSchDetail.Break_Out) & " - " & Me.ObjCD.Convert_ToString(rowSchDetail.Break_In)
                    End If
                    vrowD.Cells(Me.ColSch_End.Index).Value = Me.ObjCD.Convert_ToString(rowSchDetail._End)
                    vrowD.Cells(Me.ColSch_Duration.Index).Value = Me.ObjCD.Convert_ToString(rowSchDetail.Duration)
                    vrowD.Cells(Me.ColOverride.Index).Value = False
                End If
            End If

            'Me.UpdateTotalDuration()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdTbl_SchDetail(ByVal vrowD As DataGridViewRow)
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailRow
        Dim PrimaryShift As Double = CDbl(vrowD.Cells(Me.ColSch_PrimaryShift.Index).Value)
        Dim DateFilter As Date = CDate(vrowD.Cells(Me.ColSch_Date.Index).Value)
        Dim OTType As ClsTrans_SPL.EnumTipeOT = Me.GetTypeOT_Enum(vrowD.Cells(Me.ColSch_OTtype.Index).Value)
        Dim BreakInVal, BreakOutVal As String

        Try
            Me.TblSchDet.DefaultView.RowFilter = "[Primary Shift] = " & PrimaryShift & " AND [Date] = #" & DateFilter & "# AND [OT Type] = " & OTType
            If Me.TblSchDet.DefaultView.Count > 0 Then
                row = Me.TblSchDet.DefaultView(0).Row
                With row
                    .Start = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSch_Start.Index).Value)
                    If String.IsNullOrEmpty(vrowD.Cells(Me.ColSch_Break.Index).Value) = False Then
                        BreakOutVal = vrowD.Cells(Me.ColSch_Break.Index).Value.ToString.Remove(vrowD.Cells(Me.ColSch_Break.Index).Value.ToString.LastIndexOf("-"), _
                        vrowD.Cells(Me.ColSch_Break.Index).Value.ToString.Length - vrowD.Cells(Me.ColSch_Break.Index).Value.ToString.LastIndexOf("-"))
                        .Break_Out = Me.ObjCD.Convert_ToMinutes(BreakOutVal)
                        BreakInVal = vrowD.Cells(Me.ColSch_Break.Index).Value.ToString.Remove(0, vrowD.Cells(Me.ColSch_Break.Index).Value.ToString.LastIndexOf("-") + 1)
                        .Break_In = Me.ObjCD.Convert_ToMinutes(BreakInVal)
                    Else
                        .Break_In = 0
                        .Break_Out = 0
                    End If
                    ._End = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSch_End.Index).Value)
                    .Duration = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSch_Duration.Index).Value)
                    If vrowD.Cells(Me.ColOverride.Index).Value = False Then
                        .Override = 0
                    Else
                        .Override = 1
                    End If
                    .Remarks = vrowD.Cells(Me.ColSch_Remark.Index).Value
                End With

                Me.IsiTbl_SchEmp(PrimaryShift)
                'Me.UpdateTotalDuration()
            Else
                MsgBox("Row schedule detail not found", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdTbl_SchEmp(ByVal vrowD As DataGridViewRow)
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow
        Dim DateFilter As Date = CDate(vrowD.Cells(Me.ColSchEmp_Date.Index).Value)
        Dim PrimaryEmp As Double = CDbl(vrowD.Cells(Me.ColSchEmp_PrimaryEmp.Index).Value)
        Dim BreakInValEarly, BreakOutValEarly, BreakInValOT, BreakOutValOT As String

        Try
            Me.TblAppOT.DefaultView.RowFilter = "[Primary Employee] = " & PrimaryEmp.ToString & " AND [Date] = #" & DateFilter & "# "
            If Me.TblAppOT.DefaultView.Count > 0 Then
                row = Me.TblAppOT.DefaultView(0).Row
                With row
                    .Start_Time_Early = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSchEmp_StartEarly.Index).Value)
                    If String.IsNullOrEmpty(vrowD.Cells(Me.ColSchEmp_BreakEarly.Index).Value) = False Then
                        BreakOutValEarly = vrowD.Cells(Me.ColSchEmp_BreakEarly.Index).Value.ToString.Remove(vrowD.Cells(Me.ColSchEmp_BreakEarly.Index).Value.ToString.LastIndexOf("-"), _
                        vrowD.Cells(Me.ColSchEmp_BreakEarly.Index).Value.ToString.Length - vrowD.Cells(Me.ColSchEmp_BreakEarly.Index).Value.ToString.LastIndexOf("-"))
                        .Break_Out_Time_Early = Me.ObjCD.Convert_ToMinutes(BreakOutValEarly)
                        BreakInValEarly = vrowD.Cells(Me.ColSchEmp_BreakEarly.Index).Value.ToString.Remove(0, vrowD.Cells(Me.ColSchEmp_BreakEarly.Index).Value.ToString.LastIndexOf("-") + 1)
                        .Break_In_Time_Early = Me.ObjCD.Convert_ToMinutes(BreakInValEarly)
                    Else
                        .Break_Out_Time_Early = 0
                        .Break_In_Time_Early = 0
                    End If
                    .End_Time_Early = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSchEmp_EndEarly.Index).Value)
                    .OT_Early_Approved = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSchEmp_EarlyOTDuration.Index).Value)
                    If vrowD.Cells(Me.ColSchEmp_OverrideEarlyOT.Index).Value = False Then
                        .Override_Early_OT = 0
                    Else
                        .Override_Early_OT = 1
                    End If
                    .Start_Time = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSchEmp_StartOT.Index).Value)
                    If String.IsNullOrEmpty(vrowD.Cells(Me.ColSchEmp_BreakOT.Index).Value) = False Then
                        BreakOutValOT = vrowD.Cells(Me.ColSchEmp_BreakOT.Index).Value.ToString.Remove(vrowD.Cells(Me.ColSchEmp_BreakOT.Index).Value.ToString.LastIndexOf("-"), _
                        vrowD.Cells(Me.ColSchEmp_BreakOT.Index).Value.ToString.Length - vrowD.Cells(Me.ColSchEmp_BreakOT.Index).Value.ToString.LastIndexOf("-"))
                        .Break_Out_Time = Me.ObjCD.Convert_ToMinutes(BreakOutValOT)
                        BreakInValOT = vrowD.Cells(Me.ColSchEmp_BreakOT.Index).Value.ToString.Remove(0, vrowD.Cells(Me.ColSchEmp_BreakOT.Index).Value.ToString.LastIndexOf("-") + 1)
                        .Break_In_Time = Me.ObjCD.Convert_ToMinutes(BreakInValOT)
                    Else
                        .Break_Out_Time = 0
                        .Break_In_Time = 0
                    End If
                    .End_Time = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSchEmp_EndOT.Index).Value)
                    .OT_Aprroved = Me.ObjCD.Convert_ToMinutes(vrowD.Cells(Me.ColSchEmp_OTDuration.Index).Value)
                    If vrowD.Cells(Me.ColSchEmp_OverrideOT.Index).Value = False Then
                        .Override_OT = 0
                    Else
                        .Override_OT = 1
                    End If
                End With

                Me.UpdateTotalDuration()
            Else
                MsgBox("Row schedule employee not found", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdTbl_Approval(ByVal vrowD As DataGridViewRow)
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approvalRow
        Dim PrimaryUser As Double = CDbl(vrowD.Cells(Me.ColUser_Primary.Index).Value)

        Try
            Me.TblApp.DefaultView.RowFilter = "[Primary User] = " & PrimaryUser
            If Me.TblApp.DefaultView.Count > 0 Then
                row = Me.TblApp.DefaultView(0).Row
                With row
                    If vrowD.Cells(Me.ColUser_RequestApproval.Index).Value = True Then
                        .Req_Approval = 1
                    Else
                        .Req_Approval = 0
                    End If
                    .Remarks = vrowD.Cells(Me.ColUser_Remarks.Index).Value
                End With
            Else
                MsgBox("Row approval not found", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function HapusTbl_EmpShift(ByVal vPrimaryEmp As Double, ByVal vApakahHapusSch As Boolean) As Boolean
        Dim rowEmpShift As DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftRow
        Dim ListPrimaryShift As New List(Of Double)

        Try
            Me.TblEmpShift.DefaultView.RowFilter = "[Primary Employee] = " & vPrimaryEmp
            While Me.TblEmpShift.DefaultView.Count > 0
                rowEmpShift = Me.TblEmpShift.DefaultView(0).Row
                If ListPrimaryShift.Contains(rowEmpShift.Primary_Shift) = False Then
                    ListPrimaryShift.Add(rowEmpShift.Primary_Shift)
                End If
                Me.TblEmpShift.Rows.Remove(rowEmpShift)
            End While

            If vApakahHapusSch = True Then
                'jika setelah di hapus maka cek apakah masih ada yang menggunakan shift tsb jika tidak hapus sch setup & sch detail nya
                For Each PrimaryShift As Double In ListPrimaryShift
                    Me.TblEmpShift.DefaultView.RowFilter = "[Primary Shift] = " & PrimaryShift & " AND [Primary Employee] <> " & vPrimaryEmp
                    If Me.TblEmpShift.DefaultView.Count = 0 Then
                        Me.HapusTbl_SchSetup(PrimaryShift)
                        Me.HapusTbl_SchDetail(PrimaryShift)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub HapusTbl_SchSetup(ByVal vPrimaryShift As Double)
        Dim rowSchSet As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupRow

        Try
            Me.TblSchSetup.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift
            While Me.TblSchSetup.DefaultView.Count > 0
                rowSchSet = Me.TblSchSetup.DefaultView(0).Row
                Me.TblSchSetup.Rows.Remove(rowSchSet)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HapusTbl_SchDetail(ByVal vPrimaryShift As Double)
        Dim rowSchDetail As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailRow
        Dim PrimaryShift As Double = -1

        Try
            Me.TblSchDet.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift
            While Me.TblSchDet.DefaultView.Count > 0
                rowSchDetail = Me.TblSchDet.DefaultView(0).Row
                PrimaryShift = rowSchDetail.Primary_Shift
                Me.TblSchDet.Rows.Remove(rowSchDetail)

                Me.HapusTbl_SchEmp(PrimaryShift)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HapusTbl_SchDetail_1(ByVal vPrimaryShift As Double, ByVal vDate As Date, ByVal vEnumTypeOT As ClsTrans_SPL.EnumTipeOT)
        Dim rowSchDetail As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailRow

        Try
            Me.TblSchDet.DefaultView.RowFilter = "[Date] = #" & vDate & "# AND [Primary Shift] = " & vPrimaryShift & " AND [OT Type] = " & vEnumTypeOT
            While Me.TblSchDet.DefaultView.Count > 0
                rowSchDetail = Me.TblSchDet.DefaultView(0).Row
                Me.TblSchDet.Rows.Remove(rowSchDetail)

                Me.HapusTbl_SchEmp_3(vPrimaryShift, vDate, vEnumTypeOT)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HapusTbl_SchEmp(ByVal vPrimaryShift As Double)
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow

        Try
            Me.TblAppOT.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift
            While Me.TblAppOT.DefaultView.Count > 0
                row = Me.TblAppOT.DefaultView(0).Row
                Me.TblAppOT.Rows.Remove(row)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HapusTbl_SchEmp_1(ByVal vPrimaryEmp As Double)
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow

        Try
            Me.TblAppOT.DefaultView.RowFilter = "[Primary Employee] = " & vPrimaryEmp
            While Me.TblAppOT.DefaultView.Count > 0
                row = Me.TblAppOT.DefaultView(0).Row
                Me.TblAppOT.Rows.Remove(row)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HapusTbl_SchEmp_2(ByVal vPrimaryShift As Double, ByVal vPrimaryEmp As Double, ByVal vDate As Date)
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow

        Try
            Me.TblAppOT.DefaultView.RowFilter = "[Primary Shift] = " & vPrimaryShift.ToString & " AND [Primary Employee] = " & vPrimaryEmp.ToString & " AND [Date] = #" & vDate.Date & "#"
            While Me.TblAppOT.DefaultView.Count > 0
                row = Me.TblAppOT.DefaultView(0).Row
                Me.TblAppOT.Rows.Remove(row)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HapusTbl_SchEmp_3(ByVal vPrimaryShift As Double, ByVal vDate As Date, ByVal vEnumTypeOT As ClsTrans_SPL.EnumTipeOT)
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow

        Try
            Me.TblAppOT.DefaultView.RowFilter = "[Date] = #" & vDate & "# AND [Primary Shift] = " & vPrimaryShift
            For i As Integer = Me.TblAppOT.DefaultView.Count - 1 To 0 Step -1
                row = Me.TblAppOT.DefaultView(i).Row
                If row.OT_Early_Approved > 0 And row.OT_Aprroved > 0 Then
                    If vEnumTypeOT = ClsTrans_SPL.EnumTipeOT.Early_Overtime Then
                        row.OT_Early_Approved = 0
                        row.Start_Time_Early = 0
                        row.Break_Out_Time_Early = 0
                        row.Break_In_Time_Early = 0
                        row.End_Time_Early = 0
                    Else
                        row.OT_Aprroved = 0
                        row.Start_Time = 0
                        row.Break_Out_Time = 0
                        row.Break_In_Time = 0
                        row.End_Time = 0
                    End If
                Else
                    If vEnumTypeOT = ClsTrans_SPL.EnumTipeOT.Early_Overtime Then
                        If row.OT_Early_Approved > 0 Then
                            Me.TblAppOT.Rows.Remove(row)
                        End If
                    Else
                        If row.OT_Aprroved > 0 Then
                            Me.TblAppOT.Rows.Remove(row)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HapusTbl_Approval(ByVal vRowD As DataGridViewRow)
        Dim Dta1 As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_item_scopeTableAdapter
        Dim Tbl1 As New DseTrans_SPL_Setting.tbltrans_spl_approval_group_item_scopeDataTable
        Dim row1 As DseTrans_SPL_Setting.tbltrans_spl_approval_group_item_scopeRow
        Dim Dta2 As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_item_userTableAdapter
        Dim Tbl2 As New DseTrans_SPL_Setting.tbltrans_spl_approval_group_item_userDataTable
        Dim row2 As DseTrans_SPL_Setting.tbltrans_spl_approval_group_item_userRow
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approvalRow
        Dim rD As DataGridViewRow
        Dim ListPrimaryAuth As New List(Of Double)
        Dim ListPrimaryAuthHapus As New List(Of Double)

        Try
            Dta1.Connection.ConnectionString = Core_Module.GetConstring
            Dta2.Connection.ConnectionString = Core_Module.GetConstring

            Dta1.Fill_all(Tbl1)
            Dta2.Fill_all(Tbl2)

            For i As Integer = 0 To Me.dgvEmployee.RowCount - 2
                rD = Me.dgvEmployee.Rows(i)
                If rD.Index <> vRowD.Index Then
                    Tbl1.DefaultView.RowFilter = "[Primary Employee] = " & rD.Cells(Me.colEmpl_Primary.Index).Value
                    If Tbl1.DefaultView.Count > 0 Then
                        row1 = Tbl1.DefaultView(0).Row
                        If ListPrimaryAuth.Contains(row1.Primary_Transaksi) = False Then
                            ListPrimaryAuth.Add(row1.Primary_Transaksi)
                        End If
                    End If
                    Tbl1.DefaultView.RowFilter = "[Primary Position] = " & rD.Cells(Me.colEmpl_PrimaryPos.Index).Value
                    If Tbl1.DefaultView.Count > 0 Then
                        row1 = Tbl1.DefaultView(0).Row
                        If ListPrimaryAuth.Contains(row1.Primary_Transaksi) = False Then
                            ListPrimaryAuth.Add(row1.Primary_Transaksi)
                        End If
                    End If
                    Tbl1.DefaultView.RowFilter = "[Primary Department] = " & rD.Cells(Me.colEmpl_PrimaryDept.Index).Value
                    If Tbl1.DefaultView.Count > 0 Then
                        row1 = Tbl1.DefaultView(0).Row
                        If ListPrimaryAuth.Contains(row1.Primary_Transaksi) = False Then
                            ListPrimaryAuth.Add(row1.Primary_Transaksi)
                        End If
                    End If
                Else
                    Tbl1.DefaultView.RowFilter = "[Primary Employee] = " & rD.Cells(Me.colEmpl_Primary.Index).Value
                    If Tbl1.DefaultView.Count > 0 Then
                        row1 = Tbl1.DefaultView(0).Row
                        If ListPrimaryAuthHapus.Contains(row1.Primary_Transaksi) = False Then
                            ListPrimaryAuthHapus.Add(row1.Primary_Transaksi)
                        End If
                    End If
                    Tbl1.DefaultView.RowFilter = "[Primary Position] = " & rD.Cells(Me.colEmpl_PrimaryPos.Index).Value
                    If Tbl1.DefaultView.Count > 0 Then
                        row1 = Tbl1.DefaultView(0).Row
                        If ListPrimaryAuthHapus.Contains(row1.Primary_Transaksi) = False Then
                            ListPrimaryAuthHapus.Add(row1.Primary_Transaksi)
                        End If
                    End If
                    Tbl1.DefaultView.RowFilter = "[Primary Department] = " & rD.Cells(Me.colEmpl_PrimaryDept.Index).Value
                    If Tbl1.DefaultView.Count > 0 Then
                        row1 = Tbl1.DefaultView(0).Row
                        If ListPrimaryAuthHapus.Contains(row1.Primary_Transaksi) = False Then
                            ListPrimaryAuthHapus.Add(row1.Primary_Transaksi)
                        End If
                    End If
                End If
            Next

            Dim ListPrimaryUser As New List(Of Double)
            For Each primaryAuth As Double In ListPrimaryAuth
                Tbl2.DefaultView.RowFilter = "[Primary Transaksi] = " & primaryAuth
                If Tbl2.DefaultView.Count > 0 Then
                    For Each dv As DataRowView In Tbl2.DefaultView
                        row2 = dv.Row
                        If ListPrimaryUser.Contains(row2.Primary_User) = False Then
                            ListPrimaryUser.Add(row2.Primary_User)
                        End If
                    Next
                End If
            Next

            Dim ListPrimaryUserHapus As New List(Of Double)
            For Each primaryAuthHapus As Double In ListPrimaryAuthHapus
                Tbl2.DefaultView.RowFilter = "[Primary Transaksi] = " & primaryAuthHapus
                If Tbl2.DefaultView.Count > 0 Then
                    For Each dv As DataRowView In Tbl2.DefaultView
                        row2 = dv.Row
                        If ListPrimaryUserHapus.Contains(row2.Primary_User) = False Then
                            ListPrimaryUserHapus.Add(row2.Primary_User)
                        End If
                    Next
                End If
            Next

            For Each primaryUserHapus As Double In ListPrimaryUserHapus
                If ListPrimaryUser.Contains(primaryUserHapus) = False Then
                    Me.TblApp.DefaultView.RowFilter = "[Primary User] = " & primaryUserHapus
                    While Me.TblApp.DefaultView.Count > 0
                        row = Me.TblApp.DefaultView(0).Row
                        Me.TblApp.Rows.Remove(row)
                    End While
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbShift_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbShift.LinkClicked
        Dim i As New FormDummy
        Try
            i.vTblEnum = FormDummy.EnumTbl.shift
            i.vfrm = Me
            i.ShowDialog()
            i.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Llbsetup_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles Llbsetup.LinkClicked
        Dim i As New FormDummy
        Try
            i.vTblEnum = FormDummy.EnumTbl.setup
            i.vfrm = Me
            i.ShowDialog()
            i.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbDetail.LinkClicked
        Dim i As New FormDummy
        Try
            i.vTblEnum = FormDummy.EnumTbl.detail
            i.vfrm = Me
            i.ShowDialog()
            i.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbApp_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbApp.LinkClicked
        Dim i As New FormDummy
        Try
            i.vTblEnum = FormDummy.EnumTbl.approved
            i.vfrm = Me
            i.ShowDialog()
            i.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbAppOT_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbAppOT.LinkClicked
        Dim i As New FormDummy
        Try
            i.vTblEnum = FormDummy.EnumTbl.approved_ot
            i.vfrm = Me
            i.ShowDialog()
            i.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            'If Me.txtRemarks.Enabled = True Then
            Me.Cursor = Cursors.WaitCursor
            If Me.TabControl1.SelectedIndex = 1 Then
                If Me.TabControl2.SelectedIndex = 0 Then
                    Me.ObjItem.IsiDGVItemSchSetup(Me.dgvSchSetup, Me.txtPrimary.Text, Me, Me.TblSchSetup)
                ElseIf Me.TabControl2.SelectedIndex = 1 Then
                    Me.ObjItem.IsiDGVItemSchDetail(Me.dgvSchDetail, Me.txtPrimary.Text, Me, Me.TblSchDet)
                Else
                    Me.ObjItem.IsiDGVItemSchEmp(Me.dgvSchEmp, Me.txtPrimary.Text, Me, Me.TblAppOT)
                End If
            ElseIf Me.TabControl1.SelectedIndex = 2 Then
                Me.ObjItem.IsiDGVItemApproval(Me.dgvApproval, Me.txtPrimary.Text, Me, Me.TblApp)
            End If
            'End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub TabControl2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl2.SelectedIndexChanged
        Try
            'If Me.txtRemarks.Enabled = True Then
            Me.Cursor = Cursors.WaitCursor
            If Me.TabControl2.SelectedIndex = 0 Then
                Me.ObjItem.IsiDGVItemSchSetup(Me.dgvSchSetup, Me.txtPrimary.Text, Me, Me.TblSchSetup)
            ElseIf Me.TabControl2.SelectedIndex = 1 Then
                Me.ObjItem.IsiDGVItemSchDetail(Me.dgvSchDetail, Me.txtPrimary.Text, Me, Me.TblSchDet)
            Else
                Me.ObjItem.IsiDGVItemSchEmp(Me.dgvSchEmp, Me.txtPrimary.Text, Me, Me.TblAppOT)
            End If
            'End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Public Function GetTypeOT_String(ByVal vTypeOT As ClsTrans_SPL.EnumTipeOT) As String
        Dim Retval As String = ""

        Try
            If vTypeOT = ClsTrans_SPL.EnumTipeOT.Early_Overtime Then
                Retval = "Early Overtime"
            ElseIf vTypeOT = ClsTrans_SPL.EnumTipeOT.Overtime Then
                Retval = "Overtime"
            ElseIf vTypeOT = ClsTrans_SPL.EnumTipeOT.Holiday_Overtime Then
                Retval = "Holiday Overtime"
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Retval
        End Try
    End Function

    Public Function GetTypeOT_Enum(ByVal vTypeOTString As String) As ClsTrans_SPL.EnumTipeOT
        Dim Retval As ClsTrans_SPL.EnumTipeOT = ClsTrans_SPL.EnumTipeOT.Overtime

        Try
            If vTypeOTString = "Early Overtime" Then
                Retval = ClsTrans_SPL.EnumTipeOT.Early_Overtime
            ElseIf vTypeOTString = "Overtime" Then
                Retval = ClsTrans_SPL.EnumTipeOT.Overtime
            ElseIf vTypeOTString = "Holiday Overtime" Then
                Retval = ClsTrans_SPL.EnumTipeOT.Holiday_Overtime
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Retval
        End Try
    End Function

    Public Sub UpdateTotalDuration()
        Dim Dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approved_otTableAdapter
        Dim TblDum As New DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otDataTable
        Dim row, rowDum As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow
        Dim rowEmpShift As DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftRow
        Dim rowSchDetail As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailRow
        Dim DateFilter As Date = Now
        Dim HitungDurationOT As Double = 0

        Try
            For Each row In Me.TblAppOT
                HitungDurationOT += row.OT_Early_Approved
                HitungDurationOT += row.OT_Aprroved
            Next
            'Dta.Connection.ConnectionString = Core_Module.GetConstring
            'Me.TblEmpShift.DefaultView.RowFilter = Nothing
            'Me.TblEmpShift.DefaultView.Sort = "[Primary Employee] ASC, [Date From] ASC"
            'For ii As Integer = 0 To Me.TblEmpShift.DefaultView.Count - 1
            '    rowEmpShift = Me.TblEmpShift.DefaultView(ii).Row
            '    DateFilter = rowEmpShift.Date_From
            '    While DateFilter <= rowEmpShift.Date_To
            '        Me.TblSchDet.DefaultView.RowFilter = "[Date] = #" & DateFilter & "# AND [Primary Shift] = " & rowEmpShift.Primary_Shift
            '        Me.TblSchDet.DefaultView.Sort = "[OT Type] ASC"
            '        If Me.TblSchDet.DefaultView.Count > 0 Then
            '            For Each dv As DataRowView In Me.TblSchDet.DefaultView
            '                rowSchDetail = dv.Row
            '                HitungDurationOT += rowSchDetail.Duration
            '            Next
            '        End If

            '        DateFilter = DateFilter.AddDays(1)
            '    End While
            'Next

            Me.TxtTotalDuration.Text = Me.ObjCD.Convert_ToString(HitungDurationOT)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CustomEmployeeListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomEmployeeListToolStripMenuItem.Click
        Dim Frm As New Core_Module.FrmMaster_SPL_EmployeeCustom_Cari

        Try
            Frm.BackColorAll = Me.BackColorAll
            Frm.ShowDialog()
            If Frm.PrimaryTransaksi <> -1 Then
                Me.GetEmployeeData_ByAdd(ClsTrans_SPL.EnumCariData.Other, Frm.PrimaryTransaksi)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyToAllEmployeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApplyToAllEmployeeToolStripMenuItem.Click
        Try
            Me.GetEmployeeData_ByAdd(ClsTrans_SPL.EnumCariData.AllEmp, -1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyByDepartementToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApplyByDepartementToolStripMenuItem.Click
        Dim Frm As New FrmMasterDepartementCari
        Try
            Frm.BackColorAll = Me.BackColorAll
            Frm.ShowDialog()
            If Frm.PrimaryDepartement <> -1 Then
                Me.GetEmployeeData_ByAdd(ClsTrans_SPL.EnumCariData.dept, Frm.PrimaryDepartement)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyByJobPosirionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApplyByJobPosirionToolStripMenuItem.Click
        Dim Frm As New FrmMasterJobPosition_Cari
        Try
            Frm.BackColorAll = Me.BackColorAll
            Frm.ShowDialog()
            If Frm.mPrimary <> -1 Then
                Me.GetEmployeeData_ByAdd(ClsTrans_SPL.EnumCariData.Job, Frm.mPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'tambahan - kirim parameter tbl item desc karena jika replace data tbl item desc nya harus di clear
    Public Sub GetEmployeeData_ByAdd(ByVal vCariBerdasakan As ClsTrans_SPL.EnumCariData, ByVal vPrimary As Double)
        Dim dta As New DseTrans_SPLTableAdapters.get_employeeTableAdapter
        Dim tbl As New DseTrans_SPL.get_employeeDataTable
        Dim row As DseTrans_SPL.get_employeeRow

        'table adapter = tbltrans_spl_item_employee
        Dim Dta1 As New DseTrans_SPLTableAdapters.DataTable1TableAdapter
        Dim Tbl1 As New DseTrans_SPL.DataTable1DataTable
        Dim row1 As DseTrans_SPL.DataTable1Row
        'table adapter = tbltrans_attendance
        Dim Dta2 As New DseTrans_SPLTableAdapters.tbltrans_attendanceTableAdapter
        Dim Tbl2 As New DseTrans_SPL.tbltrans_attendanceDataTable
        'table adapter = tblmaster_spl_customemployee
        Dim Dta3 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_employee_custom_itemTableAdapter
        Dim Tbl3 As New DseTrans_SPL.tbltrans_spl_overtime_order_employee_custom_itemDataTable

        Dim rowdgv As DataGridViewRow
        Dim FrmDialog As New Core_Module.FrmAddReplace
        Dim InputKeList As Boolean = False
        Dim ObjKaryawan As New ClsMasterKaryawan
        Dim ListPrimaryAppGroup As New List(Of Double)
        Dim ListMsg_1 As String = ""
        Dim ListMsg_2 As String = ""

        Try
            FrmDialog.ShowDialog()
            If FrmDialog.DialogResult = DialogResult.OK Then 'jika add 
            ElseIf FrmDialog.DialogResult = DialogResult.Yes Then 'jika replace
                Me.dgvEmployee.Rows.Clear()
                Me.TblEmp.Clear()
                Me.TblEmpShift.Clear()
                Me.TblSchDet.Clear()
                Me.TblSchSetup.Clear()
                Me.TblApp.Clear()
            ElseIf FrmDialog.DialogResult = DialogResult.Cancel Then 'jika cancel
                Exit Sub
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvEmployee.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvEmployee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            ''filter data
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta3.Connection.ConnectionString = Core_Module.GetConstring
            If vCariBerdasakan = ClsTrans_SPL.EnumCariData.Job Then
                dta.FillBy_PrimaryPosition(tbl, vPrimary)
            ElseIf vCariBerdasakan = ClsTrans_SPL.EnumCariData.dept Then
                dta.FillBy_PrimaryDept(tbl, vPrimary)
            ElseIf vCariBerdasakan = ClsTrans_SPL.EnumCariData.Other Then
                dta.FillBy_AllEmployee(tbl)
                Dta3.FillBy_Primary(Tbl3, vPrimary)

                For i As Integer = tbl.Count - 1 To 0 Step -1
                    row = tbl.Rows(i)
                    Tbl3.DefaultView.RowFilter = "[Primary Employee] = " & row.Primary
                    If Tbl3.DefaultView.Count = 0 Then
                        tbl.Rows.Remove(row)
                    End If
                Next
            ElseIf vCariBerdasakan = ClsTrans_SPL.EnumCariData.AllEmp Then
                dta.FillBy_AllEmployee(tbl)
            End If

            ''pindahkan ke dgv
            'berdasarkan departemen atau job posisi
            For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                row = tbl.Rows(i)
                If row.Aktif = 0 Then
                    'cek data apakah sudah ada di dgv
                    'jika replace lewati pengecekan
                    InputKeList = True
                    If FrmDialog.DialogResult = DialogResult.OK Then 'jika add 
                        For Each rowdgv In Me.dgvEmployee.Rows
                            If rowdgv.Cells(Me.colEmpl_Primary.Index).Value = row.Primary Then
                                tbl.Rows.Remove(row)
                                InputKeList = False
                                Exit For
                            End If
                        Next
                    End If

                    If InputKeList = True Then
                        If ListPrimaryAppGroup.Contains(row.Primary) = False Then
                            ListPrimaryAppGroup.Add(row.Primary)
                        End If
                    End If
                Else
                    tbl.Rows.Remove(row)
                End If
            Next

            Dta1.Connection.ConnectionString = Core_Module.GetConstring
            Dta2.Connection.ConnectionString = Core_Module.GetConstring
            Dta1.Fill(Tbl1, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, ClsTrans_SPL.EnumStatus.Rejected) 'yang bukan reject)
            Dta2.Fill(Tbl2, Me.DtpTo.Value.Date, Me.DtpFrom.Value.Date)

            Me.CekSudahAllocate(tbl, -1, False, True)
            tbl.DefaultView.RowFilter = ""
            tbl.DefaultView.Sort = "[Emp_Code] DESC"
            For i As Integer = tbl.DefaultView.Count - 1 To 0 Step -1
                row = tbl.DefaultView(i).Row
                InputKeList = True
                If Me.Obj.ApaSudahAdaPadaTransaksiLain(Me.txtPrimary.Text, row.Primary, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, Tbl1) = True Then
                    If ListMsg_2 = "" Then
                        ListMsg_2 &= "Employee : " & row.Emp_Name & " (" & row.Emp_Code & ")"
                    Else
                        ListMsg_2 &= ", " & row.Emp_Name & " (" & row.Emp_Code & ")"
                    End If
                    InputKeList = False
                End If
                If Me.Obj.ApaSudahAdaAttendance(Me.txtPrimary.Text, row.Primary, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, Tbl2) = True Then
                    If ListMsg_1 = "" Then
                        ListMsg_1 &= "Employee : " & row.Emp_Name & " (" & row.Emp_Code & ")"
                    Else
                        ListMsg_1 &= ", " & row.Emp_Name & " (" & row.Emp_Code & ")"
                    End If
                    InputKeList = False
                End If

                If InputKeList = False Then
                    tbl.Rows.Remove(row)
                End If
            Next
            Me.CekApaMemilikiApprovalGroup(tbl, ListPrimaryAppGroup, True)

            For Each row In tbl
                Me.dgvEmployee.Rows.Add()
                If Me.dgvEmployee.Rows.Count > 1 Then
                    rowdgv = Me.dgvEmployee.Rows(Me.dgvEmployee.RowCount - 2)
                Else
                    rowdgv = Me.dgvEmployee.Rows(Me.dgvEmployee.RowCount)
                End If

                rowdgv.Cells(Me.colEmpl_No.Index).Value = rowdgv.Index + 1
                rowdgv.Cells(Me.colEmpl_Primary.Index).Value = row.Primary
                rowdgv.Cells(Me.colEmpl_PrimaryPos.Index).Value = row.Primary_Position
                rowdgv.Cells(Me.colEmpl_PrimaryDept.Index).Value = row.Primary_Dept
                rowdgv.Cells(Me.colEmpl_Name.Index).Value = row.Emp_Name
                rowdgv.Cells(Me.colEmpl_Code.Index).Value = row.Emp_Code
                rowdgv.Cells(Me.colEmpl_NIK.Index).Value = row.Emp_NIK
                rowdgv.Cells(Me.colEmpl_JobPosition.Index).Value = row.Position_Code & " - " & row.Position_Desc
                rowdgv.Cells(Me.colEmpl_Department.Index).Value = row.Dept_Code & " - " & row.Dept_Desc

                Me.IsiTbl_EmpShift(row.Primary)
                Me.IsiTbl_Approval(ListPrimaryAppGroup)
            Next

            Me.IsiTbl_SchSetup() 'IsiTbl_SchDetail di panggil dalam fungsi ini

            If ListMsg_1 <> "" Then
                MsgBox("Employee data is created at attendance transaction." & vbCrLf & ListMsg_1, MsgBoxStyle.Information, "Information")
            End If
            If ListMsg_2 <> "" Then
                MsgBox("Employee data is created at other transaction." & vbCrLf & ListMsg_2, MsgBoxStyle.Information, "Information")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmDialog.Dispose()
            Me.dgvEmployee.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvEmployee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi
        Dim FrmCetak As New FrmTrans_SPL_Cetak

        Try
            If Me.txtPrimary.Text = "" Or Me.txtPrimary.Text = -1 Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtPrimary.Text <> "" And Me.txtPrimary.Text <> -1 Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_SPL_Cetak") = False Then
                    MsgBox("You don't have permission to perform this action", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.Obj.Cetak(Me.txtPrimary.Text, FrmCetak.crv, FrmCetak.crvDetail) = True Then
                    FrmCetak.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmO.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Public Sub ShowSPL(ByVal vPrimarySpl As Double)
        Try
            Me.Obj.CariDetail(vPrimarySpl, Me)
            Me.baru = False
            Me.SalinKeField()
            Me.btnEdit.Focus()
        Catch ex As Exception
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function CekFormatHour(ByVal vStartOT As String, ByVal vBreakOT As String, ByVal vEndOT As String, _
    ByRef vBreakOut As String, ByRef vBreakIn As String) As Boolean

        Try
            'CEK FORMAT START OVERTIME
            If ObjCD.Is_Time_String(vStartOT) = False Then
                MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.ObjCD.Convert_ToMinutes(vStartOT) > 1439 Then
                MsgBox("Schedule in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'CEK FORMAT BREAK OVERTIME
            If String.IsNullOrEmpty(vBreakOT) = False Then
                'format untuk break harus 00.00 - 00.00
                If vBreakOT.ToString.Contains("-") = False Then
                    MsgBox("Format should be time string: hh.mm - hh.mm", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                'BREAK OUT
                vBreakOut = vBreakOT.ToString.Remove(vBreakOT.ToString.LastIndexOf("-"), vBreakOT.ToString.Length - vBreakOT.ToString.LastIndexOf("-"))
                If ObjCD.Is_Time_String(vBreakOut) = False Then
                    MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If Me.ObjCD.Convert_ToMinutes(vBreakOut) > 1439 Then
                    MsgBox("Schedule in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                'BREAK IN
                vBreakIn = vBreakOT.ToString.Remove(0, vBreakOT.ToString.LastIndexOf("-") + 1)
                If ObjCD.Is_Time_String(vBreakIn) = False Then
                    MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If Me.ObjCD.Convert_ToMinutes(vBreakIn) > 1439 Then
                    MsgBox("Schedule in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                vBreakOut = "00.00"
                vBreakIn = "00.00"
            End If

            'CEK FORMAT END OVERTIME
            If ObjCD.Is_Time_String(vEndOT) = False Then
                MsgBox("Format should be time string: hh.mm", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.ObjCD.Convert_ToMinutes(vEndOT) > 1439 Then
                MsgBox("Schedule in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function CekBatasanHitungHours(ByVal vEnumTipeOT As ClsTrans_SPL.EnumTipeOT, ByVal vStartWH As Integer, ByVal vEndWH As Integer, _
    ByVal vStartTime As Integer, ByVal vBreakOutTime As Integer, ByVal vBreakInTime As Integer, ByVal vEndTime As Integer, _
    ByVal vUseBreak As Boolean, ByVal vEnumOTEndEdit As EnumOvertimeEndEdit) As Boolean

        Try
            If vEnumTipeOT = ClsTrans_SPL.EnumTipeOT.Early_Overtime Then
                'EARLY OVERTIME #############################################################################################################################

                'START WH TO END
                If vStartWH < vEndTime Then
                    MsgBox("End not more than " & Me.ObjCD.Convert_ToString(vStartWH) & " minutes", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If vUseBreak = False Then
                    'END TO START
                    If vEndTime <= vStartTime Then
                        MsgBox("Start not more than " & Me.ObjCD.Convert_ToString(vEndTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Else
                    'END TO BREAK IN
                    If vEndTime <= vBreakInTime Then
                        MsgBox("Break in not more than " & Me.ObjCD.Convert_ToString(vEndTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    'BREAK IN TO BREAK OUT
                    If vBreakInTime <= vBreakOutTime Then
                        MsgBox("Break out not more than " & Me.ObjCD.Convert_ToString(vBreakInTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    'BREAK OUT TO START
                    If vBreakOutTime <= vStartTime Then
                        MsgBox("Start not more than " & Me.ObjCD.Convert_ToString(vBreakOutTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
                'START TO END WH

            ElseIf vEnumTipeOT = ClsTrans_SPL.EnumTipeOT.Overtime Then
                'OVERTIME #############################################################################################################################

                'END WH TO START 
                If vStartWH > vEndTime Then
                    'LINTAS HARI
                    If vEndWH > vStartTime And vStartTime >= vStartWH Then
                        MsgBox("Start not smaller than " & Me.ObjCD.Convert_ToString(vEndWH) & " minutes", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Else
                    'BUKAN LINTAS HARI
                    If vEndWH > vStartTime Then
                        MsgBox("Start not smaller than " & Me.ObjCD.Convert_ToString(vEndWH) & " minutes", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If

                If vUseBreak = False Then
                    'TANPA ISTIRAHAT

                    'START TO END 
                    If vStartWH > vEndTime Then
                        'LINTAS HARI
                        If vStartTime >= vEndTime And vStartTime < vStartWH Then
                            If vEnumOTEndEdit = EnumOvertimeEndEdit.Start_Enum Then
                                MsgBox("Start not more than " & Me.ObjCD.Convert_ToString(vEndTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                            Else
                                MsgBox("End not smaller than " & Me.ObjCD.Convert_ToString(vStartTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    Else
                        'BUKAN LINTAS HARI
                        If vStartTime >= vEndTime Then
                            If vEnumOTEndEdit = EnumOvertimeEndEdit.Start_Enum Then
                                MsgBox("Start not more than " & Me.ObjCD.Convert_ToString(vEndTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                            Else
                                MsgBox("End not smaller than " & Me.ObjCD.Convert_ToString(vStartTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                Else
                    'DENGAN ISTIRAHAT

                    'START TO BREAK OUT 
                    If vStartWH > vBreakOutTime And vBreakOutTime < vStartWH Then
                        'LINTAS HARI
                        If vStartTime >= vBreakOutTime And vStartTime < vEndWH Then
                            MsgBox("Start not more than " & Me.ObjCD.Convert_ToString(vBreakOutTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Else
                        'BUKAN LINTAS HARI
                        If vStartTime >= vBreakOutTime Then
                            If vEnumOTEndEdit = EnumOvertimeEndEdit.Start_Enum Then
                                MsgBox("Start not more than " & Me.ObjCD.Convert_ToString(vBreakOutTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                            ElseIf vEnumOTEndEdit = EnumOvertimeEndEdit.Break_Enum Then
                                MsgBox("Break out not smaller than " & Me.ObjCD.Convert_ToString(vStartTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If

                    'BREAK OUT TO BREAK IN
                    If vBreakOutTime > vBreakInTime And vBreakInTime < vStartWH Then
                        'LINTAS HARI
                        If vBreakOutTime >= vBreakInTime And vBreakOutTime < vEndWH Then
                            MsgBox("Break out not more than " & Me.ObjCD.Convert_ToString(vBreakInTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Else
                        If vStartTime >= vBreakOutTime And vBreakOutTime <= vBreakInTime And vBreakInTime >= vStartWH Then
                            'BREAK OUT LINTAS HARI DAN BREAK IN TIDAK LINTAS HARI
                            MsgBox("Break out not smaller than " & Me.ObjCD.Convert_ToString(vStartTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                            Return False
                        Else
                            'BUKAN LINTAS HARI
                            If vBreakOutTime >= vBreakInTime Then
                                MsgBox("Break out not more than " & Me.ObjCD.Convert_ToString(vBreakInTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        End If
                    End If

                    'BREAK IN TO END
                    If vBreakInTime > vEndTime And vEndTime < vStartWH Then
                        'LINTAS HARI
                        If vBreakInTime >= vEndTime And vBreakInTime < vEndWH Then
                            If vEnumOTEndEdit = EnumOvertimeEndEdit.End_Enum Then
                                MsgBox("End not smaller than " & Me.ObjCD.Convert_ToString(vBreakInTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                            Else
                                MsgBox("Break out not more than " & Me.ObjCD.Convert_ToString(vEndTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    Else
                        If vBreakOutTime >= vBreakInTime And vBreakInTime <= vEndTime And vEndTime >= vStartWH Then
                            'BREAK IN LINTAS HARI DAN END TIDAK LINTAS HARI
                            If vEnumOTEndEdit = EnumOvertimeEndEdit.End_Enum Then
                                MsgBox("End not smaller than " & Me.ObjCD.Convert_ToString(vBreakInTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                            Else
                                MsgBox("Break in not more than " & Me.ObjCD.Convert_ToString(vEndTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        Else
                            'BUKAN LINTAS HARI
                            If vBreakInTime >= vEndTime Then
                                If vEnumOTEndEdit = EnumOvertimeEndEdit.End_Enum Then
                                    MsgBox("End not smaller than " & Me.ObjCD.Convert_ToString(vBreakInTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                                Else
                                    MsgBox("Break in not more than " & Me.ObjCD.Convert_ToString(vEndTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                                End If
                                Return False
                            End If
                        End If
                    End If
                End If

                If vEnumOTEndEdit = EnumOvertimeEndEdit.End_Enum Then
                    If vStartWH <= vEndTime And vEndTime < vEndWH Then
                        MsgBox("End not more than " & Me.ObjCD.Convert_ToString(vStartWH - 1) & " minutes", MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    If vUseBreak = True Then
                        If vBreakInTime <= vEndTime And vBreakInTime <= vStartWH And vEndTime > vStartWH Then
                            MsgBox("End not smaller than " & Me.ObjCD.Convert_ToString(vBreakInTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
                ElseIf vEnumTipeOT = ClsTrans_SPL.EnumTipeOT.Holiday_Overtime Then
                    'HOLIDAY OVERTIME #############################################################################################################################

                    ''END WH TO START 
                    'If vStartWH > vEndTime Then
                    '    'LINTAS HARI
                    '    If vEndWH > vStartTime And vStartTime >= vStartWH Then
                    '        MsgBox("Start not smaller than " & Me.ObjCD.Convert_ToString(vEndWH) & " minutes", MsgBoxStyle.Information, "Information")
                    '        Return False
                    '    End If
                    'Else
                    '    'BUKAN LINTAS HARI
                    '    If vEndWH > vStartTime Then
                    '        MsgBox("Start not smaller than " & Me.ObjCD.Convert_ToString(vEndWH) & " minutes", MsgBoxStyle.Information, "Information")
                    '        Return False
                    '    End If
                    'End If

                    If vUseBreak = False Then
                        'TANPA ISTIRAHAT

                        'START TO END 
                        If vStartWH > vEndTime And vStartTime > vEndTime Then
                            'LINTAS HARI
                            If vStartTime <= vEndTime And vStartTime < vEndWH Then
                                If vEnumOTEndEdit = EnumOvertimeEndEdit.Start_Enum Then
                                    MsgBox("Start not more than " & Me.ObjCD.Convert_ToString(vEndTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                                Else
                                    MsgBox("End not smaller than " & Me.ObjCD.Convert_ToString(vStartTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                                End If
                                Return False
                            End If
                        Else
                            'BUKAN LINTAS HARI
                            If vStartTime >= vEndTime Then
                                If vEnumOTEndEdit = EnumOvertimeEndEdit.Start_Enum Then
                                    MsgBox("Start not more than " & Me.ObjCD.Convert_ToString(vEndTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                                Else
                                    MsgBox("End not smaller than " & Me.ObjCD.Convert_ToString(vStartTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                                End If
                                Return False
                            End If
                        End If
                    Else
                        'DENGAN ISTIRAHAT

                        'START TO BREAK OUT 
                        If vStartWH > vBreakOutTime And vBreakOutTime < vStartWH Then
                            'LINTAS HARI
                            If vStartTime >= vBreakOutTime And vStartTime < vStartWH Then
                                If vEnumOTEndEdit = EnumOvertimeEndEdit.Start_Enum Then
                                    MsgBox("Start not more than " & Me.ObjCD.Convert_ToString(vBreakOutTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                                Else
                                    MsgBox("Break out not smaller than " & Me.ObjCD.Convert_ToString(vStartTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                                End If
                                Return False
                            End If
                        Else
                            'BUKAN LINTAS HARI
                            If vStartTime >= vBreakOutTime Then
                                If vEnumOTEndEdit = EnumOvertimeEndEdit.Start_Enum Then
                                    MsgBox("Start not more than " & Me.ObjCD.Convert_ToString(vBreakOutTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                                ElseIf vEnumOTEndEdit = EnumOvertimeEndEdit.Break_Enum Then
                                    MsgBox("Break out not smaller than " & Me.ObjCD.Convert_ToString(vStartTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                                End If
                                Return False
                            End If
                        End If

                        'BREAK OUT TO BREAK IN
                        If vBreakOutTime > vBreakInTime And vBreakInTime < vStartWH Then
                            'LINTAS HARI
                            If vBreakOutTime >= vBreakInTime And vBreakOutTime < vStartWH Then
                                MsgBox("Break out not more than " & Me.ObjCD.Convert_ToString(vBreakInTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        Else
                            If vStartTime >= vBreakOutTime And vBreakOutTime <= vBreakInTime And vBreakInTime >= vStartWH Then
                                'BREAK OUT LINTAS HARI DAN BREAK IN TIDAK LINTAS HARI
                                MsgBox("Break out not smaller than " & Me.ObjCD.Convert_ToString(vStartTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                                Return False
                            Else
                                'BUKAN LINTAS HARI
                                If vBreakOutTime >= vBreakInTime Then
                                    MsgBox("Break out not more than " & Me.ObjCD.Convert_ToString(vBreakInTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                                    Return False
                                End If
                            End If
                        End If

                        'BREAK IN TO END
                        If vBreakInTime > vEndTime And vEndTime < vStartWH Then
                            'LINTAS HARI
                        If vBreakInTime >= vEndTime And vBreakOutTime < vStartWH Then
                            If vEnumOTEndEdit = EnumOvertimeEndEdit.End_Enum Then
                                MsgBox("End not smaller than " & Me.ObjCD.Convert_ToString(vBreakInTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                            Else
                                MsgBox("Break out not more than " & Me.ObjCD.Convert_ToString(vEndTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                        Else
                            If vBreakOutTime >= vBreakInTime And vBreakInTime <= vEndTime And vEndTime >= vStartWH Then
                                'BREAK IN LINTAS HARI DAN END TIDAK LINTAS HARI
                                If vEnumOTEndEdit = EnumOvertimeEndEdit.End_Enum Then
                                    MsgBox("End not smaller than " & Me.ObjCD.Convert_ToString(vBreakInTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                                Else
                                    MsgBox("Break in not more than " & Me.ObjCD.Convert_ToString(vEndTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                                End If
                                Return False
                            Else
                                'BUKAN LINTAS HARI
                                If vBreakInTime >= vEndTime Then
                                    If vEnumOTEndEdit = EnumOvertimeEndEdit.End_Enum Then
                                        MsgBox("End not smaller than " & Me.ObjCD.Convert_ToString(vBreakInTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                                    Else
                                        MsgBox("Break in not more than " & Me.ObjCD.Convert_ToString(vEndTime - 1) & " minutes", MsgBoxStyle.Information, "Information")
                                    End If
                                Return False
                            End If
                        End If
                    End If

                    If vEnumOTEndEdit = EnumOvertimeEndEdit.End_Enum Then
                        If vUseBreak = True Then
                            If vBreakInTime <= vEndTime And vBreakInTime <= vStartWH And vEndTime > vStartWH Then
                                MsgBox("End not smaller than " & Me.ObjCD.Convert_ToString(vBreakInTime + 1) & " minutes", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        End If
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
End Class