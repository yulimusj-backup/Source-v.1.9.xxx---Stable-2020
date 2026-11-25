Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class ClsTrans_SPL_Item
    Dim ObjCd As New Core_Module.ClsCekData

    Public Function IsiDGVItemEmployee(ByVal vDgvEmp As Windows.Forms.DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTrans_SPL, ByVal vTblEmp As DseTrans_SPL.tbltrans_spl_overtime_order_item_employeeDataTable) As Boolean
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_employeeRow
        Dim rowDGV As DataGridViewRow

        Try
            'kosongkan dgv
            vDgvEmp.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvEmp.Rows.Clear()
            ''isikan(tabel)
            For Each row In vTblEmp
                With row
                    vDgvEmp.Rows.Add("")
                    rowDGV = vDgvEmp.Rows(vDgvEmp.Rows.Count - 1)
                    rowDGV.Cells(Frm.colEmpl_No.Index).Value = vDgvEmp.Rows.Count
                    rowDGV.Cells(Frm.colEmpl_Primary.Index).Value = .Primary_Employee
                    rowDGV.Cells(Frm.colEmpl_PrimaryPos.Index).Value = .Primary_Position
                    rowDGV.Cells(Frm.colEmpl_PrimaryDept.Index).Value = .Primary_Deptartment
                    rowDGV.Cells(Frm.colEmpl_Remarks.Index).Value = .Remarks
                    rowDGV.Cells(Frm.colEmpl_Code.Index).Value = .Emp_Code
                    rowDGV.Cells(Frm.colEmpl_NIK.Index).Value = .Emp_NIK
                    rowDGV.Cells(Frm.colEmpl_Name.Index).Value = .Emp_Name
                    rowDGV.Cells(Frm.colEmpl_JobPosition.Index).Value = .Position_Code & " - " & .Position_Desc
                    rowDGV.Cells(Frm.colEmpl_Department.Index).Value = .Dept_Code & " - " & .Dept_Desc
                End With
            Next
            'set nomer
            vDgvEmp.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiDGVItemSchSetup(ByVal vDgvSchSetup As Windows.Forms.DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTrans_SPL, ByVal vTblSchSetup As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupDataTable) As Boolean
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupRow
        Dim rowDGV As DataGridViewRow

        Try
            vTblSchSetup.DefaultView.RowFilter = ""
            vTblSchSetup.DefaultView.Sort = "[Shift Code] ASC"

            'kosongkan dgv
            vDgvSchSetup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvSchSetup.Rows.Clear()
            ''isikan(tabel)
            For Each dv As DataRowView In vTblSchSetup.DefaultView
                row = dv.Row
                With row
                    vDgvSchSetup.Rows.Add("")
                    rowDGV = vDgvSchSetup.Rows(vDgvSchSetup.Rows.Count - 1)
                    rowDGV.Cells(Frm.ColSetup_No.Index).Value = vDgvSchSetup.Rows.Count
                    rowDGV.Cells(Frm.ColSetup_ShiftCode.Index).Value = .Shift_Code
                    rowDGV.Cells(Frm.ColSetup_ShiftDesc.Index).Value = .Shift_Desc
                    If .Is_Early_OT = 0 Then
                        rowDGV.Cells(Frm.ColSetup_IsEarlyOT.Index).Value = Boolean.FalseString
                    Else
                        rowDGV.Cells(Frm.ColSetup_IsEarlyOT.Index).Value = Boolean.TrueString
                    End If
                    rowDGV.Cells(Frm.ColSetup_EarlyOT.Index).Value = Me.ObjCd.Convert_ToString(.Early_Duration)
                    rowDGV.Cells(Frm.ColSetup_EarlyDelay.Index).Value = Me.ObjCd.Convert_ToString(.Early_Delay)
                    rowDGV.Cells(Frm.ColSetup_EarlyBreak.Index).Value = Me.ObjCd.Convert_ToString(.Early_Break)
                    rowDGV.Cells(Frm.ColSetup_EarlyDelayBeforeBreak.Index).Value = Me.ObjCd.Convert_ToString(.Early_Delay_Before_Break)
                    If .Is_OT = 0 Then
                        rowDGV.Cells(Frm.ColSetup_IsOT.Index).Value = Boolean.FalseString
                    Else
                        rowDGV.Cells(Frm.ColSetup_IsOT.Index).Value = Boolean.TrueString
                    End If
                    rowDGV.Cells(Frm.ColSetup_OT.Index).Value = Me.ObjCd.Convert_ToString(.OT_Duration)
                    rowDGV.Cells(Frm.ColSetup_OTDelay.Index).Value = Me.ObjCd.Convert_ToString(.OT_Delay)
                    rowDGV.Cells(Frm.ColSetup_OTBreak.Index).Value = Me.ObjCd.Convert_ToString(.OT_Break)
                    rowDGV.Cells(Frm.ColSetup_OTDelayBeforeBreak.Index).Value = Me.ObjCd.Convert_ToString(.OT_Delay_Before_Break)
                    If .Is_Holiday_OT = 0 Then
                        rowDGV.Cells(Frm.ColSetup_IsHolOT.Index).Value = Boolean.FalseString
                    Else
                        rowDGV.Cells(Frm.ColSetup_IsHolOT.Index).Value = Boolean.TrueString
                    End If
                    rowDGV.Cells(Frm.ColSetup_HolOT.Index).Value = Me.ObjCd.Convert_ToString(.Holiday_OT_Duration)
                    rowDGV.Cells(Frm.ColSetup_HolOTDelay.Index).Value = Me.ObjCd.Convert_ToString(.Holiday_OT_Delay)
                    rowDGV.Cells(Frm.ColSetup_HolOTBreak.Index).Value = Me.ObjCd.Convert_ToString(.Holiday_OT_Break)
                    rowDGV.Cells(Frm.ColSetup_HolOTDelayBeforeBreak.Index).Value = Me.ObjCd.Convert_ToString(.Holiday_Delay_Before_Break)
                    rowDGV.Cells(Frm.ColSetup_Remarks.Index).Value = .Remarks
                    rowDGV.Cells(Frm.ColSetup_PrimaryWorkshift.Index).Value = .Primary_Shift
                End With
            Next
            'set nomer
            vDgvSchSetup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiDGVItemSchDetail(ByVal vDgvSchDet As Windows.Forms.DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTrans_SPL, ByVal vTblSchDet As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailDataTable) As Boolean
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailRow
        Dim rowDGV As DataGridViewRow
        Dim ShowBreak As Boolean = False

        Try
            vTblSchDet.DefaultView.RowFilter = ""
            vTblSchDet.DefaultView.Sort = "[Shift Code] ASC, [Date] ASC, [OT Type] ASC"

            'kosongkan dgv
            vDgvSchDet.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvSchDet.Rows.Clear()
            ''isikan(tabel)
            For Each dv As DataRowView In vTblSchDet.DefaultView
                row = dv.Row
                With row
                    vDgvSchDet.Rows.Add("")
                    rowDGV = vDgvSchDet.Rows(vDgvSchDet.Rows.Count - 1)
                    rowDGV.Cells(Frm.ColSch_No.Index).Value = vDgvSchDet.RowCount
                    rowDGV.Cells(Frm.ColSch_DayOfWeek.Index).Value = Me.ObjCd.GetHari_English(._Date)
                    rowDGV.Cells(Frm.ColSch_Date.Index).Value = Format(._Date, "dd MMM yyyy")
                    rowDGV.Cells(Frm.ColSch_Shift.Index).Value = .Shift_Code
                    rowDGV.Cells(Frm.ColSch_NormalWH.Index).Value = Me.ObjCd.Convert_ToString(.Start_Time) & " - " & Me.ObjCd.Convert_ToString(.End_Time)
                    rowDGV.Cells(Frm.ColSch_OTtype.Index).Value = Frm.GetTypeOT_String(.OT_Type)
                    rowDGV.Cells(Frm.ColSch_Start.Index).Value = Me.ObjCd.Convert_ToString(.Start)
                    If .Break_Out <> 0 Or .Break_In <> 0 Then
                        rowDGV.Cells(Frm.ColSch_Break.Index).Value = Me.ObjCd.Convert_ToString(.Break_Out) & " - " & Me.ObjCd.Convert_ToString(.Break_In)
                        ShowBreak = True
                    Else
                        rowDGV.Cells(Frm.ColSch_Break.Index).Value = ""
                    End If
                    rowDGV.Cells(Frm.ColSch_End.Index).Value = Me.ObjCd.Convert_ToString(._End)
                    rowDGV.Cells(Frm.ColSch_Duration.Index).Value = Me.ObjCd.Convert_ToString(.Duration)
                    If .Override = 0 Then
                        rowDGV.Cells(Frm.ColOverride.Index).Value = Boolean.FalseString
                    Else
                        rowDGV.Cells(Frm.ColOverride.Index).Value = Boolean.TrueString
                    End If
                    rowDGV.Cells(Frm.ColSch_Remark.Index).Value = .Remarks
                    rowDGV.Cells(Frm.ColSch_PrimaryShift.Index).Value = .Primary_Shift
                End With
            Next

            If ShowBreak = True Then
                Frm.ColSch_Break.Visible = True
            Else
                Frm.ColSch_Break.Visible = False
            End If

            'set nomer
            vDgvSchDet.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiDGVItemApproval(ByVal vDgvApp As Windows.Forms.DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTrans_SPL, ByVal vTblApp As DseTrans_SPL.tbltrans_spl_overtime_order_item_approvalDataTable) As Boolean
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approvalRow
        Dim rowDGV As DataGridViewRow

        Try
            vTblApp.DefaultView.RowFilter = Nothing
            vTblApp.DefaultView.Sort = "[Kode User] ASC"

            'kosongkan dgv
            vDgvApp.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvApp.Rows.Clear()
            ''isikan(tabel)
            For Each dv As DataRowView In vTblApp.DefaultView
                row = dv.Row
                With row
                    vDgvApp.Rows.Add("")
                    rowDGV = vDgvApp.Rows(vDgvApp.Rows.Count - 1)
                    rowDGV.Cells(Frm.ColUser_No.Index).Value = vDgvApp.Rows.Count
                    If .Req_Approval = 0 Then
                        rowDGV.Cells(Frm.ColUser_RequestApproval.Index).Value = Boolean.FalseString
                    Else
                        rowDGV.Cells(Frm.ColUser_RequestApproval.Index).Value = Boolean.TrueString
                    End If
                    rowDGV.Cells(Frm.ColUser_Code.Index).Value = .Kode_User
                    rowDGV.Cells(Frm.ColUser_Name.Index).Value = .Nama_User
                    rowDGV.Cells(Frm.ColUser_Email.Index).Value = .Email_Address
                    If .Last_Delivery.Date = CDate("1/1/1950") Then
                        rowDGV.Cells(Frm.Coluser_LastDelivery.Index).Value = "-"
                    Else
                        rowDGV.Cells(Frm.Coluser_LastDelivery.Index).Value = Format(.Last_Delivery, "dd MMM yyyy") & " " & .Last_Delivery.ToShortTimeString
                    End If
                    rowDGV.Cells(Frm.ColUser_Remarks.Index).Value = .Remarks
                    rowDGV.Cells(Frm.ColUser_Approval.Index).Value = "Approval"
                    rowDGV.Cells(Frm.ColUser_Primary.Index).Value = .Primary_User
                End With
            Next
            'set nomer
            vDgvApp.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiDGVItemSchEmp(ByVal vDgvSchEmp As Windows.Forms.DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTrans_SPL, ByVal vTblAppOT As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otDataTable) As Boolean
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow
        Dim rowDGV As DataGridViewRow
        Dim ShowEarly As Boolean = False
        Dim ShowOT As Boolean = False
        Dim ShowBreakEarly As Boolean = False
        Dim ShowBreakOT As Boolean = False

        Try
            vTblAppOT.DefaultView.RowFilter = Nothing
            vTblAppOT.DefaultView.Sort = "[Emp Code] ASC, [Date] ASC"

            'kosongkan dgv
            vDgvSchEmp.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvSchEmp.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDgvSchEmp.Rows.Clear()
            ''isikan(tabel)
            For Each dv As DataRowView In vTblAppOT.DefaultView
                row = dv.Row
                With row
                    vDgvSchEmp.Rows.Add("")
                    rowDGV = vDgvSchEmp.Rows(vDgvSchEmp.Rows.Count - 1)
                    rowDGV.Cells(Frm.ColSchEmp_No.Index).Value = vDgvSchEmp.Rows.Count
                    rowDGV.Cells(Frm.ColSchEmp_Code.Index).Value = row.Emp_Code
                    rowDGV.Cells(Frm.ColSchEmp_Emp.Index).Value = row.Emp_Name
                    rowDGV.Cells(Frm.ColSchEmp_DayOfWeek.Index).Value = Me.ObjCd.GetHari_English(._Date)
                    rowDGV.Cells(Frm.ColSchEmp_IsHoliday.Index).Value = .is_holiday
                    rowDGV.Cells(Frm.ColSchEmp_Date.Index).Value = Format(._Date, "dd MMM yyyy")
                    rowDGV.Cells(Frm.ColSchEmp_Shift.Index).Value = row.Shift_Code
                    rowDGV.Cells(Frm.ColSchEmp_NormalWH.Index).Value = Me.ObjCd.Convert_ToString(.Start_WH) & " - " & Me.ObjCd.Convert_ToString(.End_WH)

                    rowDGV.Cells(Frm.ColSchEmp_StartEarly.Index).Value = Me.ObjCd.Convert_ToString(row.Start_Time_Early)
                    If .Break_Out_Time_Early <> 0 Or .Break_In_Time_Early <> 0 Then
                        rowDGV.Cells(Frm.ColSchEmp_BreakEarly.Index).Value = Me.ObjCd.Convert_ToString(.Break_Out_Time_Early) & " - " & Me.ObjCd.Convert_ToString(.Break_In_Time_Early)
                        ShowBreakEarly = True
                    Else
                        rowDGV.Cells(Frm.ColSchEmp_BreakEarly.Index).Value = ""
                    End If
                    rowDGV.Cells(Frm.ColSchEmp_EndEarly.Index).Value = Me.ObjCd.Convert_ToString(row.End_Time_Early)
                    rowDGV.Cells(Frm.ColSchEmp_EarlyOTDuration.Index).Value = Me.ObjCd.Convert_ToString(row.OT_Early_Approved)
                    If .Override_Early_OT = 0 Then
                        rowDGV.Cells(Frm.ColSchEmp_OverrideEarlyOT.Index).Value = Boolean.FalseString
                    Else
                        rowDGV.Cells(Frm.ColSchEmp_OverrideEarlyOT.Index).Value = Boolean.TrueString
                    End If

                    rowDGV.Cells(Frm.ColSchEmp_StartOT.Index).Value = Me.ObjCd.Convert_ToString(row.Start_Time)
                    If .Break_Out_Time <> 0 Or .Break_In_Time <> 0 Then
                        rowDGV.Cells(Frm.ColSchEmp_BreakOT.Index).Value = Me.ObjCd.Convert_ToString(.Break_Out_Time) & " - " & Me.ObjCd.Convert_ToString(.Break_In_Time)
                        ShowBreakOT = True
                    Else
                        rowDGV.Cells(Frm.ColSchEmp_BreakOT.Index).Value = ""
                    End If
                    rowDGV.Cells(Frm.ColSchEmp_EndOT.Index).Value = Me.ObjCd.Convert_ToString(row.End_Time)
                    rowDGV.Cells(Frm.ColSchEmp_OTDuration.Index).Value = Me.ObjCd.Convert_ToString(row.OT_Aprroved)
                    If .Override_OT = 0 Then
                        rowDGV.Cells(Frm.ColSchEmp_OverrideOT.Index).Value = Boolean.FalseString
                    Else
                        rowDGV.Cells(Frm.ColSchEmp_OverrideOT.Index).Value = Boolean.TrueString
                    End If

                    rowDGV.Cells(Frm.ColSchEmp_PrimaryShift.Index).Value = row.Primary_Shift
                    rowDGV.Cells(Frm.ColSchEmp_PrimaryEmp.Index).Value = row.Primary_Employee

                    If .OT_Early_Approved > 0 Then
                        ShowEarly = True
                    End If
                    If .OT_Aprroved > 0 Then
                        ShowOT = True
                    End If
                End With
            Next

            If ShowEarly = False Then
                Frm.ColSchEmp_StartEarly.Visible = False
                Frm.ColSchEmp_BreakEarly.Visible = False
                Frm.ColSchEmp_EndEarly.Visible = False
                Frm.ColSchEmp_EarlyOTDuration.Visible = False
                Frm.ColSchEmp_OverrideEarlyOT.Visible = False
            Else
                Frm.ColSchEmp_StartEarly.Visible = True
                Frm.ColSchEmp_BreakEarly.Visible = True
                Frm.ColSchEmp_EndEarly.Visible = True
                Frm.ColSchEmp_EarlyOTDuration.Visible = True
                Frm.ColSchEmp_OverrideEarlyOT.Visible = True
            End If

            If ShowOT = False Then
                Frm.ColSchEmp_StartOT.Visible = False
                Frm.ColSchEmp_BreakOT.Visible = False
                Frm.ColSchEmp_EndOT.Visible = False
                Frm.ColSchEmp_OTDuration.Visible = False
                Frm.ColSchEmp_OverrideOT.Visible = False
            Else
                Frm.ColSchEmp_StartOT.Visible = True
                Frm.ColSchEmp_BreakOT.Visible = True
                Frm.ColSchEmp_EndOT.Visible = True
                Frm.ColSchEmp_OTDuration.Visible = True
                Frm.ColSchEmp_OverrideOT.Visible = True
            End If

            If ShowBreakEarly = True Then
                Frm.ColSchEmp_BreakEarly.Visible = True
            Else
                Frm.ColSchEmp_BreakEarly.Visible = False
            End If
            If ShowBreakOT = True Then
                Frm.ColSchEmp_BreakOT.Visible = True
            Else
                Frm.ColSchEmp_BreakOT.Visible = False
            End If

            'set nomer
            vDgvSchEmp.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgvSchEmp.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'ByVal vTblAppOT As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otDataTable, _
    Public Function IsiTblItem(ByVal vPrimaryTransaksi As Double, _
    ByVal vTblEmp As DseTrans_SPL.tbltrans_spl_overtime_order_item_employeeDataTable, _
    ByVal vTblSchSetup As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupDataTable, _
    ByVal vTblSchDet As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailDataTable, _
    ByVal vTblApp As DseTrans_SPL.tbltrans_spl_overtime_order_item_approvalDataTable, _
    ByVal vTblEmpShift As DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftDataTable, _
    ByVal vTblAppOT As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otDataTable) As Boolean
        Dim Dta1 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_employeeTableAdapter
        Dim Dta2 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_sch_setupTableAdapter
        Dim Dta3 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_sch_detailTableAdapter
        Dim Dta4 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approvalTableAdapter
        Dim Dta5 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approved_otTableAdapter
        Dim Dta6 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_employee_shiftTableAdapter

        Try
            Dta1.Connection.ConnectionString = Core_Module.GetConstring
            Dta2.Connection.ConnectionString = Core_Module.GetConstring
            Dta3.Connection.ConnectionString = Core_Module.GetConstring
            Dta4.Connection.ConnectionString = Core_Module.GetConstring
            Dta5.Connection.ConnectionString = Core_Module.GetConstring
            Dta6.Connection.ConnectionString = Core_Module.GetConstring

            Dta1.Fill(vTblEmp, vPrimaryTransaksi)
            Dta2.Fill(vTblSchSetup, vPrimaryTransaksi)
            Dta3.Fill(vTblSchDet, vPrimaryTransaksi)
            Dta4.Fill(vTblApp, vPrimaryTransaksi)
            Dta5.Fill(vTblAppOT, vPrimaryTransaksi)
            Dta6.Fill(vTblEmpShift, vPrimaryTransaksi)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA1 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_employeeTableAdapter
        Dim DTA2 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_sch_setupTableAdapter
        Dim DTA3 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_sch_detailTableAdapter
        Dim DTA4 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approvalTableAdapter
        Dim DTA5 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approved_otTableAdapter
        Dim DTA6 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_employee_shiftTableAdapter

        Try
            'hapus data item
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            DTA3.Connection.ConnectionString = Core_Module.GetConstring
            DTA4.Connection.ConnectionString = Core_Module.GetConstring
            DTA5.Connection.ConnectionString = Core_Module.GetConstring
            DTA6.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Delete(vPrimaryTransaksi)
            DTA2.Delete(vPrimaryTransaksi)
            DTA3.Delete(vPrimaryTransaksi)
            DTA4.Delete(vPrimaryTransaksi)
            DTA5.Delete(vPrimaryTransaksi)
            DTA6.Delete(vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanItem_UsingScript(ByVal vPrimaryTransaksi As Double, ByVal i As FrmTrans_SPL) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double

        Try
            'open connection
            Conn.Open()

            'simpan item employee
            sql = ""
            StartRowIndex = 0
            i.dgvEmployee.AllowUserToAddRows = False
            While StartRowIndex < i.dgvEmployee.Rows.Count
                sql = Me.GetInsertString_ItemEmployee(StartRowIndex, BatchQty, i.dgvEmployee, vPrimaryTransaksi, i)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            'simpan item employee shift
            sql = ""
            StartRowIndex = 0
            While StartRowIndex < i.TblEmpShift.Rows.Count
                sql = Me.GetInsertString_ItemEmployeeShift(StartRowIndex, BatchQty, i.TblEmpShift, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            'simpan item schedule setup
            sql = ""
            StartRowIndex = 0
            While StartRowIndex < i.TblSchSetup.Rows.Count
                sql = Me.GetInsertString_ItemSchSetup(StartRowIndex, BatchQty, i.TblSchSetup, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            'simpan item schedule detail
            sql = ""
            StartRowIndex = 0
            While StartRowIndex < i.TblSchDet.Rows.Count
                sql = Me.GetInsertString_ItemSchDetail(StartRowIndex, BatchQty, i.TblSchDet, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            'simpan item approval 
            sql = ""
            StartRowIndex = 0
            While StartRowIndex < i.TblApp.Rows.Count
                sql = Me.GetInsertString_ItemApproval(StartRowIndex, BatchQty, i.TblApp, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            'simpan item approval ot
            sql = ""
            StartRowIndex = 0
            While StartRowIndex < i.TblAppOT.Rows.Count
                sql = Me.GetInsertString_ItemApprovalOT(StartRowIndex, BatchQty, i.TblAppOT, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemEmployee(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vDgvEmp As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal ij As FrmTrans_SPL) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","

        Try
            retval = "INSERT INTO `tbltrans_spl_overtime_order_item_employee` (`Primary Transaksi`, `Primary Employee`, `Remarks`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vDgvEmp.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vDgvEmp.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vDgvEmp.Rows(vStartRowIndex + i)

                    If String.IsNullOrEmpty(row.Cells(ij.colEmpl_Remarks.Index).Value) = True Then
                        row.Cells(ij.colEmpl_Remarks.Index).Value = ""
                    End If

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + CDbl(row.Cells(ij.colEmpl_Primary.Index).Value).ToString + "," _
                    + "'" + row.Cells(ij.colEmpl_Remarks.Index).Value.ToString + "'" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next

            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetInsertString_ItemEmployeeShift(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTblEmpShift As DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftRow

        Try
            retval = "INSERT INTO `tbltrans_spl_overtime_order_item_employee_shift` (`Primary Transaksi`, `Primary Employee`, `Primary Shift`, `Date From`, `Date To`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTblEmpShift.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTblEmpShift.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTblEmpShift.Rows(vStartRowIndex + i)

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + row.Primary_Employee.ToString + "," _
                    + row.Primary_Shift.ToString + "," _
                    + "'" + Format(row.Date_From, "yyyy-MM-dd") + "'," _
                    + "'" + Format(row.Date_To, "yyyy-MM-dd") + "'" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next

            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetInsertString_ItemSchSetup(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTblSetup As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_setupRow

        Try
            retval = "INSERT INTO `tbltrans_spl_overtime_order_item_sch_setup` (`Primary Transksi`, `Primary Shift`, `Is Early OT`, `Early Duration`, `Early Delay`, `Early Break`, `Early Delay Before Break`, `Is OT`, `OT Duration`, `OT Delay`, `OT Break`, `OT Delay Before Break`, `Is Holiday OT`, `Holiday OT Duration`, `Holiday OT Delay`, `Holiday OT Break`, `Holiday Delay Before Break`, `Remarks`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTblSetup.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTblSetup.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTblSetup.Rows(vStartRowIndex + i)

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + row.Primary_Shift.ToString + "," _
                    + row.Is_Early_OT.ToString + "," _
                    + row.Early_Duration.ToString + "," _
                    + row.Early_Delay.ToString + "," _
                    + row.Early_Break.ToString + "," _
                    + row.Early_Delay_Before_Break.ToString + "," _
                    + row.Is_OT.ToString + "," _
                    + row.OT_Duration.ToString + "," _
                    + row.OT_Delay.ToString + "," _
                    + row.OT_Break.ToString + "," _
                    + row.OT_Delay_Before_Break.ToString + "," _
                    + row.Is_Holiday_OT.ToString + "," _
                    + row.Holiday_OT_Duration.ToString + "," _
                    + row.Holiday_OT_Delay.ToString + "," _
                    + row.Holiday_OT_Break.ToString + "," _
                    + row.Holiday_Delay_Before_Break.ToString + "," _
                    + "'" + row.Remarks.ToString + "'" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next

            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetInsertString_ItemSchDetail(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTblSchDet As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailRow

        Try
            retval = "INSERT INTO `tbltrans_spl_overtime_order_item_sch_detail` (`Primary Transaksi`, `Date`, `Primary Shift`, `Start Time`, `End Time`, `OT Type`, `Start`, `Break In`, `Break Out`, `End`, `Duration`, `Override`, `Remarks`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTblSchDet.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTblSchDet.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTblSchDet.Rows(vStartRowIndex + i)

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + "'" + Format(row._Date, "yyyy-MM-dd") + "'," _
                    + row.Primary_Shift.ToString + "," _
                    + row.Start_Time.ToString + "," _
                    + row.End_Time.ToString + "," _
                    + row.OT_Type.ToString + "," _
                    + row.Start.ToString + "," _
                    + row.Break_In.ToString + "," _
                    + row.Break_Out.ToString + "," _
                    + row._End.ToString + "," _
                    + row.Duration.ToString + "," _
                    + row.Override.ToString + "," _
                    + "'" + row.Remarks.ToString + "'" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next

            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetInsertString_ItemApproval(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTblApp As DseTrans_SPL.tbltrans_spl_overtime_order_item_approvalDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approvalRow

        Try
            retval = "INSERT INTO `tbltrans_spl_overtime_order_item_approval` (`Primary Transaksi`, `Req Approval`, `Primary User`, `Email Address`, `Last Delivery`, `Remarks`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTblApp.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTblApp.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTblApp.Rows(vStartRowIndex + i)

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + row.Req_Approval.ToString + "," _
                    + row.Primary_User.ToString + "," _
                    + "'" + row.Email_Address.ToString + "'," _
                    + "'" + Format(row.Last_Delivery, "yyyy-MM-dd hh:mm") + "'," _
                    + "'" + row.Remarks.ToString + "'" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next

            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetInsertString_ItemApprovalOT(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTblAppOT As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow

        Try
            retval = "INSERT INTO `tbltrans_spl_overtime_order_item_approved_ot` (`Primary Transaksi`, `Primary Employee`, `Primary Shift`, `Date`, `OT Early Approved`, `OT Aprroved`, `Start Time Early`, `Break Out Time Early`, `Break In Time Early`, `End Time Early`, `Start Time`, `Break Out Time`, `Break In Time`, `End Time`, `Override Early OT`, `Override OT`, `Start WH`, `End WH`, `is holiday`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTblAppOT.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTblAppOT.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTblAppOT.Rows(vStartRowIndex + i)

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + row.Primary_Employee.ToString + "," _
                    + row.Primary_Shift.ToString + "," _
                    + "'" + Format(row._Date, "yyyy-MM-dd") + "'," _
                    + row.OT_Early_Approved.ToString + "," _
                    + row.OT_Aprroved.ToString + "," _
                    + row.Start_Time_Early.ToString + "," _
                    + row.Break_Out_Time_Early.ToString + "," _
                    + row.Break_In_Time_Early.ToString + "," _
                    + row.End_Time_Early.ToString + "," _
                    + row.Start_Time.ToString + "," _
                    + row.Break_Out_Time.ToString + "," _
                    + row.Break_In_Time.ToString + "," _
                    + row.End_Time.ToString + "," _
                    + row.Override_Early_OT.ToString + "," _
                    + row.Override_OT.ToString + "," _
                    + row.Start_WH.ToString + "," _
                    + row.End_WH.ToString + "," _
                    + row.is_holiday.ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next

            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function UpdateDurationOvertime(ByVal vEarlyOTDuration As Double, ByVal vOTDuration As Double, ByVal vPrimaryAttendance As Double, _
    ByVal vPrimaryOvertime As Double, ByVal vPrimaryShift As Double, ByVal vPrimaryEmp As Double, ByVal vDate As Date) As Boolean
        Dim dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approved_otTableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_OTDurationNPrimaryAttendance(vEarlyOTDuration, vOTDuration, vPrimaryAttendance, vPrimaryEmp, vPrimaryShift, vDate.Date, vPrimaryOvertime)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
