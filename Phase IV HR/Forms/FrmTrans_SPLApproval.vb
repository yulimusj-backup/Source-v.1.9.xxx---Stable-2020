Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms, System.Threading
Public Class FrmTrans_SPLApproval
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Dim PrimaryUser As Double = -1
    Dim DapatAppOverBudget As Boolean = False

    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSpl As New HR.ClsTrans_SPL
    Dim ObjEmp As New Core_Module.ClsMasterKaryawan
    Dim ObjJobPost As New Core_Module.ClsMasterJobPosition
    Dim ObjDept As New Core_Module.ClsMasterDepartement
    Dim TblSplApp As New DseTrans_SPL.tbltrans_spl_approvalDataTable
    Dim TblSplUsed As New DseTrans_SPL.ot_order_approvalDataTable
    'Dim TblSplAppDep As New DseTrans_SPL.ot_order_approval_1DataTable
    Dim TblSplAppDep2 As New DseTrans_SPL.ot_order_approval_2DataTable

    Dim BisaApproved As Boolean = True

    Dim Buffer As String = ""
    Dim EraseRowIndex As Integer = 1
    Dim EnumFillSPL As ClsTrans_SPL.EnumStatus

    Public Enum EnumStatusBudget As Integer
        OK
        OverBudget
    End Enum

    Public Enum EnumTypeFilter As Integer
        FilterEmployee
        FilterDept
        FilterPosition
    End Enum

    Private Sub LoadFilter(ByVal vPrimaryUser As Double)
        Dim Dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_approval_settingTableAdapter
        Dim Tbl As New DseTrans_SPL.tbltrans_spl_overtime_order_approval_settingDataTable
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_approval_settingRow
        Dim rowD As DataGridViewRow
        Dim Frm As New FrmTrans_SPL

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

        Try
            Me.Cursor = Cursors.WaitCursor

            Me.DgvEmpFilter.Rows.Clear()
            Me.DgvDeptFilter.Rows.Clear()
            Me.DgvPosFilter.Rows.Clear()

            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(Tbl, vPrimaryUser)
            If Tbl.Count > 0 Then
                Me.DgvEmpFilter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvEmpFilter.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.DgvDeptFilter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvDeptFilter.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.DgvPosFilter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvPosFilter.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                Me.DgvEmpFilter.AllowUserToAddRows = False
                Me.DgvDeptFilter.AllowUserToAddRows = False
                Me.DgvPosFilter.AllowUserToAddRows = False
                For Each row In Tbl
                    With row
                        If .Type_Filter = EnumTypeFilter.FilterEmployee Then
                            Me.DgvEmpFilter.Rows.Add()
                            rowD = Me.DgvEmpFilter.Rows(Me.DgvEmpFilter.RowCount - 1)
                            Frm.GetEmployeeData(.Primary_Acuan, EmpCode, EmpNIK, EmpName, PositionPrimary, PositionCode, PositionDesc, DeptPrimary, DeptCode, DeptDesc, ListPrimaryAppGroup, False)
                            rowD.Cells(Me.ColFilterEmp_No.Index).Value = Me.DgvEmpFilter.RowCount
                            rowD.Cells(Me.ColFilterEmp_Code.Index).Value = EmpCode
                            rowD.Cells(Me.ColFilterEmp_NIK.Index).Value = EmpNIK
                            rowD.Cells(Me.ColFilterEmp_Name.Index).Value = EmpName
                            rowD.Cells(Me.ColFilterEmp_JobPosition.Index).Value = PositionCode & " - " & PositionDesc
                            rowD.Cells(Me.ColFilterEmp_Dept.Index).Value = DeptCode & " - " & DeptDesc
                            rowD.Cells(Me.ColFilterEmp_Primary.Index).Value = .Primary_Acuan
                        ElseIf .Type_Filter = EnumTypeFilter.FilterDept Then
                            Me.DgvDeptFilter.Rows.Add()
                            rowD = Me.DgvDeptFilter.Rows(Me.DgvDeptFilter.Rows.Count - 1)
                            rowD.Cells(Me.ColFilterDept_No.Index).Value = Me.DgvDeptFilter.Rows.Count
                            rowD.Cells(Me.ColFilterDept_Code.Index).Value = Me.ObjDept.GetDeptCode(.Primary_Acuan)
                            rowD.Cells(Me.ColFilterDept_Desc.Index).Value = Me.ObjDept.GetDeptDesc(.Primary_Acuan)
                            rowD.Cells(Me.ColFilterDept_Primary.Index).Value = .Primary_Acuan
                        Else
                            Me.DgvPosFilter.Rows.Add()
                            rowD = Me.DgvPosFilter.Rows(Me.DgvPosFilter.Rows.Count - 1)
                            rowD.Cells(Me.ColFilterJobPos_No.Index).Value = Me.DgvPosFilter.Rows.Count
                            rowD.Cells(Me.ColFilterJobPos_Code.Index).Value = Me.ObjJobPost.GetCode(.Primary_Acuan)
                            rowD.Cells(Me.ColFilterJobPos_Desc.Index).Value = Me.ObjJobPost.GetDesc(.Primary_Acuan)
                            rowD.Cells(Me.ColFilterJobPos_Primary.Index).Value = .Primary_Acuan
                        End If
                    End With
                Next

                Me.DgvEmpFilter.AllowUserToAddRows = True
                Me.DgvDeptFilter.AllowUserToAddRows = True
                Me.DgvPosFilter.AllowUserToAddRows = True

                Me.DgvEmpFilter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.DgvEmpFilter.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                Me.DgvDeptFilter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.DgvDeptFilter.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                Me.DgvPosFilter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.DgvPosFilter.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub SaveFilter(ByVal vPrimaryUser As Double)
        Dim Dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_approval_settingTableAdapter
        Dim rowD As DataGridViewRow

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Delete(vPrimaryUser)

            For Each rowD In Me.DgvEmpFilter.Rows
                If rowD.Cells(Me.ColFilterEmp_Primary.Index).Value <> -1 Then
                    Dta.Insert(vPrimaryUser, EnumTypeFilter.FilterEmployee, rowD.Cells(Me.ColFilterEmp_Primary.Index).Value)
                End If
            Next
            For Each rowD In Me.DgvDeptFilter.Rows
                If rowD.Cells(Me.ColFilterDept_Primary.Index).Value <> -1 Then
                    Dta.Insert(vPrimaryUser, EnumTypeFilter.FilterDept, rowD.Cells(Me.ColFilterDept_Primary.Index).Value)
                End If
            Next
            For Each rowD In Me.DgvPosFilter.Rows
                If rowD.Cells(Me.ColFilterJobPos_Primary.Index).Value <> -1 Then
                    Dta.Insert(vPrimaryUser, EnumTypeFilter.FilterPosition, rowD.Cells(Me.ColFilterJobPos_Primary.Index).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahOverBudget(ByVal vPrimarySpl As Double, _
    ByVal vTblSplAppDep2 As DseTrans_SPL.ot_order_approval_2DataTable) As Boolean

        Try
            vTblSplAppDep2.DefaultView.RowFilter = "[Primary Spl] = " & vPrimarySpl & _
            " AND [Status Budget] = " & EnumStatusBudget.OverBudget
            If vTblSplAppDep2.DefaultView.Count > 0 Then
                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsiDgv_SPL(ByVal vStatusFilter As ClsTrans_SPL.EnumStatus)
        Dim dta As New DseTrans_SPLTableAdapters.tbltrans_spl_approvalTableAdapter
        Dim tbl, tbl_temp As New DseTrans_SPL.tbltrans_spl_approvalDataTable
        Dim row As DseTrans_SPL.tbltrans_spl_approvalRow
        Dim rowD As DataGridViewRow
        Dim DateFilterFrom As Date = CDate("1/1/1950")
        Dim DateFilterTo As Date = CDate("1/1/2200")
        Dim StatusReqApp As Integer = 1

        Try
            Me.Cursor = Cursors.WaitCursor

            If Me.ChkDocDate.Checked = True Then
                DateFilterFrom = Me.DtpDocDateFrom.Value.Date
                DateFilterTo = Me.DtpDocDateTo.Value.Date
            End If

            Me.TblSplApp.Clear()
            dta.Connection.ConnectionString = Core_Module.GetConstring
            'untuk Me.TblSplApp = hanya menampilkan 1 spl = 1 baris (district per spl)
            If vStatusFilter <> ClsTrans_SPL.EnumStatus.Rejected Then
                dta.ClearBeforeFill = False
                dta.Fill_SplGroup(Me.TblSplApp, ClsTrans_SPL.EnumStatus.WaitingForApproval)
                dta.Fill_SplGroup(Me.TblSplApp, ClsTrans_SPL.EnumStatus.Approved)
                dta.ClearBeforeFill = True
            Else
                dta.Fill_SplGroup(Me.TblSplApp, vStatusFilter)
            End If

            'untuk tbl_temp = 1 spl tampil lebih dari 1 baris (terkait untuk filter per user)
            dta.Fill(tbl_temp, vStatusFilter)
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_SPLApproval_ShowAll") = False Then
                If vStatusFilter = ClsTrans_SPL.EnumStatus.WaitingForApproval Then
                    tbl_temp.DefaultView.RowFilter = "[Date] >= #" & DateFilterFrom & _
                    "# AND [Date] <= #" & DateFilterTo & "# AND [Req Approval] = " & StatusReqApp & " AND [Primary User] = " & Me.PrimaryUser
                Else
                    tbl_temp.DefaultView.RowFilter = "[Date] >= #" & DateFilterFrom & _
                    "# AND [Date] <= #" & DateFilterTo & "# AND [Req Approval] = " & StatusReqApp & " AND [Primary User] = " & Me.PrimaryUser & _
                    " AND [Primary Approved By] = " & Me.PrimaryUser
                End If
            Else
                tbl_temp.DefaultView.RowFilter = "[Date] >= #" & DateFilterFrom & _
                "# AND [Date] <= #" & DateFilterTo & "# AND [Req Approval] = " & StatusReqApp
            End If

            tbl.Merge(tbl_temp.DefaultView.ToTable)

            'filter berdasarkan employee, job position, department
            Me.FilterDgv(tbl, DateFilterFrom, DateFilterTo, vStatusFilter)

            If vStatusFilter = ClsTrans_SPL.EnumStatus.WaitingForApproval Then
                Me.ColItem_Budget.Visible = True
                Me.IsiTblBudgetPerDepartment(tbl, Me.TblSplApp, Me.TblSplUsed, Me.TblSplAppDep2, DateFilterFrom, DateFilterTo, True)
            Else
                Me.ColItem_Budget.Visible = False
            End If

            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each row In tbl
                Me.dgvItem.Rows.Add()
                rowD = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                With row
                    rowD.Cells(Me.ColItem_DocumentNo.Index).Value = .No_Transaksi
                    rowD.Cells(Me.ColItem_Date.Index).Value = Format(._Date, "dd MMM yyyy")
                    rowD.Cells(Me.ColItem_StartDate.Index).Value = Format(.From, "dd MMM yyyy")
                    rowD.Cells(Me.ColItem_EndDate.Index).Value = Format(._To, "dd MMM yyyy")
                    If .Status = ClsTrans_SPL.EnumStatus.WaitingForApproval Then
                        rowD.Cells(Me.ColItem_Status.Index).Value = "WaitingForApproval"
                    ElseIf .Status = ClsTrans_SPL.EnumStatus.Approved Then
                        rowD.Cells(Me.ColItem_Status.Index).Value = "Approved"
                    Else
                        rowD.Cells(Me.ColItem_Status.Index).Value = "Rejected"
                    End If
                    rowD.Cells(Me.ColItem_Primary.Index).Value = .Primary
                    rowD.Cells(Me.ColItem_Modifiy.Index).Value = False
                    If Me.ColItem_Budget.Visible = True Then
                        If Me.ApakahOverBudget(.Primary, Me.TblSplAppDep2) = True Then
                            rowD.Cells(Me.ColItem_Budget.Index).Value = "Over"
                        Else
                            rowD.Cells(Me.ColItem_Budget.Index).Value = "OK"
                        End If
                    End If
                End With
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            If Me.dgvItem.RowCount > 0 Then
                Me.IsiField(Me.dgvItem.Item(Me.ColItem_Primary.Index, 0).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Public Sub IsiTblBudgetPerDepartment(ByVal tbl As DseTrans_SPL.tbltrans_spl_approvalDataTable, _
    ByRef vTblSplApp As DseTrans_SPL.tbltrans_spl_approvalDataTable, _
    ByRef vTblSplUsed As DseTrans_SPL.ot_order_approvalDataTable, _
    ByRef vTblSplAppDep2 As DseTrans_SPL.ot_order_approval_2DataTable, _
    ByVal PeriodFrom As Date, ByVal PeriodTo As Date, ByVal SetPeriod As Boolean)
        Dim DtaSplD As New DseTrans_SPLTableAdapters.tbltrans_spl_app_ot_detailTableAdapter
        Dim TblSplD_App, TblSplD_Wait As New DseTrans_SPL.tbltrans_spl_app_ot_detailDataTable
        Dim DtaBud As New DseTrans_SPLTableAdapters.tbl_budgetTableAdapter
        Dim TblBud As New DseTrans_SPL.tbl_budgetDataTable
        Dim DtaAtt As New DseTrans_SPLTableAdapters.tbl_attendance_otTableAdapter
        Dim TblAtt As New DseTrans_SPL.tbl_attendance_otDataTable
        Dim rowAtt As DseTrans_SPL.tbl_attendance_otRow
        Dim DtaAttDept As New DseTrans_SPLTableAdapters.tbl_attendance_dept_otTableAdapter
        Dim TblAttDept As New DseTrans_SPL.tbl_attendance_dept_otDataTable
        Dim rowAttDept As DseTrans_SPL.tbl_attendance_dept_otRow
        Dim row As DseTrans_SPL.tbltrans_spl_approvalRow
        Dim PeriodFrom_2 As Date = Now
        Dim PeriodTo_2 As Date = Now

        Try
            vTblSplUsed.Clear()
            vTblSplAppDep2.Clear()

            If tbl.Count = 0 Then
                Exit Sub
            Else
                If tbl.Count > 0 And SetPeriod = True Then
                    tbl.DefaultView.Sort = "[From] ASC"
                    row = tbl.DefaultView(0).Row
                    PeriodFrom = row.From
                    tbl.DefaultView.Sort = "[To] Desc"
                    row = tbl.DefaultView(0).Row
                    PeriodTo = row._To
                End If

                DtaSplD.Connection.ConnectionString = Core_Module.GetConstring
                DtaBud.Connection.ConnectionString = Core_Module.GetConstring
                DtaAtt.Connection.ConnectionString = Core_Module.GetConstring
                DtaAttDept.Connection.ConnectionString = Core_Module.GetConstring

                'fill overtime order yang masih waiting & approved
                PeriodFrom = Me.ObjCD.GetFirstDateofTheMonth(PeriodFrom)
                PeriodTo = Me.ObjCD.GetLasDateOfTheMonth(PeriodTo)
                DtaSplD.Fill(TblSplD_App, PeriodFrom, PeriodTo, ClsTrans_SPL.EnumStatus.Approved)
                DtaSplD.Fill(TblSplD_Wait, PeriodFrom, PeriodTo, ClsTrans_SPL.EnumStatus.WaitingForApproval)

                'fill data attendance item, untuk mendapat actual early ot & ot
                DtaAtt.Fill(TblAtt, PeriodFrom, PeriodTo)
                DtaAttDept.ClearBeforeFill = False
                Dim PeriodFilterAttDept As Date = Now
                For i As Integer = PeriodFrom.Month To PeriodTo.Month
                    PeriodFilterAttDept = CDate(i & "/1/" & PeriodFrom.Year)
                    DtaAttDept.Fill(TblAttDept, PeriodFilterAttDept, Me.ObjCD.GetLasDateOfTheMonth(PeriodFilterAttDept))
                    TblAttDept.DefaultView.RowFilter = "[date] >= #" & PeriodFilterAttDept & "# AND [date] <= #" & Me.ObjCD.GetLasDateOfTheMonth(PeriodFilterAttDept) & "#"
                    For Each dv As DataRowView In TblAttDept.DefaultView
                        rowAttDept = dv.Row
                        rowAttDept._date = PeriodFilterAttDept
                    Next
                Next

                'fill data overtime budget
                PeriodFrom_2 = Me.ObjCD.GetFisrtDateOfTheYear(PeriodFrom)
                PeriodTo_2 = Me.ObjCD.GetLastDateOfTheYear(PeriodTo)
                DtaBud.Fill(TblBud, PeriodFrom_2, PeriodTo_2)

                'dapatkan budget nya
                Me.IsiBudget(TblBud, vTblSplUsed, PeriodFrom_2, PeriodTo_2)

                ''dapatkan used nya 
                'Me.IsiUsed(TblSplD_App, vTblSplUsed, PeriodFrom, PeriodTo)

                'dapatkan detail balance perdepartment per spl
                Me.IsiSplBalanceDetail(TblSplD_App, TblSplD_Wait, vTblSplAppDep2, vTblSplApp, PeriodFrom, PeriodTo, TblAtt, TblAttDept)

                'dapatkan status per department per spl
                Me.IsiSplStatus(vTblSplUsed, vTblSplAppDep2)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiBudget(ByVal TblBud As DseTrans_SPL.tbl_budgetDataTable, _
    ByRef vTblSplUsed As DseTrans_SPL.ot_order_approvalDataTable, _
    ByVal PeriodFrom_2 As Date, ByVal PeriodTo_2 As Date)
        Dim rowBud As DseTrans_SPL.tbl_budgetRow
        Dim rowSplUsed As DseTrans_SPL.ot_order_approvalRow

        Dim Dta As New DseTrans_SPLTableAdapters.tblmaster_departemenTableAdapter
        Dim Tbl As New DseTrans_SPL.tblmaster_departemenDataTable
        Dim row As DseTrans_SPL.tblmaster_departemenRow

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(Tbl)

            For i As Integer = PeriodFrom_2.Year To PeriodTo_2.Year
                For Each row In Tbl
                    TblBud.DefaultView.RowFilter = "[Primary Department] = " & row.Primary.ToString & " AND [Period] = #" & CDate("1/1/" & i) & "#"
                    If TblBud.DefaultView.Count = 0 Then
                        rowBud = TblBud.Newtbl_budgetRow
                        rowBud.Primary_Department = row.Primary
                        rowBud.Period = CDate("1/1/" & i)
                        TblBud.Rows.Add(rowBud)
                    End If
                Next
            Next

            For Each rowBud In TblBud
                vTblSplUsed.DefaultView.RowFilter = "[Primary Dept] = " & rowBud.Primary_Department & " AND [Primary Budget] = " & rowBud.Primary
                If vTblSplUsed.DefaultView.Count = 0 Then
                    For i As Integer = 1 To 12
                        rowSplUsed = vTblSplUsed.Newot_order_approvalRow
                        rowSplUsed.Primary_Dept = rowBud.Primary_Department
                        rowSplUsed.Primary_Budget = rowBud.Primary
                        rowSplUsed.Period = CDate(i & "/1/" & rowBud.Period.Year)
                        If i = 1 Then
                            rowSplUsed.Budget = rowBud.Jan
                            rowSplUsed.Budget_Hitung = rowBud.Jan
                        ElseIf i = 2 Then
                            rowSplUsed.Budget = rowBud.Feb
                            rowSplUsed.Budget_Hitung = rowBud.Feb
                        ElseIf i = 3 Then
                            rowSplUsed.Budget = rowBud.Mar
                            rowSplUsed.Budget_Hitung = rowBud.Mar
                        ElseIf i = 4 Then
                            rowSplUsed.Budget = rowBud.Apr
                            rowSplUsed.Budget_Hitung = rowBud.Apr
                        ElseIf i = 5 Then
                            rowSplUsed.Budget = rowBud.May
                            rowSplUsed.Budget_Hitung = rowBud.May
                        ElseIf i = 6 Then
                            rowSplUsed.Budget = rowBud.Jun
                            rowSplUsed.Budget_Hitung = rowBud.Jun
                        ElseIf i = 7 Then
                            rowSplUsed.Budget = rowBud.Jul
                            rowSplUsed.Budget_Hitung = rowBud.Jul
                        ElseIf i = 8 Then
                            rowSplUsed.Budget = rowBud.Aug
                            rowSplUsed.Budget_Hitung = rowBud.Aug
                        ElseIf i = 9 Then
                            rowSplUsed.Budget = rowBud.Sep
                            rowSplUsed.Budget_Hitung = rowBud.Sep
                        ElseIf i = 10 Then
                            rowSplUsed.Budget = rowBud.Oct
                            rowSplUsed.Budget_Hitung = rowBud.Oct
                        ElseIf i = 11 Then
                            rowSplUsed.Budget = rowBud.Nov
                            rowSplUsed.Budget_Hitung = rowBud.Nov
                        ElseIf i = 12 Then
                            rowSplUsed.Budget = rowBud.Dec
                            rowSplUsed.Budget_Hitung = rowBud.Dec
                        End If
                        vTblSplUsed.Rows.Add(rowSplUsed)
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiUsed(ByVal TblSplD_App As DseTrans_SPL.tbltrans_spl_app_ot_detailDataTable, _
    ByRef vTblSplUsed As DseTrans_SPL.ot_order_approvalDataTable, _
    ByVal vDateFrom As Date, ByVal vDateTo As Date)
        Dim rowSplD_App As DseTrans_SPL.tbltrans_spl_app_ot_detailRow
        Dim rowSplUsed As DseTrans_SPL.ot_order_approvalRow
        Dim FromDate_ As Date = Now
        Dim ToDate_ As Date = Now
        Dim PeriodFilter As Date = Now

        Try
            'For Each rowSplD_App In TblSplD_App
            '    PeriodFilter = rowSplD_App._Date.Month & "/1/" & rowSplD_App._Date.Year
            '    vTblSplUsed.DefaultView.RowFilter = "[Period] = #" & PeriodFilter & "# AND [Primary Dept] = " & rowSplD_App.primary_departemen
            '    If vTblSplUsed.DefaultView.Count > 0 Then
            '        rowSplUsed = vTblSplUsed.DefaultView(0).Row
            '        If rowSplD_App.Primary_Attendance <> -1 Then
            '            rowSplUsed.Used += rowSplD_App.Total_OT_Actual
            '        Else
            '            rowSplUsed.Used += rowSplD_App.Total_OT_Approved
            '        End If
            '    Else
            '        rowSplUsed = vTblSplUsed.Newot_order_approvalRow
            '        rowSplUsed.Primary_Dept = rowSplD_App.primary_departemen
            '        rowSplUsed.Period = PeriodFilter
            '        rowSplUsed.Budget = 0
            '        If rowSplD_App.Primary_Attendance <> -1 Then
            '            rowSplUsed.Used += rowSplD_App.Total_OT_Actual
            '        Else
            '            rowSplUsed.Used += rowSplD_App.Total_OT_Approved
            '        End If
            '        vTblSplUsed.Rows.Add(rowSplUsed)
            '    End If
            'Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetOTAttendance(ByVal TblAtt As DseTrans_SPL.tbl_attendance_otDataTable, _
    ByVal vPrimaryEmp As Double, ByVal vDate As Date, ByRef vApakahAdaAttendance As Boolean) As Double
        Dim rowAtt As DseTrans_SPL.tbl_attendance_otRow
        Dim OT_Attendance As Double = 0
        vApakahAdaAttendance = False
        Try
            TblAtt.DefaultView.RowFilter = "[primary karyawan] = " & vPrimaryEmp.ToString & " AND [date] = #" & vDate.Date & "#"
            If TblAtt.DefaultView.Count > 0 Then
                rowAtt = TblAtt.DefaultView(0).Row
                vApakahAdaAttendance = True
                If rowAtt.Primary_Overtime_Order <> -1 Then
                    OT_Attendance = rowAtt.Total_OT_Attendance
                End If
            End If

            Return OT_Attendance
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub IsiSplBalanceDetail(ByVal TblSplD_App As DseTrans_SPL.tbltrans_spl_app_ot_detailDataTable, _
    ByVal TblSplD_Wait As DseTrans_SPL.tbltrans_spl_app_ot_detailDataTable, _
    ByRef vTblSplAppDep2 As DseTrans_SPL.ot_order_approval_2DataTable, _
    ByRef vTblSplApp As DseTrans_SPL.tbltrans_spl_approvalDataTable, _
    ByVal vDateFrom As Date, ByVal vDateTo As Date, _
    ByVal vTblAtt As DseTrans_SPL.tbl_attendance_otDataTable, _
    ByVal vTblAttDept As DseTrans_SPL.tbl_attendance_dept_otDataTable)
        Dim rowSplAppDep2 As DseTrans_SPL.ot_order_approval_2Row
        Dim rowSplD_App, rowSplD_Wait As DseTrans_SPL.tbltrans_spl_app_ot_detailRow
        Dim row As DseTrans_SPL.tbltrans_spl_approvalRow
        Dim rowAttDept As DseTrans_SPL.tbl_attendance_dept_otRow
        Dim PeriodFilter As Date = Now
        Dim OT_Attendance As Double = 0
        Dim ApakahAdaAttendance As Boolean = False
        Dim OT_ProposedORUsed As Double = 0

        Try
            vTblSplApp.DefaultView.RowFilter = "[to] >= #" & vDateFrom.Date & "# AND [from] <= #" & vDateTo.Date _
            & "# AND [Status] <> " & ClsTrans_SPL.EnumStatus.Rejected
            vTblSplApp.DefaultView.Sort = "[Date] ASC, [Urutan] ASC"
            'TblSplApp = tbl header spl
            For Each dv As DataRowView In vTblSplApp.DefaultView
                row = dv.Row

                If row.Status = ClsTrans_SPL.EnumStatus.Approved Then
                    TblSplD_App.DefaultView.RowFilter = "[Primary Transaksi] = " & row.Primary
                    For Each dv1 As DataRowView In TblSplD_App.DefaultView
                        rowSplD_App = dv1.Row

                        OT_Attendance = Me.GetOTAttendance(vTblAtt, rowSplD_App.Primary_Employee, rowSplD_App._Date, ApakahAdaAttendance)

                        PeriodFilter = rowSplD_App._Date.Month & "/1/" & rowSplD_App._Date.Year
                        vTblSplAppDep2.DefaultView.RowFilter = "[Period] = #" & PeriodFilter & "# AND [Primary Spl] = " & _
                        rowSplD_App.Primary_Transaksi & " AND [Primary Dept] = " & rowSplD_App.primary_departemen
                        If vTblSplAppDep2.DefaultView.Count > 0 Then
                            rowSplAppDep2 = vTblSplAppDep2.DefaultView(0).Row
                            rowSplAppDep2.Proposed += rowSplD_App.Total_OT_Approved
                            If ApakahAdaAttendance = False Then
                                rowSplAppDep2.Used += rowSplD_App.Total_OT_Approved
                            Else
                                rowSplAppDep2.Used += OT_Attendance
                            End If
                            rowSplAppDep2.Balance = 0
                            'OT_ProposedORUsed = rowSplAppDep2.Used
                        Else
                            rowSplAppDep2 = vTblSplAppDep2.Newot_order_approval_2Row
                            rowSplAppDep2.Primary_Spl = rowSplD_App.Primary_Transaksi
                            rowSplAppDep2.Primary_Dept = rowSplD_App.primary_departemen
                            rowSplAppDep2.Department = rowSplD_App.Dept_Desc
                            rowSplAppDep2.Period = PeriodFilter
                            rowSplAppDep2.Spl_No = row.No_Transaksi
                            rowSplAppDep2.Spl_Date = row._Date
                            rowSplAppDep2.Proposed = rowSplD_App.Total_OT_Approved
                            If ApakahAdaAttendance = False Then
                                rowSplAppDep2.Used = rowSplD_App.Total_OT_Approved
                            Else
                                rowSplAppDep2.Used = OT_Attendance
                            End If
                            'OT_ProposedORUsed = rowSplAppDep2.Used
                            rowSplAppDep2.Balance = 0
                            rowSplAppDep2.Status_Budget = EnumStatusBudget.OK
                            rowSplAppDep2.Created_Date = row.Created_Date
                            rowSplAppDep2.Urutan = row.Urutan
                            rowSplAppDep2.Apakah_Approved = 1
                            vTblSplAppDep2.Rows.Add(rowSplAppDep2)
                        End If

                        'Attendace Total OT per Department
                        vTblAttDept.DefaultView.RowFilter = "[date] = #" & PeriodFilter & "# AND [Primary] = " & rowSplD_App.primary_departemen & " AND [Sudah Add Row] = 0"
                        If vTblAttDept.DefaultView.Count > 0 Then
                            rowAttDept = vTblAttDept.DefaultView(0).Row
                            'rowAttDept.Total_OT_Att -= OT_ProposedORUsed
                            If rowAttDept.Sudah_Add_Row = 0 Then
                                rowSplAppDep2 = vTblSplAppDep2.Newot_order_approval_2Row
                                rowSplAppDep2.Primary_Spl = -1
                                rowSplAppDep2.Primary_Dept = rowSplD_App.primary_departemen
                                rowSplAppDep2.Department = rowSplD_App.Dept_Desc
                                rowSplAppDep2.Period = PeriodFilter
                                rowSplAppDep2.Spl_No = "Attendance"
                                rowSplAppDep2.Spl_Date = PeriodFilter
                                rowSplAppDep2.Proposed = 0
                                rowSplAppDep2.Used = rowAttDept.Total_OT_Att
                                rowSplAppDep2.Balance = 0
                                rowSplAppDep2.Status_Budget = EnumStatusBudget.OK
                                rowSplAppDep2.Created_Date = CDate("1/1/1950")
                                rowSplAppDep2.Urutan = -1
                                rowSplAppDep2.Apakah_Approved = 2
                                vTblSplAppDep2.Rows.Add(rowSplAppDep2)

                                rowAttDept.Sudah_Add_Row = 1
                            End If
                        End If
                    Next
                ElseIf row.Status = ClsTrans_SPL.EnumStatus.WaitingForApproval Then
                    'TblSplD_Wait = tbl item approval ot dengan status waiting for approval
                    TblSplD_Wait.DefaultView.RowFilter = "[Primary Transaksi] = " & row.Primary
                    For Each dv1 As DataRowView In TblSplD_Wait.DefaultView
                        rowSplD_Wait = dv1.Row

                        PeriodFilter = rowSplD_Wait._Date.Month & "/1/" & rowSplD_Wait._Date.Year
                        vTblSplAppDep2.DefaultView.RowFilter = "[Period] = #" & PeriodFilter & "# AND [Primary Spl] = " & _
                        rowSplD_Wait.Primary_Transaksi & " AND [Primary Dept] = " & rowSplD_Wait.primary_departemen
                        If vTblSplAppDep2.DefaultView.Count > 0 Then
                            rowSplAppDep2 = vTblSplAppDep2.DefaultView(0).Row
                            rowSplAppDep2.Proposed += rowSplD_Wait.Total_OT_Approved
                            'OT_ProposedORUsed = rowSplAppDep2.Used
                            rowSplAppDep2.Balance = 0
                        Else
                            rowSplAppDep2 = vTblSplAppDep2.Newot_order_approval_2Row
                            rowSplAppDep2.Primary_Spl = rowSplD_Wait.Primary_Transaksi
                            rowSplAppDep2.Primary_Dept = rowSplD_Wait.primary_departemen
                            rowSplAppDep2.Department = rowSplD_Wait.Dept_Desc
                            rowSplAppDep2.Period = PeriodFilter
                            rowSplAppDep2.Spl_No = row.No_Transaksi
                            rowSplAppDep2.Spl_Date = row._Date
                            rowSplAppDep2.Proposed = rowSplD_Wait.Total_OT_Approved
                            'OT_ProposedORUsed = rowSplAppDep2.Used
                            rowSplAppDep2.Balance = 0
                            rowSplAppDep2.Status_Budget = EnumStatusBudget.OK
                            rowSplAppDep2.Created_Date = row.Created_Date
                            rowSplAppDep2.Urutan = row.Urutan
                            vTblSplAppDep2.Rows.Add(rowSplAppDep2)
                        End If

                        'Attendace Total OT per Department
                        vTblAttDept.DefaultView.RowFilter = "[date] = #" & PeriodFilter & "# AND [Primary] = " & rowSplD_Wait.primary_departemen & " AND [Sudah Add Row] = 0"
                        If vTblAttDept.DefaultView.Count > 0 Then
                            rowAttDept = vTblAttDept.DefaultView(0).Row
                            'rowAttDept.Total_OT_Att -= OT_ProposedORUsed
                            If rowAttDept.Sudah_Add_Row = 0 Then
                                rowSplAppDep2 = vTblSplAppDep2.Newot_order_approval_2Row
                                rowSplAppDep2.Primary_Spl = -1
                                rowSplAppDep2.Primary_Dept = rowSplD_Wait.primary_departemen
                                rowSplAppDep2.Department = rowSplD_Wait.Dept_Desc
                                rowSplAppDep2.Period = PeriodFilter
                                rowSplAppDep2.Spl_No = "Attendance"
                                rowSplAppDep2.Spl_Date = PeriodFilter
                                rowSplAppDep2.Proposed = 0
                                rowSplAppDep2.Used = rowAttDept.Total_OT_Att
                                rowSplAppDep2.Balance = 0
                                rowSplAppDep2.Status_Budget = EnumStatusBudget.OK
                                rowSplAppDep2.Created_Date = CDate("1/1/1950")
                                rowSplAppDep2.Urutan = -1
                                rowSplAppDep2.Apakah_Approved = 2
                                vTblSplAppDep2.Rows.Add(rowSplAppDep2)

                                rowAttDept.Sudah_Add_Row = 1
                            End If
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiSplStatus(ByRef vTblSplUsed As DseTrans_SPL.ot_order_approvalDataTable, _
    ByRef vTblSplAppDep2 As DseTrans_SPL.ot_order_approval_2DataTable)
        Dim rowSplUsed As DseTrans_SPL.ot_order_approvalRow
        Dim rowSplAppDep2 As DseTrans_SPL.ot_order_approval_2Row
        Dim SisaBalance As Double = 0

        Try
            vTblSplAppDep2.DefaultView.RowFilter = Nothing
            vTblSplAppDep2.DefaultView.Sort = "[Apakah Approved] DESC, [Spl Date] ASC, [Urutan] ASC"
            For Each dv As DataRowView In vTblSplAppDep2.DefaultView
                rowSplAppDep2 = dv.Row
                SisaBalance = 0

                vTblSplUsed.DefaultView.RowFilter = "[Period] = #" & rowSplAppDep2.Period & "# AND [Primary Dept] = " & rowSplAppDep2.Primary_Dept
                If vTblSplUsed.DefaultView.Count > 0 Then
                    rowSplUsed = vTblSplUsed.DefaultView(0).Row
                    If rowSplAppDep2.Apakah_Approved = 0 Then
                        rowSplUsed.Budget_Hitung -= rowSplAppDep2.Proposed
                    Else
                        rowSplUsed.Budget_Hitung -= rowSplAppDep2.Used
                    End If
                    SisaBalance = rowSplUsed.Budget_Hitung
                Else
                    If rowSplAppDep2.Apakah_Approved = 0 Then
                        SisaBalance -= rowSplAppDep2.Proposed
                    Else
                        SisaBalance -= rowSplAppDep2.Used
                    End If
                End If

                rowSplAppDep2.Balance = SisaBalance
                If SisaBalance < 0 Then
                    rowSplAppDep2.Status_Budget = EnumStatusBudget.OverBudget
                Else
                    rowSplAppDep2.Status_Budget = EnumStatusBudget.OK
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterDgv(ByRef tbl As DseTrans_SPL.tbltrans_spl_approvalDataTable, ByVal DateFilterFrom As DateTime, _
    ByVal DateFilterTo As DateTime, ByVal vStatusSpl As ClsTrans_SPL.EnumStatus)
        Dim Dta1 As New DseTrans_SPLTableAdapters.tbltrans_spl_approval_employeeTableAdapter
        Dim Tbl1 As New DseTrans_SPL.tbltrans_spl_approval_employeeDataTable
        Dim row1 As DseTrans_SPL.tbltrans_spl_approval_employeeRow
        Dim row As DseTrans_SPL.tbltrans_spl_approvalRow
        Dim Filter As String = ""
        Dim ListPrimarySPL As New List(Of Double)

        Try
            'ISI FILTER BERDASARKAN DGV EMPLOYEE, JOB POSITION, DEPARTMENT
            Dta1.Connection.ConnectionString = Core_Module.GetConstring
            Dta1.Fill(Tbl1, DateFilterFrom, DateFilterTo, vStatusSpl)
            Tbl1.DefaultView.RowFilter = Me.Filtering()
            If Tbl1.DefaultView.Count > 0 Then
                For Each dv As DataRowView In Tbl1.DefaultView
                    row1 = dv.Row
                    If ListPrimarySPL.Contains(row1.Primary) = False Then
                        ListPrimarySPL.Add(row1.Primary)
                    End If
                Next
            End If

            If ListPrimarySPL.Count = 0 Then
                tbl.Clear()
            Else
                For i As Integer = tbl.Count - 1 To 0 Step -1
                    row = tbl.Rows(i)
                    If ListPrimarySPL.Contains(row.Primary) = False Then
                        tbl.Rows.Remove(row)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Filtering() As String
        Dim FilterAll As String = ""
        Dim FilterEmployee As String = ""
        Dim FilterDepartemen As String = ""
        Dim FilterPosisi As String = ""

        Try
            For Each dr As DataGridViewRow In Me.DgvEmpFilter.Rows
                With dr
                    If dr.Cells(Me.ColFilterEmp_Primary.Index).Value <> -1 And String.IsNullOrEmpty(dr.Cells(Me.ColFilterEmp_Primary.Index).Value) = False Then
                        If FilterEmployee = "" Then
                            FilterEmployee = "[Primary Employee] = " & .Cells(Me.ColFilterEmp_Primary.Index).Value.ToString
                        Else
                            FilterEmployee &= " OR [Primary Employee] = " & .Cells(Me.ColFilterEmp_Primary.Index).Value.ToString
                        End If
                    End If
                End With
            Next
            For Each dr As DataGridViewRow In Me.DgvDeptFilter.Rows
                With dr
                    If dr.Cells(Me.ColFilterDept_Primary.Index).Value <> -1 And String.IsNullOrEmpty(dr.Cells(Me.ColFilterDept_Primary.Index).Value) = False Then
                        If FilterDepartemen = "" Then
                            FilterDepartemen = "[primary departemen] = " & .Cells(Me.ColFilterDept_Primary.Index).Value.ToString
                        Else
                            FilterDepartemen &= " OR [primary departemen] = " & .Cells(Me.ColFilterDept_Primary.Index).Value.ToString
                        End If
                    End If
                End With
            Next
            For Each dr As DataGridViewRow In Me.DgvPosFilter.Rows
                With dr
                    If dr.Cells(Me.ColFilterJobPos_Primary.Index).Value <> -1 And String.IsNullOrEmpty(dr.Cells(Me.ColFilterJobPos_Primary.Index).Value) = False Then
                        If FilterPosisi = "" Then
                            FilterPosisi = "[primary position] = " & .Cells(Me.ColFilterJobPos_Primary.Index).Value.ToString
                        Else
                            FilterPosisi &= " OR [primary position] = " & .Cells(Me.ColFilterJobPos_Primary.Index).Value.ToString
                        End If
                    End If
                End With
            Next

            If FilterEmployee <> "" Then
                If FilterAll <> "" Then
                    FilterAll &= "AND (" & FilterEmployee & ")"
                Else
                    FilterAll &= "(" & FilterEmployee & ")"
                End If
            End If
            If FilterDepartemen <> "" Then
                If FilterAll <> "" Then
                    FilterAll &= "AND (" & FilterDepartemen & ")"
                Else
                    FilterAll &= "(" & FilterDepartemen & ")"
                End If
            End If
            If FilterPosisi <> "" Then
                If FilterAll <> "" Then
                    FilterAll &= "AND (" & FilterPosisi & ")"
                Else
                    FilterAll &= "(" & FilterPosisi & ")"
                End If
            End If

            Return FilterAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub FrmTrans_SPLApproval_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            Me.DapatAppOverBudget = Me.ObjSpl.ApproveOverBudget(Me.PrimaryUser)

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_SPLApproval_ShowAll") = False Then
                Me.BisaApproved = False
            Else
                Me.BisaApproved = True
            End If

            If Me.RboWaitingForApproval.Checked = True Then
                Me.IsiDgv_SPL(ClsTrans_SPL.EnumStatus.WaitingForApproval)
            ElseIf Me.RboApproved.Checked = True Then
                Me.IsiDgv_SPL(ClsTrans_SPL.EnumStatus.Approved)
            Else
                Me.IsiDgv_SPL(ClsTrans_SPL.EnumStatus.Rejected)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SPLApproval_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.SaveFilter(Me.PrimaryUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SPLApproval_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.PrimaryUser = Me.ObjUser.GetPrimary(Me.UserCode)

            Me.DapatAppOverBudget = Me.ObjSpl.ApproveOverBudget(Me.PrimaryUser)
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_SPLApproval_ShowAll") = False Then
                Me.BisaApproved = False
            Else
                Me.BisaApproved = True
            End If

            Me.LoadFilter(Me.PrimaryUser)

            Me.colEmpl_PrimaryDept.Visible = False
            Me.colEmpl_PrimaryPos.Visible = False
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSchDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSchEmp.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvEmpFilter.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvDeptFilter.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvPosFilter.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.dgvEmployee.ReadOnly = True
            Me.dgvSchDetail.ReadOnly = True
            Me.dgvSchEmp.ReadOnly = True
            Me.dgvSchEmp.MultiSelect = False

            Me.dgvEmployee.AllowUserToDeleteRows = False
            Me.dgvSchDetail.AllowUserToDeleteRows = False
            Me.dgvSchEmp.AllowUserToDeleteRows = False

            Me.RboWaitingForApproval.Checked = True
            Me.ChkDocDate.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboWaitingForApproval_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboWaitingForApproval.CheckedChanged
        Try
            If Me.Visible = True And Me.RboWaitingForApproval.Checked = True Then
                Me.EnumFillSPL = ClsTrans_SPL.EnumStatus.WaitingForApproval
                Me.IsiDgv_SPL(Me.EnumFillSPL)
            End If

            Me.BtnWait.Enabled = False
            Me.BtnApprove.Enabled = True
            Me.BtnReject.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboApproved_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboApproved.CheckedChanged
        Try
            If Me.Visible = True And Me.RboApproved.Checked = True Then
                Me.EnumFillSPL = ClsTrans_SPL.EnumStatus.Approved
                Me.IsiDgv_SPL(Me.EnumFillSPL)
            End If

            Me.BtnWait.Enabled = True
            Me.BtnApprove.Enabled = False
            Me.BtnReject.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboRejected_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboRejected.CheckedChanged
        Try
            If Me.Visible = True And Me.RboRejected.Checked = True Then
                Me.EnumFillSPL = ClsTrans_SPL.EnumStatus.Rejected
                Me.IsiDgv_SPL(Me.EnumFillSPL)
            End If

            Me.BtnWait.Enabled = True
            Me.BtnApprove.Enabled = True
            Me.BtnReject.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnApprove.Click
        Dim Hasil As MsgBoxResult
        Dim ListPrimarySPL As New List(Of Double)

        Try
            If Me.ObjSpl.ApakahDataBenar_Approval(ListPrimarySPL, Me.dgvItem, Me.EnumFillSPL, ClsTrans_SPL.EnumStatus.Approved, Me.DapatAppOverBudget, Me.UserCode) = True Then
                Hasil = MsgBox("Approve ?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    Dim ij As New Core_Module.DialogPassword
                    ij.txtUserName.Text = Me.UserCode
                    ij.ShowDialog()

                    If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                        For Each row As DataGridViewRow In Me.dgvItem.Rows
                            If row.Cells(Me.ColItem_Modifiy.Index).Value = True And ListPrimarySPL.Contains(row.Cells(Me.ColItem_Primary.Index).Value) = True Then
                                Me.ObjSpl.UpdateStatusApproval(Me.PrimaryUser, ClsTrans_SPL.EnumStatus.Approved, row.Cells(Me.ColItem_Primary.Index).Value)
                            End If
                        Next

                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Me.KosongkanField()
                        Me.IsiDgv_SPL(Me.EnumFillSPL)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnReject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnReject.Click
        Dim Hasil As MsgBoxResult
        Dim ListPrimarySPL As New List(Of Double)

        Try
            If Me.ObjSpl.ApakahDataBenar_Approval(ListPrimarySPL, Me.dgvItem, Me.EnumFillSPL, ClsTrans_SPL.EnumStatus.Rejected, Me.DapatAppOverBudget, Me.UserCode) = True Then
                Hasil = MsgBox("Reject ?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    Dim ij As New Core_Module.DialogPassword
                    ij.txtUserName.Text = Me.UserCode
                    ij.ShowDialog()

                    If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                        For Each row As DataGridViewRow In Me.dgvItem.Rows
                            If row.Cells(Me.ColItem_Modifiy.Index).Value = True And ListPrimarySPL.Contains(row.Cells(Me.ColItem_Primary.Index).Value) = True Then
                                Me.ObjSpl.UpdateStatusApproval(Me.PrimaryUser, ClsTrans_SPL.EnumStatus.Rejected, row.Cells(Me.ColItem_Primary.Index).Value)
                            End If
                        Next

                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Me.KosongkanField()
                        Me.IsiDgv_SPL(Me.EnumFillSPL)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnWait_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnWait.Click
        Dim Hasil As MsgBoxResult
        Dim ListPrimarySPL As New List(Of Double)

        Try
            If Me.ObjSpl.ApakahDataBenar_Approval(ListPrimarySPL, Me.dgvItem, Me.EnumFillSPL, ClsTrans_SPL.EnumStatus.WaitingForApproval, Me.DapatAppOverBudget, Me.UserCode) = True Then
                Hasil = MsgBox("Wait for approval ?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    Dim ij As New Core_Module.DialogPassword
                    ij.txtUserName.Text = Me.UserCode
                    ij.ShowDialog()

                    If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                        For Each row As DataGridViewRow In Me.dgvItem.Rows
                            If row.Cells(Me.ColItem_Modifiy.Index).Value = True And ListPrimarySPL.Contains(row.Cells(Me.ColItem_Primary.Index).Value) = True Then
                                Me.ObjSpl.UpdateStatusApproval(-1, ClsTrans_SPL.EnumStatus.WaitingForApproval, row.Cells(Me.ColItem_Primary.Index).Value)
                            End If
                        Next

                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Me.KosongkanField()
                        Me.IsiDgv_SPL(Me.EnumFillSPL)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFilter.Click
        Try
            If Me.RboWaitingForApproval.Checked = True Then
                Me.IsiDgv_SPL(ClsTrans_SPL.EnumStatus.WaitingForApproval)
            ElseIf Me.RboApproved.Checked = True Then
                Me.IsiDgv_SPL(ClsTrans_SPL.EnumStatus.Approved)
            ElseIf Me.RboRejected.Checked = True Then
                Me.IsiDgv_SPL(ClsTrans_SPL.EnumStatus.Rejected)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbClearFilter_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbClearFilter.LinkClicked
        Try
            Me.ChkDocDate.Checked = True

            Me.DgvEmpFilter.Rows.Clear()
            Me.DgvDeptFilter.Rows.Clear()
            Me.DgvPosFilter.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkDocDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkDocDate.CheckedChanged
        Try
            If Me.ChkDocDate.Checked = True Then
                Me.DtpDocDateFrom.Enabled = True
                Me.DtpDocDateTo.Enabled = True
                Me.DtpDocDateFrom.Value = Me.ObjCD.GetFirstDateofTheMonth(Now)
                Me.DtpDocDateTo.Value = Me.ObjCD.GetLasDateOfTheMonth(Now)
            Else
                Me.DtpDocDateFrom.Enabled = False
                Me.DtpDocDateTo.Enabled = False
                Me.DtpDocDateFrom.Value = Now
                Me.DtpDocDateTo.Value = Now
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim i As New FrmTrans_SPLApprovalBudget
        Dim i2 As New FrmTrans_SPLApprovalNotes

        Try
            If e.ColumnIndex = Me.ColItem_Budget.Index Then
                i.PrimarySPL = CDbl(Me.dgvItem.Item(Me.ColItem_Primary.Index, e.RowIndex).Value)
                i.TblSplAppDep2 = Me.TblSplAppDep2
                i.TblSplUsed = Me.TblSplUsed
                i.BackColorAll = Me.BackColorAll
                i.ShowDialog()
            ElseIf e.ColumnIndex = Me.ColItem_StatusNotes.Index Then
                i2.mPrimaryTrans = Me.dgvItem(Me.ColItem_Primary.Index, e.RowIndex).Value
                i2.FrmUtama = Me
                i2.BackColorAll = Me.BackColorAll
                i2.ShowDialog()
            ElseIf e.ColumnIndex = Me.ColItem_Print.Index Then
                Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi
                Dim FrmCetak As New FrmTrans_SPL_Cetak

                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_SPL_Cetak") = False Then
                    MsgBox("You don't have permission to perform this action", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjSpl.Cetak(Me.dgvItem.Item(Me.ColItem_Primary.Index, e.RowIndex).Value, FrmCetak.crv, FrmCetak.crvDetail) = True Then
                    FrmCetak.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            i2.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Dim PrimarySPL As Double = -1

        Try
            If Me.dgvItem.SelectedCells.Count > 0 Then
                PrimarySPL = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex).Cells(Me.ColItem_Primary.Index).Value
                If String.IsNullOrEmpty(Me.txtPrimary.Text) = False Then
                    If Me.txtPrimary.Text <> PrimarySPL And PrimarySPL > 0 Then
                        Me.IsiField(PrimarySPL)
                    End If
                End If
            Else
                Me.KosongkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiField(ByVal vPrimarySPL As Double)
        If Me.dgvEmployee.ColumnCount = 0 Then Exit Sub
        If Me.dgvSchDetail.ColumnCount = 0 Then Exit Sub
        If Me.dgvSchEmp.ColumnCount = 0 Then Exit Sub

        Dim Frm As New FrmTrans_SPL
        Dim Dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_employeeTableAdapter
        Dim Tbl As New DseTrans_SPL.tbltrans_spl_overtime_order_item_employeeDataTable
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_employeeRow
        Dim Dta1 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_sch_detailTableAdapter
        Dim Tbl1 As New DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailDataTable
        Dim row1 As DseTrans_SPL.tbltrans_spl_overtime_order_item_sch_detailRow
        Dim Dta2 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_orderTableAdapter
        Dim Tbl2 As New DseTrans_SPL.tbltrans_spl_overtime_orderDataTable
        Dim row2 As DseTrans_SPL.tbltrans_spl_overtime_orderRow
        Dim Dta3 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approved_otTableAdapter
        Dim Tbl3 As New DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otDataTable
        Dim row3 As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow

        Dim rowD As DataGridViewRow
        Dim ShowEarly As Boolean = False
        Dim ShowOT As Boolean = False
        Dim ShowBreak As Boolean = False
        Dim ShowBreakEarly As Boolean = False
        Dim ShowBreakOT As Boolean = False
        Me.Cursor = Cursors.WaitCursor

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta1.Connection.ConnectionString = Core_Module.GetConstring
            Dta2.Connection.ConnectionString = Core_Module.GetConstring
            Dta3.Connection.ConnectionString = Core_Module.GetConstring

            Dta.Fill(Tbl, vPrimarySPL)
            Dta2.Fill(Tbl2, vPrimarySPL)
            Dta3.Fill(Tbl3, vPrimarySPL)

            If Tbl2.Count > 0 Then
                row2 = Tbl2.Rows(0)
                With row2
                    Me.txtPrimary.Text = .Primary
                    Me.txtDocumentNo.Text = .No_Transaksi
                    Me.dtpDate.Value = Format(._Date, "dd MMM yyyy")
                    Me.TxtReason.Text = .Reason
                    Me.DtpFrom.Value = Format(.From, "dd MMM yyyy")
                    Me.DtpTo.Value = Format(._To, "dd MMM yyyy")
                    Me.txtTotalDuration.Text = Me.ObjCD.Convert_ToString(.Total_Duration)
                    Me.txtRemarks.Text = .Remarks
                End With
            End If

            'ISI DGV EMPLOYEE
            If Tbl.Count > 0 Then
                'kosongkan dgv
                Me.dgvEmployee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvEmployee.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.dgvEmployee.Rows.Clear()
                ''isikan(tabel)
                For Each row In Tbl
                    With row
                        Me.dgvEmployee.Rows.Add()
                        rowD = Me.dgvEmployee.Rows(Me.dgvEmployee.Rows.Count - 1)
                        rowD.Cells(Me.colEmpl_No.Index).Value = Me.dgvEmployee.Rows.Count
                        rowD.Cells(Me.colEmpl_Primary.Index).Value = .Primary_Employee
                        rowD.Cells(Me.colEmpl_PrimaryPos.Index).Value = .Primary_Position
                        rowD.Cells(Me.colEmpl_PrimaryDept.Index).Value = .Primary_Deptartment
                        rowD.Cells(Me.colEmpl_Remarks.Index).Value = .Remarks
                        rowD.Cells(Me.colEmpl_Code.Index).Value = .Emp_Code
                        rowD.Cells(Me.colEmpl_NIK.Index).Value = .Emp_NIK
                        rowD.Cells(Me.colEmpl_Name.Index).Value = .Emp_Name
                        rowD.Cells(Me.colEmpl_JobPosition.Index).Value = .Position_Code & " - " & .Position_Desc
                        rowD.Cells(Me.colEmpl_Department.Index).Value = .Dept_Code & " - " & .Dept_Desc
                        rowD.Cells(Me.colEmpl_TotalHours.Index).Value = Me.ObjCD.Convert_ToString(Me.ObjSpl.GetTotalHoursEmployee(.Primary_Employee, Tbl3))
                    End With
                Next
                'set nomer
                Me.dgvEmployee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvEmployee.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End If

            'ISI DGV SCEDULE DETAIL
            If Dta1.Fill(Tbl1, vPrimarySPL) > 0 Then
                Me.dgvSchDetail.Rows.Clear()
                Me.dgvSchDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvSchDetail.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Tbl1.DefaultView.RowFilter = ""
                Tbl1.DefaultView.Sort = "[Shift Code] ASC, [Date] ASC, [OT Type] ASC"
                For Each dv As DataRowView In Tbl1.DefaultView
                    row1 = dv.Row
                    Me.dgvSchDetail.Rows.Add()
                    rowD = Me.dgvSchDetail.Rows(Me.dgvSchDetail.RowCount - 1)
                    With row1
                        rowD.Cells(Me.ColSch_No.Index).Value = Me.dgvSchDetail.RowCount
                        rowD.Cells(Me.ColSch_DayOfWeek.Index).Value = Me.ObjCD.GetHari_English(._Date)
                        rowD.Cells(Me.ColSch_Date.Index).Value = Format(._Date, "dd MMM yyyy")
                        rowD.Cells(Me.ColSch_Shift.Index).Value = .Shift_Code
                        rowD.Cells(Me.ColSch_NormalWH.Index).Value = Me.ObjCD.Convert_ToString(.Start_Time) & " - " & Me.ObjCD.Convert_ToString(.End_Time)
                        rowD.Cells(Me.ColSch_OTtype.Index).Value = Frm.GetTypeOT_String(.OT_Type)
                        rowD.Cells(Me.ColSch_Start.Index).Value = Me.ObjCD.Convert_ToString(.Start)
                        If .Break_Out <> 0 And .Break_In <> 0 Then
                            rowD.Cells(Me.ColSch_Break.Index).Value = Me.ObjCD.Convert_ToString(.Break_Out) & " - " & Me.ObjCD.Convert_ToString(.Break_In)
                            ShowBreak = True
                        Else
                            rowD.Cells(Me.ColSch_Break.Index).Value = ""
                        End If
                        rowD.Cells(Me.ColSch_End.Index).Value = Me.ObjCD.Convert_ToString(._End)
                        rowD.Cells(Me.ColSch_Duration.Index).Value = Me.ObjCD.Convert_ToString(.Duration)
                        If .Override = 0 Then
                            rowD.Cells(Me.ColOverride.Index).Value = False
                        Else
                            rowD.Cells(Me.ColOverride.Index).Value = True
                        End If
                        rowD.Cells(Me.ColSch_Remark.Index).Value = .Remarks
                        rowD.Cells(Me.ColSch_PrimaryShift.Index).Value = .Primary_Shift
                    End With
                Next

                If ShowBreak = True Then
                    Me.ColSch_Break.Visible = True
                Else
                    Me.ColSch_Break.Visible = False
                End If

                Me.dgvSchDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvSchDetail.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End If

            If Tbl3.Count > 0 Then
                Tbl3.DefaultView.RowFilter = ""
                Tbl3.DefaultView.Sort = "[Emp Code] ASC, [Date] ASC"

                'kosongkan dgv
                Me.dgvSchEmp.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvSchEmp.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.dgvSchEmp.Rows.Clear()
                ''isikan(tabel)
                For Each dv As DataRowView In Tbl3.DefaultView
                    row3 = dv.Row
                    With row3
                        Me.dgvSchEmp.Rows.Add("")
                        rowD = Me.dgvSchEmp.Rows(Me.dgvSchEmp.RowCount - 1)
                        rowD.Cells(Me.ColSchEmp_No.Index).Value = Me.dgvSchEmp.Rows.Count
                        rowD.Cells(Me.ColSchEmp_Code.Index).Value = .Emp_Code
                        rowD.Cells(Me.ColSchEmp_Emp.Index).Value = .Emp_Name
                        rowD.Cells(Me.ColSchEmp_DayOfWeek.Index).Value = Me.ObjCD.GetHari_English(._Date)
                        rowD.Cells(Me.ColSchEmp_Date.Index).Value = Format(._Date, "dd MMM yyyy")
                        rowD.Cells(Me.ColSchEmp_Shift.Index).Value = .Shift_Code
                        rowD.Cells(Me.ColSchEmp_NormalWH.Index).Value = Me.ObjCD.Convert_ToString(.Start_WH) & " - " & Me.ObjCD.Convert_ToString(.End_WH)

                        rowD.Cells(Me.ColSchEmp_StartEarly.Index).Value = Me.ObjCD.Convert_ToString(.Start_Time_Early)
                        If .Break_Out_Time_Early <> 0 And .Break_In_Time_Early <> 0 Then
                            rowD.Cells(Me.ColSchEmp_BreakEarly.Index).Value = Me.ObjCD.Convert_ToString(.Break_Out_Time_Early) & " - " & Me.ObjCD.Convert_ToString(.Break_In_Time_Early)
                            ShowBreakEarly = True
                        Else
                            rowD.Cells(Me.ColSchEmp_BreakEarly.Index).Value = ""
                        End If
                        rowD.Cells(Me.ColSchEmp_EndEarly.Index).Value = Me.ObjCD.Convert_ToString(.End_Time_Early)
                        rowD.Cells(Me.ColSchEmp_EarlyOTDuration.Index).Value = Me.ObjCD.Convert_ToString(.OT_Early_Approved)
                        If .Override_Early_OT = 0 Then
                            rowD.Cells(Me.ColSchEmp_OverrideEarlyOT.Index).Value = Boolean.FalseString
                        Else
                            rowD.Cells(Me.ColSchEmp_OverrideEarlyOT.Index).Value = Boolean.TrueString
                        End If

                        rowD.Cells(Me.ColSchEmp_StartOT.Index).Value = Me.ObjCD.Convert_ToString(.Start_Time)
                        If .Break_Out_Time <> 0 And .Break_In_Time <> 0 Then
                            rowD.Cells(Me.ColSchEmp_BreakOT.Index).Value = Me.ObjCD.Convert_ToString(.Break_Out_Time) & " - " & Me.ObjCD.Convert_ToString(.Break_In_Time)
                            ShowBreakOT = True
                        Else
                            rowD.Cells(Me.ColSchEmp_BreakOT.Index).Value = ""
                        End If
                        rowD.Cells(Me.ColSchEmp_EndOT.Index).Value = Me.ObjCD.Convert_ToString(.End_Time)
                        rowD.Cells(Me.ColSchEmp_OTDuration.Index).Value = Me.ObjCD.Convert_ToString(.OT_Aprroved)
                        If .Override_OT = 0 Then
                            rowD.Cells(Me.ColSchEmp_OverrideOT.Index).Value = Boolean.FalseString
                        Else
                            rowD.Cells(Me.ColSchEmp_OverrideOT.Index).Value = Boolean.TrueString
                        End If

                        rowD.Cells(Me.ColSchEmp_PrimaryShift.Index).Value = .Primary_Shift
                        rowD.Cells(Me.ColSchEmp_PrimaryEmp.Index).Value = .Primary_Employee

                        If .OT_Early_Approved > 0 Then
                            ShowEarly = True
                        End If
                        If .OT_Aprroved > 0 Then
                            ShowOT = True
                        End If
                    End With
                Next

                If ShowEarly = False Then
                    Me.ColSchEmp_StartEarly.Visible = False
                    Me.ColSchEmp_BreakEarly.Visible = False
                    Me.ColSchEmp_EndEarly.Visible = False
                    Me.ColSchEmp_EarlyOTDuration.Visible = False
                    Me.ColSchEmp_OverrideEarlyOT.Visible = False
                Else
                    Me.ColSchEmp_StartEarly.Visible = True
                    Me.ColSchEmp_BreakEarly.Visible = True
                    Me.ColSchEmp_EndEarly.Visible = True
                    Me.ColSchEmp_EarlyOTDuration.Visible = True
                    Me.ColSchEmp_OverrideEarlyOT.Visible = True
                End If

                If ShowOT = False Then
                    Me.ColSchEmp_StartOT.Visible = False
                    Me.ColSchEmp_BreakOT.Visible = False
                    Me.ColSchEmp_EndOT.Visible = False
                    Me.ColSchEmp_OTDuration.Visible = False
                    Me.ColSchEmp_OverrideOT.Visible = False
                Else
                    Me.ColSchEmp_StartOT.Visible = True
                    Me.ColSchEmp_BreakOT.Visible = True
                    Me.ColSchEmp_EndOT.Visible = True
                    Me.ColSchEmp_OTDuration.Visible = True
                    Me.ColSchEmp_OverrideOT.Visible = True
                End If

                If ShowBreakEarly = True Then
                    Me.ColSchEmp_BreakEarly.Visible = True
                Else
                    Me.ColSchEmp_BreakEarly.Visible = False
                End If
                If ShowBreakOT = True Then
                    Me.ColSchEmp_BreakOT.Visible = True
                Else
                    Me.ColSchEmp_BreakOT.Visible = False
                End If

                'set nomer
                Me.dgvSchEmp.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvSchEmp.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'Me.Dispose()
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub KosongkanField()

        Try
            Me.txtPrimary.Text = -1
            Me.txtDocumentNo.Text = ""
            Me.dtpDate.Value = Now
            Me.TxtReason.Text = ""
            Me.DtpFrom.Value = Now
            Me.DtpTo.Value = Now
            Me.txtTotalDuration.Text = "00.00"
            Me.txtRemarks.Text = ""

            'Me.dgvItem.Rows.Clear()
            Me.dgvEmployee.Rows.Clear()
            Me.dgvSchDetail.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError

    End Sub

    Private Sub DgpDeptFilter_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvDeptFilter.DataError

    End Sub

    Private Sub DgvPosFilter_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvPosFilter.DataError

    End Sub

    Private Sub DgvEmpFilter_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvEmpFilter.DataError

    End Sub

    Private Sub dgvEmployee_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvEmployee.DataError

    End Sub

    Private Sub dgvSchDetail_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSchDetail.DataError

    End Sub

    Private Sub dgvSchEmp_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSchEmp.DataError

    End Sub

    Private Sub DgvEmpFilter_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvEmpFilter.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colEmpl_Code.Index Or e.ColumnIndex = Me.colEmpl_Remarks.Index Then
                Me.Buffer = Me.DgvEmpFilter.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemEmployee(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvEmpFilter.Rows.Count - 1
                RowDGV = Me.DgvEmpFilter.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColFilterEmp_Primary.Index).Value = vPrimary Then
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

    Private Sub DgvEmpFilter_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvEmpFilter.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvEmpFilter.Rows(e.RowIndex)
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
        Dim Frm As New FrmTrans_SPL

        Try
            If IsNothing(Me.DgvEmpFilter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If Me.DgvEmpFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                    Exit Sub
                End If

                If e.ColumnIndex = Me.colEmpl_Code.Index Then
                    'FUNGSI COL EMPLOYEE CODE
                    Me.DgvEmpFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvEmpFilter.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.ObjEmp.ApakahSudahAda(Me.DgvEmpFilter.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimaryAcuan = Me.ObjEmp.GetPrimaryByCode(Me.DgvEmpFilter.Item(e.ColumnIndex, e.RowIndex).Value)
                        index = Me.GetIndexItemEmployee(PrimaryAcuan, e.RowIndex)
                        If index < 0 Then
                            'jika employee ditambahkan maka isi dgv sch setup, dgv sch detail, dgv approval
                            If Frm.GetEmployeeData(PrimaryAcuan, EmpCode, EmpNIK, EmpName, PositionPrimary, PositionCode, PositionDesc, DeptPrimary, DeptCode, DeptDesc, ListPrimaryAppGroup, False) = False Then
                                GoTo DeleteRows
                            End If
                            Me.DgvEmpFilter.Item(Me.ColFilterEmp_Primary.Index, e.RowIndex).Value = PrimaryAcuan
                            Me.DgvEmpFilter.Item(Me.ColFilterEmp_Code.Index, e.RowIndex).Value = EmpCode
                            Me.DgvEmpFilter.Item(Me.ColFilterEmp_NIK.Index, e.RowIndex).Value = EmpNIK
                            Me.DgvEmpFilter.Item(Me.ColFilterEmp_Name.Index, e.RowIndex).Value = EmpName
                            Me.DgvEmpFilter.Item(Me.ColFilterEmp_JobPosition.Index, e.RowIndex).Value = PositionCode & " - " & PositionDesc
                            Me.DgvEmpFilter.Item(Me.ColFilterEmp_Dept.Index, e.RowIndex).Value = DeptCode & " - " & DeptDesc
                        Else
                            If e.RowIndex <> index Then
                                MsgBox("Duplication employee code, please input another code.", MsgBoxStyle.Information, "Information")
                                GoTo DeleteRows
                            End If
                        End If
                    Else
                        Dim ij As New FrmMasterKaryawanCari
                        ij.txtCari.Text = Me.DgvEmpFilter.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.BackColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        If ij.PrimaryEmploye <> -1 Then
                            index = Me.GetIndexItemEmployee(ij.PrimaryEmploye, e.RowIndex)
                            If index < 0 Then
                                'jika employee ditambahkan maka isi dgv sch setup, dgv sch detail, dgv approval
                                If Frm.GetEmployeeData(ij.PrimaryEmploye, EmpCode, EmpNIK, EmpName, PositionPrimary, PositionCode, PositionDesc, DeptPrimary, DeptCode, DeptDesc, ListPrimaryAppGroup, False) = False Then
                                    GoTo DeleteRows
                                End If
                                Me.DgvEmpFilter.Item(Me.ColFilterEmp_Primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                Me.DgvEmpFilter.Item(Me.ColFilterEmp_Code.Index, e.RowIndex).Value = EmpCode
                                Me.DgvEmpFilter.Item(Me.ColFilterEmp_NIK.Index, e.RowIndex).Value = EmpNIK
                                Me.DgvEmpFilter.Item(Me.ColFilterEmp_Name.Index, e.RowIndex).Value = EmpName
                                Me.DgvEmpFilter.Item(Me.ColFilterEmp_JobPosition.Index, e.RowIndex).Value = PositionCode & " - " & PositionDesc
                                Me.DgvEmpFilter.Item(Me.ColFilterEmp_Dept.Index, e.RowIndex).Value = DeptCode & " - " & DeptDesc
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication employee code, please input another code.", MsgBoxStyle.Information, "Information")
                                    GoTo DeleteRows
                                Else
                                    'jika employee ditambahkan maka isi dgv sch setup, dgv sch detail, dgv approval
                                    If Frm.GetEmployeeData(ij.PrimaryEmploye, EmpCode, EmpNIK, EmpName, PositionPrimary, PositionCode, PositionDesc, DeptPrimary, DeptCode, DeptDesc, ListPrimaryAppGroup, False) = False Then
                                        GoTo DeleteRows
                                    End If
                                    Me.DgvEmpFilter.Item(Me.ColFilterEmp_Primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    Me.DgvEmpFilter.Item(Me.ColFilterEmp_Code.Index, e.RowIndex).Value = EmpCode
                                    Me.DgvEmpFilter.Item(Me.ColFilterEmp_NIK.Index, e.RowIndex).Value = EmpNIK
                                    Me.DgvEmpFilter.Item(Me.ColFilterEmp_Name.Index, e.RowIndex).Value = EmpName
                                    Me.DgvEmpFilter.Item(Me.ColFilterEmp_JobPosition.Index, e.RowIndex).Value = PositionCode & " - " & PositionDesc
                                    Me.DgvEmpFilter.Item(Me.ColFilterEmp_Dept.Index, e.RowIndex).Value = DeptCode & " - " & DeptDesc
                                End If
                            End If
                        ElseIf ij.PrimaryEmploye = -1 Then
                            GoTo DeleteRows
                        End If

                        ij = Nothing
                    End If

                End If
            Else
DeleteRows:
                If String.IsNullOrEmpty(Me.Buffer) = True Then
                    If Me.DgvEmpFilter.Rows.Count - 1 <> e.RowIndex Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BgwEmployee.RunWorkerAsync()
                    End If
                Else
                    Me.DgvEmpFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvEmpFilter_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvEmpFilter.RowsAdded
        If Me.DgvEmpFilter.RowCount <= 0 Then Exit Sub
        Dim row As DataGridViewRow = Me.DgvEmpFilter.Rows(Me.DgvEmpFilter.RowCount - 1)

        Try
            row.Cells(Me.ColFilterEmp_No.Index).Value = Me.DgvEmpFilter.RowCount
            row.Cells(Me.ColFilterEmp_Code.Index).Value = ""
            row.Cells(Me.ColFilterEmp_NIK.Index).Value = ""
            row.Cells(Me.ColFilterEmp_Name.Index).Value = ""
            row.Cells(Me.ColFilterEmp_JobPosition.Index).Value = ""
            row.Cells(Me.ColFilterEmp_Dept.Index).Value = ""
            row.Cells(Me.ColFilterEmp_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvEmpFilter_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvEmpFilter.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvEmpFilter, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvEmpFilter_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvEmpFilter.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvEmpFilter, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvPosFilter_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvPosFilter.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvPosFilter.Item(Me.ColFilterJobPos_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColFilterJobPos_Code.Index Then
                    Me.Buffer = Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColFilterJobPos_Code.Index Then
                    Me.Buffer = Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemJobPosition(ByVal vPrimary As Double, _
         ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvPosFilter.Rows.Count - 1
                RowDGV = Me.DgvPosFilter.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColFilterJobPos_Primary.Index).Value = vPrimary Then
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

    Private Sub DgvPosFilter_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvPosFilter.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvPosFilter.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryJobPost As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            Me.DgvPosFilter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvPosFilter.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            With row
                If IsNothing(Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.Buffer = Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColFilterJobPos_Code.Index Then
                        Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjJobPost.ApakahSudahAda(Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryJobPost = Me.ObjJobPost.GetPrimary(Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemJobPosition(PrimaryJobPost, e.RowIndex)
                            If index < 0 Then
                                Me.DgvPosFilter.Item(Me.ColFilterJobPos_Primary.Index, e.RowIndex).Value = PrimaryJobPost
                                Me.DgvPosFilter.Item(Me.ColFilterJobPos_Code.Index, e.RowIndex).Value = Me.ObjJobPost.GetCode(PrimaryJobPost)
                                Me.DgvPosFilter.Item(Me.ColFilterJobPos_Desc.Index, e.RowIndex).Value = Me.ObjJobPost.GetDesc(PrimaryJobPost)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvPosFilter.Item(Me.ColFilterJobPos_Primary.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwJobPosition.RunWorkerAsync()
                                    Else
                                        Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvPosFilter)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterJobPosition_Cari
                            ij.txtCari.Text = Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.mPrimary <> -1 Then
                                index = Me.GetIndexItemJobPosition(ij.mPrimary, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvPosFilter.Item(Me.ColFilterJobPos_Primary.Index, e.RowIndex).Value = ij.mPrimary
                                    Me.DgvPosFilter.Item(Me.ColFilterJobPos_Code.Index, e.RowIndex).Value = Me.ObjJobPost.GetCode(ij.mPrimary)
                                    Me.DgvPosFilter.Item(Me.ColFilterJobPos_Desc.Index, e.RowIndex).Value = Me.ObjJobPost.GetDesc(ij.mPrimary)
                                Else
                                    If Me.Buffer = Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value Then
                                        Exit Sub
                                    End If
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwJobPosition.RunWorkerAsync()
                                        Else
                                            Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.DgvPosFilter)
                                        Exit Sub
                                    Else
                                        Me.DgvPosFilter.Item(Me.ColFilterJobPos_Primary.Index, e.RowIndex).Value = ij.mPrimary
                                        Me.DgvPosFilter.Item(Me.ColFilterJobPos_Code.Index, e.RowIndex).Value = Me.ObjJobPost.GetCode(ij.mPrimary)
                                        Me.DgvPosFilter.Item(Me.ColFilterJobPos_Desc.Index, e.RowIndex).Value = Me.ObjJobPost.GetDesc(ij.mPrimary)
                                    End If
                                End If
                            ElseIf ij.mPrimary = -1 Then
                                If Me.DgvPosFilter.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwJobPosition.RunWorkerAsync()
                                    Else
                                        Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvPosFilter)
                                    Exit Sub
                                End If
                                If Me.DgvPosFilter.Item(Me.ColFilterJobPos_Primary.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwJobPosition.RunWorkerAsync()
                                    Else
                                        Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvPosFilter)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvPosFilter)
                    End If
                Else
                    If e.ColumnIndex = Me.ColFilterJobPos_Primary.Index Then
                        If IsNothing(Me.Buffer) = False Then
                            Me.EraseRowIndex = e.RowIndex
                            Me.BgwJobPosition.RunWorkerAsync()
                        End If
                        Me.DgvPosFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvPosFilter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvPosFilter.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Sub DgvPosFilter_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvPosFilter.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvPosFilter, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvPosFilter_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvPosFilter.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvPosFilter.Rows(Me.DgvPosFilter.Rows.Count - 1)
            dr.Cells(Me.ColFilterJobPos_Primary.Index).Value = -1
            dr.Cells(Me.ColFilterJobPos_No.Index).Value = Me.DgvPosFilter.Rows.Count
            dr.Cells(Me.ColFilterJobPos_Code.Index).Value = ""
            dr.Cells(Me.ColFilterJobPos_Desc.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvPosFilter_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvPosFilter.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvPosFilter, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvDeptFilter_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvDeptFilter.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvDeptFilter.Item(Me.ColFilterDept_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColFilterDept_Code.Index Then
                    Me.Buffer = Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColFilterDept_Code.Index Then
                    Me.Buffer = Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemDepartment(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvDeptFilter.Rows.Count - 1
                RowDGV = Me.DgvDeptFilter.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColFilterDept_Primary.Index).Value = vPrimary Then
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

    Private Sub DgvDeptFilter_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvDeptFilter.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvDeptFilter.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryDept As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            Me.DgvDeptFilter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvDeptFilter.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            With row
                If IsNothing(Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.Buffer = Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColFilterDept_Code.Index Then
                        Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjDept.ApakahSudahAda(Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryDept = Me.ObjDept.GetPrimary(Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemDepartment(PrimaryDept, e.RowIndex)
                            If index < 0 Then
                                Me.DgvDeptFilter.Item(Me.ColFilterDept_Primary.Index, e.RowIndex).Value = PrimaryDept
                                Me.DgvDeptFilter.Item(Me.ColFilterDept_Code.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(PrimaryDept)
                                Me.DgvDeptFilter.Item(Me.ColFilterDept_Desc.Index, e.RowIndex).Value = Me.ObjDept.GetDeptDesc(PrimaryDept)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvDeptFilter.Item(Me.ColFilterDept_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwDepartment.RunWorkerAsync()
                                    Else
                                        Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvDeptFilter)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterDepartementCari
                            ij.txtCari.Text = Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryDepartement <> -1 Then
                                index = Me.GetIndexItemDepartment(ij.PrimaryDepartement, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvDeptFilter.Item(Me.ColFilterDept_Primary.Index, e.RowIndex).Value = ij.PrimaryDepartement
                                    Me.DgvDeptFilter.Item(Me.ColFilterDept_Code.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(ij.PrimaryDepartement)
                                    Me.DgvDeptFilter.Item(Me.ColFilterDept_Desc.Index, e.RowIndex).Value = Me.ObjDept.GetDeptDesc(ij.PrimaryDepartement)
                                Else
                                    If Me.Buffer = Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value Then
                                        Exit Sub
                                    End If
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwDepartment.RunWorkerAsync()
                                        Else
                                            Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.DgvDeptFilter)
                                        Exit Sub
                                    Else
                                        Me.DgvDeptFilter.Item(Me.ColFilterDept_Primary.Index, e.RowIndex).Value = ij.PrimaryDepartement
                                        Me.DgvDeptFilter.Item(Me.ColFilterDept_Code.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(ij.PrimaryDepartement)
                                        Me.DgvDeptFilter.Item(Me.ColFilterDept_Desc.Index, e.RowIndex).Value = Me.ObjDept.GetDeptDesc(ij.PrimaryDepartement)
                                    End If
                                End If
                            ElseIf ij.PrimaryDepartement = -1 Then
                                If Me.DgvDeptFilter.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwDepartment.RunWorkerAsync()
                                    Else
                                        Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvDeptFilter)
                                    Exit Sub
                                End If
                                If Me.DgvDeptFilter.Item(Me.ColFilterDept_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwDepartment.RunWorkerAsync()
                                    Else
                                        Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvDeptFilter)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvDeptFilter)
                    End If
                Else
                    If e.ColumnIndex = Me.ColFilterDept_Code.Index Then
                        If IsNothing(Me.Buffer) = False Then
                            Me.EraseRowIndex = e.RowIndex
                            Me.BgwDepartment.RunWorkerAsync()
                        End If
                        Me.DgvDeptFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvDeptFilter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvDeptFilter.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Sub DgvDeptFilter_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvDeptFilter.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvDeptFilter, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvDeptFilter_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvDeptFilter.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvDeptFilter.Rows(Me.DgvDeptFilter.Rows.Count - 1)
            dr.Cells(Me.ColFilterDept_Primary.Index).Value = -1
            dr.Cells(Me.ColFilterDept_No.Index).Value = Me.DgvDeptFilter.Rows.Count
            dr.Cells(Me.ColFilterDept_Code.Index).Value = ""
            dr.Cells(Me.ColFilterDept_Desc.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvDeptFilter_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvDeptFilter.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvDeptFilter, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDepartment_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwDepartment.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDepartment_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwDepartment.RunWorkerCompleted
        Try
            If Me.DgvDeptFilter.RowCount > 0 Then
                Me.DgvDeptFilter.Rows.RemoveAt(Me.EraseRowIndex)
                Me.ObjCD.SetNumber(Me.DgvDeptFilter)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwJobPosition_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwJobPosition.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwJobPosition_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwJobPosition.RunWorkerCompleted
        Try
            If Me.DgvPosFilter.RowCount > 0 Then
                Me.DgvPosFilter.Rows.RemoveAt(Me.EraseRowIndex)
                Me.ObjCD.SetNumber(Me.DgvPosFilter)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwEmployee_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwEmployee.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwEmployee_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwEmployee.RunWorkerCompleted
        Try
            If Me.DgvEmpFilter.RowCount > 0 Then
                Me.DgvEmpFilter.Rows.RemoveAt(Me.EraseRowIndex)
                Me.ObjCD.SetNumber(Me.DgvEmpFilter)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub txtRemarks_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.Leave
        Me.txtRemarks.Text = Me.txtRemarks.Text.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
    End Sub
End Class