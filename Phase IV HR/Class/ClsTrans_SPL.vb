Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class ClsTrans_SPL
    Dim mPrimary As Double = -1
    Dim mNoTransaksi As String = ""
    Dim mUrutan As Integer = -1
    Dim mDate As Date = #1/1/1950#
    Dim mReason As String = ""
    Dim mAppliedBy As String = ""
    Dim mAppliedByPosition As String = ""
    Dim mPrimaryApprovedBy As Double = -1
    Dim mApprovedTime As DateTime = #1/1/1950#
    Dim mDateFrom As Date = #1/1/1950#
    Dim mDateTo As Date = #1/1/1950#
    Dim mTotalDuration As Double = 0
    Dim mStatusNotes As String = ""
    Dim mStatus As Integer = 0
    Dim mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now

    Dim ObjItem As New ClsTrans_SPL_Item
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUser As New Core_Module.ClsUser

    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Enum EnumSearch As Integer
        DocNo
        Tgl
        EmpCode
        EmpName
        Remarks
    End Enum

    Public Enum EnumStatus As Integer
        WaitingForApproval
        Approved
        Rejected
    End Enum

    Enum EnumTipeOT As Integer
        Early_Overtime
        Overtime
        Holiday_Overtime
    End Enum

    Enum EnumCariData As Integer
        AllEmp
        dept
        Job
        Other
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
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

    Public Property Urutan() As Integer
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Integer)
            mUrutan = value
        End Set
    End Property

    Public Property Date_() As Date
        Get
            Return mDate
        End Get
        Set(ByVal value As Date)
            mDate = value
        End Set
    End Property

    Public Property Reason() As String
        Get
            Return mReason
        End Get
        Set(ByVal value As String)
            mReason = value
        End Set
    End Property

    Public Property AppliedBy() As String
        Get
            Return mAppliedBy
        End Get
        Set(ByVal value As String)
            mAppliedBy = value
        End Set
    End Property

    Public Property PrimaryApprovedBy() As String
        Get
            Return mPrimaryApprovedBy
        End Get
        Set(ByVal value As String)
            mPrimaryApprovedBy = value
        End Set
    End Property

    Public Property ApprovedTime() As Date
        Get
            Return mApprovedTime
        End Get
        Set(ByVal value As Date)
            mApprovedTime = value
        End Set
    End Property

    Public Property AppliedByPosition() As String
        Get
            Return mAppliedByPosition
        End Get
        Set(ByVal value As String)
            mAppliedByPosition = value
        End Set
    End Property

    Public Property DateFrom_() As Date
        Get
            Return mDateFrom
        End Get
        Set(ByVal value As Date)
            mDateFrom = value
        End Set
    End Property

    Public Property DateTo_() As Date
        Get
            Return mDateTo
        End Get
        Set(ByVal value As Date)
            mDateTo = value
        End Set
    End Property

    Public Property TotalDuration() As Double
        Get
            Return mTotalDuration
        End Get
        Set(ByVal value As Double)
            mTotalDuration = value
        End Set
    End Property

    Public Property StatusNotes() As String
        Get
            Return mStatusNotes
        End Get
        Set(ByVal value As String)
            mStatusNotes = value
        End Set
    End Property

    Public Property Status() As Integer
        Get
            Return mStatus
        End Get
        Set(ByVal value As Integer)
            mStatus = value
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
            Return Me.mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return Me.mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return Me.mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return Me.mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return Me.mTimestamp
        End Get
        Set(ByVal value As Date)
            Me.mTimestamp = value
        End Set
    End Property

    Public Sub ResetObject()
        Try
            mPrimary = -1
            mNoTransaksi = ""
            mUrutan = -1
            mDate = #1/1/1950#
            mReason = ""
            mAppliedBy = ""
            mAppliedByPosition = ""
            mPrimaryApprovedBy = -1
            mApprovedTime = #1/1/1950#
            mDateFrom = #1/1/1950#
            mDateTo = #1/1/1950#
            mTotalDuration = 0
            mStatusNotes = ""
            mStatus = 0
            mRemarks = ""

            mCreatedBy = ""
            mCreatedDate = Now
            mModifiedBy = ""
            mModifiedDate = Now
            mTimestamp = Now
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahBolehedit(ByVal vUserCode As String, ByVal vPrimaryTransaksi As Double, Optional ByVal vShowMsg As Boolean = True) As Boolean
        Dim Dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_scalarTableAdapter

        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_SPL_Edit") = False Then
                If vShowMsg = True Then
                    MsgBox("You don't have permission to perform this action", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '1.strict edit
            Dim UserCode As String = Dta.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If UserCode <> vUserCode Then
                If Me.ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_SPL_EditUserLain") = False Then
                    If vShowMsg = True Then
                        MsgBox("Only user with ID:" + UserCode + " can edit this data.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            ''2.status aktif tidak boleh di edit
            Dim status As Integer = Me.GetStatus(vPrimaryTransaksi)
            If status = EnumStatus.Rejected Then
                'jika reject
                If vShowMsg = True Then
                    MsgBox("Overtime order status is rejected, you can't edit this data", MsgBoxStyle.Information, "Information")
                End If
                Return False
            ElseIf status = EnumStatus.Approved Then
                'jika approved
                If vShowMsg = True Then
                    MsgBox("Overtime order status is approved, you can't edit this data", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '4.tidak boleh jika sudah ada di transaksi loan payment

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vUserCode As String, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim Dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_scalarTableAdapter

        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_SPL_Hapus") = False Then
                MsgBox("You don't have permission to perform this action", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1.strict edit
            Dim UserCode As String = Dta.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If UserCode <> vUserCode Then
                If Me.ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_SPL_HapusUserLain") = False Then
                    MsgBox("Only user with ID:" + UserCode + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Dim status As Integer = Me.GetStatus(vPrimaryTransaksi)
            If status = EnumStatus.Rejected Then
                MsgBox("Overtime order status is rejected, you can't delete this data", MsgBoxStyle.Information, "Information")
                'jika reject
                Return False
            ElseIf status = EnumStatus.Approved Then
                'jika approved
                MsgBox("Overtime order status is approved, you can't delete this data", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Private Function ApakahDataBenar(ByVal Frm As FrmTrans_SPL, ByVal vBaru As Boolean) As Boolean
        Dim i As New FrmTrans_SPL

        Try
            '1.tidak boleh double
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.NoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mDate, Me.mPrimary)
                End If
            Else
                Dim NoTransaksiLama As String = Me.GetNoTransaksiByPrimary(Me.mPrimary)
                If NoTransaksiLama <> Me.mNoTransaksi Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mDate, Me.mPrimary)
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Function

    Public Function ApakahDataBenar_Approval(ByRef vListPrimarySpl As List(Of Double), ByVal vDgvItem As DataGridView, _
    ByVal vStatusSPL_Bef As EnumStatus, ByVal vStatusSPL_Aft As EnumStatus, ByVal vDapatAppOverBudget As Boolean, _
    ByVal vUserCode As String) As Boolean
        'table adapter = tbltrans_spl_item_employee
        Dim Dta As New DseTrans_SPLTableAdapters.DataTable1TableAdapter
        Dim Tbl As New DseTrans_SPL.DataTable1DataTable
        Dim row As DseTrans_SPL.DataTable1Row
        'table adapter = tbltrans_attendance
        Dim Dta2 As New DseTrans_SPLTableAdapters.tbltrans_attendanceTableAdapter
        Dim Tbl2 As New DseTrans_SPL.tbltrans_attendanceDataTable
        'table adapter = tbltrans_spl_item_employee_shift
        Dim Dta3 As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_employee_shiftTableAdapter
        Dim Tbl3 As New DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftDataTable
        'table adapter = tbl_allocate
        Dim Dta4 As New DseTrans_SPLTableAdapters.allocate_employeeTableAdapter
        Dim Tbl4 As New DseTrans_SPL.allocate_employeeDataTable

        Dim i As New FrmTrans_SPLApproval
        Dim rowD As DataGridViewRow
        Dim FilterDateFrom As Date = Now
        Dim FilterDateTo As Date = Now
        Dim ListMsg_1 As String = ""
        Dim ListMsg_2 As String = ""
        Dim ListMsg_3 As String = ""
        Dim Ketemu_1 As Boolean = False
        Dim Ketemu_2 As Boolean = False
        Dim Ketemu_3 As Boolean = False
        Dim PrimarySPL_Bef As Double = -1
        Dim ListSPLOverBudget As String = ""

        Try
            If vDgvItem.Rows.Count > 0 Then
                For ii As Integer = 0 To vDgvItem.RowCount - 1
                    rowD = vDgvItem.Rows(ii)
                    If rowD.Cells(i.ColItem_Modifiy.Index).Value = True Then
                        If vStatusSPL_Aft = EnumStatus.Approved Then
                            If vDapatAppOverBudget = False And rowD.Cells(i.ColItem_Budget.Index).Value = "Over" Then
                                If ListSPLOverBudget = "" Then
                                    ListSPLOverBudget = "Overtime Order : " & rowD.Cells(i.ColItem_DocumentNo.Index).Value
                                Else
                                    ListSPLOverBudget &= ", " & rowD.Cells(i.ColItem_DocumentNo.Index).Value
                                End If
                                GoTo SkipRows
                            End If
                        End If

                        If vListPrimarySpl.Contains(rowD.Cells(i.ColItem_Modifiy.Index).Value) = False Then
                            vListPrimarySpl.Add(rowD.Cells(i.ColItem_Primary.Index).Value)
                            If FilterDateFrom > CDate(rowD.Cells(i.ColItem_StartDate.Index).Value) Then
                                FilterDateFrom = CDate(rowD.Cells(i.ColItem_StartDate.Index).Value)
                            End If
                            If FilterDateTo < CDate(rowD.Cells(i.ColItem_EndDate.Index).Value) Then
                                FilterDateTo = CDate(rowD.Cells(i.ColItem_EndDate.Index).Value)
                            End If
                        End If

SkipRows:
                    End If
                Next

                If vListPrimarySpl.Count > 0 Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_SPLApproval_Authorize") = False Then
                        MsgBox("You don't have permission to perform this action", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If

                If ListSPLOverBudget <> "" Then
                    MsgBox("You can't approve over budget spl." & vbCrLf & ListSPLOverBudget, MsgBoxStyle.Information, "Information")
                    If vListPrimarySpl.Count = 0 Then
                        Return False
                    End If
                End If
                If vListPrimarySpl.Count = 0 Then
                    MsgBox("Please checked modify data", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                Return False
            End If

            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta2.Connection.ConnectionString = Core_Module.GetConstring
            Dta3.Connection.ConnectionString = Core_Module.GetConstring
            Dta4.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill_All(Tbl, FilterDateFrom.Date, FilterDateTo.Date) 'yang bukan reject
            Dta2.Fill(Tbl2, FilterDateTo.Date, FilterDateFrom.Date)

            For Each rowD In vDgvItem.Rows
                If rowD.Cells(i.ColItem_Modifiy.Index).Value = True Then
                    If vListPrimarySpl.Contains(rowD.Cells(i.ColItem_Primary.Index).Value) = True Then
                        If vStatusSPL_Bef = EnumStatus.Rejected Then
                            Dta3.Fill(Tbl3, rowD.Cells(i.ColItem_Primary.Index).Value)
                            Dta4.Fill_All(Tbl4, CDate(rowD.Cells(i.ColItem_StartDate.Index).Value), CDate(rowD.Cells(i.ColItem_EndDate.Index).Value))
                        End If

                        PrimarySPL_Bef = rowD.Cells(i.ColItem_Primary.Index).Value
                        Tbl.DefaultView.RowFilter = "[Primary] = " & rowD.Cells(i.ColItem_Primary.Index).Value
                        For Each dv As DataRowView In Tbl.DefaultView
                            row = dv.Row
                            '#CEK PER EMPOLYEE, PERIODE, SPL#
                            Ketemu_1 = False
                            Ketemu_2 = False
                            Ketemu_3 = False
                            If vStatusSPL_Aft = EnumStatus.WaitingForApproval Then
                                If vStatusSPL_Bef = EnumStatus.Approved Then
                                    'approved > wait for approval # tidak bisa di wait for approval jika salah satu employee dari spl sudah di buat attendancenya
                                    If Me.ApaSudahAdaAttendance(rowD.Cells(i.ColItem_Primary.Index).Value, row.Primary_Employee, row.From, row._To, Tbl2) = True Then
                                        vListPrimarySpl.Remove(rowD.Cells(i.ColItem_Primary.Index).Value)
                                        Ketemu_1 = True
                                    End If
                                ElseIf vStatusSPL_Bef = EnumStatus.Rejected Then
                                    'rejected > wait for approval 
                                    '# tidak bisa di wait for approval jika salah satu employee dari spl sudah di buat attendancenya
                                    If Me.ApaSudahAdaAttendance(rowD.Cells(i.ColItem_Primary.Index).Value, row.Primary_Employee, row.From, row._To, Tbl2) = True Then
                                        vListPrimarySpl.Remove(rowD.Cells(i.ColItem_Primary.Index).Value)
                                        Ketemu_1 = True
                                    End If
                                    If Ketemu_1 = False Then
                                        '# tidak bisa di wait for approval jika salah satu employee dari spl sudah di buat lagi di other spl 
                                        If Me.ApaSudahAdaPadaTransaksiLain(rowD.Cells(i.ColItem_Primary.Index).Value, row.Primary_Employee, row.From, row._To, Tbl) = True Then
                                            vListPrimarySpl.Remove(rowD.Cells(i.ColItem_Primary.Index).Value)
                                            Ketemu_2 = True
                                        End If
                                    End If
                                    'pastikan shift nya masih sama jika ubah status reject ke waiting approval
                                    If Ketemu_1 = False And Ketemu_2 = False Then
                                        If Me.ApaShiftKonsisten(rowD.Cells(i.ColItem_Primary.Index).Value, row.Primary_Employee, Tbl3, Tbl4) = False Then
                                            vListPrimarySpl.Remove(rowD.Cells(i.ColItem_Primary.Index).Value)
                                            Ketemu_3 = True
                                        End If
                                    End If
                                End If
                            ElseIf vStatusSPL_Aft = EnumStatus.Approved Then
                                If vStatusSPL_Bef = EnumStatus.WaitingForApproval Then
                                    'wait for approval > approved # tidak ada batasan
                                ElseIf vStatusSPL_Bef = EnumStatus.Rejected Then
                                    'rejected > approved 
                                    '# tidak bisa di approved jika salah satu employee dari spl sudah di buat attendancenya
                                    If Me.ApaSudahAdaAttendance(rowD.Cells(i.ColItem_Primary.Index).Value, row.Primary_Employee, row.From, row._To, Tbl2) = True Then
                                        vListPrimarySpl.Remove(rowD.Cells(i.ColItem_Primary.Index).Value)
                                        Ketemu_1 = True
                                    End If
                                    '# tidak bisa di approved jika salah satu employee dari spl sudah di buat lagi di other spl 
                                    If Ketemu_1 = False Then
                                        If Me.ApaSudahAdaPadaTransaksiLain(rowD.Cells(i.ColItem_Primary.Index).Value, row.Primary_Employee, row.From, row._To, Tbl) = True Then
                                            vListPrimarySpl.Remove(rowD.Cells(i.ColItem_Primary.Index).Value)
                                            Ketemu_2 = True
                                        End If
                                    End If
                                    'pastikan shift nya masih sama jika ubah status reject ke approved
                                    If Ketemu_1 = False And Ketemu_2 = False Then
                                        If Me.ApaShiftKonsisten(rowD.Cells(i.ColItem_Primary.Index).Value, row.Primary_Employee, Tbl3, Tbl4) = False Then
                                            vListPrimarySpl.Remove(rowD.Cells(i.ColItem_Primary.Index).Value)
                                            Ketemu_3 = True
                                        End If
                                    End If
                                End If
                            ElseIf vStatusSPL_Aft = EnumStatus.Rejected Then
                                If vStatusSPL_Bef = EnumStatus.WaitingForApproval Then
                                    'wait for approval > rejected # tidak ada batasan
                                ElseIf vStatusSPL_Bef = EnumStatus.Approved Then
                                    'approved > rejected # tidak bisa di rejected jika salah satu employee dari spl sudah di buat attendancenya
                                    If Me.ApaSudahAdaAttendance(rowD.Cells(i.ColItem_Primary.Index).Value, row.Primary_Employee, row.From, row._To, Tbl2) = True Then
                                        vListPrimarySpl.Remove(rowD.Cells(i.ColItem_Primary.Index).Value)
                                        Ketemu_1 = True
                                    End If
                                End If
                            End If

                            If Ketemu_1 = True Then
                                If ListMsg_1 = "" Then
                                    ListMsg_1 &= vbCrLf & "Attendance : " & rowD.Cells(i.ColItem_DocumentNo.Index).Value & " - Employee : " & row.Full_Name
                                Else
                                    If PrimarySPL_Bef <> rowD.Cells(i.ColItem_Primary.Index).Value Then
                                        ListMsg_1 &= vbCrLf & ", Attendance : " & rowD.Cells(i.ColItem_DocumentNo.Index).Value & " - Employee : " & row.Full_Name
                                    Else
                                        ListMsg_1 &= ", " & row.Full_Name
                                    End If
                                End If
                            End If
                            If Ketemu_2 = True Then
                                If ListMsg_2 = "" Then
                                    ListMsg_2 &= vbCrLf & "Overtime Order : " & rowD.Cells(i.ColItem_DocumentNo.Index).Value & " - Employee : " & row.Full_Name
                                Else
                                    If PrimarySPL_Bef <> rowD.Cells(i.ColItem_Primary.Index).Value Then
                                        ListMsg_2 &= ", Overtime Order : " & rowD.Cells(i.ColItem_DocumentNo.Index).Value & " - Employee : " & row.Full_Name
                                    Else
                                        ListMsg_2 &= ", " & row.Full_Name
                                    End If
                                End If
                            End If
                            If Ketemu_3 = True Then
                                If ListMsg_3 = "" Then
                                    ListMsg_3 &= vbCrLf & "Overtime Order : " & rowD.Cells(i.ColItem_DocumentNo.Index).Value & " - Employee : " & row.Full_Name
                                Else
                                    If PrimarySPL_Bef <> rowD.Cells(i.ColItem_Primary.Index).Value Then
                                        ListMsg_3 &= ", Overtime Order : " & rowD.Cells(i.ColItem_DocumentNo.Index).Value & " - Employee : " & row.Full_Name
                                    Else
                                        ListMsg_3 &= ", " & row.Full_Name
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next

            If ListMsg_1 <> "" Then
                MsgBox("Can't change status because this data is refering to attendance transaction." & ListMsg_1, MsgBoxStyle.Information, "Information")
            End If
            If ListMsg_2 <> "" Then
                MsgBox("Can't change status because this data is refering to other transaction." & ListMsg_2, MsgBoxStyle.Information, "Information")
            End If
            If ListMsg_3 <> "" Then
                MsgBox("Can't change status because this data inconsistent employee shift." & ListMsg_3, MsgBoxStyle.Information, "Information")
            End If

            If vListPrimarySpl.Count = 0 Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Function

    Public Function ApaSudahAdaPadaTransaksiLain(ByVal vPrimarySpl As Double, ByVal vPrimaryEmp As Double, _
    ByVal vDateFrom As Date, ByVal vDateTo As Date, ByVal Tbl As DseTrans_SPL.DataTable1DataTable) As Boolean
        Dim Tbl2 As New DseTrans_SPL.DataTable1DataTable
        Dim row As DseTrans_SPL.DataTable1Row

        Try
            Tbl2 = Tbl.Copy
            'Tbl2.DefaultView.RowFilter = "[Primary] <> " & vPrimarySpl & " AND [Primary Employee] = " & vPrimaryEmp & _
            '" AND [From] <= #" & vDateTo.Date & "# AND [To] >= #" & vDateFrom.Date & "# AND [Status] <> " & EnumStatus.Rejected
            Tbl2.DefaultView.RowFilter = "[Primary] <> " & vPrimarySpl & " AND [Primary Employee] = " & vPrimaryEmp & _
            " AND [Date] <= #" & vDateTo.Date & "# AND [Date] >= #" & vDateFrom.Date & "# AND [Status] <> " & EnumStatus.Rejected
            If Tbl2.DefaultView.Count > 0 Then
                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApaSudahAdaAttendance(ByVal vPrimarySpl As Double, ByVal vPrimaryEmp As Double, ByVal vDateFrom As Date, _
    ByVal vDateTo As Date, ByVal Tbl As DseTrans_SPL.tbltrans_attendanceDataTable) As Boolean
        Dim row2 As DseTrans_SPL.tbltrans_attendanceRow

        Try
            Tbl.DefaultView.RowFilter = "[Primary] <> " & vPrimarySpl & " AND [primary karyawan] = " & vPrimaryEmp & _
            " AND [start date] <= #" & vDateTo.Date & "# AND [end date] >= #" & vDateFrom.Date & "#"
            If Tbl.DefaultView.Count > 0 Then
                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApaShiftKonsisten(ByVal vPrimarySpl As Double, ByVal vPrimaryEmp As Double, _
    ByVal vTblShift As DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftDataTable, _
    ByVal vTblAllocate As DseTrans_SPL.allocate_employeeDataTable) As Boolean
        Dim TblAllocate As New DseTrans_SPL.allocate_employeeDataTable
        Dim rowAllocate As DseTrans_SPL.allocate_employeeRow
        Dim rowShift As DseTrans_SPL.tbltrans_spl_overtime_order_item_employee_shiftRow
        Dim DateFilter As Date = Now

        Try
            vTblAllocate.DefaultView.RowFilter = "[Primary] = " & vPrimaryEmp
            TblAllocate.Merge(vTblAllocate.DefaultView.ToTable)

            vTblShift.DefaultView.RowFilter = "[Primary Transaksi] = " & vPrimarySpl.ToString & " AND [Primary Employee] = " & vPrimaryEmp
            vTblShift.DefaultView.Sort = "[Primary Employee] ASC, [Date From] ASC"
            If vTblShift.DefaultView.Count > 0 Then
                For Each dv As DataRowView In vTblShift.DefaultView
                    rowShift = dv.Row
                    DateFilter = rowShift.Date_From
                    While DateFilter <= rowShift.Date_To
                        TblAllocate.DefaultView.RowFilter = "[start date] <= #" & DateFilter & "# AND [end date] >= #" & DateFilter & "#"
                        If TblAllocate.DefaultView.Count > 0 Then
                            'bandingkan employee shift pada spl masih sama dengan yang di allocate sekarang
                            rowAllocate = TblAllocate.DefaultView(0).Row
                            If rowShift.Primary_Shift <> rowAllocate.Primary_Workshift Then
                                Return False
                            End If
                        Else
                            'kalo di tblallocate berarti employee tsb tidak punya shift 
                            Return False
                        End If

                        DateFilter = DateFilter.AddDays(1)
                    End While
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_scalarTableAdapter
        Dim Dummy As System.Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTglTransaksi(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vPrimaryTrans As Double) As String
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_scalarTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim JmlDb As Nullable(Of Double)
        Try
            'cari tanggal sampai dan dari
            TanggalEnd = Me.ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = Me.ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            JmlDb = DTA.ScalarQuery_GetMaxUrutan(TanggalStart, TanggalEnd, vPrimaryTrans)
            If JmlDb.HasValue = True Then
                NoUrut = CStr(CInt(JmlDb.Value)) + 1
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            'pastikan no urut terdiri atas 6 huruf
            While NoUrut.Length < 6
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return NoUrut + "/" + "SPL" + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmTrans_SPL) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_orderTableAdapter
        Dim TBL As New DseTrans_SPL.tbltrans_spl_overtime_orderDataTable
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_orderRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mUrutan = .Urutan
                    Me.mDate = ._Date
                    Me.mReason = .Reason
                    Me.mAppliedBy = .Applied_By
                    Me.mAppliedByPosition = .Applied_By_Position
                    Me.mPrimaryApprovedBy = .Primary_Approved_By
                    Me.mApprovedTime = .Approved_time
                    Me.mDateFrom = .From
                    Me.mDateTo = ._To
                    Me.mTotalDuration = .Total_Duration
                    Me.mStatusNotes = .Status_Notes
                    Me.mStatus = .Status
                    Me.mRemarks = .Remarks

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp

                    Me.ObjItem.IsiTblItem(.Primary, vFrm.TblEmp, vFrm.TblSchSetup, vFrm.TblSchDet, vFrm.TblApp, vFrm.TblEmpShift, vFrm.TblAppOT)
                    Me.ObjItem.IsiDGVItemEmployee(vFrm.dgvEmployee, .Primary, vFrm, vFrm.TblEmp)
                    'Me.ObjItem.IsiDGVItemSchSetup(vFrm.dgvSchSetup, .Primary, vFrm, vFrm.TblSchSetup)
                    'Me.ObjItem.IsiDGVItemSchDetail(vFrm.dgvSchDetail, .Primary, vFrm, vFrm.TblSchDet)
                    'Me.ObjItem.IsiDGVItemSchEmp(vFrm.dgvSchEmp, .Primary, vFrm, vFrm.TblAppOT)
                    'Me.ObjItem.IsiDGVItemApproval(vFrm.dgvApproval, .Primary, vFrm, vFrm.TblApp)
                End With
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_scalarTableAdapter
        Dim dummy As Nullable(Of Long)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJml(vNoTransaksi)
            If dummy.HasValue Then
                If dummy.Value > 0 Then
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
            dta.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim d As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_scalarTableAdapter
        Dim n As New Nullable(Of Double)

        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            n = d.ScalarQuery_GetPrimary(vNoTransaksi)
            If n.HasValue Then
                Return n.Value
            End If

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            d.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksiByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_scalarTableAdapter
        Dim Dummy As String

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetNoTransaksi(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTimeStampByPrimary(ByVal vPrimary As Double) As Date
        Dim dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_scalarTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetTimestamp(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return "1/1/1950"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, _
   ByVal frm As FrmTrans_SPL) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_SimpanBaru, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If CInt(Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value) = CInt(Me.vPrimary_Antrian) Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, frm)
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal frm As FrmTrans_SPL) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_orderTableAdapter
        Dim BufferDoc As String = Me.mNoTransaksi

        Try
            If Me.ApakahDataBenar(frm, True) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring

                Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(0, 6))
                If DTA.Insert(Me.mNoTransaksi, Me.mUrutan, Me.mDate, Me.mReason, Me.mAppliedBy, Me.mAppliedByPosition, _
                Me.mPrimaryApprovedBy, Me.mApprovedTime, Me.mDateFrom, Me.mDateTo, Me.mTotalDuration, _
                 Me.mRemarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = Me.GetPrimary(Me.mNoTransaksi)
                    Me.mTimestamp = Me.GetTimeStampByPrimary(Me.mPrimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan item
                    Me.ObjItem.SimpanItem_UsingScript(Me.Primary, frm)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, _
   ByVal frm As FrmTrans_SPL) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If CInt(Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value) = CInt(Me.vPrimary_Antrian) Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode, frm)
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

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal frm As FrmTrans_SPL) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_orderTableAdapter
        Dim BufferDoc As String = Me.mNoTransaksi

        Try
            If Me.ApakahDataBenar(frm, False) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring

                Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(0, 6))
                If DTA.Update(Me.mNoTransaksi, Me.mUrutan, Me.mDate, Me.mReason, Me.mAppliedBy, Me.mAppliedByPosition, _
                Me.mPrimaryApprovedBy, Me.mApprovedTime, Me.mDateFrom, Me.mDateTo, Me.mTotalDuration, Me.mRemarks, _
                vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = Me.GetTimeStampByPrimary(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus item
                    Me.ObjItem.HapusItem(Me.mPrimary)

                    'simpan item
                    Me.ObjItem.SimpanItem_UsingScript(Me.Primary, frm)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, ByVal frm As FrmTrans_SPL) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_Hapus, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If CInt(Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value) = CInt(Me.vPrimary_Antrian) Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, vUserCode, frm)
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

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal frm As FrmTrans_SPL) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_orderTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahBolehHapus(vUserCode, vPrimary) = True Then
                DTA.Delete(vPrimary)
                Me.ObjItem.HapusItem(vPrimary)

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, ByVal vDGV As DataGridView, _
    ByVal vTBL As DseTrans_SPL.tbltrans_spl_overtime_order_fillDataTable, ByVal SearchCriteria As EnumSearch, _
    ByVal tsl As ToolStripLabel, ByVal vTanggalDari As Date, _
    ByVal vTanggalSampai As Date, ByVal vFrm As FrmTrans_SPL, Optional ByVal vJml As Integer = 0) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_fillTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearch.DocNo Then
                QueryResult = DTA.Fill_Like(vTBL, Criteria + "%", "%", #1/1/1950#, #1/1/2200#, "%", "%")
            ElseIf SearchCriteria = EnumSearch.Tgl Then
                QueryResult = DTA.Fill_Like(vTBL, "%", "%", vTanggalDari, vTanggalSampai, "%", "%")
            ElseIf SearchCriteria = EnumSearch.EmpCode Then
                QueryResult = DTA.Fill_Like(vTBL, "%", "%", #1/1/1950#, #1/1/2200#, Criteria + "%", "%")
            ElseIf SearchCriteria = EnumSearch.EmpName Then
                QueryResult = DTA.Fill_Like(vTBL, "%", "%", #1/1/1950#, #1/1/2200#, "%", Criteria + "%")
            ElseIf SearchCriteria = EnumSearch.Remarks Then
                QueryResult = DTA.Fill_Like(vTBL, "%", Criteria + "%", #1/1/1950#, #1/1/2200#, "%", "%")
            End If

            'update dgv
            vDGV.DataSource = vTBL.DefaultView
            vDGV.Refresh()
            If vDGV.ColumnCount > 0 Then
                vDGV.Columns(0).Visible = False
                vDGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                vDGV.Columns(3).DefaultCellStyle.Format = "dd MMM yyyy"
                vDGV.Columns(4).DefaultCellStyle.Format = "dd MMM yyyy"
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vJml = QueryResult

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal vCriteria As String, ByVal vDGV As DataGridView, ByVal vSearchCriteria As EnumSearch, _
    ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date) As Double
        Dim Dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_fillTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try

            Dta.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchCriteria = EnumSearch.DocNo Then
                Dummy = Dta.ScalarQuery_GetJmlLike(vCriteria + "%", "%", #1/1/1950#, #1/1/2200#, "%", "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSearch.Tgl Then
                Dummy = Dta.ScalarQuery_GetJmlLike("%", "%", vTanggalDari, vTanggalSampai, "%", "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSearch.EmpCode Then
                Dummy = Dta.ScalarQuery_GetJmlLike("%", "%", #1/1/1950#, #1/1/2200#, vCriteria + "%", "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSearch.EmpName Then
                Dummy = Dta.ScalarQuery_GetJmlLike("%", "%", #1/1/1950#, #1/1/2200#, "%", vCriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchCriteria = EnumSearch.Remarks Then
                Dummy = Dta.ScalarQuery_GetJmlLike("%", vCriteria + "%", #1/1/1950#, #1/1/2200#, "%", "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function UpdateStatusApproval(ByVal vPrimaryApp As Double, ByVal vStatus As ClsTrans_SPL.EnumStatus, _
    ByVal vPrimarySpl As Double) As Boolean
        Dim Dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_orderTableAdapter
        Dim UpdDate As DateTime = CDate("1/1/1950")

        Try
            If vPrimaryApp <> -1 Then
                UpdDate = Now
            End If

            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.UpdateQuery_StatusSPL(vPrimaryApp, UpdDate, vStatus, vPrimarySpl)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function UpdateStatusNotes(ByVal vStatusNotes As String, _
    ByVal vPrimarySpl As Double) As Boolean
        Dim Dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_orderTableAdapter

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.UpdateQuery_StatusNotesSPL(vStatusNotes, vPrimarySpl)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApproveOverBudget(ByVal vPrimaryUser) As Boolean
        Dim Dta As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_item_userTableAdapter
        Dim Retval As Nullable(Of Integer)

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Retval = Dta.ScalarQuery_GetApproveOverBudgetOT(vPrimaryUser)
            If Retval.HasValue = True Then
                If Retval.Value = 0 Then
                    Return False
                End If
            Else
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetStatus(ByVal vPrimary As String) As Integer
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_scalarTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try
            Dummy = DTA.ScalarQuery_GetStatus(vPrimary)
            If Dummy.HasValue = True Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Public Function GetStatusNotes(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_scalarTableAdapter
        Dim Dummy As String = ""

        Try
            Dummy = DTA.ScalarQuery_GetStatusNotes(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function ApaSudahDiBuatAttendance(ByVal vPrimaryEmp As Double, ByVal vDate As Date) As Boolean
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approved_otTableAdapter
        Dim Dummy As Nullable(Of Long)

        Try
            Dummy = DTA.ScalarQuery_GetJmlAttendanceItem(vPrimaryEmp, vDate.Date)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function Cetak(ByVal vPrimary As Double, ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vCrvDetail As CrystalDecisions.Windows.Forms.CrystalReportViewer, Optional ByVal vCreatePDFFile As Boolean = False) As Boolean
        Dim report As New CrpTrans_SPL
        Dim ds As New DseTrans_SPLCetak
        Dim Dta As New DseTrans_SPLCetakTableAdapters.DataTable1TableAdapter
        Dim Tbl As New DseTrans_SPLCetak.DataTable1DataTable
        Dim row As DseTrans_SPLCetak.DataTable1Row
        Dim TblC1 As DseTrans_SPLCetak.tbl_cetakDataTable = ds.tbl_cetak
        Dim rowC1 As DseTrans_SPLCetak.tbl_cetakRow
        Dim DtaC2 As New DseTrans_SPLCetakTableAdapters.tbl_cetak_employeeTableAdapter
        Dim TblC2 As DseTrans_SPLCetak.tbl_cetak_employeeDataTable = ds.tbl_cetak_employee
        Dim rowC2 As DseTrans_SPLCetak.tbl_cetak_employeeRow
        Dim OvertimeOrderNo As String = ""

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            DtaC2.Connection.ConnectionString = Core_Module.GetConstring

            Dta.Fill(Tbl, vPrimary)
            DtaC2.Fill(TblC2, vPrimary)

            If Tbl.Count > 0 Then
                Tbl.DefaultView.RowFilter = ""
                Tbl.DefaultView.Sort = "[Date Item] ASC, [Shift Desc] ASC, [OT Type] ASC"
                For Each dv As DataRowView In Tbl.DefaultView
                    row = dv.Row
                    rowC1 = TblC1.Newtbl_cetakRow
                    With rowC1
                        OvertimeOrderNo = row.No_Transaksi
                        .No_Transaksi = row.No_Transaksi
                        ._Date = row._Date
                        .Reason = row.Reason
                        .Overtime_Period = Format(row.From.Date, "dd MMM yyyy") & " - " & Format(row._To.Date, "dd MMM yyyy")
                        .Total_Duration = Me.ObjCD.Convert_ToString(row.Total_Duration)
                        .Applied_By = row.Applied_By
                        .Applied_By_Position = row.Applied_By_Position
                        .Approved_By = row.Approved_By
                        .Approved_By_Position = row.Approved_By_Position
                        .Day_of_Week = Me.ObjCD.GetHari_English(row.Date_Item.DayOfWeek)
                        .Date_Item = row.Date_Item
                        .Shift = row.Shift_Desc
                        .Start = Me.ObjCD.Convert_ToString(row.Start)
                        If row.Break_In > 0 Then
                            .Break_In = Me.ObjCD.Convert_ToString(row.Break_In)
                        End If
                        If row.Break_Out > 0 Then
                            .Break_Out = Me.ObjCD.Convert_ToString(row.Break_Out)
                        End If
                        ._End = Me.ObjCD.Convert_ToString(row._End)
                        .Duration = Me.ObjCD.Convert_ToString(row.Duration)
                    End With
                    TblC1.Rows.Add(rowC1)
                Next
            End If

            report.SetDataSource(ds)
            'Report.SetParameterValue(Report.Parameter_LoanType.ParameterFieldName, vTypeLoan)
            vCrv.ReportSource = report
            vCrv.Show()

            If vCreatePDFFile = True Then
                If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Overtime Order - " + OvertimeOrderNo.Replace("/", "-") + ".pdf") = True Then
                    'hapus filenya
                    My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\Overtime Order - " + OvertimeOrderNo.Replace("/", "-") + ".pdf", FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                report.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Application.StartupPath & "\Overtime Order - " + OvertimeOrderNo.Replace("/", "-") + ".pdf")

                'dispose report karena saat kirim email slip gaji terjadi load report failed
                report.Dispose()
            End If

            Me.CetakDetail(vPrimary, vCrvDetail, vCreatePDFFile)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CetakDetail(ByVal vPrimary As Double, ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    Optional ByVal vCreatePDFFile As Boolean = False) As Boolean
        Dim report As New CrpTrans_SPL_DetailEmp
        Dim ds As New DseTrans_SPLCetak
        Dim Dta As New DseTrans_SPLCetakTableAdapters.tbl_cetak_2TableAdapter
        Dim Tbl As New DseTrans_SPLCetak.tbl_cetak_2DataTable
        Dim row As DseTrans_SPLCetak.tbl_cetak_2Row
        Dim TblC1 As DseTrans_SPLCetak.tbl_cetakDataTable = ds.tbl_cetak
        Dim rowC1 As DseTrans_SPLCetak.tbl_cetakRow
        Dim OvertimeOrderNo As String = ""

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring

            Dta.Fill(Tbl, vPrimary)

            Dim JmlLoop1, JmlLoop2 As Integer
            If Tbl.Count > 0 Then
                For Each row In Tbl
                    JmlLoop1 = 2
                    JmlLoop2 = 2

                    If row.OT_Early_Approved > 0 And row.OT_Aprroved > 0 Then
                        JmlLoop1 = 1
                        JmlLoop2 = 2
                    Else
                        If row.OT_Early_Approved > 0 Then
                            JmlLoop1 = 1
                            JmlLoop2 = 1
                        ElseIf row.OT_Aprroved > 0 Then
                            JmlLoop1 = 2
                            JmlLoop2 = 2
                        End If
                    End If

                    OvertimeOrderNo = row.No_Transaksi

                    For i As Integer = JmlLoop1 To JmlLoop2
                        rowC1 = TblC1.Newtbl_cetakRow
                        With rowC1
                            .Urutan = JmlLoop1
                            .No_Transaksi = row.No_Transaksi
                            ._Date = row._Date
                            .Emp_Code = row.Emp_Code
                            .Emp_NIK = row.Emp_NIK
                            .Emp_Name = row.Emp_Name
                            .Dept_Code = row.Dept_Code
                            .Dept_Desc = row.Dept_Desc

                            .Day_of_Week = Me.ObjCD.GetHari_English(row.Date_Day.DayOfWeek)
                            .Date_Item = row.Date_Day

                            If i = 1 Then
                                .Start = Me.ObjCD.Convert_ToString(row.Start_Time_Early)
                                If row.Break_Out_Time_Early > 0 Then
                                    .Break_Out = Me.ObjCD.Convert_ToString(row.Break_Out_Time_Early)
                                End If
                                If row.Break_In_Time_Early > 0 Then
                                    .Break_In = Me.ObjCD.Convert_ToString(row.Break_In_Time_Early)
                                End If
                                ._End = Me.ObjCD.Convert_ToString(row.End_Time_Early)
                                .Duration = Me.ObjCD.Convert_ToString(row.OT_Early_Approved)
                            ElseIf i = 2 Then
                                .Start = Me.ObjCD.Convert_ToString(row.Start_Time)
                                If row.Break_Out_Time > 0 Then
                                    .Break_Out = Me.ObjCD.Convert_ToString(row.Break_Out_Time)
                                End If
                                If row.Break_In_Time > 0 Then
                                    .Break_In = Me.ObjCD.Convert_ToString(row.Break_In_Time)
                                End If
                                ._End = Me.ObjCD.Convert_ToString(row.End_Time)
                                .Duration = Me.ObjCD.Convert_ToString(row.OT_Aprroved)
                            End If

                            .Applied_By = row.Applied_By
                            .Applied_By_Position = row.Applied_By_Position
                            .Approved_By = row.Approved_By
                            .Approved_By_Position = row.Approved_By_Position

                            .Total_Duration = Me.ObjCD.Convert_ToString(row.Total_OT_Emp)
                        End With
                        TblC1.Rows.Add(rowC1)
                    Next
                Next
            End If

            report.SetDataSource(ds)
            'Report.SetParameterValue(Report.Parameter_LoanType.ParameterFieldName, vTypeLoan)
            vCrv.ReportSource = report
            vCrv.Show()

            If vCreatePDFFile = True Then
                If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Overtime Order Detail - " + OvertimeOrderNo.Replace("/", "-") + ".pdf") = True Then
                    'hapus filenya
                    My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\Overtime Order Detail - " + OvertimeOrderNo.Replace("/", "-") + ".pdf", FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                report.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Application.StartupPath & "\Overtime Order Detail - " + OvertimeOrderNo.Replace("/", "-") + ".pdf")

                'dispose report karena saat kirim email slip gaji terjadi load report failed
                report.Dispose()
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function UpdateLastDeliveryApprovalReq(ByVal vPrimarySPL As Double, ByVal vPrimaryUser As Double) As Boolean
        Dim dta As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approvalTableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_LastDelivery(Now, vPrimarySPL, vPrimaryUser)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetTotalHoursEmployee(ByVal vPrimaryEmp As Double, ByVal TblEmpOT As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otDataTable) As Double
        Dim row As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_otRow
        Dim TotalHours As Double = 0

        Try
            TblEmpOT.DefaultView.RowFilter = "[Primary Employee] = " & vPrimaryEmp
            For Each dv As DataRowView In TblEmpOT.DefaultView
                row = dv.Row
                TotalHours += row.OT_Early_Approved + row.OT_Aprroved
            Next

            Return TotalHours
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return TotalHours
        End Try
    End Function
End Class
