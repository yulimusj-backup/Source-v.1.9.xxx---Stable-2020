Option Explicit On
Option Strict Off

Partial Public Class ClsMasterWorkShift

    Public Enum MetodeDeleteAllocation As Integer
        AllData
        SatuPerSatu
    End Enum

    Public Sub KelolaDataKaryawan(ByVal TBLKaryawan As DseMasterWorkshift_Allocation.tblmaster_karyawanDataTable)
        Dim DTAKaryawan As New DseMasterWorkshift_AllocationTableAdapters.tblmaster_karyawanTableAdapter
        Try
            DTAKaryawan.Connection.ConnectionString = Core_Module.GetConstring
            TBLKaryawan.Rows.Clear()
            DTAKaryawan.Fill(TBLKaryawan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAKaryawan.Dispose()
        End Try
    End Sub

    Public Sub KelolaDataKaryawan(ByVal TBLKaryawan As DseMasterWorkshift_Allocation.tblmaster_karyawanDataTable, _
    ByVal ListDataFilter As List(Of Double), ByVal DGVEmployee As DataGridView, ByVal TBL_JP As DseMasterWorkshift_Allocation.tblmaster_jobpositionDataTable, _
    ByVal Clb_JP As CheckedListBox, ByVal TBL_Dept As DseMasterWorkshift_Allocation.tblmaster_departemenDataTable, _
    ByVal Clb_dept As CheckedListBox, ByVal TBL_Formula As DseMasterWorkshift_Allocation.tblmaster_salary_formula1DataTable, _
    ByVal Clb_Formula As CheckedListBox, ByVal frm As Core_Module.FrmMasterWorkShift_Allocate)
        Dim DTAKaryawan As New DseMasterWorkshift_AllocationTableAdapters.tblmaster_karyawanTableAdapter
        Dim RowKaryawan As DseMasterWorkshift_Allocation.tblmaster_karyawanRow
        Dim DTAMasterSalary As New DseMasterWorkshift_AllocationTableAdapters.tblmaster_salary_formulaTableAdapter
        Dim TBLMasterSalary As New DseMasterWorkshift_Allocation.tblmaster_salary_formulaDataTable
        Dim RowMasterSalary As DseMasterWorkshift_Allocation.tblmaster_salary_formulaRow
        Dim EnumTipeDataFilter As Double
        Dim ij As Integer = 0
        Dim Ketemu As Boolean = False
        Dim Primary As Double = -1
        Dim RV As Data.DataRowView
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Try
            DTAKaryawan.Connection.ConnectionString = Core_Module.GetConstring
            DTAMasterSalary.Connection.ConnectionString = DTAKaryawan.Connection.ConnectionString

            TBLKaryawan.Rows.Clear()
            DTAKaryawan.Fill(TBLKaryawan)

            If ListDataFilter.Count > 0 Then
                For Each EnumTipeDataFilter In ListDataFilter
                    If EnumTipeDataFilter = Core_Module.FrmMasterWorkShift_Allocate.EnumLaporan.Departemen Then
                        ij = 0
                        While ij <= TBLKaryawan.Rows.Count - 1
                            RowKaryawan = TBLKaryawan.Rows(ij)
                            Ketemu = False
                            If Clb_dept.CheckedItems.Count > 0 Then
                                For Each RV In Clb_dept.CheckedItems
                                    Primary = RV.Item("Primary")
                                    If RowKaryawan.primary_departemen = Primary Then
                                        Ketemu = True
                                        Exit For
                                    End If
                                Next
                                If Ketemu = False Then
                                    TBLKaryawan.Rows.RemoveAt(ij)
                                Else
                                    ij += 1
                                End If
                            Else
                                ij += 1
                            End If
                        End While
                    End If
                    If EnumTipeDataFilter = Core_Module.FrmMasterWorkShift_Allocate.EnumLaporan.Employee Then
                        ij = 0
                        DGVEmployee.AllowUserToAddRows = False
                        DGVEmployee.AllowUserToDeleteRows = False
                        While ij <= TBLKaryawan.Rows.Count - 1
                            RowKaryawan = TBLKaryawan.Rows(ij)
                            Ketemu = False
                            If DGVEmployee.Rows.Count > 0 Then
                                For Each RowDGV In DGVEmployee.Rows
                                    Primary = RowDGV.Cells(frm.ColPrimaryCode.Index).Value
                                    If RowKaryawan.Primary = Primary Then
                                        Ketemu = True
                                        Exit For
                                    End If
                                Next
                            End If
                            If Ketemu = False Then
                                TBLKaryawan.Rows.RemoveAt(ij)
                            Else
                                ij += 1
                            End If
                        End While
                        DGVEmployee.AllowUserToAddRows = True
                        DGVEmployee.AllowUserToDeleteRows = True
                    End If
                    If EnumTipeDataFilter = Core_Module.FrmMasterWorkShift_Allocate.EnumLaporan.JobPosition Then
                        ij = 0
                        While ij <= TBLKaryawan.Rows.Count - 1
                            RowKaryawan = TBLKaryawan.Rows(ij)
                            Ketemu = False
                            If Clb_JP.CheckedItems.Count > 0 Then
                                For Each RV In Clb_JP.CheckedItems
                                    Primary = RV.Item("primary")
                                    If RowKaryawan.primary_position = Primary Then
                                        Ketemu = True
                                        Exit For
                                    End If
                                Next
                                If Ketemu = False Then
                                    TBLKaryawan.Rows.RemoveAt(ij)
                                Else
                                    ij += 1
                                End If
                            Else
                                ij += 1
                            End If
                        End While
                    End If
                    If EnumTipeDataFilter = Core_Module.FrmMasterWorkShift_Allocate.EnumLaporan.SalaryFormula Then
                        'MsgBox("Salary Formula masih dipending", MsgBoxStyle.Information, "Information")
                        DTAMasterSalary.Fill(TBLMasterSalary)
                        ij = 0
                        While ij <= TBLKaryawan.Rows.Count - 1
                            RowKaryawan = TBLKaryawan.Rows(ij)
                            Ketemu = False
                            Filter = "[Primary Karyawan] = " + RowKaryawan.Primary.ToString + ""
                            TBLMasterSalary.DefaultView.RowFilter = Filter
                            If TBLMasterSalary.DefaultView.Count > 0 Then
                                RowMasterSalary = TBLMasterSalary.DefaultView(0).Row
                                If Clb_Formula.CheckedItems.Count > 0 Then
                                    For Each RV In Clb_Formula.CheckedItems
                                        Primary = RV.Item("primary")
                                        If RowMasterSalary.primary = Primary Then
                                            Ketemu = True
                                            Exit For
                                        End If
                                    Next
                                    If Ketemu = False Then
                                        TBLKaryawan.Rows.RemoveAt(ij)
                                    Else
                                        ij += 1
                                    End If
                                Else
                                    ij += 1
                                End If
                            Else
                                TBLKaryawan.Rows.RemoveAt(ij)
                            End If
                        End While
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAKaryawan.Dispose()
            DTAMasterSalary.Dispose()
            TBLMasterSalary.Dispose()
        End Try
    End Sub

    Public Sub FilterDataAllocation(ByVal TBLKaryawan As DseMasterWorkshift_Allocation.tblmaster_karyawanDataTable, _
    ByVal TBL_Allocate As DseMasterWorkshift_Allocation.tblmaster_workshift_itemallocateDataTable, _
    ByVal dtpend As Date, ByVal dtpstart As Date, ByVal OutputDGV As Core_Module.DseMasterWorkshift_Allocation.Output_DGVDataTable)
        Dim DTA_Allocate As New DseMasterWorkshift_AllocationTableAdapters.tblmaster_workshift_itemallocateTableAdapter
        Dim Row_Allocate As DseMasterWorkshift_Allocation.tblmaster_workshift_itemallocateRow
        Dim TBL_UnSelection As New DseMasterWorkshift_Allocation.tblmaster_workshift_itemallocateDataTable
        Dim RowKaryawan As DseMasterWorkshift_Allocation.tblmaster_karyawanRow
        Dim Row_Output As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        Dim Filter As String = ""
        Dim Filter_2 As String = ""
        Dim DV As Data.DataRowView
        Dim ij As Integer = 0

        Try
            dtpstart = Format(dtpstart, "MM/dd/yyyy") + " " + "00:00:00"
            dtpend = Format(dtpend, "MM/dd/yyyy") + " " + "23:59:59"

            DTA_Allocate.Connection.ConnectionString = Core_Module.GetConstring
            TBL_Allocate.Rows.Clear()
            DTA_Allocate.FillBy_DurationStartEndDateAllocate(TBL_Allocate, dtpend, dtpstart)
            OutputDGV.Rows.Clear()

            If TBLKaryawan.Rows.Count > 0 Then
                For Each RowKaryawan In TBLKaryawan.Rows
                    Filter = "[primary entity] = " + RowKaryawan.Primary.ToString + ""
                    TBL_Allocate.DefaultView.RowFilter = Filter
                    If TBL_Allocate.DefaultView.Count > 0 Then
                        For i As Integer = 0 To TBL_Allocate.DefaultView.Count - 1
                            Row_Allocate = TBL_Allocate.DefaultView(i).Row
                            Row_Output = OutputDGV.NewOutput_DGVRow
                            With Row_Output
                                .Employee_Code = Row_Allocate.Employee_Code
                                .Employee_Name = Row_Allocate.Employee_Name
                                .Workshift_Code = Row_Allocate.Workshift_Code
                                .Workshift_Name = Row_Allocate.Workshift_Description
                                .Start_Date = Row_Allocate.start_date
                                .End_Date = Row_Allocate.end_date
                                .Remarks = Row_Allocate.remarks
                                .Primary_WS = Row_Allocate.primary_workshift
                                .Primary_Entity = Row_Allocate.primary_entity
                                .Tipe_Entity = Row_Allocate.tipe_entity
                                .Primary = Row_Allocate.primary
                            End With
                            OutputDGV.Rows.Add(Row_Output)
                        Next
                    End If
                Next

                '1. cek ada karyawan yang ketinggal belum di allocate
                For Each RowKaryawan In TBLKaryawan.Rows
                    Filter = "[Primary Entity] = " + RowKaryawan.Primary.ToString + ""
                    OutputDGV.DefaultView.RowFilter = Filter
                    If OutputDGV.DefaultView.Count = 0 Then
                        'Input DGV Allocation = nilai default, belum punya workshift
                        Row_Output = OutputDGV.NewOutput_DGVRow
                        With Row_Output
                            .Employee_Code = RowKaryawan.Code
                            .Employee_Name = RowKaryawan.Full_Name
                            .Workshift_Code = ""
                            .Workshift_Name = ""
                            .Start_Date = dtpstart
                            .End_Date = dtpend
                            .Remarks = ""
                            .Primary_Entity = RowKaryawan.Primary
                            .Primary_WS = -1
                            .Tipe_Entity = 0
                            .Primary = -1
                        End With
                        OutputDGV.Rows.Add(Row_Output)
                    End If
                Next
            End If

            OutputDGV.DefaultView.RowFilter = ""
            TBL_Allocate.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Allocate.Dispose()
        End Try
    End Sub

    'Fungsi Filter Data
    Public Sub FilterDataAllocation_Obs(ByVal TBLKaryawan As DseMasterWorkshift_Allocation.tblmaster_karyawanDataTable, _
    ByVal TBL_Allocate As DseMasterWorkshift_Allocation.tblmaster_workshift_itemallocateDataTable, _
    ByVal dtpend As Date, ByVal dtpstart As Date, ByVal OutputDGV As Core_Module.DseMasterWorkshift_Allocation.Output_DGVDataTable)
        Dim DTA_Allocate As New DseMasterWorkshift_AllocationTableAdapters.tblmaster_workshift_itemallocateTableAdapter
        Dim Row_Allocate As DseMasterWorkshift_Allocation.tblmaster_workshift_itemallocateRow
        Dim TBL_UnSelection As New DseMasterWorkshift_Allocation.tblmaster_workshift_itemallocateDataTable
        Dim RowKaryawan As DseMasterWorkshift_Allocation.tblmaster_karyawanRow
        Dim Row_Output As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        Dim Filter As String = ""
        Dim Filter_2 As String = ""
        Dim DV As Data.DataRowView
        Dim ij As Integer = 0
        Try
            DTA_Allocate.Connection.ConnectionString = Core_Module.GetConstring
            TBL_Allocate.Rows.Clear()
            DTA_Allocate.Fill(TBL_Allocate)
            OutputDGV.Rows.Clear()

            dtpstart = Format(dtpstart, "MM/dd/yyyy") + " " + "00:00:00"
            dtpend = Format(dtpend, "MM/dd/yyyy") + " " + "23:59:59"
            DTA_Allocate.FillBy_NotDurationStartEndAllocate(TBL_UnSelection, dtpend, dtpstart)
            If TBLKaryawan.Rows.Count > 0 Then
                For Each RowKaryawan In TBLKaryawan.Rows
                    Filter = "[primary entity] = " + RowKaryawan.Primary.ToString + ""
                    TBL_Allocate.DefaultView.RowFilter = ""
                    TBL_Allocate.DefaultView.RowFilter = Filter
                    If TBL_Allocate.DefaultView.Count > 0 Then
                        ij = 0
                        While ij <= TBL_Allocate.DefaultView.Count - 1
                            DV = TBL_Allocate.DefaultView(ij)
                            Row_Allocate = DV.Row
                            Filter_2 = "[primary] = " + Row_Allocate.primary.ToString + ""
                            TBL_UnSelection.DefaultView.RowFilter = ""
                            TBL_UnSelection.DefaultView.RowFilter = Filter_2
                            If TBL_UnSelection.DefaultView.Count > 0 Then
                                TBL_Allocate.Rows.Remove(Row_Allocate)
                            Else
                                'Input DGV Allocation
                                Row_Output = OutputDGV.NewOutput_DGVRow
                                With Row_Output
                                    .Employee_Code = Row_Allocate.Employee_Code
                                    .Employee_Name = Row_Allocate.Employee_Name
                                    .Workshift_Code = Row_Allocate.Workshift_Code
                                    .Workshift_Name = Row_Allocate.Workshift_Description
                                    .Start_Date = Row_Allocate.start_date
                                    .End_Date = Row_Allocate.end_date
                                    .Remarks = Row_Allocate.remarks
                                    .Primary_WS = Row_Allocate.primary_workshift
                                    .Primary_Entity = Row_Allocate.primary_entity
                                    .Tipe_Entity = Row_Allocate.tipe_entity
                                    .Primary = Row_Allocate.primary
                                End With
                                OutputDGV.Rows.Add(Row_Output)
                                ij += 1
                            End If
                        End While
                    Else
                        'Input DGV Allocation = nilai default, belum punya workshift
                        Row_Output = OutputDGV.NewOutput_DGVRow
                        With Row_Output
                            .Employee_Code = RowKaryawan.Code
                            .Employee_Name = RowKaryawan.Full_Name
                            .Workshift_Code = ""
                            .Workshift_Name = ""
                            .Start_Date = dtpstart
                            .End_Date = dtpend
                            .Remarks = ""
                            .Primary_Entity = RowKaryawan.Primary
                            .Primary_WS = -1
                            .Tipe_Entity = 0
                            .Primary = -1
                        End With
                        OutputDGV.Rows.Add(Row_Output)
                    End If
                Next
                '1. cek ada karyawan yang ketinggal belum di allocate
                For Each RowKaryawan In TBLKaryawan.Rows
                    Filter = "[Primary Entity] = " + RowKaryawan.Primary.ToString + ""
                    OutputDGV.DefaultView.RowFilter = Filter
                    If OutputDGV.DefaultView.Count > 0 Then
                    Else
                        'Input DGV Allocation = nilai default, belum punya workshift
                        Row_Output = OutputDGV.NewOutput_DGVRow
                        With Row_Output
                            .Employee_Code = RowKaryawan.Code
                            .Employee_Name = RowKaryawan.Full_Name
                            .Workshift_Code = ""
                            .Workshift_Name = ""
                            .Start_Date = dtpstart
                            .End_Date = dtpend
                            .Remarks = ""
                            .Primary_Entity = RowKaryawan.Primary
                            .Primary_WS = -1
                            .Tipe_Entity = 0
                            .Primary = -1
                        End With
                        OutputDGV.Rows.Add(Row_Output)
                    End If
                Next
            End If
            OutputDGV.DefaultView.RowFilter = ""
            TBL_Allocate.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Allocate.Dispose()
            TBL_UnSelection.Dispose()
        End Try
    End Sub

    Public Sub KeyInDGV_Allocation(ByVal OutputDGV As Core_Module.DseMasterWorkshift_Allocation.Output_DGVDataTable, _
    ByVal DGVAllocate As DataGridView, ByVal i As Core_Module.FrmMasterWorkShift_Allocate)
        Dim RowOut As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        Dim RowOut_Lanjutan As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        Dim OutputDGV_Dummy As New Core_Module.DseMasterWorkshift_Allocation.Output_DGVDataTable
        Dim RowOut_Dummy As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        'Dim RowOut_Dummy_1 As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        Dim RowDGV As DataGridViewRow
        Dim RV As Data.DataRowView
        Dim ijk As Integer = 0
        'Dim AdaTambahan_Baris As Boolean = False
        Try
            DGVAllocate.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGVAllocate.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            DGVAllocate.Rows.Clear()
            If OutputDGV.Rows.Count > 0 Then
                OutputDGV.DefaultView.Sort = "[Employee Code], [Start Date], [End Date], [Workshift Code]"
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                '------Optimasi Koding 28 Jan 2015----------
                For idx As Integer = 0 To OutputDGV.DefaultView.Count - 1
                    RowOut = OutputDGV.DefaultView(idx).Row
                    ijk = idx - 1
                    With RowOut
                        If idx = 0 Then
                            'Data Awal
                            'a. Cek start date nya sama ga dengan dtp start
                            If i.DtpStart.Value.Date < RowOut.Start_Date Then
                                'Input Baris Awal Bolong
                                RowOut_Dummy = OutputDGV_Dummy.NewOutput_DGVRow
                                With RowOut_Dummy
                                    .Employee_Code = RowOut.Employee_Code
                                    .Employee_Name = RowOut.Employee_Name
                                    .Workshift_Code = ""
                                    .Workshift_Name = ""
                                    .Start_Date = i.DtpStart.Value.Date
                                    .End_Date = RowOut.Start_Date - TimeSpan.FromDays(1)
                                    .Remarks = ""
                                    .Primary_Entity = RowOut.Primary_Entity
                                    .Primary_WS = -1
                                    .Tipe_Entity = 0
                                    .Primary = -1
                                End With
                                OutputDGV_Dummy.Rows.Add(RowOut_Dummy)
                            End If
                            'Input baris Awal
                            RowOut_Dummy = OutputDGV_Dummy.NewOutput_DGVRow
                            RowOut_Dummy.Employee_Code = .Employee_Code
                            RowOut_Dummy.Employee_Name = .Employee_Name
                            RowOut_Dummy.Start_Date = .Start_Date
                            RowOut_Dummy.End_Date = .End_Date
                            RowOut_Dummy.Workshift_Code = .Workshift_Code
                            RowOut_Dummy.Workshift_Name = .Workshift_Name
                            RowOut_Dummy.Remarks = .Remarks
                            RowOut_Dummy.Primary_Entity = .Primary_Entity
                            RowOut_Dummy.Primary_WS = .Primary_WS
                            RowOut_Dummy.Tipe_Entity = .Tipe_Entity
                            RowOut_Dummy.Primary = .Primary
                            OutputDGV_Dummy.Rows.Add(RowOut_Dummy)
                        Else
                            'Cek Pergantian Employee
                            'a. Cek dulu baris sebelumnya, lakukan pengecekan enddate-nya sesuai ga dengan DTPEndDate, jika tidak sesuai maka munculkan baris baru..... 
                            RowOut_Lanjutan = OutputDGV.DefaultView(ijk).Row
                            If RowOut_Lanjutan.Primary_Entity <> RowOut.Primary_Entity Then
                                If i.DtpEnd.Value.Date > RowOut_Lanjutan.End_Date Then
                                    'Input Baris Akhir Bolong
                                    RowOut_Dummy = OutputDGV_Dummy.NewOutput_DGVRow
                                    With RowOut_Dummy
                                        .Employee_Code = RowOut_Lanjutan.Employee_Code
                                        .Employee_Name = RowOut_Lanjutan.Employee_Name
                                        .Workshift_Code = ""
                                        .Workshift_Name = ""
                                        .Start_Date = RowOut_Lanjutan.End_Date + TimeSpan.FromDays(1)
                                        .End_Date = i.DtpEnd.Value.Date
                                        .Remarks = ""
                                        .Primary_Entity = RowOut_Lanjutan.Primary_Entity
                                        .Primary_WS = -1
                                        .Tipe_Entity = 0
                                        .Primary = -1
                                    End With
                                    OutputDGV_Dummy.Rows.Add(RowOut_Dummy)
                                End If
                                'b. Cek baris saat ini, lakukan pengecekan startdate-nya sesuai ga dengan DTPstartdate, jika tidak sesuai maka munculkan baris baru..... 
                                If i.DtpStart.Value.Date < RowOut.Start_Date Then
                                    'Input Baris Awal Bolong
                                    RowOut_Dummy = OutputDGV_Dummy.NewOutput_DGVRow
                                    With RowOut_Dummy
                                        .Employee_Code = RowOut.Employee_Code
                                        .Employee_Name = RowOut.Employee_Name
                                        .Workshift_Code = ""
                                        .Workshift_Name = ""
                                        .Start_Date = i.DtpStart.Value.Date
                                        .End_Date = RowOut.Start_Date - TimeSpan.FromDays(1)
                                        .Remarks = ""
                                        .Primary_Entity = RowOut.Primary_Entity
                                        .Primary_WS = -1
                                        .Tipe_Entity = 0
                                        .Primary = -1
                                    End With
                                    OutputDGV_Dummy.Rows.Add(RowOut_Dummy)
                                End If
                            End If
                            'Masuk datanya
                            RowOut_Dummy = OutputDGV_Dummy.NewOutput_DGVRow
                            RowOut_Dummy.Employee_Code = .Employee_Code
                            RowOut_Dummy.Employee_Name = .Employee_Name
                            RowOut_Dummy.Start_Date = .Start_Date
                            RowOut_Dummy.End_Date = .End_Date
                            RowOut_Dummy.Workshift_Code = .Workshift_Code
                            RowOut_Dummy.Workshift_Name = .Workshift_Name
                            RowOut_Dummy.Remarks = .Remarks
                            RowOut_Dummy.Primary_Entity = .Primary_Entity
                            RowOut_Dummy.Primary_WS = .Primary_WS
                            RowOut_Dummy.Tipe_Entity = .Tipe_Entity
                            RowOut_Dummy.Primary = .Primary
                            OutputDGV_Dummy.Rows.Add(RowOut_Dummy)
                            'Cek ada yang bolong di antara periode end date baris ini dan start date baris selanjutnya
                            If ijk <= OutputDGV.Rows.Count - 1 Then
                                RowOut_Lanjutan = OutputDGV.DefaultView(ijk).Row
                                If RowOut_Lanjutan.Primary_Entity = RowOut.Primary_Entity Then
                                    If RowOut.Start_Date - TimeSpan.FromDays(1) >= RowOut_Lanjutan.End_Date + TimeSpan.FromDays(1) Then
                                        'Input DGV Allocation = nilai default, belum punya workshift
                                        RowOut_Dummy = OutputDGV_Dummy.NewOutput_DGVRow
                                        With RowOut_Dummy
                                            .Employee_Code = RowOut.Employee_Code
                                            .Employee_Name = RowOut.Employee_Name
                                            .Workshift_Code = ""
                                            .Workshift_Name = ""
                                            .Start_Date = RowOut_Lanjutan.End_Date + TimeSpan.FromDays(1)
                                            .End_Date = RowOut.Start_Date - TimeSpan.FromDays(1)
                                            .Remarks = ""
                                            .Primary_Entity = RowOut.Primary_Entity
                                            .Primary_WS = -1
                                            .Tipe_Entity = 0
                                            .Primary = -1
                                        End With
                                        OutputDGV_Dummy.Rows.Add(RowOut_Dummy)
                                    End If
                                End If
                            End If
                            '----------------------------------------------------------------------------------------------
                        End If
                    End With
                Next
                '----------------------------------------------------------------------------------------------
                'Pindah in OutputDGV_Dummy ke OutputDGV
                OutputDGV.DefaultView.Sort = ""
                OutputDGV.Rows.Clear()
                DGVAllocate.Rows.Clear()
                OutputDGV_Dummy.DefaultView.Sort = "[Employee Code], [Start Date], [End Date], [Workshift Code]"
                For Each RV In OutputDGV_Dummy.DefaultView
                    RowOut_Dummy = RV.Row
                    RowOut = OutputDGV.NewOutput_DGVRow
                    For index As Integer = 0 To OutputDGV.Columns.Count - 1
                        RowOut.Item(index) = RowOut_Dummy.Item(index)
                    Next
                    OutputDGV.Rows.Add(RowOut)
                    'Key In DGV Allocate
                    With RowOut_Dummy
                        DGVAllocate.Rows.Add()
                        RowDGV = DGVAllocate.Rows(DGVAllocate.Rows.Count - 1)
                        RowDGV.Cells(i.ColPrimary.Index).Value = .Primary
                        RowDGV.Cells(i.ColEntityTipe.Index).Value = .Tipe_Entity
                        RowDGV.Cells(i.ColPrimaryEntity.Index).Value = .Primary_Entity
                        RowDGV.Cells(i.ColPrimaryWorkShift.Index).Value = .Primary_WS
                        RowDGV.Cells(i.ColRemarks.Index).Value = .Remarks
                        RowDGV.Cells(i.ColWorkShiftName.Index).Value = .Workshift_Name
                        RowDGV.Cells(i.ColWorkshiftCode.Index).Value = .Workshift_Code
                        RowDGV.Cells(i.ColEndDate.Index).Value = Format(.End_Date, "dd MMM yyyy")
                        RowDGV.Cells(i.ColStartDate.Index).Value = Format(.Start_Date, "dd MMM yyyy")
                        RowDGV.Cells(i.ColEmployeeName.Index).Value = .Employee_Name
                        RowDGV.Cells(i.ColEmployeeCode.Index).Value = .Employee_Code
                    End With
                Next
                '----------------------------------------------------------------------------------------------
                'Data Akhir
                'b. Cek end date nya sama ga dengan dtp end
                RowOut = OutputDGV.Rows(OutputDGV.Rows.Count - 1)
                If i.DtpEnd.Value.Date > RowOut.End_Date Then
                    'Input Baris Akhir Bolong
                    RowOut_Dummy = OutputDGV.NewOutput_DGVRow
                    With RowOut_Dummy
                        .Employee_Code = RowOut.Employee_Code
                        .Employee_Name = RowOut.Employee_Name
                        .Workshift_Code = ""
                        .Workshift_Name = ""
                        .Start_Date = RowOut.End_Date + TimeSpan.FromDays(1)
                        .End_Date = i.DtpEnd.Value.Date
                        .Remarks = ""
                        .Primary_Entity = RowOut.Primary_Entity
                        .Primary_WS = -1
                        .Tipe_Entity = 0
                        .Primary = -1
                    End With
                    OutputDGV.Rows.Add(RowOut_Dummy)
                    'Key In DGV Allocate
                    With RowOut_Dummy
                        DGVAllocate.Rows.Add()
                        RowDGV = DGVAllocate.Rows(DGVAllocate.Rows.Count - 1)
                        RowDGV.Cells(i.ColPrimary.Index).Value = .Primary
                        RowDGV.Cells(i.ColEntityTipe.Index).Value = .Tipe_Entity
                        RowDGV.Cells(i.ColPrimaryEntity.Index).Value = .Primary_Entity
                        RowDGV.Cells(i.ColPrimaryWorkShift.Index).Value = .Primary_WS
                        RowDGV.Cells(i.ColRemarks.Index).Value = .Remarks
                        RowDGV.Cells(i.ColWorkShiftName.Index).Value = .Workshift_Name
                        RowDGV.Cells(i.ColWorkshiftCode.Index).Value = .Workshift_Code
                        RowDGV.Cells(i.ColEndDate.Index).Value = Format(.End_Date, "dd MMM yyyy")
                        RowDGV.Cells(i.ColStartDate.Index).Value = Format(.Start_Date, "dd MMM yyyy")
                        RowDGV.Cells(i.ColEmployeeName.Index).Value = .Employee_Name
                        RowDGV.Cells(i.ColEmployeeCode.Index).Value = .Employee_Code
                    End With
                End If
                '----------------------------------------------------------------------------------------------
                '-------Koding Sebelum Optimasi-------
                'For Each RV In OutputDGV.DefaultView
                '    RowOut = RV.Row
                '    DGVAllocate.Rows.Add()
                '    RowDGV = DGVAllocate.Rows(DGVAllocate.Rows.Count - 1)
                '    With RowOut
                '        RowDGV.Cells(i.ColPrimary.Index).Value = .Primary
                '        RowDGV.Cells(i.ColEntityTipe.Index).Value = .Tipe_Entity
                '        RowDGV.Cells(i.ColPrimaryEntity.Index).Value = .Primary_Entity
                '        RowDGV.Cells(i.ColPrimaryWorkShift.Index).Value = .Primary_WS
                '        RowDGV.Cells(i.ColRemarks.Index).Value = .Remarks
                '        RowDGV.Cells(i.ColWorkShiftName.Index).Value = .Workshift_Name
                '        RowDGV.Cells(i.ColWorkshiftCode.Index).Value = .Workshift_Code
                '        RowDGV.Cells(i.ColEndDate.Index).Value = Format(.End_Date, "dd MMM yyyy")
                '        RowDGV.Cells(i.ColStartDate.Index).Value = Format(.Start_Date, "dd MMM yyyy")
                '        RowDGV.Cells(i.ColEmployeeName.Index).Value = .Employee_Name
                '        RowDGV.Cells(i.ColEmployeeCode.Index).Value = .Employee_Code
                '        '------------------------------------------------------------
                '        'add data yg sudah ngurut by employee code dan workshift code
                '        RowOut_Dummy = OutputDGV_Dummy.NewOutput_DGVRow
                '        RowOut_Dummy.Employee_Code = .Employee_Code
                '        RowOut_Dummy.Employee_Name = .Employee_Name
                '        RowOut_Dummy.Start_Date = .Start_Date
                '        RowOut_Dummy.End_Date = .End_Date
                '        RowOut_Dummy.Workshift_Code = .Workshift_Code
                '        RowOut_Dummy.Workshift_Name = .Workshift_Name
                '        RowOut_Dummy.Remarks = .Remarks
                '        RowOut_Dummy.Primary_Entity = .Primary_Entity
                '        RowOut_Dummy.Primary_WS = .Primary_WS
                '        RowOut_Dummy.Tipe_Entity = .Tipe_Entity
                '        RowOut_Dummy.Primary = .Primary
                '        OutputDGV_Dummy.Rows.Add(RowOut_Dummy)
                '    End With
                'Next
                'OutputDGV.DefaultView.Sort = ""
                'OutputDGV.Rows.Clear()
                'For Each RowOut_Dummy In OutputDGV_Dummy.Rows
                '    RowOut = OutputDGV.NewOutput_DGVRow
                '    For index As Integer = 0 To OutputDGV.Columns.Count - 1
                '        RowOut.Item(index) = RowOut_Dummy.Item(index)
                '    Next
                '    OutputDGV.Rows.Add(RowOut)
                'Next
                ''2. Cek range waktu filter dengan DGV ada yang bolong ditengah-tengah ga, kalo ada diinput masuk
                ''a. Cek data yang dengan range waktu bolong
                'For idx As Integer = 0 To OutputDGV.Rows.Count - 1
                '    RowOut = OutputDGV.Rows(idx)
                '    ijk = idx + 1
                '    If ijk <= OutputDGV.Rows.Count - 1 Then
                '        RowOut_Dummy = OutputDGV.Rows(ijk)
                '        If RowOut_Dummy.Primary_Entity = RowOut.Primary_Entity Then
                '            If RowOut_Dummy.Start_Date - TimeSpan.FromDays(1) >= RowOut.End_Date + TimeSpan.FromDays(1) Then
                '                'Input DGV Allocation = nilai default, belum punya workshift
                '                RowOut_Dummy_1 = OutputDGV_Dummy.NewOutput_DGVRow
                '                With RowOut_Dummy_1
                '                    .Employee_Code = RowOut.Employee_Code
                '                    .Employee_Name = RowOut.Employee_Name
                '                    .Workshift_Code = ""
                '                    .Workshift_Name = ""
                '                    .Start_Date = RowOut.End_Date.Date + TimeSpan.FromDays(1)
                '                    .End_Date = RowOut_Dummy.Start_Date.Date - TimeSpan.FromDays(1)
                '                    .Remarks = ""
                '                    .Primary_Entity = RowOut.Primary_Entity
                '                    .Primary_WS = -1
                '                    .Tipe_Entity = 0
                '                    .Primary = -1
                '                End With
                '                OutputDGV_Dummy.Rows.Add(RowOut_Dummy_1)
                '                AdaTambahan_Baris = True
                '            End If
                '        End If
                '    End If
                'Next
                ''b. urutkan data dummy, sesuaikan DGV Allocate
                'If AdaTambahan_Baris = True Then
                '    OutputDGV_Dummy.DefaultView.Sort = "[Employee Code], [Start Date], [End Date], [Workshift Code]"
                '    OutputDGV.Rows.Clear()
                '    DGVAllocate.Rows.Clear()
                '    For Each RV In OutputDGV_Dummy.DefaultView
                '        RowOut_Dummy = RV.Row
                '        DGVAllocate.Rows.Add()
                '        RowDGV = DGVAllocate.Rows(DGVAllocate.Rows.Count - 1)
                '        With RowOut_Dummy
                '            RowDGV.Cells(i.ColPrimary.Index).Value = .Primary
                '            RowDGV.Cells(i.ColEntityTipe.Index).Value = .Tipe_Entity
                '            RowDGV.Cells(i.ColPrimaryEntity.Index).Value = .Primary_Entity
                '            RowDGV.Cells(i.ColPrimaryWorkShift.Index).Value = .Primary_WS
                '            RowDGV.Cells(i.ColRemarks.Index).Value = .Remarks
                '            RowDGV.Cells(i.ColWorkShiftName.Index).Value = .Workshift_Name
                '            RowDGV.Cells(i.ColWorkshiftCode.Index).Value = .Workshift_Code
                '            RowDGV.Cells(i.ColEndDate.Index).Value = Format(.End_Date, "dd MMM yyyy")
                '            RowDGV.Cells(i.ColStartDate.Index).Value = Format(.Start_Date, "dd MMM yyyy")
                '            RowDGV.Cells(i.ColEmployeeName.Index).Value = .Employee_Name
                '            RowDGV.Cells(i.ColEmployeeCode.Index).Value = .Employee_Code
                '            '------------------------------------------------------------
                '            'add data yg sudah ngurut by employee code dan workshift code
                '            RowOut = OutputDGV.NewOutput_DGVRow
                '            RowOut.Employee_Code = .Employee_Code
                '            RowOut.Employee_Name = .Employee_Name
                '            RowOut.Start_Date = .Start_Date
                '            RowOut.End_Date = .End_Date
                '            RowOut.Workshift_Code = .Workshift_Code
                '            RowOut.Workshift_Name = .Workshift_Name
                '            RowOut.Remarks = .Remarks
                '            RowOut.Primary_Entity = .Primary_Entity
                '            RowOut.Primary_WS = .Primary_WS
                '            RowOut.Tipe_Entity = .Tipe_Entity
                '            RowOut.Primary = .Primary
                '            OutputDGV.Rows.Add(RowOut)
                '        End With
                '    Next
                '    '-----------
                'End If
            End If
            DGVAllocate.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGVAllocate.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            OutputDGV_Dummy.Dispose()
        End Try
    End Sub

    'Fungsi Query Data Hapus, dan Simpan
    Public Sub HapusItemAllocate(ByVal vPrimaryAuto As Double)
        Dim DTA As New DseMasterWorkshift_AllocationTableAdapters.tblmaster_workshift_itemallocateTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery_HapusAllocate(vPrimaryAuto)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub SimpanItemAllocate(ByVal vPrimaryEntity As Double, ByVal vTipeEntity As Integer, ByVal vEndDate As Date, ByVal vStartDate As Date, _
    ByVal vPrimaryWorkshift As Double, ByVal vRemarks As String)
        Dim DTA As New DseMasterWorkshift_AllocationTableAdapters.tblmaster_workshift_itemallocateTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.InsertQuery_SimpanAllocate(vPrimaryEntity, vTipeEntity, vStartDate, vEndDate, vPrimaryWorkshift, vRemarks)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    'Fungsi Delete, dan Insert data Allocation
    Public Sub DeleteDataAllocation(ByVal OutputDGV As Core_Module.DseMasterWorkshift_Allocation.Output_DGVDataTable, _
    ByVal DtpEnd As Date, ByVal DtpStart As Date, Optional ByVal MetodeDelete As Integer = 0)
        Dim RowOutput As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        Dim TglDariAllocate As Date = #1/1/1950#
        Dim TglSampaiAllocate As Date = #1/1/1950#
        Dim PrimaryWS As Double = -1
        'Dim RowOutput_Dummy As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        'Dim OutputDGV_Dummy As New Core_Module.DseMasterWorkshift_Allocation.Output_DGVDataTable
        'Dim ij As Integer = 0
        'Dim DV As Data.DataRowView
        Try
            If OutputDGV.Rows.Count > 0 Then
                For Each RowOutput In OutputDGV.Rows
                    PrimaryWS = -1
                    TglDariAllocate = #1/1/1950#
                    TglSampaiAllocate = #1/1/1950#
                    If RowOutput.Primary_WS <> -1 Then
                        If DtpStart <= RowOutput.Start_Date And DtpEnd >= RowOutput.End_Date Then
                            TglDariAllocate = #1/1/1950#
                            TglSampaiAllocate = #1/1/1950#
                            PrimaryWS = -1
                            Me.HapusItemAllocate(RowOutput.Primary)
                        ElseIf DtpStart >= RowOutput.Start_Date And DtpEnd >= RowOutput.End_Date Then
                            TglDariAllocate = RowOutput.Start_Date
                            TglSampaiAllocate = #1/1/1950#
                            PrimaryWS = RowOutput.Primary_WS
                            Me.HapusItemAllocate(RowOutput.Primary)
                            Me.SimpanItemAllocate(RowOutput.Primary_Entity, RowOutput.Tipe_Entity, _
                            DtpStart - TimeSpan.FromDays(1), TglDariAllocate, PrimaryWS, RowOutput.Remarks)
                        ElseIf DtpStart <= RowOutput.Start_Date And DtpEnd <= RowOutput.End_Date Then
                            TglDariAllocate = #1/1/1950#
                            TglSampaiAllocate = RowOutput.End_Date
                            PrimaryWS = RowOutput.Primary_WS
                            Me.HapusItemAllocate(RowOutput.Primary)
                            Me.SimpanItemAllocate(RowOutput.Primary_Entity, RowOutput.Tipe_Entity, _
                             TglSampaiAllocate, DtpEnd + TimeSpan.FromDays(1), PrimaryWS, RowOutput.Remarks)
                        ElseIf DtpStart >= RowOutput.Start_Date And DtpEnd <= RowOutput.End_Date Then
                            TglDariAllocate = RowOutput.Start_Date
                            TglSampaiAllocate = RowOutput.End_Date
                            PrimaryWS = RowOutput.Primary_WS
                            Me.HapusItemAllocate(RowOutput.Primary)
                            Me.SimpanItemAllocate(RowOutput.Primary_Entity, RowOutput.Tipe_Entity, _
                            DtpStart - TimeSpan.FromDays(1), TglDariAllocate, PrimaryWS, RowOutput.Remarks)
                            Me.SimpanItemAllocate(RowOutput.Primary_Entity, RowOutput.Tipe_Entity, _
                            TglSampaiAllocate, DtpEnd + TimeSpan.FromDays(1), PrimaryWS, RowOutput.Remarks)
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'OutputDGV_Dummy.Dispose()
        End Try
    End Sub

    Public Sub InsertDataAllocation(ByVal OutputDGV As Core_Module.DseMasterWorkshift_Allocation.Output_DGVDataTable, _
    ByVal PrimaryWorkshift As Double)
        Dim RowOutput As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        Try
            If OutputDGV.Rows.Count > 0 Then
                For Each RowOutput In OutputDGV.Rows
                    'Me.SimpanItemAllocate(RowOutput.Primary_Entity, RowOutput.Tipe_Entity, dtpend, dtpstart, PrimaryWorkshift, RowOutput.Remarks)
                    Me.SimpanItemAllocate(RowOutput.Primary_Entity, RowOutput.Tipe_Entity, RowOutput.End_Date, RowOutput.Start_Date, PrimaryWorkshift, RowOutput.Remarks)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetPrimary(ByVal vCodeWS As String, ByVal vTipeWS As ClsMasterWorkShift.EnumTipeWorkshift) As Double
        Dim DTA As New Core_Module.DseMasterWorkshift_AllocationTableAdapters.tblmaster_workshiftTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimary(vCodeWS, vTipeWS)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDescription(ByVal vPrimaryWS As Double) As String
        Dim DTA As New Core_Module.DseMasterWorkshift_AllocationTableAdapters.tblmaster_workshiftTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetDescription(vPrimaryWS)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
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

    'Public Function ApakahDataBolehEditAllocate(ByVal i As Core_Module.FrmMasterWorkShift_Allocate, _
    'ByVal TBLSetUlang As Core_Module.DseMasterWorkshift_Allocation.Output_DGVDataTable) As Boolean
    '    Dim DTA1 As New DseMasterWorkshift_AllocationTableAdapters.DataTable1TableAdapter
    '    Dim TBL1 As New DseMasterWorkshift_Allocation.DataTable1DataTable
    '    Dim Row1 As DseMasterWorkshift_Allocation.DataTable1Row
    '    Dim Retval As String = ""
    '    Dim Filter As String = ""
    '    Dim RowDGV As DataGridViewRow
    '    Dim DV As Data.DataRowView
    '    Dim Ket As String
    '    Dim ik As Integer = 0
    '    Dim ListKet As New List(Of String)
    '    Dim ListPrimaryBaris As New List(Of Double)
    '    Dim Hasil As MsgBoxResult
    '    Dim CounterRow As Integer = 0
    '    Dim PeriodFrom As Date = #1/1/1950#
    '    Dim PeriodTo As Date = #1/1/1950#
    '    Dim RowSetUlang As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
    '    Try
    '        'Notes : Batasan boleh allocate dan hapus allocate sama
    '        '1. cek data DGV Allocate tidak boleh kosong 
    '        If i.DgvWorkshift.Rows.Count <= 0 Then
    '            MsgBox("Please filter data allocation.", MsgBoxStyle.Information, "Information")
    '            Return False
    '        End If
    '        '2. Cek Tanggal Dari Sampai
    '        If i.DtpStart.Value.Date > i.DtpEnd.Value.Date Then
    '            MsgBox("Period from should not be greater than period to.", MsgBoxStyle.Information, "Information")
    '            Return False
    '        End If
    '        '3. cek batasan berkaitan dengan attendance
    '        TBLSetUlang.Rows.Clear()
    '        DTA1.Fill(TBL1)
    '        For Each RowDGV In i.DgvWorkshift.Rows
    '            My.Application.ChangeCulture("id-ID")
    '            My.Application.ChangeUICulture("id-ID")
    '            Dim tglstart As Date = CDate(RowDGV.Cells(i.ColStartDate.Index).Value)
    '            Dim tglend As Date = CDate(RowDGV.Cells(i.ColEndDate.Index).Value)
    '            My.Application.ChangeCulture("en-US")
    '            My.Application.ChangeUICulture("en-US")
    '            Filter = "[primary karyawan] = " + RowDGV.Cells(i.ColPrimaryEntity.Index).Value.ToString + _
    '            " AND [date] >= #" + tglstart + "# AND [date] <= #" + _
    '            tglend + _
    '            "# AND [primary workshift] = " + RowDGV.Cells(i.ColPrimaryWorkShift.Index).Value.ToString + ""
    '            TBL1.DefaultView.RowFilter = Filter
    '            TBL1.DefaultView.Sort = "[Date]"
    '            '-------Parameter------
    '            PeriodFrom = #1/1/1950#
    '            PeriodTo = #1/1/1950#
    '            CounterRow = 0
    '            '----------------------
    '            If TBL1.DefaultView.Count > 0 Then
    '                ListKet.Clear()
    '                ListPrimaryBaris.Clear()
    '                If ListPrimaryBaris.Contains(CDbl(RowDGV.Cells(i.ColPrimary.Index).Value)) = False Then
    '                    ListPrimaryBaris.Add(CDbl(RowDGV.Cells(i.ColPrimary.Index).Value))
    '                End If
    '                For Each DV In TBL1.DefaultView
    '                    Row1 = DV.Row
    '                    If ListKet.Contains(Row1.no_transaksi) = False Then
    '                        ListKet.Add(Row1.no_transaksi)
    '                    End If
    '                    '---Period Data ke dalam TBLSetUlang---
    '                    If CounterRow = 0 Then
    '                        If TBL1.DefaultView.Count - 1 = CounterRow Then
    '                            PeriodFrom = Row1._date
    '                            PeriodTo = Row1._date
    '                        Else
    '                            PeriodFrom = Row1._date
    '                        End If
    '                    Else
    '                        If TBL1.DefaultView.Count - 1 = CounterRow Then
    '                            PeriodTo = Row1._date
    '                        End If
    '                    End If
    '                    CounterRow += 1
    '                    '---------------------------------------
    '                Next
    '                '---Masukan Data ke dalam TBLSetUlang---
    '                RowSetUlang = TBLSetUlang.NewOutput_DGVRow
    '                RowSetUlang.Start_Date = PeriodFrom
    '                RowSetUlang.End_Date = PeriodTo
    '                RowSetUlang.Employee_Code = RowDGV.Cells(i.ColEmployeeCode.Index).Value.ToString
    '                RowSetUlang.Employee_Name = RowDGV.Cells(i.ColEmployeeName.Index).Value.ToString
    '                RowSetUlang.Workshift_Code = RowDGV.Cells(i.ColWorkshiftCode.Index).Value.ToString
    '                RowSetUlang.Workshift_Name = RowDGV.Cells(i.ColWorkShiftName.Index).Value.ToString
    '                RowSetUlang.Remarks = RowDGV.Cells(i.ColRemarks.Index).Value.ToString
    '                RowSetUlang.Primary_Entity = CDbl(RowDGV.Cells(i.ColPrimaryEntity.Index).Value)
    '                RowSetUlang.Primary_WS = CDbl(RowDGV.Cells(i.ColPrimaryWorkShift.Index).Value)
    '                RowSetUlang.Tipe_Entity = 0
    '                RowSetUlang.Primary = -1
    '                TBLSetUlang.Rows.Add(RowSetUlang)
    '                '---------------------------------------
    '                ik += 1
    '                Retval = "Row " + ik.ToString + ". Attendance no :"
    '                For Each Ket In ListKet
    '                    If Retval.EndsWith(":") = True Then
    '                        Retval += " " + Ket.ToString
    '                    Else
    '                        Retval += ", " + Ket.ToString
    '                    End If
    '                Next
    '                Retval += vbCrLf
    '            End If
    '        Next
    '        If Retval <> "" Then
    '            Hasil = MsgBox("This Allocate item cannot update at :" + _
    '            Retval.ToString + "because this data refer to transaction allocate" + vbCrLf + "do you want to continue ?", MsgBoxStyle.YesNo, "Confirm")
    '            If Hasil = MsgBoxResult.No Then
    '                Return False
    '            End If
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA1.Dispose()
    '        TBL1.Dispose()
    '        ListKet = Nothing
    '    End Try
    'End Function

    Public Function ApakahDataBolehHapusAllocate(ByVal OutputDGV As Core_Module.DseMasterWorkshift_Allocation.Output_DGVDataTable, _
    ByVal i As Core_Module.FrmMasterWorkShift_Allocate, Optional ByVal UntukKebutuhanFilterPadaBtnAllocate As Boolean = False) As Boolean
        Dim DTA1 As New DseMasterWorkshift_AllocationTableAdapters.DataTable1TableAdapter
        Dim TBL1 As New DseMasterWorkshift_Allocation.DataTable1DataTable
        Dim Row1 As DseMasterWorkshift_Allocation.DataTable1Row
        Dim DTA2 As New DseMasterWorkshift_AllocationTableAdapters.tbl_overtime_orderTableAdapter
        Dim TBL2 As New DseMasterWorkshift_Allocation.tbl_overtime_orderDataTable
        Dim Row2 As DseMasterWorkshift_Allocation.tbl_overtime_orderRow
        Dim RowOutput As Core_Module.DseMasterWorkshift_Allocation.Output_DGVRow
        Dim Retval As String = ""
        Dim Filter As String = ""
        Dim RowDGV As DataGridViewRow
        Dim DV As Data.DataRowView
        Dim Ket As String
        Dim PrimaryBaris As Double
        Dim ik As Integer = 0
        Dim ListKet As New List(Of String)
        Dim ListPrimaryBaris As New List(Of Double)
        Dim Hasil As MsgBoxResult
        Dim count As Integer = 0
        Dim FilterDateFrom As Date = CDate("1/1/2200")
        Dim FilterDateTo As Date = CDate("1/1/1950")

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            'Notes : Batasan boleh allocate dan hapus allocate sama
            '1. cek data DGV Allocate tidak boleh kosong 
            If i.DgvWorkshift.Rows.Count <= 0 Then
                MsgBox("Please filter data allocation.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            '2. Cek Tanggal Dari Sampai
            If i.DtpStart.Value.Date > i.DtpEnd.Value.Date Then
                MsgBox("Period from should not be greater than period to.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            '3. cek batasan berkaitan dengan attendance
            For Each RowDGV In i.DgvWorkshift.Rows
                If CDate(RowDGV.Cells(i.ColStartDate.Index).Value) < FilterDateFrom Then
                    FilterDateFrom = CDate(RowDGV.Cells(i.ColStartDate.Index).Value)
                End If
                If CDate(RowDGV.Cells(i.ColEndDate.Index).Value) > FilterDateTo Then
                    FilterDateTo = CDate(RowDGV.Cells(i.ColEndDate.Index).Value)
                End If
            Next

            DTA1.FillBy_DateFromTo(TBL1, FilterDateFrom.Date, FilterDateTo.Date)
            DTA2.FillBy_DateFromTo(TBL2, FilterDateFrom.Date, FilterDateTo.Date)
            'DTA1.FillBy_DateFromTo(TBL1, i.DtpStart.Value.Date, i.DtpEnd.Value.Date)
            'DTA1.Fill(TBL1)
            For Each RowDGV In i.DgvWorkshift.Rows
                My.Application.ChangeCulture("id-ID")
                My.Application.ChangeUICulture("id-ID")
                Dim tglstart As Date = CDate(RowDGV.Cells(i.ColStartDate.Index).Value)
                Dim tglend As Date = CDate(RowDGV.Cells(i.ColEndDate.Index).Value)
                My.Application.ChangeCulture("en-US")
                My.Application.ChangeUICulture("en-US")
                Filter = "[primary karyawan] = " + RowDGV.Cells(i.ColPrimaryEntity.Index).Value.ToString + _
                " AND [date] >= #" + tglstart + "# AND [date] <= #" + _
                tglend + _
                "# AND [primary workshift] = " + RowDGV.Cells(i.ColPrimaryWorkShift.Index).Value.ToString + ""
                TBL1.DefaultView.RowFilter = Filter
                If TBL1.DefaultView.Count > 0 Then
                    count += 1
                    ListKet.Clear()
                    'ListPrimaryBaris.Clear()
                    If ListPrimaryBaris.Contains(CDbl(RowDGV.Cells(i.ColPrimary.Index).Value)) = False Then
                        ListPrimaryBaris.Add(CDbl(RowDGV.Cells(i.ColPrimary.Index).Value))
                    End If
                    For Each DV In TBL1.DefaultView
                        Row1 = DV.Row
                        If ListKet.Contains(Row1.no_transaksi) = False Then
                            ListKet.Add(Row1.no_transaksi)
                        End If
                    Next
                    ik += 1
                    If String.IsNullOrEmpty(Retval) = False Then
                        Retval += "Row " + CStr(RowDGV.Index + 1) + ". because refering to attendance no :"
                    Else
                        Retval = "Row " + CStr(RowDGV.Index + 1) + ". because refering to attendance no :"
                    End If
                    For Each Ket In ListKet
                        If Retval.EndsWith(":") = True Then
                            Retval += " " + Ket.ToString
                        Else
                            Retval += ", " + Ket.ToString
                        End If
                    Next
                    Retval += vbCrLf
                End If

                If count > 10 Then
                    Retval += "................" + vbNewLine
                    Exit For
                End If
            Next

            If count < 11 Then
                'cek juga untuk overtime ordernya
                For Each RowDGV In i.DgvWorkshift.Rows
                    My.Application.ChangeCulture("id-ID")
                    My.Application.ChangeUICulture("id-ID")
                    Dim tglstart As Date = CDate(RowDGV.Cells(i.ColStartDate.Index).Value)
                    Dim tglend As Date = CDate(RowDGV.Cells(i.ColEndDate.Index).Value)
                    My.Application.ChangeCulture("en-US")
                    My.Application.ChangeUICulture("en-US")
                    Filter = "[Primary Employee] = " + RowDGV.Cells(i.ColPrimaryEntity.Index).Value.ToString + _
                    " AND [To] >= #" + tglstart + "# AND [From] <= #" + tglend + _
                    "# AND [Primary Shift] = " + RowDGV.Cells(i.ColPrimaryWorkShift.Index).Value.ToString + ""
                    TBL2.DefaultView.RowFilter = Filter
                    If TBL2.DefaultView.Count > 0 Then
                        count += 1
                        ListKet.Clear()
                        'ListPrimaryBaris.Clear()
                        If ListPrimaryBaris.Contains(CDbl(RowDGV.Cells(i.ColPrimary.Index).Value)) = False Then
                            ListPrimaryBaris.Add(CDbl(RowDGV.Cells(i.ColPrimary.Index).Value))
                        End If
                        For Each DV In TBL2.DefaultView
                            Row2 = DV.Row
                            If ListKet.Contains(Row2.No_Transaksi) = False Then
                                ListKet.Add(Row2.No_Transaksi)
                            End If
                        Next
                        ik += 1
                        If String.IsNullOrEmpty(Retval) = False Then
                            Retval += "Row " + CStr(RowDGV.Index + 1) + ". because refering to overtime order no :"
                        Else
                            Retval = "Row " + CStr(RowDGV.Index + 1) + ". because refering to overtime order no :"
                        End If
                        For Each Ket In ListKet
                            If Retval.EndsWith(":") = True Then
                                Retval += " " + Ket.ToString
                            Else
                                Retval += ", " + Ket.ToString
                            End If
                        Next
                        Retval += vbCrLf
                    End If

                    If count > 10 Then
                        Retval += "................" + vbNewLine
                        Exit For
                    End If
                Next
            End If

            If UntukKebutuhanFilterPadaBtnAllocate = False Then
                If Retval <> "" Then
                    'Hasil = MsgBox("This Allocate item cannot update at :" + _
                    'Retval.ToString + "because this data refer to transaction allocate" + vbCrLf + "do you want to continue ?", MsgBoxStyle.YesNo, "Confirm")
                    Hasil = MsgBox("This allocate row item cannot delete at :" + vbCrLf + _
                    Retval.ToString + vbCrLf + "do you want to continue ?", MsgBoxStyle.YesNo, "Confirm")
                    If Hasil = MsgBoxResult.No Then
                        Return False
                    Else
                        If ListPrimaryBaris.Count > 0 Then
                            For Each PrimaryBaris In ListPrimaryBaris
                                Filter = "[Primary] = " + PrimaryBaris.ToString + ""
                                OutputDGV.DefaultView.RowFilter = Filter
                                If OutputDGV.DefaultView.Count > 0 Then
                                    RowOutput = OutputDGV.DefaultView(0).Row
                                    OutputDGV.Rows.Remove(RowOutput)
                                End If
                            Next
                            OutputDGV.DefaultView.RowFilter = ""
                        End If
                    End If
                End If
            Else
                If ListPrimaryBaris.Count > 0 Then
                    For Each PrimaryBaris In ListPrimaryBaris
                        Filter = "[Primary] = " + PrimaryBaris.ToString + ""
                        OutputDGV.DefaultView.RowFilter = Filter
                        If OutputDGV.DefaultView.Count > 0 Then
                            RowOutput = OutputDGV.DefaultView(0).Row
                            OutputDGV.Rows.Remove(RowOutput)
                        End If
                    Next
                    OutputDGV.DefaultView.RowFilter = ""
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
            ListKet = Nothing
            ListPrimaryBaris = Nothing
        End Try
    End Function
    'Fungsi simpan data dgv allocate
    Public Function GetPrimaryAutoAllocate(ByVal vPrimaryEntity As Double, ByVal vTipeEntity As Integer, ByVal vStartDate As Date, ByVal vEndDate As Date, _
    ByVal vPrimaryWorkshift As Double) As Double
        Dim DTA As New Core_Module.DseMasterWorkshift_AllocationTableAdapters.tblmaster_workshift_itemallocate1TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryAutoAllocate(vPrimaryEntity, vTipeEntity, vStartDate, vEndDate, vPrimaryWorkshift)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahDataBenar_ItemAllocate_Save(ByVal RowDGV As DataGridViewRow, _
    ByVal i As Core_Module.FrmMasterWorkShift_Allocate) As Boolean
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            If CDate(RowDGV.Cells(i.ColStartDate.Index).Value) > _
            CDate(RowDGV.Cells(i.ColEndDate.Index).Value) Then
                Return False
            End If
            If IsDate(RowDGV.Cells(i.ColStartDate.Index).Value) = False Then
                Return False
            End If
            If IsDate(RowDGV.Cells(i.ColEndDate.Index).Value) = False Then
                Return False
            End If
            If RowDGV.Cells(i.ColPrimaryWorkShift.Index).Value = "-1" Then
                Return False
            End If
            If ObjCD.CekDataKosong(RowDGV.Cells(i.ColRemarks.Index).Value, False) = False Then
                'MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'sampe sini
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Public Sub UpdateDataPerItemAllocate(ByVal RowDGV As DataGridViewRow, ByVal vPrimaryAuto As Double, _
    ByVal i As Core_Module.FrmMasterWorkShift_Allocate)
        Dim DTA As New DseMasterWorkshift_AllocationTableAdapters.tblmaster_workshift_itemallocate1TableAdapter
        Dim Remarks As String = ""
        Dim TglStart As Date = #1/1/1950#
        Dim TglEnd As Date = #1/1/1950#
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            My.Application.ChangeCulture("id-ID")
            My.Application.ChangeUICulture("id-ID")
            TglStart = RowDGV.Cells(i.ColStartDate.Index).Value
            TglEnd = RowDGV.Cells(i.ColEndDate.Index).Value
            My.Application.ChangeCulture("en-US")
            My.Application.ChangeUICulture("en-US")
            DTA.UpdateQuery_StartDate(TglStart, vPrimaryAuto)
            DTA.UpdateQuery_EndDate(TglEnd, vPrimaryAuto)
            DTA.UpdateQuery_PrimaryWorkshift(RowDGV.Cells(i.ColPrimaryWorkShift.Index).Value, vPrimaryAuto)
            If String.IsNullOrEmpty(RowDGV.Cells(i.ColRemarks.Index).Value) = False Then
                DTA.UpdateQuery_Remarks(RowDGV.Cells(i.ColRemarks.Index).Value, vPrimaryAuto)
            Else
                DTA.UpdateQuery_Remarks("", vPrimaryAuto)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    'Batasan DGV Allocate
    Public Function GetJmlAttendance(ByVal vPrimaryWS As Double, ByVal vDate As Date, ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New Core_Module.DseMasterWorkshift_AllocationTableAdapters.tbltrans_attendance_itemTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Dim Retval As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlBy_WS_Date_Kry(vPrimaryWS, vDate, vPrimaryEmployee)
            If Dummy.HasValue Then
                Retval = Dummy.Value
            Else
                Retval = 0
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNomerAttendance(ByVal vPrimaryWS As Double, ByVal vDate As Date, ByVal vPrimaryEmployee As Double) As String
        Dim DTA As New Core_Module.DseMasterWorkshift_AllocationTableAdapters.tbltrans_attendance_itemTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNomerAttendance(vPrimaryWS, vDate, vPrimaryEmployee)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
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

    Public Function GetJmlBatasanGeserLuarStartDate(ByVal vStartBaru As Date, ByVal vStartLama As Date, ByVal vPrimaryAuto As Double, _
    ByVal vPrimaryEntity As Double) As Double
        Dim DTA As New Core_Module.DseMasterWorkshift_AllocationTableAdapters.tblmaster_workshift_itemallocate1TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlAllocateByGeser_Luar_Start_Date(vStartBaru, vStartLama, vPrimaryAuto, vPrimaryEntity)
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

    Public Function GetJmlBatasanGeserLuarEndDate(ByVal vEndBaru As Date, ByVal vEndLama As Date, ByVal vPrimaryAuto As Double, _
    ByVal vPrimaryEntity As Double) As Double
        Dim DTA As New Core_Module.DseMasterWorkshift_AllocationTableAdapters.tblmaster_workshift_itemallocate1TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlAllocateByGeser_Luar_End_Date(vEndBaru, vEndLama, vPrimaryAuto, vPrimaryEntity)
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

    Public Sub FungsiFillDataAttendanceItem(ByVal TBLItem As DseMasterWorkshift_Allocation.tbltrans_attendance_headerDataTable, _
    ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vPrimaryEmployee As Double)
        Dim DTA As New DseMasterWorkshift_AllocationTableAdapters.tbltrans_attendance_headerTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBLItem, vStartDate, vEndDate, vPrimaryEmployee)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub FungsiFillDataOvertimeOrderItem(ByVal TBLItem As DseMasterWorkshift_Allocation.tbl_overtime_order_headerDataTable, _
    ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vPrimaryEmployee As Double)
        Dim DTA As New DseMasterWorkshift_AllocationTableAdapters.tbl_overtime_order_headerTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill_2(TBLItem, vPrimaryEmployee, vStartDate, vEndDate)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    'Batasan DGV Allocate
    Public Function GetJmlOvertimeOrder(ByVal vPrimaryWS As Double, ByVal vDate As Date, ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New Core_Module.DseMasterWorkshift_AllocationTableAdapters.tbl_overtime_orderTableAdapter
        Dim Dummy As System.Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlBy_WS_Date_Kry_2(vPrimaryWS, vPrimaryEmployee, vDate)
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

    Public Function GetNomerOvertimeOrder(ByVal vPrimaryWS As Double, ByVal vDate As Date, ByVal vPrimaryEmployee As Double) As String
        Dim Dta As New DseMasterWorkshift_AllocationTableAdapters.tbl_overtime_order_headerTableAdapter
        Dim Tbl As New DseMasterWorkshift_Allocation.tbl_overtime_order_headerDataTable
        Dim row As DseMasterWorkshift_Allocation.tbl_overtime_order_headerRow
        'Dim DTA As New Core_Module.DseMasterWorkshift_AllocationTableAdapters.tbl_overtime_orderTableAdapter
        Dim Dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill_3(Tbl, vPrimaryEmployee, vDate, vPrimaryWS)
            For Each row In Tbl
                If Dummy = "" Then
                    Dummy = row.No_Transaksi
                Else
                    Dummy &= ", " & row.No_Transaksi
                End If
            Next
            'Dummy = Dta.ScalarQuery_GetNomorOvertimeOrder(vPrimaryWS, vPrimaryEmployee, vDate)
            'If String.IsNullOrEmpty(Dummy) = False Then
            '    Return Dummy
            'Else
            '    Return ""
            'End If
            Return Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
