Option Explicit On
Option Strict Off

Public Class ClsMasterWorkUnitItem
    Public Function IsiDgvEmployee(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
          ByVal i As FrmMasterWorkUnit, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseMasterWorkUnitTableAdapters.DataTableEmployeeTableAdapter
        Dim TBLItem As New DseMasterWorkUnit.DataTableEmployeeDataTable
        Dim rowItem As DseMasterWorkUnit.DataTableEmployeeRow
        Dim DefaultPrecision As Integer = 0
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItem.FillBy_PrimaryWorkUnit(TBLItem, vPrimary)
            vDGVItem.Rows.Clear()

            vDGVItem.AllowUserToAddRows = False
            For Each rowItem In TBLItem.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColEmployee_No.Index).Value = ""
                    .Cells(i.ColEmployee_ShiftCode.Index).Value = rowItem.Code_Shift
                    .Cells(i.ColEmployee_EmployeeCode.Index).Value = rowItem.Code_Employee
                    .Cells(i.ColEmployee_Name.Index).Value = rowItem.Full_Name
                    .Cells(i.ColEmployee_Remarks.Index).Value = rowItem.remarks
                    .Cells(i.ColEmployee_Primary.Index).Value = rowItem.primary_employee
                    .Cells(i.ColEmployee_PrimaryShift.Index).Value = rowItem.primary_work_shift
                End With
            Next
            vDGVItem.AllowUserToAddRows = False

            'set number
            vObjCD.SetNumber(vDGVItem)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Public Function IsiDgvMachine(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
          ByVal i As FrmMasterWorkUnit, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseMasterWorkUnitTableAdapters.DataTableMachineTableAdapter
        Dim TBLItem As New DseMasterWorkUnit.DataTableMachineDataTable
        Dim rowItem As DseMasterWorkUnit.DataTableMachineRow
        Dim DefaultPrecision As Integer = 0
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItem.FillBy_PrimaryWorkUnit(TBLItem, vPrimary)
            vDGVItem.Rows.Clear()

            vDGVItem.AllowUserToAddRows = False
            For Each rowItem In TBLItem.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColMachine_No.Index).Value = ""
                    .Cells(i.ColMachine_MachineCode.Index).Value = rowItem.Machine_Data_Code
                    .Cells(i.ColMachine_Desc.Index).Value = rowItem.Description
                    .Cells(i.ColMachine_Remarks.Index).Value = rowItem.Remarks
                    .Cells(i.ColMachine_PrimaryMachine.Index).Value = rowItem.primary_machine
                End With
            Next
            vDGVItem.AllowUserToAddRows = False

            'set number
            vObjCD.SetNumber(vDGVItem)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Function


    Public Function IsiDgvSetup(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
          ByVal i As FrmMasterWorkUnit, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseMasterWorkUnitTableAdapters.DataTableSetuptimeTableAdapter
        Dim TBLItem As New DseMasterWorkUnit.DataTableSetuptimeDataTable
        Dim rowItem As DseMasterWorkUnit.DataTableSetuptimeRow
        Dim DefaultPrecision As Integer = 0
        Dim objCD As New Core_Module.ClsCekData
        Dim objP As New Core_Module.ClsMasterProses
        Dim objProd As New Core_Module.clsMasterProduct
        Dim objSetup As New Core_Module.ClsMasterSetupTime
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItem.Fill(TBLItem, vPrimary)
            vDGVItem.Rows.Clear()

            vDGVItem.AllowUserToAddRows = False
            For Each rowItem In TBLItem.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColSetupTime_No.Index).Value = ""
                    .Cells(i.ColSetupTime_SetupTimeDuration.Index).Value = objCD.Convert_ToString(rowItem.Duration)
                    .Cells(i.ColSetupTime_PrimaryCodeNo1.Index).Value = rowItem.Primary_Product_1
                    .Cells(i.ColSetupTime_CodeNo1.Index).Value = objProd.GetProductCode(rowItem.Primary_Product_1)
                    .Cells(i.ColSetupTime_Desc1.Index).Value = objProd.GetDescription(rowItem.Primary_Product_1)
                    .Cells(i.ColSetupTime_PrimaryCodeNo2.Index).Value = rowItem.Primary_Product_2
                    .Cells(i.ColSetupTime_CodeNo2.Index).Value = objProd.GetProductCode(rowItem.Primary_Product_2)
                    .Cells(i.ColSetupTime_Desc2.Index).Value = objProd.GetDescription(rowItem.Primary_Product_2)
                    .Cells(i.ColSetupTime_PrimaryProcess1.Index).Value = rowItem.Primary_Process_1
                    .Cells(i.ColSetupTime_PrimaryProses2.Index).Value = rowItem.Primary_Process_2
                    .Cells(i.ColSetupTime_ProcessCode1.Index).Value = objP.GetCode(rowItem.Primary_Process_1)
                    .Cells(i.ColSetupTime_ProcessCode2.Index).Value = objP.GetCode(rowItem.Primary_Process_2)
                    .Cells(i.ColSetupTime_Remarks.Index).Value = rowItem.Remarks
                    .Cells(i.ColSetupTime_PrimarySetupTime.Index).Value = rowItem.Primary_Setup_Time
                    .Cells(i.ColSetupTime_Code.Index).Value = objSetup.GetCode(rowItem.Primary_Setup_Time)
                End With
            Next
            vDGVItem.AllowUserToAddRows = False

            'set number
            vObjCD.SetNumber(vDGVItem)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Public Function HapusItemEmployee(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunit_item_employeeTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function HapusItemMachine(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunit_item_machineTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function



    Public Function HapusItemSetup(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunit_item_setuptimeTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanItemEmployee(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmMasterWorkUnit) As Boolean
        Dim DTAItem As New DseMasterWorkUnitTableAdapters.tblmaster_workunit_item_employeeTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItem.AllowUserToAddRows = False
            For Each rowDGV In vDGVItem.Rows
                With rowDGV
                    DTAItem.Insert(vPrimary, .Cells(i.ColEmployee_PrimaryShift.Index).Value, .Cells(i.ColEmployee_Primary.Index).Value, .Cells(i.ColEmployee_Remarks.Index).Value)
                End With
            Next
            vDGVItem.AllowUserToAddRows = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function SimpanItemMachine(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmMasterWorkUnit) As Boolean
        Dim DTAItem As New DseMasterWorkUnitTableAdapters.tblmaster_workunit_item_machineTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItem.AllowUserToAddRows = False
            For Each rowDGV In vDGVItem.Rows
                With rowDGV
                    DTAItem.Insert(vPrimary, .Cells(i.ColMachine_PrimaryMachine.Index).Value, .Cells(i.ColMachine_Remarks.Index).Value)
                End With
            Next
            vDGVItem.AllowUserToAddRows = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function


    Public Function SimpanItemSetup(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal i As FrmMasterWorkUnit) As Boolean
        Dim DTAItem As New DseMasterWorkUnitTableAdapters.tblmaster_workunit_item_setuptimeTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItem.AllowUserToAddRows = False
            For Each rowDGV In vDGVItem.Rows
                With rowDGV
                    Dim Duration As Integer = 0
                    Duration = ObjCD.Convert_ToMinutes(.Cells(i.ColSetupTime_SetupTimeDuration.Index).Value)
                    DTAItem.Insert(vPrimary, .Cells(i.ColSetupTime_PrimaryProcess1.Index).Value, _
                    .Cells(i.ColSetupTime_PrimaryProses2.Index).Value, _
                    .Cells(i.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                    .Cells(i.ColSetupTime_PrimaryCodeNo2.Index).Value, _
                    .Cells(i.ColSetupTime_PrimarySetupTime.Index).Value, _
                    Duration, _
                    .Cells(i.ColSetupTime_Remarks.Index).Value)
                End With
            Next
            vDGVItem.AllowUserToAddRows = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function


    Public Function SimpanItemSetup2(ByVal vDGVItem As DataGridView, _
    ByVal i As FrmMasterWorkUnit_AddSetupTime) As Boolean
        Dim DTAItem As New DseMasterWorkUnitTableAdapters.tblmaster_workunit_item_setuptimeTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItem.AllowUserToAddRows = False
            For Each rowDGV In vDGVItem.Rows
                With rowDGV
                    Dim Duration As Integer = 0
                    Duration = ObjCD.Convert_ToMinutes(.Cells(i.ColSetupTime_SetupTimeDuration.Index).Value)
                    DTAItem.Insert(.Cells(i.ColSetupTime_PrimaryWorkUnit.Index).Value, .Cells(i.ColSetupTime_PrimaryProcess1.Index).Value, _
                    .Cells(i.ColSetupTime_PrimaryProses2.Index).Value, _
                    .Cells(i.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                    .Cells(i.ColSetupTime_PrimaryCodeNo2.Index).Value, _
                    .Cells(i.ColSetupTime_PrimarySetupTime.Index).Value, _
                    Duration, _
                    .Cells(i.ColSetupTime_Remarks.Index).Value)
                End With
            Next
            vDGVItem.AllowUserToAddRows = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function
End Class
