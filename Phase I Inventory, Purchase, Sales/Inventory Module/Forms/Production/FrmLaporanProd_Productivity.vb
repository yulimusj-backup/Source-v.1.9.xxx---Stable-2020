Option Explicit On
Option Strict Off

Public Class FrmLaporanProd_Productivity
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCD As New Core_Module.ClsCekData
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim objMO As New ClsTrans_ManufacturingOrder
    Dim MO_No As String = ""
    Dim Buffer As String = ""
    Dim index As Integer = -1
    Dim ObjLicense As New Core_Module.ClsLicense

    Enum EnumFormat As Integer
        Productivity_by_JO
        Productivity_Recap
        Work_Unit_Recap
        Work_Unit_Detail
        Process_Recap
        Process_Detail
    End Enum

    Enum EnumStatus As Integer
        All
        Finished_Only
        On_Progress
    End Enum

    Enum EnumQtyProduce As Integer
        LebihKecil
        SamaDengan
        LebihBesar
    End Enum

    Enum EnumQtyResult As Integer
        LebihKecil
        SamaDengan
        LebihBesar
    End Enum

    Enum EnumGP As Integer
        LebihKecil
        SamaDengan
        LebihBesar
    End Enum

    Enum EnumProdRate As Integer
        LebihKecil
        SamaDengan
        LebihBesar
    End Enum

    Private Sub FrmLaporanProd_Productivity_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.dgvProduct.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvMO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.CboFormat.SelectedIndex = EnumFormat.Productivity_by_JO
            Me.CboStatus.SelectedIndex = EnumStatus.All
            Me.CboQtyProduce.SelectedIndex = EnumQtyProduce.LebihKecil
            Me.CboQtyResult.SelectedIndex = EnumQtyResult.LebihKecil
            Me.CboGP.SelectedIndex = EnumGP.LebihKecil
            Me.CboProdRate.SelectedIndex = EnumProdRate.LebihKecil
            Me.CboQtyProduce.Enabled = False
            Me.CboQtyResult.Enabled = False
            Me.CboGP.Enabled = False
            Me.CboProdRate.Enabled = False
            Me.TxtToleranceQtyProduce.Enabled = False
            Me.TxtToleranceQtyResult.Enabled = False
            Me.TxtToleranceGP.Enabled = False
            Me.TxtToleranceProdRate.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShow.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.dtpFrom.Value.Date <= Me.dtpTo.Value.Date Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.Cetak()
            Else
                MsgBox("Starting period must be smaller than ending period", MsgBoxStyle.Information, "Information")
                Me.dtpFrom.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterByProduct(ByRef vTbl As Data.DataTable)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
            For Each dr As DataGridViewRow In Me.dgvProduct.Rows
                With dr
                    If dr.Cells(Me.colProduct_Primary.Index).Value <> -1 Then
                        If dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Group" Then
                            If filter = "" Then
                                filter = "[Primary Group] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Group] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Brand" Then
                            If filter = "" Then
                                filter = "[Primary Brand] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Brand] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Sub Brand" Then
                            If filter = "" Then
                                filter = "[Primary Sub Brand] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Sub Brand] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Category" Then
                            If filter = "" Then
                                filter = "[Primary Category] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Category] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                        If dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Product" Then
                            If filter = "" Then
                                filter = "[Primary Product WO] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            Else
                                filter += " OR [Primary Product WO] = " & .Cells(Me.colProduct_Primary.Index).Value.ToString
                            End If
                        End If
                    End If
                End With
            Next
            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                rowDummy = tblDummy.DefaultView(i).Row
                rowTbl = vTbl.NewRow
                For ii As Integer = 0 To tblDummy.Columns.Count - 1
                    rowTbl.Item(ii) = rowDummy.Item(ii)
                Next
                vTbl.Rows.Add(rowTbl)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblDummy.Dispose()
        End Try
    End Sub

    Private Sub FilterByMO(ByRef vTbl As Data.DataTable)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
            MO_No = ""
            For Each dr As DataGridViewRow In Me.DgvMO.Rows
                With dr
                    If dr.Cells(Me.colMO_Primary.Index).Value <> -1 Then
                        If filter = "" Then
                            filter = "[Primary] = " & .Cells(Me.colMO_Primary.Index).Value.ToString
                            MO_No = .Cells(Me.colMO_DocNo.Index).Value.ToString
                        Else
                            filter += " OR [Primary] = " & .Cells(Me.colMO_Primary.Index).Value.ToString
                            MO_No += ", " & .Cells(Me.colMO_DocNo.Index).Value.ToString
                        End If
                    End If
                End With
            Next

            If filter = "" Then
                MO_No = "All"
            End If

            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            If tblDummy.DefaultView.Count > 0 Then
                For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                    rowDummy = tblDummy.DefaultView(i).Row
                    rowTbl = vTbl.NewRow
                    For ii As Integer = 0 To tblDummy.Columns.Count - 1
                        rowTbl.Item(ii) = rowDummy.Item(ii)
                    Next
                    vTbl.Rows.Add(rowTbl)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblDummy.Dispose()
        End Try
    End Sub

    Private Sub FilterByTolerance(ByRef vTbl As Data.DataTable)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""

        Try
            tblDummy = vTbl.Copy
            If Me.ChkToleranceQtyProduce.Checked = True Then
                If Me.CboQtyProduce.SelectedIndex = EnumQtyProduce.LebihKecil Then
                    filter = "[Tolerance Qty Produce] < -"
                ElseIf Me.CboProdRate.SelectedIndex = EnumQtyProduce.SamaDengan Then
                    filter = "[Tolerance Qty Produce] = "
                Else
                    filter = "[Tolerance Qty Produce] > -"
                End If
                filter &= CDec(Me.TxtToleranceQtyProduce.Text)
            End If
            If Me.ChkToleranceQtyResult.Checked = True Then
                If Me.CboQtyResult.SelectedIndex = EnumQtyResult.LebihKecil Then
                    filter = "[Tolerance Qty Result] < -"
                ElseIf Me.CboQtyResult.SelectedIndex = EnumQtyResult.SamaDengan Then
                    filter = "[Tolerance Qty Result] = "
                Else
                    filter = "[Tolerance Qty Result] > "
                End If
                filter &= CDec(Me.TxtToleranceQtyResult.Text)
            End If
            If Me.ChkToleranceGP.Checked = True Then
                If filter <> "" Then
                    filter &= " AND "
                End If
                If Me.CboGP.SelectedIndex = EnumGP.LebihKecil Then
                    filter = "[Tolerance GP] < -"
                ElseIf Me.CboGP.SelectedIndex = EnumGP.SamaDengan Then
                    filter = "[Tolerance GP] = "
                Else
                    filter = "[Tolerance GP] > "
                End If
                filter &= CDec(Me.TxtToleranceGP.Text)
            End If
            If Me.ChkToleranceProdRate.Checked = True Then
                If filter <> "" Then
                    filter &= " AND "
                End If
                If Me.CboProdRate.SelectedIndex = EnumProdRate.LebihKecil Then
                    filter = "[Tolerance ProdRate] < -"
                ElseIf Me.CboProdRate.SelectedIndex = EnumProdRate.SamaDengan Then
                    filter = "[Tolerance ProdRate] = "
                Else
                    filter = "[Tolerance ProdRate] > "
                End If
                filter &= CDec(Me.TxtToleranceProdRate.Text)
            End If

            tblDummy.DefaultView.RowFilter = filter
            vTbl.Rows.Clear()
            If tblDummy.DefaultView.Count > 0 Then
                For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                    rowDummy = tblDummy.DefaultView(i).Row
                    rowTbl = vTbl.NewRow
                    For ii As Integer = 0 To tblDummy.Columns.Count - 1
                        rowTbl.Item(ii) = rowDummy.Item(ii)
                    Next
                    vTbl.Rows.Add(rowTbl)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblDummy.Dispose()
        End Try
    End Sub

    'Private Sub Cetak2()
    '    Try
    '        If Me.CboFormat.SelectedIndex = EnumFormat.Productivity_by_JO Then
    '            'Productivity by JO
    '        ElseIf Me.CboFormat.SelectedIndex = EnumFormat.Productivity_Recap Then
    '            'Productivity(Recap)
    '        ElseIf Me.CboFormat.SelectedIndex = EnumFormat.Work_Unit_Recap Then
    '            'Work Unit Recap
    '        ElseIf Me.CboFormat.SelectedIndex = EnumFormat.Work_Unit_Detail Then
    '            'Work Unit Detail
    '        ElseIf Me.CboFormat.SelectedIndex = EnumFormat.Process_Recap Then
    '            'Process(Recap)
    '        ElseIf Me.CboFormat.SelectedIndex = EnumFormat.Process_Detail Then
    '            'Process(Detail)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub Cetak()
        Dim Ds As New DseLaporanProductivity
        Dim Dta As New DseLaporanProductivityTableAdapters.DataTableTableAdapter
        Dim Tbl, TblDummy As New DseLaporanProductivity.DataTableDataTable
        Dim TblC As DseLaporanProductivity.Tbl_CetakDataTable = Ds.Tbl_Cetak
        Dim Row, RowDummy As DseLaporanProductivity.DataTableRow
        Dim RowC As DseLaporanProductivity.Tbl_CetakRow
        Dim StatusWO_1, StatusWO_2, Level_WO As Integer
        Dim TypeFilter As String = ""

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Me.CboStatus.SelectedIndex = EnumStatus.All Then
                StatusWO_1 = 3 'on progress
                StatusWO_2 = 4 'finish
            ElseIf Me.CboStatus.SelectedIndex = EnumStatus.On_Progress Then
                StatusWO_1 = 3 'on progress
                StatusWO_2 = 3 'on progress
            ElseIf Me.CboStatus.SelectedIndex = EnumStatus.Finished_Only Then
                StatusWO_1 = 4 'finish
                StatusWO_2 = 4 'finish
            End If

            If Me.CboFormat.SelectedIndex = EnumFormat.Productivity_Recap Then
                'hanya tampilkan level wo yang paling terakhir
                Level_WO = 0
                Dta.FillBy_TglSampaiDariNLevelWO(Tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, StatusWO_1, StatusWO_2, Level_WO)
            Else
                Dta.FillBy_TglSampaiDari(Tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, StatusWO_1, StatusWO_2)
            End If

            FilterByMO(Tbl)
            FilterByProduct(Tbl)

            If Me.CboFormat.SelectedIndex = EnumFormat.Work_Unit_Recap Then
                TypeFilter = "Work_Unit"
            ElseIf Me.CboFormat.SelectedIndex = EnumFormat.Process_Recap Then
                TypeFilter = "Process"
            End If

            If TypeFilter <> "" Then
                TblDummy = Tbl.Copy
                Tbl.Rows.Clear()

                For Each RowDummy In TblDummy.Rows
                    With RowDummy
                        If TypeFilter = "Work_Unit" Then
                            Tbl.DefaultView.RowFilter = "[Primary Work Unit] = " & .Primary_Work_Unit & " AND [Primary Product WO] = " & .Primary_Product_WO
                        Else
                            Tbl.DefaultView.RowFilter = "[Primary Process] = " & .Primary_Process & " AND [Primary Product WO] = " & .Primary_Product_WO
                        End If
                        If Tbl.DefaultView.Count > 0 Then
                            Row = Tbl.DefaultView(0).Row
                            Row.Qty_Actual_Prod += .Qty_Actual_Prod
                            Row.Qty_Target_Prod += .Qty_Target_Prod
                            Row.Qty_Actual_Result += .Qty_Actual_Result
                            Row.Qty_Target_Result += .Qty_Target_Result
                            Row.Actual_Reject_Qty += .Actual_Reject_Qty
                            Row.Plan_Reject_Qty += .Plan_Reject_Qty
                            'Row.Actual_Good_Perct = (Row.Qty_Actual_Result * 100) / Row.Qty_Actual_Prod
                            'Row.Plan_Good_Perct = (Row.Qty_Target_Result * 100) / Row.Qty_Target_Prod
                            Row.Actual_Setup_Time += .Actual_Setup_Time
                            Row.Plan_Setup_Time += .Plan_Setup_Time
                            Row.Actual_Work_Hour += .Actual_Work_Hour
                            Row.Plan_Work_Hour += .Plan_Work_Hour
                            Row.Actual_Delay_Time += .Actual_Delay_Time
                            Row.Plan_Delay_Time += .Plan_Delay_Time
                            Row.Actual_Total_Time += .Actual_Total_Time
                            Row.Plan_Total_Time += .Plan_Total_Time
                            'Row.Actual_Prod_Rate += .Actual_Prod_Rate
                            'Row.Plan_Prod_Rate += .Plan_Prod_Rate
                        Else
                            Row = Tbl.NewDataTableRow
                            For ii As Integer = 0 To TblDummy.Columns.Count - 1
                                Row.Item(ii) = RowDummy.Item(ii)
                            Next
                            Tbl.Rows.Add(Row)
                        End If
                    End With
                Next
            End If

            For Each Row In Tbl.Rows
                With Row
                    RowC = TblC.NewTbl_CetakRow
                    RowC.Plan_End_Date = .Plan_End_Time
                    RowC.MO_Number = .MO_Number
                    RowC.Tanggal_MO = .Tanggal_MO
                    RowC.JO_Number = .JO_Number
                    RowC.Actual_Finish_Date = .Actual_Finish_Date
                    RowC.Target_Finish_Date = .Target_Finish_Date
                    RowC.Target_Finish_Date_R1 = .Target_Finish_Date_R1
                    RowC.Target_Finish_Date_R2 = .Target_Finish_Date_R2
                    RowC.WO_Number = .WO_Number
                    RowC.Status_WO = .Status_WO
                    RowC.Level_WO = .Level_WO
                    RowC.Product_Code = .Product_Code
                    RowC.Unit_Code_WO = .Unit_Code_WO
                    RowC.__Decimal__WO = .__Decimal__WO
                    RowC.work_unit_code = .work_unit_code
                    RowC.description = .description
                    RowC.Process_Code = .Process_Code
                    RowC.Process_Desc = .Process_Desc
                    'berkaitan dengan qty
                    RowC.Qty_Target_Prod = .Qty_Target_Prod
                    RowC.Qty_Actual_Prod = .Qty_Actual_Prod
                    RowC.Qty_Target_Result = .Qty_Target_Result
                    RowC.Qty_Actual_Result = .Qty_Actual_Result
                    RowC.Plan_Reject_Qty = .Qty_Target_Prod - .Qty_Target_Result
                    RowC.Actual_Reject_Qty = .Qty_Actual_Prod - .Qty_Actual_Result
                    If .Qty_Actual_Prod <> 0 And .Qty_Target_Prod <> 0 Then
                        RowC.Progress = (.Qty_Actual_Prod / .Qty_Target_Prod) * 100
                    End If
                    If .Plan_Work_Hour > 0 Then
                        RowC.Plan_Prod_Rate = Me.ObjCD.DigitGrouping(.Qty_Target_Prod / (.Plan_Work_Hour / 60), .__Decimal__WO, Core_Module.ClsCekData.EnumTipeRounding.Down)
                    End If
                    If .Actual_Work_Hour > 0 Then
                        RowC.Actual_Prod_Rate = Me.ObjCD.DigitGrouping(.Qty_Actual_Prod / (.Actual_Work_Hour / 60), .__Decimal__WO, Core_Module.ClsCekData.EnumTipeRounding.Down)
                    End If
                    If .Qty_Target_Result <> 0 And .Qty_Target_Prod <> 0 Then
                        RowC.Plan_Good_Perct = (.Qty_Target_Result * 100) / .Qty_Target_Prod
                    End If
                    If .Qty_Actual_Result <> 0 And .Qty_Actual_Prod <> 0 Then
                        RowC.Actual_Good_Perct = (.Qty_Actual_Result * 100) / .Qty_Actual_Prod
                    End If
                    'berkaitan dengan time
                    RowC.Plan_Setup_Time = Me.ObjCD.Convert_ToString(.Plan_Setup_Time)
                    RowC.Actual_Setup_Time = Me.ObjCD.Convert_ToString(.Actual_Setup_Time)
                    RowC.Plan_Work_Hour = Me.ObjCD.Convert_ToString(.Plan_Work_Hour)
                    RowC.Actual_Work_Hour = Me.ObjCD.Convert_ToString(.Actual_Work_Hour)
                    RowC.Plan_Delay_Time = Me.ObjCD.Convert_ToString(.Plan_Delay_Time)
                    RowC.Actual_Delay_Time = Me.ObjCD.Convert_ToString(.Actual_Delay_Time)
                    RowC.Plan_Total_Time = Me.ObjCD.Convert_ToString(.Plan_Total_Time)
                    RowC.Actual_Total_Time = Me.ObjCD.Convert_ToString(.Actual_Total_Time)
                    'tolerance berkaitan untuk filter 
                    If .Qty_Target_Prod <> 0 And .Qty_Actual_Prod <> 0 Then
                        RowC.Tolerance_Qty_Produce = Me.ObjCD.DigitGrouping(((.Qty_Actual_Prod / .Qty_Target_Prod) * 100) - 100, 2)
                    ElseIf .Qty_Target_Prod <> 0 And .Qty_Actual_Prod = 0 Then
                        RowC.Tolerance_Qty_Produce = -100
                    ElseIf .Qty_Target_Prod = 0 And .Qty_Actual_Prod <> 0 Then
                        RowC.Tolerance_Qty_Produce = 100
                    End If
                    If .Qty_Target_Result <> 0 And .Qty_Actual_Result <> 0 Then
                        RowC.Tolerance_Qty_Result = Me.ObjCD.DigitGrouping(((.Qty_Actual_Result / .Qty_Target_Result) * 100) - 100, 2)
                    ElseIf .Qty_Target_Result <> 0 And .Qty_Actual_Result = 0 Then
                        RowC.Tolerance_Qty_Result = -100
                    ElseIf .Qty_Target_Result = 0 And .Qty_Actual_Result <> 0 Then
                        RowC.Tolerance_Qty_Result = 100
                    End If
                    If RowC.Plan_Good_Perct <> 0 And RowC.Actual_Good_Perct <> 0 Then
                        RowC.Tolerance_GP = Me.ObjCD.DigitGrouping(((RowC.Actual_Good_Perct / RowC.Plan_Good_Perct) * 100) - 100, 2)
                    ElseIf RowC.Plan_Good_Perct <> 0 And RowC.Actual_Good_Perct = 0 Then
                        RowC.Tolerance_GP = -100
                    ElseIf RowC.Plan_Good_Perct = 0 And RowC.Actual_Good_Perct <> 0 Then
                        RowC.Tolerance_GP = 100
                    End If
                    If .Plan_Prod_Rate <> 0 And .Actual_Prod_Rate <> 0 Then
                        RowC.Tolerance_ProdRate = Me.ObjCD.DigitGrouping(((.Actual_Prod_Rate / .Plan_Prod_Rate) * 100) - 100, 2)
                    ElseIf .Plan_Prod_Rate <> 0 And .Actual_Prod_Rate = 0 Then
                        RowC.Tolerance_ProdRate = -100
                    ElseIf .Plan_Prod_Rate = 0 And .Actual_Prod_Rate <> 0 Then
                        RowC.Tolerance_ProdRate = 100
                    End If

                    If Me.CboFormat.SelectedIndex = EnumFormat.Work_Unit_Recap Then
                        RowC.Group = .work_unit_code
                    ElseIf Me.CboFormat.SelectedIndex = EnumFormat.Process_Recap Then
                        RowC.Group = .Process_Code
                    End If

                    RowC.primary_mo = .Primary
                    RowC.primary_process = .Primary_Process
                    RowC.identifier_jo = .Identifier_JO
                    RowC.urutan = 0
                    TblC.Rows.Add(RowC)
                End With
            Next

            If Me.ChkToleranceQtyProduce.Checked = True Or Me.ChkToleranceQtyResult.Checked = True _
            Or Me.ChkToleranceGP.Checked = True Or Me.ChkToleranceProdRate.Checked = True Then
                FilterByTolerance(TblC)
            End If

            Me.UrutkanBerdasarkanProcess(TblC)

            'format cetakan
            If Me.CboFormat.SelectedIndex = EnumFormat.Productivity_by_JO Then
                Dim Report As New CrpLaporanProductivity_ByJODetail
                Report.SetDataSource(Ds)
                Report.SetParameterValue(Report.Parameter_ParamMO_DateFrom.ParameterFieldName, Me.dtpFrom.Value.Date)
                Report.SetParameterValue(Report.Parameter_ParamMO_DateTo.ParameterFieldName, Me.dtpTo.Value.Date)
                Report.SetParameterValue(Report.Parameter_ParamMO_No.ParameterFieldName, MO_No)
                Crv.ReportSource = Report
            ElseIf Me.CboFormat.SelectedIndex = EnumFormat.Productivity_Recap Then
                Dim Report As New CrpLaporanProductivity_ByJORecap
                Report.SetDataSource(Ds)
                Report.SetParameterValue(Report.Parameter_ParamMO_DateFrom.ParameterFieldName, Me.dtpFrom.Value.Date)
                Report.SetParameterValue(Report.Parameter_ParamMO_DateTo.ParameterFieldName, Me.dtpTo.Value.Date)
                Report.SetParameterValue(Report.Parameter_ParamMO_No.ParameterFieldName, MO_No)
                Crv.ReportSource = Report
            ElseIf Me.CboFormat.SelectedIndex = EnumFormat.Work_Unit_Recap _
            Or Me.CboFormat.SelectedIndex = EnumFormat.Process_Recap Then
                Dim Report As New CrpLaporanProductivity_Recap
                Report.SetDataSource(Ds)
                Report.SetParameterValue(Report.Parameter_ParamMO_DateFrom.ParameterFieldName, Me.dtpFrom.Value.Date)
                Report.SetParameterValue(Report.Parameter_ParamMO_DateTo.ParameterFieldName, Me.dtpTo.Value.Date)
                Report.SetParameterValue(Report.Parameter_ParamMO_No.ParameterFieldName, MO_No)
                If Me.CboFormat.SelectedIndex = EnumFormat.Work_Unit_Recap Then
                    Report.SetParameterValue(Report.Parameter_Group.ParameterFieldName, "Work Unit")
                Else
                    Report.SetParameterValue(Report.Parameter_Group.ParameterFieldName, "Process")
                End If
                Crv.ReportSource = Report
            ElseIf Me.CboFormat.SelectedIndex = EnumFormat.Work_Unit_Detail _
            Or Me.CboFormat.SelectedIndex = EnumFormat.Process_Detail Then
                Dim Report As New CrpLaporanProductivity_Detail
                Report.SetDataSource(Ds)
                Report.SetParameterValue(Report.Parameter_ParamMO_DateFrom.ParameterFieldName, Me.dtpFrom.Value.Date)
                Report.SetParameterValue(Report.Parameter_ParamMO_DateTo.ParameterFieldName, Me.dtpTo.Value.Date)
                Report.SetParameterValue(Report.Parameter_ParamMO_No.ParameterFieldName, MO_No)
                If Me.CboFormat.SelectedIndex = EnumFormat.Work_Unit_Detail Then
                    Report.SetParameterValue(Report.Parameter_ParamGroup.ParameterFieldName, "Work Unit")
                Else
                    Report.SetParameterValue(Report.Parameter_ParamGroup.ParameterFieldName, "Process")
                End If
                Crv.ReportSource = Report
                End If

                Crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            Tbl.Dispose()
            Ds.Dispose()
        End Try
    End Sub

    Private Sub UrutkanBerdasarkanProcess(ByVal TblC As DseLaporanProductivity.Tbl_CetakDataTable)
        Dim rowC As DseLaporanProductivity.Tbl_CetakRow
        Dim TblDummy As New DseLaporanProductivity.Tbl_CetakDataTable
        Dim rowDummy As DseLaporanProductivity.Tbl_CetakRow
        Dim DTAUrutanProses As New DseLaporanProductivityTableAdapters.tbltrans_mo_processorderTableAdapter
        Dim TBLUrutanProses As New DseLaporanProductivity.tbltrans_mo_processorderDataTable
        Dim rowUrutanProses As DseLaporanProductivity.tbltrans_mo_processorderRow
        Dim DummyIdentifierJO As String = ""
        Dim DV As Data.DataRowView

        Try
            DTAUrutanProses.Connection.ConnectionString = Core_Module.GetConstring
            DTAUrutanProses.Fill(TBLUrutanProses, Me.dtpTo.Value, Me.dtpFrom.Value)

            'urutankan berdasarkan primary mo, identifier jo, primary proses
            TblC.DefaultView.RowFilter = ""
            TblC.DefaultView.Sort = "[primary mo], [identifier jo], [primary process]"
            While TblC.DefaultView.Count > 0
                rowC = TblC.DefaultView(0).Row
                If DummyIdentifierJO <> rowC.identifier_jo Then
                    DummyIdentifierJO = rowC.identifier_jo
                    'isikan urutan proses
                    Me.IsikanUrutanProses(rowC.identifier_jo, rowC.primary_mo, TBLUrutanProses, TblC)
                End If
                'pindahkan ke tbl dummy
                rowDummy = TblDummy.NewTbl_CetakRow
                For ij As Integer = 0 To TblDummy.Columns.Count - 1
                    rowDummy.Item(ij) = rowC.Item(ij)
                Next
                TblDummy.Rows.Add(rowDummy)
                TblC.Rows.Remove(rowC)
            End While

            'urutkan tblDummy berdasarkan tanggal mo, no mo, no jo, urutan process, no wo
            TblC.Rows.Clear()
            TblDummy.DefaultView.Sort = "[Tanggal MO], [MO Number], [JO Number], urutan, [WO Number]"
            'pindahkan ke tblC
            For Each dv In TblDummy.DefaultView
                rowDummy = DV.Row
                rowC = TblC.NewTbl_CetakRow
                For ij As Integer = 0 To TblC.Columns.Count - 1
                    rowC.Item(ij) = rowDummy.Item(ij)
                Next
                TblC.Rows.Add(rowC)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblDummy.Rows.Clear()
            TblDummy.Dispose()
        End Try
    End Sub

    Private Sub IsikanUrutanProses(ByVal vIdentifierJO As String, ByVal vPrimaryMO As Double, _
    ByVal TBLUrutanProses As DseLaporanProductivity.tbltrans_mo_processorderDataTable, _
    ByVal TblC As DseLaporanProductivity.Tbl_CetakDataTable)
        Dim DV As Data.DataRowView
        Dim rowC As DseLaporanProductivity.Tbl_CetakRow
        Dim rowUrutanProses As DseLaporanProductivity.tbltrans_mo_processorderRow
        Dim idx As Integer = 0

        Try
            For Each DV In TblC.DefaultView
                rowC = DV.Row
                If rowC.primary_mo = vPrimaryMO And rowC.identifier_jo = vIdentifierJO Then
                    For Each rowUrutanProses In TBLUrutanProses.Rows
                        If rowUrutanProses.primary_mo = vPrimaryMO Then
                            If rowUrutanProses.identifier_jo = vIdentifierJO Then
                                If rowUrutanProses.primary_process = rowC.primary_process Then
                                    rowC.urutan = rowUrutanProses.urutan
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                Else
                    Exit For
                End If
            Next

            'buang baris pada tbl urutan proses sehingga process mencari menjadi lebih cepat
            While idx <= TBLUrutanProses.Rows.Count - 1
                rowUrutanProses = TBLUrutanProses.Rows(idx)
                If rowUrutanProses.primary_mo < vPrimaryMO Then
                    TBLUrutanProses.Rows.Remove(rowUrutanProses)
                ElseIf rowUrutanProses.primary_mo = vPrimaryMO Then
                    If rowUrutanProses.identifier_jo = vIdentifierJO Then
                        TBLUrutanProses.Rows.Remove(rowUrutanProses)
                    Else
                        idx += 1
                    End If
                Else
                    Exit While
                End If 
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub TxtToleranceQtyProduce_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtToleranceQtyProduce.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtToleranceQtyProduce, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceQtyProduce_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtToleranceQtyProduce.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtToleranceQtyProduce, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceQtyProduce_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtToleranceQtyProduce.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtToleranceQtyProduce, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceQtyProduce_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtToleranceQtyProduce.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtToleranceQtyProduce, sender, e, 2, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceQtyResult_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtToleranceQtyResult.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtToleranceQtyResult, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceQtyResult_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtToleranceQtyResult.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtToleranceQtyResult, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceQtyResult_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtToleranceQtyResult.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtToleranceQtyResult, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceQtyResult_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtToleranceQtyResult.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtToleranceQtyResult, sender, e, 2, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceGP_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtToleranceGP.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtToleranceGP, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceGP_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtToleranceGP.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtToleranceGP, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceGP_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtToleranceGP.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtToleranceGP, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceGP_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtToleranceGP.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtToleranceGP, sender, e, 2, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceProdRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtToleranceProdRate.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtToleranceProdRate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceProdRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtToleranceProdRate.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtToleranceProdRate, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceProdRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtToleranceProdRate.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtToleranceProdRate, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtToleranceProdRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtToleranceProdRate.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtToleranceProdRate, sender, e, 2, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvProduct.CellBeginEdit
        Try
            If Me.colProduct_GroupLevel.Index Or Me.colProduct_Code.Index Then
                Me.Buffer = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProduct.CellEndEdit

        Try
            If IsNothing(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If IsNothing(Me.Buffer) = True Then
                        Me.index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                        Exit Sub
                    End If
                    If Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colProduct_GroupLevel.Index Then
                        cellEndEditGroupLevel(sender, e)
                    End If
                    If e.ColumnIndex = Me.colProduct_Code.Index Then
                        cellEndEditCode(sender, e)
                    End If
                Else
                    If IsNothing(Me.Buffer) = False Then
                        If String.IsNullOrEmpty(Buffer) = False Then
                            Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    Me.index = e.RowIndex
                    Me.BgwProduct.RunWorkerAsync()
                End If
                'Else
                '    If IsNothing(Me.buffer) = False Then
                '        If String.IsNullOrEmpty(buffer) = False Then
                '            Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                '            Exit Sub
                '        End If
                '    End If
                '    Me.index = e.RowIndex
                '    Me.BgwProduct.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditGroupLevel(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindGroup As New Core_Module.FrmCariMasterProductGroup
        Dim frmFindCategory As New Core_Module.frmCariMasterCategory
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            Select Case Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value
                Case "Group"
                    frmFindGroup.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 3
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                    End If
                Case "Brand"
                    frmFindGroup.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 2
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                    End If
                Case "Sub Brand"
                    frmFindGroup.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 1
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                    End If
                Case "Category"
                    frmFindCategory.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindCategory.ShowDialog()
                    If frmFindCategory.PrimaryCategory > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindCategory.PrimaryCategory
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objCategory.GetCategoryCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                Case "Product"
                    frmFindProduct.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindProduct.ShowDialog()
                    If frmFindProduct.PrimaryProduct > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindProduct.PrimaryProduct
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objProduct.GetProductCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
            End Select
            If Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexProduct(e.RowIndex, Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value, Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value & " " & Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.index = e.RowIndex
                    Me.BgwProduct.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cellEndEditCode(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindGroup As New Core_Module.FrmCariMasterProductGroup
        Dim frmFindCategory As New Core_Module.frmCariMasterCategory
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
            Select Case Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value
                Case "Group"
                    If Me.objGroup.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 3) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 3)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 3
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
                Case "Brand"
                    If Me.objGroup.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 2) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 2
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
                Case "Sub Brand"
                    If Me.objGroup.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 1) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 1)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 1
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
                Case "Category"
                    If Me.objCategory.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objCategory.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindCategory.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindCategory.ShowDialog()
                        If frmFindCategory.PrimaryCategory > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindCategory.PrimaryCategory
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objCategory.GetCategoryCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
                Case "Product"
                    If Me.objProduct.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objProduct.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindProduct.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindProduct.ShowDialog()
                        If frmFindProduct.PrimaryProduct > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindProduct.PrimaryProduct
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objProduct.GetProductCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                        End If
                    End If
            End Select
            If Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexProduct(e.RowIndex, Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value, Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value & " " & Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.index = e.RowIndex
                    Me.BgwProduct.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getIndexProduct(ByVal vRowException As Integer, ByVal vGroupLevel As String, ByVal vPrimary As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvProduct.Rows
                With dr
                    If Me.dgvProduct.Rows.IndexOf(dr) <> vRowException Then
                        If dr.Cells(Me.colProduct_Primary.Index).Value <> -1 Then
                            If dr.Cells(Me.colProduct_GroupLevel.Index).Value = vGroupLevel And dr.Cells(Me.colProduct_Primary.Index).Value = vPrimary Then
                                Return Me.dgvProduct.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvProduct_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvProduct.DataError

    End Sub

    Private Sub dgvProduct_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvProduct.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvProduct.Rows(Me.dgvProduct.Rows.Count - 1)
            dr.Cells(Me.colProduct_No.Index).Value = Me.dgvProduct.Rows.Count
            dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Group"
            dr.Cells(Me.colProduct_Desc.Index).Value = ""
            dr.Cells(Me.colProduct_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvProduct.RowsRemoved
        Try
            Me.objCd.SetNumber(Me.dgvProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvProduct.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMO_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvMO.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.colMO_DocNo.Index Then
                e.Cancel = True
            Else
                Me.Buffer = Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getIndexMO(ByVal vRowException As Integer, ByVal vPrimary As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvMO.Rows
                With dr
                    If Me.dgvMO.Rows.IndexOf(dr) <> vRowException Then
                        If .Cells(Me.colMO_Primary.Index).Value <> -1 Then
                            If .Cells(Me.colMO_Primary.Index).Value = vPrimary Then
                                Return Me.dgvMO.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvMO_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvMO.CellEndEdit
        Dim frm As New FrmTrans_ManufacturingOrder_Cari
        Try
            If IsNothing(Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.Buffer = Me.DgvMO(e.ColumnIndex, e.RowIndex).Value Then
                        Me.index = e.RowIndex
                        Me.BgwMO.RunWorkerAsync()
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colMO_DocNo.Index Then
                        If Me.objMO.Apakahsudahada(Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Me.DgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value = Me.objMO.GetPrimary(Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.DgvMO.Item(Me.colMO_Date.Index, e.RowIndex).Value = Format(Me.objMO.GetDate(Me.DgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value), "dd MMM yyyy")
                            If getIndexMO(e.RowIndex, Me.DgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value) <> -1 Then
                                MsgBox("Manufacturing Order " & Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                If IsNothing(Me.Buffer) = False Then
                                    If String.IsNullOrEmpty(Buffer) = False Then
                                        Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        Me.DgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value = Me.objMO.GetPrimary(Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value)
                                        Me.DgvMO.Item(Me.colMO_Date.Index, e.RowIndex).Value = Format(Me.objMO.GetDate(Me.DgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value), "dd MMM yyyy")
                                        Exit Sub
                                    End If
                                End If
                                Me.index = e.RowIndex
                                Me.BgwMO.RunWorkerAsync()
                            End If
                        Else
                            frm.txtCari.Text = Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value
                            frm.ShowDialog()
                            If frm.Primary > 0 Then
                                Me.DgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value = frm.Primary
                                Me.DgvMO.Item(Me.colMO_DocNo.Index, e.RowIndex).Value = Me.objMO.GetNoTransaksi(Me.DgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value)
                                Me.DgvMO.Item(Me.colMO_Date.Index, e.RowIndex).Value = Format(Me.objMO.GetDate(Me.DgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value), "dd MMM yyyy")
                                If getIndexMO(e.RowIndex, Me.DgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value) <> -1 Then
                                    MsgBox("Manufacturing Order " & Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                    If IsNothing(Me.Buffer) = False Then
                                        If String.IsNullOrEmpty(Buffer) = False Then
                                            Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            Me.DgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value = Me.objMO.GetPrimary(Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value)
                                            Me.DgvMO.Item(Me.colMO_Date.Index, e.RowIndex).Value = Format(Me.objMO.GetDate(Me.DgvMO.Item(Me.colMO_Primary.Index, e.RowIndex).Value), "dd MMM yyyy")
                                            Exit Sub
                                        End If
                                    End If
                                    Me.index = e.RowIndex
                                    Me.BgwMO.RunWorkerAsync()
                                End If
                            Else
                                If IsNothing(Me.Buffer) = False Then
                                    If String.IsNullOrEmpty(Buffer) = False Then
                                        Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        Exit Sub
                                    End If
                                End If
                                Me.index = e.RowIndex
                                Me.BgwMO.RunWorkerAsync()
                            End If
                        End If
                    End If
                    'Else
                    '    If IsNothing(Me.Buffer) = False Then
                    '        If String.IsNullOrEmpty(Buffer) = False Then
                    '            Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    '            Exit Sub
                    '        End If
                    '    End If
                    '    Me.index = e.RowIndex
                    '    Me.BgwMO.RunWorkerAsync()
                End If
            Else
                If IsNothing(Me.Buffer) = False Then
                    If String.IsNullOrEmpty(Buffer) = False Then
                        Me.DgvMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
                Me.index = e.RowIndex
                Me.BgwMO.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMO_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvMO.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvMO.Rows(Me.dgvMO.Rows.Count - 1)
            dr.Cells(Me.colMO_No.Index).Value = Me.dgvMO.Rows.Count
            dr.Cells(Me.colMO_DocNo.Index).Value = ""
            dr.Cells(Me.colMO_Date.Index).Value = ""
            dr.Cells(Me.colMO_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMO_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvMO.RowsRemoved
        Try
            Me.objCd.SetNumber(Me.dgvMO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMO_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvMO.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvMO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwProduct_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwProduct.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwProduct_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwProduct.RunWorkerCompleted
        Try
            Me.dgvProduct.Rows.RemoveAt(Me.index)
            Me.ObjCD.SetNumber(Me.dgvProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwMO_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwMO.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwMO_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwMO.RunWorkerCompleted
        Try
            Me.DgvMO.Rows.RemoveAt(Me.index)
            Me.ObjCD.SetNumber(Me.DgvMO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkToleranceQtyProduce_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkToleranceQtyProduce.CheckedChanged
        Try
            If Me.ChkToleranceQtyProduce.Checked = False Then
                Me.CboQtyProduce.Enabled = False
                Me.TxtToleranceQtyProduce.Enabled = False
            Else
                Me.CboQtyProduce.Enabled = True
                If Me.CboProdRate.SelectedIndex <> EnumQtyProduce.SamaDengan Then
                    Me.TxtToleranceQtyProduce.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkToleranceQtyResult_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkToleranceQtyResult.CheckedChanged
        Try
            If Me.ChkToleranceQtyResult.Checked = False Then
                Me.CboQtyResult.Enabled = False
                Me.TxtToleranceQtyResult.Enabled = False
            Else
                Me.CboQtyResult.Enabled = True
                If Me.CboQtyResult.SelectedIndex <> EnumQtyResult.SamaDengan Then
                    Me.TxtToleranceQtyResult.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkToleranceGP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkToleranceGP.CheckedChanged
        Try
            If Me.ChkToleranceGP.Checked = False Then
                Me.CboGP.Enabled = False
                Me.TxtToleranceGP.Enabled = False
            Else
                Me.CboGP.Enabled = True
                If Me.CboGP.SelectedIndex <> EnumGP.SamaDengan Then
                    Me.TxtToleranceGP.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkToleranceProdRate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkToleranceProdRate.CheckedChanged
        Try
            If Me.ChkToleranceProdRate.Checked = False Then
                Me.CboProdRate.Enabled = False
                Me.TxtToleranceProdRate.Enabled = False
            Else
                Me.CboProdRate.Enabled = True
                If Me.CboProdRate.SelectedIndex <> EnumProdRate.SamaDengan Then
                    Me.TxtToleranceProdRate.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboQtyProduce_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboQtyProduce.SelectedIndexChanged
        Try
            If Me.CboQtyProduce.SelectedIndex = EnumProdRate.SamaDengan Then
                Me.TxtToleranceQtyProduce.Enabled = False
                Me.TxtToleranceQtyProduce.Text = "0.00"
            Else
                Me.TxtToleranceQtyProduce.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboQtyResult_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboQtyResult.SelectedIndexChanged
        Try
            If Me.CboQtyResult.SelectedIndex = EnumProdRate.SamaDengan Then
                Me.TxtToleranceQtyResult.Enabled = False
                Me.TxtToleranceQtyResult.Text = "0.00"
            Else
                Me.TxtToleranceQtyResult.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboGP_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboGP.SelectedIndexChanged
        Try
            If Me.CboGP.SelectedIndex = EnumProdRate.SamaDengan Then
                Me.TxtToleranceGP.Enabled = False
                Me.TxtToleranceGP.Text = "0.00"
            Else
                Me.TxtToleranceGP.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboProdRate_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboProdRate.SelectedIndexChanged
        Try
            If Me.CboProdRate.SelectedIndex = EnumProdRate.SamaDengan Then
                Me.TxtToleranceProdRate.Enabled = False
                Me.TxtToleranceProdRate.Text = "0.00"
            Else
                Me.TxtToleranceProdRate.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class