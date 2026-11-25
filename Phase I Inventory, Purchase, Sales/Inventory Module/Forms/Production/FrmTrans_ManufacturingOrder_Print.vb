''@aul
''20150902 1419
Option Explicit On
Option Strict Off
Public Class FrmTrans_ManufacturingOrder_Print

    Public PrimaryMO As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public PaperSize As FrmTrans_ManufacturingOrder_PrintOption.EnumPaperSize = FrmTrans_ManufacturingOrder_PrintOption.EnumPaperSize.A5
    Dim rptA4 As New CrpCetak_ManufacturingOrderA4
    Dim rptBanyakA4 As New CrpCetak_ManufacturingOrder_BatchA4

    Dim rpt As New CrpCetak_ManufacturingOrder
    Dim rptBanyak As New CrpCetak_ManufacturingOrder_Batch
    Public val As New List(Of Double)
    Public pilihan As Boolean = False

    Private Sub FrmTrans_ManufacturingOrder_Print_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim UseCustomPrint As String = ""

        Try
            UseCustomPrint = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ManufacturingOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If UseCustomPrint = "1" Then

            Else
                If pilihan = False Then
                    Cetak()
                Else
                    cetakBanyak()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cetak()
        Dim DS As New DseTrans_ManufacturingOrder_Print

        Dim DTAH As New DseTrans_ManufacturingOrder_PrintTableAdapters.DetailTableAdapter
        Dim DTAWO As New DseTrans_ManufacturingOrder_PrintTableAdapters.tbltrans_mo_item_woTableAdapter

        Dim tblH As New DseTrans_ManufacturingOrder_Print.DetailDataTable
        Dim tblCetak As DseTrans_ManufacturingOrder_Print.CetakDataTable = DS.Cetak
        Dim tblWO As New DseTrans_ManufacturingOrder_Print.tbltrans_mo_item_woDataTable

        Dim rowH As DseTrans_ManufacturingOrder_Print.DetailRow
        Dim rowWO As DseTrans_ManufacturingOrder_Print.tbltrans_mo_item_woRow
        Dim rowCetak As DseTrans_ManufacturingOrder_Print.CetakRow


        Try
            DTAH.Connection.ConnectionString = Core_Module.GetConstring
            DTAWO.Connection.ConnectionString = DTAH.Connection.ConnectionString

            tblCetak.Rows.Clear()
            DTAH.Fill(tblH, Me.PrimaryMO)
            DTAWO.Fill(tblWO, Me.PrimaryMO)



            If tblH.Rows.Count > 0 Then
                For Each rowH In tblH.Rows
                    With rowH
                        tblWO.DefaultView.RowFilter = Nothing
                        tblWO.DefaultView.RowFilter = "[Identifier JO] = '" & .Identifier & _
                        "' And [Primary Product] = " & .Primary_Product_Code_RM
                        If tblWO.DefaultView.Count = 0 Then
                            rowCetak = tblCetak.NewCetakRow
                            rowCetak.Primary_MO = Me.PrimaryMO
                            rowCetak.MO_Number = .MO_Number
                            rowCetak.MO_Date = ._Date
                            rowCetak.Remarks = .Remarks
                            rowCetak.JO_Number = .JO_Number
                            rowCetak.Primary_Product = .Primary_Product
                            rowCetak.Product_Code = .Product_Code_FG
                            rowCetak.Product_Desc = .Product_Desc_FG
                            rowCetak.Primary_Work_Unit = -1
                            rowCetak.Work_Unit = ""
                            rowCetak.Primary_Process = -1
                            rowCetak.Process_Code = ""
                            rowCetak.Qty_Target_Result = .Target_Qty
                            rowCetak.Qty_Target_Produce = .Produce_Qty
                            If .Actual_Finish_Date <> #1/1/1950# Then
                                rowCetak.Actual_Finish_Date = Format(CDate(.Actual_Finish_Date), "dd MMM yyyy")
                            End If
                            If .Target_Finish_Date <> #1/1/1950# Then
                                rowCetak.Target_Finish_Date = Format(CDate(.Target_Finish_Date), "dd MMM yyyy")
                            End If
                            If .Target_Finish_Date_R1 <> #1/1/1950# Then
                                rowCetak.Target_Finish_Date_Rev_1 = Format(CDate(.Target_Finish_Date_R1), "dd MMM yyyy")
                            End If
                            If .Target_Finish_Date_R2 <> #1/1/1950# Then
                                rowCetak.Target_Finisih_Date_Rev_2 = Format(CDate(.Target_Finish_Date_R2), "dd MMM yyyy")
                            End If
                            rowCetak.Good_Percentage_Target = 0
                            rowCetak.Good_Percentage_Actual = 0
                            rowCetak.Primary_RM_Code = .Primary_Product_Code_RM
                            rowCetak.RM_Code = .Product_Code_RM
                            rowCetak.RM_Desc = .Product_Desc_RM
                            rowCetak.Qty_Target = .Target_Qty_RM
                            rowCetak.Qty_Required = .Required_Qty
                            rowCetak.Qty_Allocated = .Allocated_Qty
                            rowCetak.Qty_Additional = .Additional_Qty
                            rowCetak.GP_Actual = .GP_Actual
                            rowCetak.GP_Target = .GP_Target
                            rowCetak.Remarks_Item = .Remarks_item
                            rowCetak.Identifier_JO = .Identifier
                            rowCetak.DeffPrec_RM = .DeffPrecItem
                            rowCetak.DeffPrec = .Decimal_Product
                            tblCetak.Rows.Add(rowCetak)
                        End If
                    End With
                Next
            End If


            If tblCetak.Rows.Count > 0 Then
                For Each rowCetak In tblCetak.Rows
                    With rowCetak
                        tblWO.DefaultView.RowFilter = Nothing
                        tblWO.DefaultView.RowFilter = "[Identifier JO] = '" & .Identifier_JO & _
                        "' And [Level] = 0 "
                        If tblWO.DefaultView.Count > 0 Then
                            rowWO = tblWO.DefaultView(0).Row
                            rowCetak.Good_Percentage_Target = rowWO.Plan_Good_Perct
                            rowCetak.Good_Percentage_Actual = rowWO.Actual_Good_Perct
                            If rowWO.Primary_Work_Unit <> -1 Then
                                rowCetak.Primary_Work_Unit = rowWO.Primary_Work_Unit
                                rowCetak.Work_Unit = rowWO.work_unit_code
                            End If

                            If rowWO.Primary_Process <> -1 Then
                                rowCetak.Primary_Process = rowWO.Primary_Process
                                rowCetak.Process_Code = rowWO.Process_Code
                            End If
                        End If
                    End With
                Next
            End If

            If Me.PaperSize = FrmTrans_ManufacturingOrder_PrintOption.EnumPaperSize.A5 Then
                rpt.SetDataSource(DS)
                Me.CrvMO.ReportSource = rpt
            ElseIf Me.PaperSize = FrmTrans_ManufacturingOrder_PrintOption.EnumPaperSize.A4 Then
                rptA4.SetDataSource(DS)
                Me.CrvMO.ReportSource = rptA4
            End If
            
            Me.CrvMO.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAH.Dispose()
            tblH.Dispose()
            DTAWO.Dispose()
            tblWO.Dispose()
            tblCetak.Dispose()
        End Try
    End Sub

    Public Sub cetakBanyak()
        Dim DS As New DseTrans_ManufacturingOrder_Print
        Dim DTAH As New DseTrans_ManufacturingOrder_PrintTableAdapters.DetailTableAdapter
        Dim DTAWO As New DseTrans_ManufacturingOrder_PrintTableAdapters.tbltrans_mo_item_woTableAdapter

        Dim tblH As New DseTrans_ManufacturingOrder_Print.DetailDataTable
        Dim tblCetak As DseTrans_ManufacturingOrder_Print.CetakDataTable = DS.Cetak
        Dim tblCetakD As New DseTrans_ManufacturingOrder_Print.CetakDataTable
        Dim tblWO As New DseTrans_ManufacturingOrder_Print.tbltrans_mo_item_woDataTable

        Dim rowH As DseTrans_ManufacturingOrder_Print.DetailRow
        Dim rowWO As DseTrans_ManufacturingOrder_Print.tbltrans_mo_item_woRow
        Dim rowCetak As DseTrans_ManufacturingOrder_Print.CetakRow
        Dim rowCetakD As DseTrans_ManufacturingOrder_Print.CetakRow




        Dim lastCode As Double = -1
        Try
            DTAH.Connection.ConnectionString = Core_Module.GetConstring
            DTAWO.Connection.ConnectionString = DTAH.Connection.ConnectionString
            DTAH.ClearBeforeFill = False
            DTAWO.ClearBeforeFill = False
            For Each primary As Double In Me.val
                tblCetak.Rows.Clear()
                DTAH.Fill(tblH, primary)
                DTAWO.Fill(tblWO, primary)
            Next

            If tblH.Rows.Count > 0 Then
                For Each rowH In tblH.Rows
                    With rowH
                        tblWO.DefaultView.RowFilter = Nothing
                        tblWO.DefaultView.RowFilter = "[Identifier JO] = '" & .Identifier & _
                        "' And [Primary Product] = " & .Primary_Product_Code_RM
                        If tblWO.DefaultView.Count = 0 Then
                            rowCetak = tblCetak.NewCetakRow
                            rowCetak.Primary_MO = .Primary_MO
                            rowCetak.MO_Number = .MO_Number
                            rowCetak.MO_Date = ._Date
                            rowCetak.Remarks = .Remarks
                            rowCetak.JO_Number = .JO_Number
                            rowCetak.Primary_Product = .Primary_Product
                            rowCetak.Product_Code = .Product_Code_FG
                            rowCetak.Product_Desc = .Product_Desc_FG
                            rowCetak.Primary_Work_Unit = -1
                            rowCetak.Work_Unit = ""
                            rowCetak.Primary_Process = -1
                            rowCetak.Process_Code = ""
                            rowCetak.Qty_Target_Result = .Target_Qty
                            rowCetak.Qty_Target_Produce = .Produce_Qty
                            If .Actual_Finish_Date <> #1/1/1950# Then
                                rowCetak.Actual_Finish_Date = Format(CDate(.Actual_Finish_Date), "dd MMM yyyy")
                            End If
                            If .Target_Finish_Date <> #1/1/1950# Then
                                rowCetak.Target_Finish_Date = Format(CDate(.Target_Finish_Date), "dd MMM yyyy")
                            End If
                            If .Target_Finish_Date_R1 <> #1/1/1950# Then
                                rowCetak.Target_Finish_Date_Rev_1 = Format(CDate(.Target_Finish_Date_R1), "dd MMM yyyy")
                            End If
                            If .Target_Finish_Date_R2 <> #1/1/1950# Then
                                rowCetak.Target_Finisih_Date_Rev_2 = Format(CDate(.Target_Finish_Date_R2), "dd MMM yyyy")
                            End If
                            rowCetak.Good_Percentage_Target = 0
                            rowCetak.Good_Percentage_Actual = 0
                            rowCetak.Primary_RM_Code = .Primary_Product_Code_RM
                            rowCetak.RM_Code = .Product_Code_RM
                            rowCetak.RM_Desc = .Product_Desc_RM
                            rowCetak.Qty_Target = .Target_Qty_RM
                            rowCetak.Qty_Required = .Required_Qty
                            rowCetak.Qty_Allocated = .Allocated_Qty
                            rowCetak.Qty_Additional = .Additional_Qty
                            rowCetak.GP_Actual = .GP_Actual
                            rowCetak.GP_Target = .GP_Target
                            rowCetak.Remarks_Item = .Remarks_item
                            rowCetak.Identifier_JO = .Identifier
                            rowCetak.DeffPrec_RM = .DeffPrecItem
                            rowCetak.DeffPrec = .Decimal_Product
                            tblCetak.Rows.Add(rowCetak)

                            rowCetakD = tblCetakD.NewCetakRow
                            rowCetakD.Primary_MO = .Primary_MO
                            rowCetakD.MO_Number = .MO_Number
                            rowCetakD.MO_Date = ._Date
                            rowCetakD.Remarks = .Remarks
                            rowCetakD.JO_Number = .JO_Number
                            rowCetakD.Primary_Product = .Primary_Product
                            rowCetakD.Product_Code = .Product_Code_FG
                            rowCetakD.Product_Desc = .Product_Desc_FG
                            rowCetakD.Primary_Work_Unit = -1
                            rowCetakD.Work_Unit = ""
                            rowCetakD.Primary_Process = -1
                            rowCetakD.Process_Code = ""
                            rowCetakD.Qty_Target_Result = .Target_Qty
                            rowCetakD.Qty_Target_Produce = .Produce_Qty
                            If .Actual_Finish_Date <> #1/1/1950# Then
                                rowCetakD.Actual_Finish_Date = Format(CDate(.Actual_Finish_Date), "dd MMM yyyy")
                            End If
                            If .Target_Finish_Date <> #1/1/1950# Then
                                rowCetakD.Target_Finish_Date = Format(CDate(.Target_Finish_Date), "dd MMM yyyy")
                            End If
                            If .Target_Finish_Date_R1 <> #1/1/1950# Then
                                rowCetakD.Target_Finish_Date_Rev_1 = Format(CDate(.Target_Finish_Date_R1), "dd MMM yyyy")
                            End If
                            If .Target_Finish_Date_R2 <> #1/1/1950# Then
                                rowCetakD.Target_Finisih_Date_Rev_2 = Format(CDate(.Target_Finish_Date_R2), "dd MMM yyyy")
                            End If
                            rowCetakD.Good_Percentage_Target = 0
                            rowCetakD.Good_Percentage_Actual = 0
                            rowCetakD.Primary_RM_Code = .Primary_Product_Code_RM
                            rowCetakD.RM_Code = .Product_Code_RM
                            rowCetakD.RM_Desc = .Product_Desc_RM
                            rowCetakD.Qty_Target = .Target_Qty_RM
                            rowCetakD.Qty_Required = .Required_Qty
                            rowCetakD.Qty_Allocated = .Allocated_Qty
                            rowCetakD.Qty_Additional = .Additional_Qty
                            rowCetakD.GP_Actual = .GP_Actual
                            rowCetakD.GP_Target = .GP_Target
                            rowCetakD.Remarks_Item = .Remarks_item
                            rowCetakD.Identifier_JO = .Identifier
                            rowCetakD.DeffPrec_RM = .DeffPrecItem
                            rowCetakD.DeffPrec = .Decimal_Product
                            tblCetakD.Rows.Add(rowCetakD)
                        End If
                    End With
                Next
            End If


            If tblCetak.Rows.Count > 0 Then
                For Each rowCetak In tblCetak.Rows
                    With rowCetak
                        tblWO.DefaultView.RowFilter = Nothing
                        tblWO.DefaultView.RowFilter = "[Identifier JO] = '" & .Identifier_JO & _
                        "' And [Level] = 0 "
                        If tblWO.DefaultView.Count > 0 Then
                            rowWO = tblWO.DefaultView(0).Row
                            rowCetak.Good_Percentage_Target = rowWO.Plan_Good_Perct
                            rowCetak.Good_Percentage_Actual = rowWO.Actual_Good_Perct
                            If rowWO.Primary_Work_Unit <> -1 Then
                                rowCetak.Primary_Work_Unit = rowWO.Primary_Work_Unit
                                rowCetak.Work_Unit = rowWO.work_unit_code
                            End If

                            If rowWO.Primary_Process <> -1 Then
                                rowCetak.Primary_Process = rowWO.Primary_Process
                                rowCetak.Process_Code = rowWO.Process_Code
                            End If
                        End If
                    End With
                    lastCode = rowCetak.Primary_MO
                Next
            End If

            rptBanyak.SetDataSource(DS)
            rptBanyak.SetParameterValue(rptBanyak.Parameter_LastCode.ParameterFieldName, lastCode)
            Me.CrvMO.ReportSource = rptBanyak
            If Me.PaperSize = FrmTrans_ManufacturingOrder_PrintOption.EnumPaperSize.A5 Then
                rptBanyak.SetDataSource(DS)
                rptBanyak.SetParameterValue(rptBanyak.Parameter_LastCode.ParameterFieldName, lastCode)
                Me.CrvMO.ReportSource = rptBanyak
            Else
                rptBanyakA4.SetDataSource(DS)
                rptBanyakA4.SetParameterValue(rptBanyakA4.Parameter_LastCode.ParameterFieldName, lastCode)
                Me.CrvMO.ReportSource = rptBanyakA4
            End If
            Me.CrvMO.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAH.Dispose()
            tblH.Dispose()
            DTAWO.Dispose()
            tblWO.Dispose()
            tblCetak.Dispose()
        End Try

    End Sub

End Class