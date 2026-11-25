Option Explicit On
Option Strict Off

Public Class ClsMasterProductionRate_Item

    Dim objGroup As New ClsMasterProductGroup
    Dim objCategory As New ClsMasterCategory
    Dim objProduct As New clsMasterProduct
    Dim objCd As New Core_Module.ClsCekData

    Public Function SimpanItem(ByVal vFrm As FrmMasterProductionRate, ByVal vPrimaryTrans As Double) As Boolean
        Dim dtaWU As New DseMasterProductionRateTableAdapters.tblmsater_production_rate_item_workunitTableAdapter
        Dim dtaProduct As New DseMasterProductionRateTableAdapters.tbltrans_production_rate_item_productTableAdapter
        Try
            dtaWU.Connection.ConnectionString = Core_Module.GetConstring
            dtaProduct.Connection.ConnectionString = Core_Module.GetConstring
            For Each r As DataGridViewRow In vFrm.DgvWorkUnit.Rows
                If IsNothing(r.Cells(vFrm.ColWorkUnit_PrimaryWorkUnit.Index)) = False Then
                    If String.IsNullOrEmpty(r.Cells(vFrm.ColWorkUnit_PrimaryWorkUnit.Index).Value) = False Then
                        If r.Cells(vFrm.ColWorkUnit_PrimaryWorkUnit.Index).Value <> "-1" Then
                            Dim remarks As String = ""
                            If IsNothing(r.Cells(vFrm.ColWorkUnit_Remarks.Index).Value) = False Then
                                If String.IsNullOrEmpty(r.Cells(vFrm.ColWorkUnit_Remarks.Index).Value) = False Then
                                    remarks = r.Cells(vFrm.ColWorkUnit_Remarks.Index).Value
                                End If
                            End If
                            If dtaWU.InsertQuery(vPrimaryTrans, r.Cells(vFrm.ColWorkUnit_Priority.Index).Value, r.Cells(vFrm.ColWorkUnit_PrimaryWorkUnit.Index).Value, _
                                    r.Cells(vFrm.ColWorkUnit_Rate.Index).Value, r.Cells(vFrm.ColWorkUnit_Efficiency.Index).Value, r.Cells(vFrm.ColWorkUnit_EffectiveRate.Index).Value, remarks) <= 0 Then
                                MsgBox("Failed to save some work unit item(" & r.Cells(vFrm.ColWorkUnit_Code.Index).Value & ").", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        End If
                    End If
                End If
            Next
            Dim level As Double = -1
            For Each r As DataGridViewRow In vFrm.dgvItemProductGroup.Rows
                If IsNothing(r.Cells(vFrm.ColProductGroup_PrimaryAcuan.Index)) = False Then
                    If String.IsNullOrEmpty(r.Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value) = False Then
                        If r.Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value <> "-1" Then
                            Select Case r.Cells(vFrm.ColProductGroup_Level.Index).Value
                                Case "Product Group"
                                    level = 0
                                Case "Brand"
                                    level = 1
                                Case "Sub Brand"
                                    level = 2
                                Case "Category"
                                    level = 3
                                Case "Product"
                                    level = 4
                            End Select
                            Dim remarks As String = ""
                            If IsNothing(r.Cells(vFrm.ColProductGroup_Remarks.Index).Value) = False Then
                                If String.IsNullOrEmpty(r.Cells(vFrm.ColProductGroup_Remarks.Index).Value) = False Then
                                    remarks = r.Cells(vFrm.ColProductGroup_Remarks.Index).Value
                                End If
                            End If
                            If dtaProduct.InsertQuery(vPrimaryTrans, level, r.Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value, remarks) <= 0 Then
                                MsgBox("Failed to save some product group item.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        End If
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTrans As Double) As Boolean
        Dim dtaWU As New DseMasterProductionRateTableAdapters.tblmsater_production_rate_item_workunitTableAdapter
        Dim dtaProd As New DseMasterProductionRateTableAdapters.tbltrans_production_rate_item_productTableAdapter
        Try
            dtaWU.Connection.ConnectionString = Core_Module.GetConstring
            dtaProd.Connection.ConnectionString = Core_Module.GetConstring
            If dtaWU.DeleteQuery(vPrimaryTrans) <= 0 Then
                Return False
            End If
            If dtaProd.DeleteQuery(vPrimaryTrans) <= 0 Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function IsiDgvItem(ByVal vFrm As FrmMasterProductionRate, ByVal vPrimaryTrans As Double) As Boolean
        Dim dtaWU As New DseMasterProductionRateTableAdapters.DataTable2TableAdapter
        Dim tblWU As New DseMasterProductionRate.DataTable2DataTable
        Dim rowWU As DseMasterProductionRate.DataTable2Row
        Dim dtaP As New DseMasterProductionRateTableAdapters.tbltrans_production_rate_item_productTableAdapter
        Dim tblP As New DseMasterProductionRate.tbltrans_production_rate_item_productDataTable
        Dim rowP As DseMasterProductionRate.tbltrans_production_rate_item_productRow
        Try
            dtaWU.Connection.ConnectionString = Core_Module.GetConstring
            dtaWU.Fill(tblWU, vPrimaryTrans)
            vFrm.DgvWorkUnit.Rows.Clear()
            vFrm.DgvWorkUnit.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowWU In tblWU
                Dim dr As DataGridViewRow = vFrm.DgvWorkUnit.Rows(vFrm.DgvWorkUnit.Rows.Add())
                With dr
                    .Cells(vFrm.ColWorkUnit_Priority.Index).Value = Me.objCd.DigitGrouping(rowWU.priority, 0)
                    .Cells(vFrm.ColWorkUnit_Code.Index).Value = rowWU.Code
                    .Cells(vFrm.ColWorkUnit_Desc.Index).Value = rowWU.Desc
                    .Cells(vFrm.ColWorkUnit_EffectiveRate.Index).Value = Me.objCd.BuangNolDibelakangDecimal(Me.objCd.DigitGrouping(rowWU.effective_rate, 6))
                    .Cells(vFrm.ColWorkUnit_Efficiency.Index).Value = Me.objCd.DigitGrouping(rowWU.eficiency, 2)
                    .Cells(vFrm.ColWorkUnit_Rate.Index).Value = Me.objCd.BuangNolDibelakangDecimal(Me.objCd.DigitGrouping(rowWU.rate, 6))
                    .Cells(vFrm.ColWorkUnit_Remarks.Index).Value = rowWU.remarks
                    .Cells(vFrm.ColWorkUnit_PrimaryWorkUnit.Index).Value = rowWU.primary_work_unit
                End With
            Next
            vFrm.DgvWorkUnit.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dtaP.Connection.ConnectionString = Core_Module.GetConstring
            dtaP.Fill(tblP, vPrimaryTrans)
            vFrm.dgvItemProductGroup.Rows.Clear()
            vFrm.dgvItemProductGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowP In tblP
                Dim dr As DataGridViewRow = vFrm.dgvItemProductGroup.Rows(vFrm.dgvItemProductGroup.Rows.Add)
                With dr
                    .Cells(vFrm.ColProductGroup_No.Index).Value = vFrm.dgvItemProductGroup.Rows.Count
                    Select Case rowP.group_level
                        Case 0
                            .Cells(vFrm.ColProductGroup_Level.Index).Value = "Product Group"
                            .Cells(vFrm.ColProductGroup_CodeNo.Index).Value = Me.objGroup.GetGroupCode(rowP.primary_ref_party)
                            .Cells(vFrm.ColProductGroup_Desc.Index).Value = Me.objGroup.GetGroupDesc(rowP.primary_ref_party)
                        Case 1
                            .Cells(vFrm.ColProductGroup_Level.Index).Value = "Brand"
                            .Cells(vFrm.ColProductGroup_CodeNo.Index).Value = Me.objGroup.GetGroupCode(rowP.primary_ref_party)
                            .Cells(vFrm.ColProductGroup_Desc.Index).Value = Me.objGroup.GetGroupDesc(rowP.primary_ref_party)
                        Case 2
                            .Cells(vFrm.ColProductGroup_Level.Index).Value = "Sub Brand"
                            .Cells(vFrm.ColProductGroup_CodeNo.Index).Value = Me.objGroup.GetGroupCode(rowP.primary_ref_party)
                            .Cells(vFrm.ColProductGroup_Desc.Index).Value = Me.objGroup.GetGroupDesc(rowP.primary_ref_party)
                        Case 3
                            .Cells(vFrm.ColProductGroup_Level.Index).Value = "Category"
                            .Cells(vFrm.ColProductGroup_CodeNo.Index).Value = Me.objCategory.GetCategoryCode(rowP.primary_ref_party)
                            .Cells(vFrm.ColProductGroup_Desc.Index).Value = Me.objCategory.GetDesc(rowP.primary_ref_party)
                        Case 4
                            .Cells(vFrm.ColProductGroup_Level.Index).Value = "Product"
                            .Cells(vFrm.ColProductGroup_CodeNo.Index).Value = Me.objProduct.GetProductCode(rowP.primary_ref_party)
                            .Cells(vFrm.ColProductGroup_Desc.Index).Value = Me.objProduct.GetDescription(rowP.primary_ref_party)
                    End Select
                    .Cells(vFrm.ColProductGroup_Remarks.Index).Value = rowP.remarks
                    .Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value = rowP.primary_ref_party
                End With
            Next
            vFrm.dgvItemProductGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

End Class
