Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_ExplainRequired
    Public vRow As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row
    Public vRefDate As Date = Today
    Public ObjCD As Core_Module.ClsCekData
    Public vDPM As Double = 0

    Dim OrderStrategy As Core_Module.clsMasterProduct.EnumOrderStrategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty
    Dim ProcurementStrategy As Core_Module.clsMasterProduct.EnumProcurementStrategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainRequired_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim FS As Decimal = 0
        Dim LD As Double = 0
        Dim QtyCompensate As Decimal = 0

        Try
            With Me.vRow
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Free stock calculation"
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "FS = OH + POB + GIT + MOB - SOB - Pending DO - P.Req.OB -  RM.Req - Spare SOB - Spare RM. Req - Spare P.Req" '- Inq
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_on_hand, .default_prec_unit) + _
                " + " + Me.ObjCD.DigitGrouping(.qty_pob, .default_prec_unit) + _
                " + " + Me.ObjCD.DigitGrouping(.qty_git, .default_prec_unit) + _
                " + " + Me.ObjCD.DigitGrouping(.qty_mob, .default_prec_unit) + _
                " - " + Me.ObjCD.DigitGrouping(.qty_sob, .default_prec_unit) + _
                " - " + Me.ObjCD.DigitGrouping(.qty_pending_do, .default_prec_unit) + _
                " - " + Me.ObjCD.DigitGrouping(.qty_prob, .default_prec_unit) + _
                " - " + Me.ObjCD.DigitGrouping(.qty_rm_req, .default_prec_unit) + _
                " - " + Me.ObjCD.DigitGrouping(.MP_Spare_Qty_SOB, .default_prec_unit) + _
                " - " + Me.ObjCD.DigitGrouping(.MP_Spare_Qty_Rm_Req, .default_prec_unit) + _
                " - " + Me.ObjCD.DigitGrouping(.MP_Spare_Qty_PrOB, .default_prec_unit)
                '"- " + Me.ObjCD.DigitGrouping(.qty_inquiry, .default_prec_unit) + _
                FS = Me.ObjCD.DigitGrouping(.qty_on_hand + .qty_pob + .qty_git + .qty_mob - .qty_sob - .qty_pending_do - .qty_inquiry - .qty_prob - .qty_rm_req - _
                .MP_Spare_Qty_PrOB - .MP_Spare_Qty_Rm_Req - .MP_Spare_Qty_SOB, .default_prec_unit)

                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_on_hand + .qty_pob + .qty_git + .qty_mob - _
                .qty_sob - .qty_pending_do - .qty_inquiry - .qty_prob - .qty_rm_req - .MP_Spare_Qty_PrOB - .MP_Spare_Qty_Rm_Req - .MP_Spare_Qty_SOB, .default_prec_unit) + " " + .Unit_Code
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + vbCrLf + "where:"
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "FS = Free Stock Qty"
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "OH = On Hand Qty"
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "POB = Purchase Order Balance Qty"
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "MOB = Manufacture Order Balance Qty"
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "GIT = Goods In Transit Qty"
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "SOB = Sales Order Balance Qty"
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Pending DO = Pending DO Qty"
                'Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Inq = Sales Inquiry Qty"
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "P.Req.OB = Purchase Request Order Balance Qty"
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "RM.Req = Raw Material Request Qty"
                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "=====================================================================" + vbCrLf

                

                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Required Qty calculation:"
                OrderStrategy = .order_strategy
                ProcurementStrategy = .procurement_strategy

                If ProcurementStrategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Order point date calculation:"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= Ref. Date + Rounddown((FS-SS)*DPM/AvgPM)) - LT"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Format(vRefDate, "dd MMM yyyy") + " + Rounddown((" + Me.ObjCD.DigitGrouping(FS, .default_prec_unit) + _
                    "-" + Me.ObjCD.DigitGrouping(.qty_safety_stock, .default_prec_unit) + ")*" + Me.vDPM.ToString + "/" + _
                    Me.ObjCD.DigitGrouping(.qty_avg_montly_demand, .default_prec_unit) + ")) - " + ObjCD.BuangNolDibelakangDecimal(.lead_time_days)
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Format(.order_point, "dd MMM yyyy")
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + vbCrLf + "where:"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Ref. Date = Reference date used as start date for date calculation"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "SS = Safety Stock Qty"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "DPM = Days Per Month, the amount of total days in 1 month used in calculation"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "AvgPM = Average usage per month for the product"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "LT = Lead time (days)"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "=====================================================================" + vbCrLf

                    'r1-iwan
                    If vRefDate.Date > .order_point.Date Then
                        If FS > 0 Then
                            LD = Me.ObjCD.RoundDown1((vRefDate.Date - .order_point.Date).TotalDays, 0)
                            Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Compensation for late order:"
                            Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Late Duration = Order Point - Reference Date = " + Format(vRefDate, "dd MMM yyyy") + " - " + Format(.order_point, "dd MMM yyyy") + _
                            " = " + Me.ObjCD.RoundDown1((vRefDate.Date - .order_point.Date).TotalDays, 0).ToString + " days"
                            Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Qty Compensate = (LD * AvgPM)/DPM -> Maximum value is Safety Stock Qty(" + Me.ObjCD.DigitGrouping(.qty_safety_stock, .default_prec_unit) + " " + .Unit_Code + ")"
                            Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= (" + Me.ObjCD.DigitGrouping(LD, 0) + " * " + _
                            Me.ObjCD.DigitGrouping(.qty_avg_montly_demand, .default_prec_unit) + ")/" + Me.vDPM.ToString
                            QtyCompensate = Me.ObjCD.DigitGrouping((LD * .qty_avg_montly_demand) / Me.vDPM, .default_prec_unit, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                            Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping((LD * .qty_avg_montly_demand) / Me.vDPM, .default_prec_unit, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                            If QtyCompensate > .qty_safety_stock Then
                                QtyCompensate = .qty_safety_stock
                                Me.TxtExplain.Text = Me.TxtExplain.Text + "-> " + Me.ObjCD.DigitGrouping(.qty_safety_stock, .default_prec_unit) + " " + .Unit_Code
                            End If
                            Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + vbCrLf + "where:"
                            Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "LD = Late Duration"
                        Else
                            Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Compensation for late order:"
                            Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Cause free stock < 0, the Qty Compensante = Qty Safety Stock = " + Me.ObjCD.DigitGrouping(.qty_safety_stock, .default_prec_unit) + " " + .Unit_Code
                            QtyCompensate = .qty_safety_stock
                        End If

                        'Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Cause qty compensate + qty required > qty optimum -> "
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "=====================================================================" + vbCrLf
                    End If

                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Procurement Strategy: Build To Stock"
                    If OrderStrategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty Or _
                    OrderStrategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Minimum_Order_Qty Or _
                    OrderStrategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty Then
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Order Strategy: Order by optimum order qty"
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Required Qty = Optimum Order Qty = " + Me.ObjCD.DigitGrouping(.qty_order_optimum, .default_prec_unit)
                    Else
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Order Strategy: Order to optimum stock qty"
                        If FS > .qty_safety_stock Then
                            Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Because FS > SS, the the following condition apply:"
                            If QtyCompensate = 0 Then
                                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Required Qty = Optimum Stock Qty - SS "
                                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_optimum, .default_prec_unit) + " - " + Me.ObjCD.DigitGrouping(.qty_safety_stock, .default_prec_unit)
                                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_required, .default_prec_unit) + " " + .Unit_Code
                            Else
                                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Required Qty = Optimum Stock Qty - SS + Qty Compensate"
                                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_optimum, .default_prec_unit) + " - " + Me.ObjCD.DigitGrouping(.qty_safety_stock, .default_prec_unit) _
                                + " + " + Me.ObjCD.DigitGrouping(QtyCompensate, .default_prec_unit)
                                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_required, .default_prec_unit) + " " + .Unit_Code
                            End If
                        Else
                            If QtyCompensate = 0 Then
                                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Because FS <= SS, the the following condition apply:"
                                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Required Qty = Optimum Stock Qty - SS + (SS - FS)"
                                If FS < 0 Then
                                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_optimum, .default_prec_unit) + " - (" + Me.ObjCD.DigitGrouping(FS, .default_prec_unit) + ")"
                                Else
                                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_optimum, .default_prec_unit) + " - " + Me.ObjCD.DigitGrouping(FS, .default_prec_unit)
                                End If
                                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_required, .default_prec_unit) + " " + .Unit_Code
                            Else
                                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Because FS <= SS, the the following condition apply:"
                                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Required Qty = Optimum Stock Qty - SS + (SS - FS) + Qty Compensate"
                                If FS < 0 Then
                                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_optimum, .default_prec_unit) + " - (" + Me.ObjCD.DigitGrouping(FS, .default_prec_unit) + ")" _
                                    + " + " + Me.ObjCD.DigitGrouping(QtyCompensate, .default_prec_unit)
                                Else
                                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_optimum, .default_prec_unit) + " - " + Me.ObjCD.DigitGrouping(FS, .default_prec_unit) _
                                    + " + " + Me.ObjCD.DigitGrouping(QtyCompensate, .default_prec_unit)
                                End If
                                Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_required, .default_prec_unit) + " " + .Unit_Code
                            End If
                        End If
                    End If
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "=====================================================================" + vbCrLf

                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Next order point date calculation:"
                    If QtyCompensate > 0 Then
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= Reference Date + LT + Rounddown((Order Qty*DPM)/AvgPM)) - LT"
                    Else
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= Order point + LT + Rounddown((Order Qty*DPM)/AvgPM)) - LT"
                    End If
                    If QtyCompensate > 0 Then
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Format(vRefDate, "dd MMM yyyy") + " + Rounddown((" + Me.ObjCD.DigitGrouping(.qty_order, .default_prec_unit) + _
                        "*" + Me.vDPM.ToString + ")/" + _
                        Me.ObjCD.DigitGrouping(.qty_avg_montly_demand, .default_prec_unit) + "))"
                    Else
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Format(.order_point, "dd MMM yyyy") + " + Rounddown((" + Me.ObjCD.DigitGrouping(.qty_order, .default_prec_unit) + _
                        "*" + Me.vDPM.ToString + ")/" + _
                        Me.ObjCD.DigitGrouping(.qty_avg_montly_demand, .default_prec_unit) + "))"
                    End If
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Format(.next_order_point, "dd MMM yyyy")
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "=====================================================================" + vbCrLf
                Else
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Procurement Strategy: Build To Order"
                    If FS <= .qty_safety_stock Then
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Because FS <= SS, the the following condition apply:"
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Required Qty = SS - FS"
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_safety_stock, .default_prec_unit) + " - " + Me.ObjCD.DigitGrouping(FS, .default_prec_unit)
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "= " + Me.ObjCD.DigitGrouping(.qty_required, .default_prec_unit) + " " + .Unit_Code
                    Else
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Because FS > SS, the the following condition apply:"
                        Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Required Qty = 0" + " " + .Unit_Code
                    End If
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "=====================================================================" + vbCrLf
                End If
            End With

            Me.Button1.Focus()
            If Me.TxtExplain.Text.Length > 0 Then
                Me.TxtExplain.Select(0, 0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class