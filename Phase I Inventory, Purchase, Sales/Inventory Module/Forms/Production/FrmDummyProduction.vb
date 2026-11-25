Public Class FrmDummyProduction
    Dim ObjHakAkses As New Core_Module.ClsHakAkses

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ij As New FrmTrans_ManufacturingOrder

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ij As New FrmTrans_ProductionSchedule

        Try
            ij.BackColorAll = Core_Module.GetSMESSBackColorAll
            ij.UserCode = "USER"
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    'Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
    '    Dim ij As New FrmTransaksiStockBundling2

    '    Try
    '        ij.ObjStockBundling.TipeBundling = Inventory_Module.ClsTransaksiStockBundling.EnumTipeBundling.Purchase
    '        ij.ShowDialog()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ij.Dispose()
    '    End Try
    'End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim ij As New FrmLaporanProd_PickList

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim ij As New FrmTrans_ManufacturingOrder_SetFinishJO

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim ij As New Core_Module.FrmMasterMachineData

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim ij As New Core_Module.frmMasterProduct

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim ij As New Core_Module.FrmMasterWorkUnit

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim ij As New Core_Module.FrmMasterMachineGroup

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim ij As New Core_Module.FrmMasterWorkShift

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ij As New Core_Module.FrmMasterWorkUnit_Workshift

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim ij As New Core_Module.FrmMasterKompetensi

        Try
            ij.Text = "Custom Value"
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim ij As New Core_Module.FrmMasterReject

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim ij As New Core_Module.FrmMasterDownTime

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ij As New FrmLaporanProd_ProductionOutput

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Dim ij As New FrmLaporanProd_RejectReport

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Dim ij As New FrmLaporanProd_RawMaterialUsage
        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Dim ij As New FrmLaporanProd_Downtime

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Dim ij As New FrmLaporanProd_Productivity

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ij As New FrmLaporanProd_MOStatus

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Dim ij As New FrmTrans_ManufacturingOrder_Authorize

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        Dim ij As New FrmTrans_SPL

        Try
            ij.UserCode = "ARGEN"
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        Dim ij As New FrmTrans_SPL_Authorize

        Try
            ij.UserCode = "USER"
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Dim ij As New FrmTrans_SPL_Authorize_Access

        Try
            ij.UserCode = "ARGEN"
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        Dim ij As New FrmTransMO_WorkShift_Allocate

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        Dim ij As New Core_Module.FrmMasterWorkUnitGroup
        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)

    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
        Dim ij As New FrmLaporanProd_RejectReport
        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button10_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim ij As New FrmLaporanProd_PlanningReport
        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub FrmDummyProduction_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'Me.UpdateNumbering()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateHakAkses()
        Dim DTA As New Core_Module.DSUserTableAdapters.tblmasterhakaksesTableAdapter
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim ObjSalaryForm As New Core_Module.ClsMasterSalaryFormula

        Try
            Me.ObjHakAkses.SetConnectionString(DTA)

            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ProductionSchedule_SendTakeControl", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ProductionSchedule_SendTakeControl", "Hak akses untuk ambil/tarik kontrol jadwal produksi.", "FrmTrans_ProductionSchedule", 0)
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjSetting = Nothing
            ObjSalaryForm = Nothing
        End Try
    End Sub

    'Private Sub UpdateNumbering()
    '    Dim objNum As New Core_Module.ClsMasterTransactionNumbering

    '    Try
    '        If objNum.UpdateNumbering("tbltrans_spl", "Production Overtime Order", _
    '                   Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SPL) = False Then
    '            Exit Sub
    '        End If
    '        If objNum.UpdateNumbering("tbltrans_mo_item_wo", "Work Order", _
    '        Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.WO) = False Then
    '            Exit Sub
    '        End If
    '        If objNum.UpdateNumbering("tbltrans_mo_item", "Job Order", _
    '        Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.JO) = False Then
    '            Exit Sub
    '        End If
    '        If objNum.UpdateNumbering("tbltrans_mo", "Manufacturing Order", _
    '      Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO) = False Then
    '            Exit Sub
    '        End If
    '        If objNum.UpdateNumbering("tbltrans_sch_downtime", "Schedule Downtime", _
    '        Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime) = False Then
    '            Exit Sub
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        objNum = Nothing
    '    End Try
    'End Sub

    Private Sub Button19_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim ij As New FrmSetWU
        'Try
        '    ij.ShowDialog()
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    ij.Dispose()
        'End Try
    End Sub

    Private Sub Button19_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim ij As New FrmTrans_SchDownTime

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button20_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Dim ij As New Core_Module.FrmMasterSetupTime

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        Dim ij As New FrmSettingProduksi

        Try
            ij.BackColorAll = Core_Module.GetSMESSBackColorAll
            ij.UserCode = "USER"

            ij.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        Dim ij As New FrmTrans_work_unitionSchedule_EditOT

        Try
            ij.BackColorAll = Core_Module.GetSMESSBackColorAll
            ij.UserCode = "USER"

            ij.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub
End Class