Option Explicit On
Option Strict Off
Public Class FrmTrans_ProductionSchedule_NewSch2_ColumnSetting
    Public UserCode As String = ""
    Public CalledFromMainGantt As Boolean = False
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim objCd As New Core_Module.ClsCekData
    Dim buffer As Object
    Dim dgvIndex As Integer
    Dim objConfig As New ClsMasterWorkUnitConfiguration
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser

    Private Sub FrmTrans_ProductionSchedule_NewSch2_ColumnSetting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ShowWO, ShowCodeNo, ShowDesc, ShowDayStart, ShowDayFinish, ShowProcess, ShowProgress, ShowRate, ShowQty, ShowUnit, ShowWH, ShowInputLocation, ShowOutputLocation, ShowActualGood, ShowActualReject, ShowActualProduce As String
        Try
            Me.dgvItemConfig.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If CalledFromMainGantt = False Then
                Me.Text = "Column Settings"
                Me.TabControl1.TabPages.Remove(TabPage2)
                ShowWO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowWO, Me.UserCode)
                ShowCodeNo = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowCodeNo, Me.UserCode)
                ShowDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowDesc, Me.UserCode)
                ShowDayStart = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowDayStart, Me.UserCode)
                ShowDayFinish = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowDayFinish, Me.UserCode)
                ShowProcess = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowProcess, Me.UserCode)
                ShowRate = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowRate, Me.UserCode)
                ShowQty = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowQty, Me.UserCode)
                ShowUnit = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowUnit, Me.UserCode)
                ShowWH = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowWH, Me.UserCode)
                ShowInputLocation = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowInputLocation, Me.UserCode)
                ShowOutputLocation = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowOutputLocation, Me.UserCode)
            Else
                Me.LoadConfig()
                ShowWO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowWO, Me.UserCode)
                ShowCodeNo = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowCodeNo, Me.UserCode)
                ShowDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowDesc, Me.UserCode)
                ShowDayStart = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowDayStart, Me.UserCode)
                ShowDayFinish = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowDayFinish, Me.UserCode)
                ShowProcess = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowProcess, Me.UserCode)
                ShowProgress = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowProgress, Me.UserCode)
                ShowRate = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowRate, Me.UserCode)
                ShowQty = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowQty, Me.UserCode)
                ShowUnit = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowUnit, Me.UserCode)
                ShowWH = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowWH, Me.UserCode)
                ShowInputLocation = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowInputLocation, Me.UserCode)
                ShowOutputLocation = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowOutputLocation, Me.UserCode)
                ShowActualGood = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowActualGoodQty, Me.UserCode)
                ShowActualReject = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowActualRejectQty, Me.UserCode)
                ShowActualProduce = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowActualProduceQty, Me.UserCode)
            End If

            If CalledFromMainGantt = False Then
                Me.ChkShowProgress.Visible = False
            End If

            If ShowWO = "1" Then
                Me.ChkWONo.Checked = True
            Else
                Me.ChkWONo.Checked = False
            End If
            If ShowCodeNo = "1" Then
                Me.ChkCodeNo.Checked = True
            Else
                Me.ChkCodeNo.Checked = False
            End If
            If ShowDesc = "1" Then
                Me.ChkDesc.Checked = True
            Else
                Me.ChkDesc.Checked = False
            End If
            If ShowDayStart = "1" Then
                Me.ChkDayStart.Checked = True
            Else
                Me.ChkDayStart.Checked = False
            End If
            If ShowDayFinish = "1" Then
                Me.ChkDayFinish.Checked = True
            Else
                Me.ChkDayFinish.Checked = False
            End If
            If ShowProcess = "1" Then
                Me.ChkShowProcess.Checked = True
            Else
                Me.ChkShowProcess.Checked = False
            End If
            If ShowProgress = "1" Then
                Me.ChkShowProgress.Checked = True
            Else
                Me.ChkShowProgress.Checked = False
            End If
            If ShowRate = "1" Then
                Me.ChkRate.Checked = True
            Else
                Me.ChkRate.Checked = False
            End If
            If ShowQty = "1" Then
                Me.ChkQty.Checked = True
            Else
                Me.ChkQty.Checked = False
            End If
            If ShowUnit = "1" Then
                Me.ChkUnit.Checked = True
            Else
                Me.ChkUnit.Checked = False
            End If
            If ShowWH = "1" Then
                Me.ChkWorkHour.Checked = True
            Else
                Me.ChkWorkHour.Checked = False
            End If
            If ShowInputLocation = "1" Then
                Me.ChkInputLocation.Checked = True
            Else
                Me.ChkInputLocation.Checked = False
            End If
            If ShowOutputLocation = "1" Then
                Me.ChkOutputLocation.Checked = True
            Else
                Me.ChkOutputLocation.Checked = False
            End If
            If ShowActualGood = "1" Then
                Me.ChkGoodQty.Checked = True
            Else
                Me.ChkGoodQty.Checked = False
            End If
            If ShowActualReject = "1" Then
                Me.ChkRejectQty.Checked = True
            Else
                Me.ChkRejectQty.Checked = False
            End If
            If ShowActualProduce = "1" Then
                Me.ChkProduceQty.Checked = True
            Else
                Me.ChkProduceQty.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                If Me.CalledFromMainGantt = False Then
                    If Me.ChkWONo.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowWO, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowWO, "0", Me.UserCode)
                    End If
                    If Me.ChkCodeNo.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowCodeNo, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowCodeNo, "0", Me.UserCode)
                    End If
                    If Me.ChkDesc.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowDesc, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowDesc, "0", Me.UserCode)
                    End If
                    If Me.ChkDayStart.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowDayStart, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowDayStart, "0", Me.UserCode)
                    End If
                    If Me.ChkDayFinish.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowDayFinish, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowDayFinish, "0", Me.UserCode)
                    End If
                    If Me.ChkShowProcess.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowProcess, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowProcess, "0", Me.UserCode)
                    End If
                    If Me.ChkRate.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowRate, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowRate, "0", Me.UserCode)
                    End If
                    If Me.ChkQty.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowQty, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowQty, "0", Me.UserCode)
                    End If
                    If Me.ChkUnit.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowUnit, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowUnit, "0", Me.UserCode)
                    End If
                    If Me.ChkWorkHour.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowWH, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowWH, "0", Me.UserCode)
                    End If

                    If Me.ChkInputLocation.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowInputLocation, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowInputLocation, "0", Me.UserCode)
                    End If
                    If Me.ChkOutputLocation.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowOutputLocation, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowOutputLocation, "0", Me.UserCode)
                    End If

                Else
                    If Me.ChkWONo.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowWO, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowWO, "0", Me.UserCode)
                    End If
                    If Me.ChkCodeNo.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowCodeNo, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowCodeNo, "0", Me.UserCode)
                    End If
                    If Me.ChkDesc.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowDesc, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowDesc, "0", Me.UserCode)
                    End If
                    If Me.ChkDayStart.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowDayStart, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowDayStart, "0", Me.UserCode)
                    End If
                    If Me.ChkDayFinish.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowDayFinish, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowDayFinish, "0", Me.UserCode)
                    End If
                    If Me.ChkShowProcess.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowProcess, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowProcess, "0", Me.UserCode)
                    End If
                    If Me.ChkShowProgress.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowProgress, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowProgress, "0", Me.UserCode)
                    End If
                    If Me.ChkRate.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowRate, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowRate, "0", Me.UserCode)
                    End If
                    If Me.ChkQty.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowQty, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowQty, "0", Me.UserCode)
                    End If
                    If Me.ChkUnit.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowUnit, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowUnit, "0", Me.UserCode)
                    End If
                    If Me.ChkWorkHour.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowWH, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowWH, "0", Me.UserCode)
                    End If

                    If Me.ChkInputLocation.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowInputLocation, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowInputLocation, "0", Me.UserCode)
                    End If
                    If Me.ChkOutputLocation.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowOutputLocation, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowOutputLocation, "0", Me.UserCode)
                    End If

                    If Me.ChkGoodQty.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowActualGoodQty, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowActualGoodQty, "0", Me.UserCode)
                    End If
                    If Me.ChkRejectQty.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowActualRejectQty, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowActualRejectQty, "0", Me.UserCode)
                    End If
                    If Me.ChkProduceQty.Checked = True Then
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowActualProduceQty, "1", Me.UserCode)
                    Else
                        Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowActualProduceQty, "0", Me.UserCode)
                    End If

                    Me.SaveConfig()
                End If
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemConfig_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemConfig.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColCode.Index Then
                e.Cancel = True
                Exit Sub
            End If
            Me.buffer = Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''ok-rosadin-r5b
    Dim objMasterLine As New Core_Module.ClsMasterMachineLine
    Private Sub dgvItemConfig_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemConfig.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColCode.Index Then
                If IsNothing(Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If String.IsNullOrEmpty(Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                        If Me.buffer.ToString.ToUpper = Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper Then
                            Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer.ToString
                            Exit Sub
                        End If
                        Dim code As String = Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value.ToString
                        If Me.objMasterLine.ApakahSudahAda(code) = True Then
                            Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                            Me.objConfig.FillPropertyByCode(code.ToUpper, Me.dgvItemConfig.Item(Me.ColPrimaryMasterLine.Index, e.RowIndex).Value, Me.dgvItemConfig.Item(Me.ColDesc.Index, e.RowIndex).Value, Me.dgvItemConfig.Item(Me.ColRemarks.Index, e.RowIndex).Value)
                        Else
                            Dim f As New Core_Module.FrmMasterMachineLine_Cari
                            f.BackColorAll = Me.BackColorAll
                            f.txtCari.Text = Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value
                            f.ShowDialog()
                            If f.Primary > 0 Then
                                Me.objConfig.FillPropertyByPrimary(f.Primary, Me.dgvItemConfig.Item(Me.ColCode.Index, e.RowIndex).Value, Me.dgvItemConfig.Item(Me.ColDesc.Index, e.RowIndex).Value, Me.dgvItemConfig.Item(Me.ColRemarks.Index, e.RowIndex).Value)
                                Me.dgvItemConfig.Item(Me.ColPrimaryMasterLine.Index, e.RowIndex).Value = f.Primary
                            Else
                                If String.IsNullOrEmpty(Me.buffer.ToString) = False Then
                                    Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer.ToString.ToUpper
                                Else
                                    Me.dgvIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                End If
                            End If
                        End If
                        If Me.dgvItemConfig.Item(Me.ColPrimaryMasterLine.Index, e.RowIndex).Value.ToString <> "-1" Then
                            If IsExistConfig(Me.dgvItemConfig.Item(Me.ColPrimaryMasterLine.Index, e.RowIndex).Value, e.RowIndex) Then
                                MsgBox("Master line " & Me.dgvItemConfig.Item(Me.ColCode.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                If String.IsNullOrEmpty(Me.buffer.ToString) = False Then
                                    Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer.ToString.ToUpper
                                    Me.objConfig.FillPropertyByCode(Me.buffer.ToString.ToUpper, Me.dgvItemConfig.Item(Me.ColPrimaryMasterLine.Index, e.RowIndex).Value, Me.dgvItemConfig.Item(Me.ColDesc.Index, e.RowIndex).Value, Me.dgvItemConfig.Item(Me.ColRemarks.Index, e.RowIndex).Value)
                                Else
                                    Me.dgvIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                End If
                            End If
                        End If
                    Else
                        If String.IsNullOrEmpty(Me.buffer.ToString) = False Then
                            Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer.ToString.ToUpper
                        Else
                            Me.dgvIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
                Else
                    If String.IsNullOrEmpty(Me.buffer.ToString) = False Then
                        Me.dgvItemConfig.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer.ToString.ToUpper
                    Else
                        Me.dgvIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''ok-rosadin-r5b
    Function IsExistConfig(ByVal vPrimaryConfig As Double, ByVal vCurrentIndex As Integer) As Boolean
        Try
            For Each row As DataGridViewRow In Me.dgvItemConfig.Rows
                If row.Index <> vCurrentIndex Then
                    If row.Cells(Me.ColPrimaryMasterLine.Index).Value = vPrimaryConfig Then
                        Return True
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''ok-rosadin-r5b
    Private Sub dgvItemConfig_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemConfig.RowsAdded
        Try
            Dim r As DataGridViewRow = Me.dgvItemConfig.Rows(Me.dgvItemConfig.Rows.Count - 1)
            r.Cells(Me.ColNo.Index).Value = Me.dgvItemConfig.Rows.Count.ToString
            r.Cells(Me.ColPrimaryMasterLine.Index).Value = -1
            r.Cells(Me.ColCode.Index).Value = String.Empty
            r.Cells(Me.ColDesc.Index).Value = String.Empty
            r.Cells(Me.ColRemarks.Index).Value = String.Empty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemConfig_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemConfig.RowsRemoved
        Try
            Me.objCd.SetNumber(Me.dgvItemConfig)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemConfig_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemConfig.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvItemConfig)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try

            ''ok-rosadin-r5b
            If (Me.dgvItemConfig.AllowUserToAddRows = True And Me.dgvItemConfig.RowCount = 1) Or (Me.dgvIndex = Me.dgvItemConfig.Rows.Count - 1) Then
            Else
                Me.dgvItemConfig.Rows.RemoveAt(Me.dgvIndex)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''ok-rosadin-r5b
    Sub LoadConfig()
        Dim d As New DseMasterWorkUnitConfigurationTableAdapters.DataTable3TableAdapter
        Dim t As New DseMasterWorkUnitConfiguration.DataTable3DataTable
        Dim r As DseMasterWorkUnitConfiguration.DataTable3Row
        Dim r1 As DataGridViewRow
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            Me.Cursor = Cursors.WaitCursor
            If d.Fill(t, Me.UserCode) > 0 Then
                Me.dgvItemConfig.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                For Each r In t
                    r1 = Me.dgvItemConfig.Rows(Me.dgvItemConfig.Rows.Add())
                    r1.Cells(Me.ColPrimaryMasterLine.Index).Value = r.primary_config
                    r1.Cells(Me.ColCode.Index).Value = r.code
                    r1.Cells(Me.ColDesc.Index).Value = r.description
                    r1.Cells(Me.ColRemarks.Index).Value = r.remarks
                    r1.Cells(Me.ColNo.Index).Value = r1.Index + 1
                Next
                Me.dgvItemConfig.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    ''ok-rosadin-r5b
    Sub SaveConfig()
        Dim d As New DseMasterWorkUnitConfigurationTableAdapters.DataTable2TableAdapter
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            Me.Cursor = Cursors.WaitCursor
            d.DeleteQuery(Me.UserCode)
            For Each r As DataGridViewRow In Me.dgvItemConfig.Rows
                If r.Cells(Me.ColPrimaryMasterLine.Index).Value.ToString <> "-1" Then
                    d.InsertQuery(Me.UserCode, r.Cells(Me.ColPrimaryMasterLine.Index).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Public Function GetListLinePerUser(ByVal vUserCode As String) As List(Of Double)
        Dim v As New List(Of Double)
        Dim d As New Core_Module.DseMasterMachineLineTableAdapters.tbltrans_filter_lineconfigTableAdapter
        Dim t As New Core_Module.DseMasterMachineLine.tbltrans_filter_lineconfigDataTable
        Dim r As Core_Module.DseMasterMachineLine.tbltrans_filter_lineconfigRow
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            If d.Fill(t, vUserCode) > 0 Then
                For Each r In t
                    v.Add(r.primary_config)
                Next
            End If
            Return v
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return v
        End Try
    End Function

    Public Function GetListMasterLinePerUser(ByVal vUserCode As String) As List(Of Double)
        Dim v As New List(Of Double)
        Dim d As New Core_Module.DseMasterMachineLineTableAdapters.tbltrans_filter_lineconfigTableAdapter
        Dim t As New Core_Module.DseMasterMachineLine.tbltrans_filter_lineconfigDataTable
        Dim r As Core_Module.DseMasterMachineLine.tbltrans_filter_lineconfigRow
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            If d.Fill(t, vUserCode) > 0 Then
                For Each r In t
                    v.Add(r.primary_config)
                Next
            End If
            Return v
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return v
        End Try
    End Function
End Class