Option Explicit On
Option Strict Off

Public Class FrmMasterAsset
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim Baru As Boolean = True

    Dim AkunAsset As ClsMasterAsset.EnumAkunAsset
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjCekdata As New ClsCekData
    Dim ObjAsset As New ClsMasterAsset
    Dim ObjAssetGroup As New ClsMasterAssetGroup
    Dim ObjAkun As New ClsMasterAkun1
    Dim ObjUnit As New ClsMasterUnit
    Dim DefPrec_Qty As Double = 2
    Dim DefPrec_Amount As Double = Me.ObjCurr.GetDecimal_DefaultCurrency()

    Private Function BlokSimbol(ByVal obj As TextBox, ByVal e As Windows.Forms.KeyPressEventArgs) As Boolean
        Try
            If e.KeyChar = "'" Or _
            e.KeyChar = "%" Or _
            e.KeyChar = "*" Or _
            e.KeyChar = "#" Then
                obj.ReadOnly = True
            Else
                obj.ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CumaAngka(ByVal kode As String, Optional ByVal spasi As Boolean = False) As Boolean
        Select Case kode
            Case 40 To 57
                Return False
            Case 8
                Return False
            Case 32
                If spasi Then
                    Return False
                Else
                    Return True
                End If
            Case Else
                Return True
        End Select
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.ObjAsset
                .AssetCode = Me.txtAssetCode.Text
                If Me.TxtStatusEnum.Text = ClsMasterAsset.EnumStatusAsset.On_Progress Then
                    .StatusAsset = ClsMasterAsset.EnumStatusAsset.On_Progress
                ElseIf Me.TxtStatusEnum.Text = ClsMasterAsset.EnumStatusAsset.Active Then
                    .StatusAsset = ClsMasterAsset.EnumStatusAsset.Active
                ElseIf Me.TxtStatusEnum.Text = ClsMasterAsset.EnumStatusAsset.Dispose Then
                    .StatusAsset = ClsMasterAsset.EnumStatusAsset.Dispose
                End If
                .Desc = Me.txtDescription.Text
                .PrimaryAssetGroupCode = Me.txtPrimaryGroup.Text
                .Qty = Me.TxtQty.Text
                .PrimaryUnit = Me.TxtPrimaryUnit.Text
                .Remarks = Me.txtRemarks.Text
                If Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.Define_Value Then
                    .AcquicitionType = ClsMasterAsset.EnumAcquisitionType.Define_Value
                ElseIf Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.BuyOrBuild Then
                    .AcquicitionType = ClsMasterAsset.EnumAcquisitionType.BuyOrBuild
                End If
                If Me.ChkAcquisitionFinish.Checked = True Then
                    .AcquicitionFinish = ClsMasterAsset.EnumAcquisitionFinish._true
                ElseIf Me.ChkAcquisitionFinish.Checked = False Then
                    .AcquicitionFinish = ClsMasterAsset.EnumAcquisitionFinish._false
                End If
                If Me.ChkDepThisAsset.Checked = True Then
                    .DepreciationThisAsset = ClsMasterAsset.EnumDepreciateThisAsset._true
                Else
                    .DepreciationThisAsset = ClsMasterAsset.EnumDepreciateThisAsset._false
                End If
                If Me.cboDepreciationMethod.Text = "Straight Line" Then
                    .DepreciationMethod = ClsMasterAsset.EnumDepreciationMethod.Straight_Line
                ElseIf Me.cboDepreciationMethod.Text = "Sum Year Digits" Then
                    .DepreciationMethod = ClsMasterAsset.EnumDepreciationMethod.Sum_year_Digits
                ElseIf Me.cboDepreciationMethod.Text = "Double Declining Balance" Then
                    .DepreciationMethod = ClsMasterAsset.EnumDepreciationMethod.Double_Declining_Balance
                ElseIf Me.cboDepreciationMethod.Text = "No Depreciation Method" Then
                    .DepreciationMethod = ClsMasterAsset.EnumDepreciationMethod.No_Depreciation_Method
                End If
                .DepreciationLife = Me.TxtDepLife.Text
                .AcquiredValue = Me.txtAcquiredValue.Text
                .SalvageValue = Me.TxtSalvageValue.Text
                .BookValue = Me.TxtBookValue.Text
                If Me.cboDepreciateFrequency.SelectedIndex = ClsMasterAsset.EnumFrequency.Monthly Then
                    .Frequency = ClsMasterAsset.EnumFrequency.Monthly
                ElseIf Me.cboDepreciateFrequency.SelectedIndex = ClsMasterAsset.EnumFrequency.Yearly Then
                    .Frequency = ClsMasterAsset.EnumFrequency.Yearly
                End If
                .StartDepreciateOn = Me.dtpStart.Value
                .DepreciateOn = Me.DtpDepOn.Value
                .PrimaryAkunOnProcess = Me.TxtPrimaryAkun_OnProcess.Text
                .PrimaryAkunAsset = Me.TxtPrimaryAkun_Asset.Text
                .PrimaryAkunDepreciationAsset = Me.TxtPrimaryAkun_Dep_Expense.Text
                .PrimaryAkunAccumDepreciation = Me.TxtPrimaryAkun_Accum_Dep.Text
                .PrimaryAkunDisposalIncome = Me.TxtPrimaryAkun_Disposal_Income.Text
                .PrimaryAkunGainLossDisposal = Me.TxtPrimaryAkun_Gain_Loss.Text
                .PrimaryJournalSetfinish = Me.TxtPrimaryJurnal_SetFinish.Text
                .TanggalSetfinish = Me.DtpSetFinish.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjAsset
                Me.txtPrimaryAsset.Text = .Primary
                Me.txtAssetCode.Text = .AssetCode
                If .StatusAsset = ClsMasterAsset.EnumStatusAsset.On_Progress Then
                    Me.TxtStatus.Text = "On Progress"
                ElseIf .StatusAsset = ClsMasterAsset.EnumStatusAsset.Active Then
                    Me.TxtStatus.Text = "Active"
                ElseIf .StatusAsset = ClsMasterAsset.EnumStatusAsset.Dispose Then
                    Me.TxtStatus.Text = "Dispose"
                End If
                Me.TxtStatusEnum.Text = .StatusAsset
                Me.txtDescription.Text = .Desc
                Me.txtPrimaryGroup.Text = .PrimaryAssetGroupCode
                Me.txtGroupCode.Text = Me.ObjAssetGroup.GetGroupCode(.PrimaryAssetGroupCode)
                Me.TxtQty.Text = Me.ObjCekdata.DigitGrouping(.Qty, Me.DefPrec_Qty)
                Me.TxtPrimaryUnit.Text = .PrimaryUnit
                Me.TxtUnit.Text = Me.ObjUnit.GetUnitCode(.PrimaryUnit)
                Me.txtRemarks.Text = .Remarks
                If .AcquicitionType = ClsMasterAsset.EnumAcquisitionType.Define_Value Then
                    Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.Define_Value
                ElseIf .AcquicitionType = ClsMasterAsset.EnumAcquisitionType.BuyOrBuild Then
                    Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.BuyOrBuild
                End If
                If .AcquicitionFinish = ClsMasterAsset.EnumAcquisitionFinish._true Then
                    Me.ChkAcquisitionFinish.Checked = True
                Else
                    Me.ChkAcquisitionFinish.Checked = False
                End If
                If .DepreciationThisAsset = ClsMasterAsset.EnumDepreciateThisAsset._true Then
                    Me.ChkDepThisAsset.Checked = True
                Else
                    Me.ChkDepThisAsset.Checked = False
                End If
                If .DepreciationMethod = ClsMasterAsset.EnumDepreciationMethod.Straight_Line Then
                    Me.cboDepreciationMethod.SelectedIndex = ClsMasterAsset.EnumDepreciationMethod.Straight_Line
                ElseIf .DepreciationMethod = ClsMasterAsset.EnumDepreciationMethod.Sum_year_Digits Then
                    Me.cboDepreciationMethod.SelectedIndex = ClsMasterAsset.EnumDepreciationMethod.Sum_year_Digits
                ElseIf .DepreciationMethod = ClsMasterAsset.EnumDepreciationMethod.Double_Declining_Balance Then
                    Me.cboDepreciationMethod.SelectedIndex = ClsMasterAsset.EnumDepreciationMethod.Double_Declining_Balance
                ElseIf .DepreciationMethod = ClsMasterAsset.EnumDepreciationMethod.No_Depreciation_Method Then
                    If Me.cboDepreciationMethod.Items.Contains("No Depreciation Method") = False Then
                        Me.cboDepreciationMethod.Items.Add("No Depreciation Method")
                    End If
                    Me.cboDepreciationMethod.SelectedIndex = ClsMasterAsset.EnumDepreciationMethod.No_Depreciation_Method
                End If
                Me.TxtDepLife.Text = Me.ObjCekdata.DigitGrouping(.DepreciationLife, 0)
                Me.txtAcquiredValue.Text = Me.ObjCekdata.DigitGrouping(.AcquiredValue, Me.DefPrec_Amount)
                Me.TxtSalvageValue.Text = Me.ObjCekdata.DigitGrouping(.SalvageValue, Me.DefPrec_Amount)
                Me.TxtBookValue.Text = Me.ObjCekdata.DigitGrouping(.BookValue, Me.DefPrec_Amount)

                If .Frequency = ClsMasterAsset.EnumFrequency.Monthly Then
                    Me.cboDepreciateFrequency.SelectedIndex = ClsMasterAsset.EnumFrequency.Monthly
                    Me.LblLife1.Text = "Month"
                    Me.dtpStart.CustomFormat = "MMM yyyy"
                    Me.dtpStart.Format = DateTimePickerFormat.Custom
                    Me.DtpDepOn.Visible = False
                    Me.Label14.Visible = False
                ElseIf .Frequency = ClsMasterAsset.EnumFrequency.Yearly Then
                    Me.cboDepreciateFrequency.SelectedIndex = ClsMasterAsset.EnumFrequency.Yearly
                    Me.LblLife1.Text = "Year"
                    Me.dtpStart.CustomFormat = "yyyy"
                    Me.dtpStart.Format = DateTimePickerFormat.Custom
                    Me.DtpDepOn.Visible = True
                    Me.Label14.Visible = True
                End If
                Me.dtpStart.Value = .StartDepreciateOn
                Me.DtpDepOn.Value = .DepreciateOn
                Me.TxtPrimaryAkun_OnProcess.Text = .PrimaryAkunOnProcess
                Me.TxtAkunOnProcess.Text = Me.ObjAkun.GetAccountCode(.PrimaryAkunOnProcess)
                Me.TxtAccDesc_OnProgress.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(.PrimaryAkunOnProcess)
                Me.TxtPrimaryAkun_Asset.Text = .PrimaryAkunAsset
                Me.TxtAkunAsset.Text = Me.ObjAkun.GetAccountCode(.PrimaryAkunAsset)
                Me.TxtAccDesc_asset.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(.PrimaryAkunAsset)
                Me.TxtPrimaryAkun_Dep_Expense.Text = .PrimaryAkunDepreciationAsset
                Me.TxtAkunDepExpense.Text = Me.ObjAkun.GetAccountCode(.PrimaryAkunDepreciationAsset)
                Me.TxtAccDesc_dept_Expense.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(.PrimaryAkunDepreciationAsset)
                Me.TxtPrimaryAkun_Accum_Dep.Text = .PrimaryAkunAccumDepreciation
                Me.TxtAkunAccumDep.Text = Me.ObjAkun.GetAccountCode(.PrimaryAkunAccumDepreciation)
                Me.TxtAccDesc_accum_dept.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(.PrimaryAkunAccumDepreciation)
                Me.TxtPrimaryAkun_Disposal_Income.Text = .PrimaryAkunDisposalIncome
                Me.TxtAkunDisposalIncome.Text = Me.ObjAkun.GetAccountCode(.PrimaryAkunDisposalIncome)
                Me.TxtAccDesc_DisposalIncome.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(.PrimaryAkunDisposalIncome)
                Me.TxtPrimaryAkun_Gain_Loss.Text = .PrimaryAkunGainLossDisposal
                Me.TxtAkunGainLoss.Text = Me.ObjAkun.GetAccountCode(.PrimaryAkunGainLossDisposal)
                Me.TxtAccDesc_GainLoss.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(.PrimaryAkunGainLossDisposal)
                Me.TxtPrimaryJurnal_SetFinish.Text = .PrimaryJournalSetfinish
                Me.DtpSetFinish.Value = .TanggalSetfinish
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim ObjJurnal As New ClsTransaksiJournal1
        Dim ObjStockInOut As New clsTransaksiStockInOut1
        Try
            'tidak boleh kosong
            KoleksiNamaData.Add("Code")
            KoleksiTextBox.Add(Me.txtAssetCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.txtDescription)
            KoleksiNamaData.Add("Group Code")
            KoleksiTextBox.Add(Me.txtGroupCode)
            KoleksiNamaData.Add("Qty")
            KoleksiTextBox.Add(Me.TxtQty)
            KoleksiNamaData.Add("Unit")
            KoleksiTextBox.Add(Me.TxtUnit)
            If Me.TxtAkunOnProcess.Enabled = True Then
                KoleksiNamaData.Add("On Proses Account")
                KoleksiTextBox.Add(Me.TxtAkunOnProcess)
            End If
            KoleksiNamaData.Add("Asset Account")
            KoleksiTextBox.Add(Me.TxtAkunAsset)
            If Me.TxtAkunDepExpense.Enabled = True Then
                KoleksiNamaData.Add("Depreciation Ecpense Account")
                KoleksiTextBox.Add(Me.TxtAkunDepExpense)
                KoleksiNamaData.Add("Accumulate Depreciation Account")
                KoleksiTextBox.Add(Me.TxtAkunAccumDep)
            End If
            KoleksiNamaData.Add("Disposal Income Account")
            KoleksiTextBox.Add(Me.TxtAkunGainLoss)
            KoleksiNamaData.Add("Gain / Loss Disposal Account")
            KoleksiTextBox.Add(Me.TxtAkunGainLoss)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'tidak boleh double
            If Me.baru = True Then
                If Me.ObjAsset.ApakahSudahAda(Me.txtAssetCode.Text) = True Then
                    MsgBox("Duplication in asset code, please change your code", MsgBoxStyle.Information, "Information")
                    Me.txtAssetCode.Focus()
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.ObjAsset.GetCode(Me.txtPrimaryAsset.Text)
                If CodeLama <> Me.txtAssetCode.Text Then
                    If Me.ObjAsset.ApakahSudahAda(Me.txtAssetCode.Text) = True Then
                        MsgBox("Duplication in asset code, please change your code", MsgBoxStyle.Information, "Information")
                        Me.txtAssetCode.Focus()
                        Return False
                    End If
                End If
                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses(edit)
                Dim OldTimeStamp As Date = Me.ObjAsset.GetTimestamp(Me.txtPrimaryAsset.Text)
                If Me.ObjAsset.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjAsset.CariDetail(Me.ObjAsset.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'berkaitan dengan angka
            If Me.TxtQty.Text < 1 Then
                MsgBox("Qty must have value", MsgBoxStyle.Information, "Information")
                Me.TxtQty.Focus()
                Return False
            End If
            If Me.cboAcquisitionType.SelectedIndex <> ClsMasterAsset.EnumAcquisitionType.BuyOrBuild Then
                If Me.txtAcquiredValue.Text < 1 Then
                    MsgBox("Acquired Value must have value", MsgBoxStyle.Information, "Information")
                    Me.txtAcquiredValue.Focus()
                    Return False
                End If
            End If
            If Me.Baru = False Then
                If Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.BuyOrBuild And _
                    Me.ChkAcquisitionFinish.Checked = True Then
                    If Me.TxtSalvageValue.Text < 1 Then
                        MsgBox("Salvage Value must have value", MsgBoxStyle.Information, "Information")
                        Me.TxtSalvageValue.Focus()
                        Return False
                    End If
                End If
            End If

            'tanggal finis tidak boleh lebih dari tanggal depresiasi
            If Me.DtpSetFinish.Enabled = True Then
                If Me.cboDepreciateFrequency.SelectedIndex = ClsMasterAsset.EnumFrequency.Monthly Then
                    If Me.DtpSetFinish.Value.Month > Me.dtpStart.Value.Month And _
                    Me.DtpSetFinish.Value.Year > Me.dtpStart.Value.Year Then
                        MsgBox("Acquisition finish date don't more than depreciate date")
                        Me.DtpSetFinish.Value = Now
                        Me.DtpSetFinish.Focus()
                    End If
                ElseIf Me.cboDepreciateFrequency.SelectedIndex = ClsMasterAsset.EnumFrequency.Yearly Then
                    If Me.DtpSetFinish.Value.Month > Me.DtpDepOn.Value.Month And _
                    Me.DtpSetFinish.Value.Year > Me.dtpStart.Value.Year Then
                        MsgBox("Acquisition finish date don't more than depreciate date")
                        Me.DtpSetFinish.Value = Now
                        Me.DtpSetFinish.Focus()
                    End If
                End If
            End If

            'tanggal minimum
            Dim MaxDate As Date
            MaxDate = Me.ObjAsset.GetMAXDateDepreciation()
            If Me.Baru = True Then
                If Me.DtpSetFinish.Enabled = True Then
                    If Me.cboDepreciateFrequency.SelectedIndex = ClsMasterAsset.EnumFrequency.Monthly Then
                        If Me.dtpStart.Value.Month < MaxDate.Month And _
                       Me.dtpStart.Value.Year < MaxDate.Year Then
                            MsgBox("depreciate date don't more than depreciate date")
                            Me.DtpSetFinish.Value = Now
                            Me.DtpSetFinish.Focus()
                        End If
                    ElseIf Me.cboDepreciateFrequency.SelectedIndex = ClsMasterAsset.EnumFrequency.Yearly Then
                        If Me.DtpDepOn.Value.Month < MaxDate.Month And _
                        Me.dtpStart.Value.Year < MaxDate.Year Then
                            MsgBox("depreciate date don't more than depreciate date")
                            Me.DtpSetFinish.Value = Now
                            Me.DtpSetFinish.Focus()
                        End If
                    End If
                End If
            End If


            'pastikan asset mempunyai minimal 1 transaksi akusisi material usage atau other
            'keadaan buy/build dan set finis di ceklist
            If Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.BuyOrBuild And _
            Me.ChkAcquisitionFinish.Checked = True Then
                Dim JmlMaterialUsage As Double = ObjStockInOut.GetJmlByAsset(Me.txtPrimaryAsset.Text)
                Dim JmlOTher As Double = ObjJurnal.GetJmlByTipeNPrimaryAsset(Me.txtPrimaryAsset.Text, ClsTransaksiJournal1.EnumTipeJurnal2.AssetAcquistion_Other)

                If JmlMaterialUsage = 0 And JmlOTher = 0 Then
                    MsgBox("Please uncheck acquistion finish, because this asset doesn't have an any acquistion transaction", MsgBoxStyle.Information, _
                    "Information")
                    Return False
                End If
            End If

            'jika cek finis
            If Me.txtAcquiredValue.Enabled = True Then
                If CDec(Me.txtAcquiredValue.Text) < CDec(Me.TxtSalvageValue.Text) Then
                    MsgBox("Acquired value is shouldn't be smaller than salvage value .")
                    Me.TxtSalvageValue.Focus()
                    Return False
                End If
            End If


            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            ObjJurnal = Nothing
            ObjStockInOut = Nothing
        End Try
    End Function

    Private Sub FrmMasterAsset_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtAssetCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjAsset.SimpanBaru(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjAsset.SimpanPerubahan(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAsset_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.NonAktifkanDetail()
            Me.TabControl1.TabPages.RemoveAt(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Dim ObjJournal As New ClsTransaksiJournal1
        Try
            Me.ObjCekdata.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.GroupBox1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)

            'cek transaksi yang bersangkutan
            If Me.CekAdaTransaksi(Me.txtPrimaryAsset.Text) = True Then
                Me.cboAcquisitionType.Enabled = False
            End If

            'cek depresiasi
            If Me.ChkDepThisAsset.Checked = False Then
                Me.BtnFindAccDepAsset.Enabled = False
                Me.btnFindAccAccum.Enabled = False
                Me.TxtAkunDepExpense.Enabled = False
                Me.TxtAkunAccumDep.Enabled = False
                Me.TxtAccDesc_dept_Expense.Enabled = False
                Me.TxtAccDesc_accum_dept.Enabled = False
                Me.TxtPrimaryAkun_Dep_Expense.Text = -1
                Me.TxtPrimaryAkun_Accum_Dep.Text = -1
                Me.TxtAccDesc_dept_Expense.Text = ""
                Me.TxtAccDesc_accum_dept.Text = ""
                Me.TxtAkunDepExpense.Text = ""
                Me.TxtAkunAccumDep.Text = ""
                If Me.cboDepreciationMethod.Items.Contains("No Depreciation Method") = False Then
                    Me.cboDepreciationMethod.Items.Add("No Depreciation Method")
                End If
            ElseIf Me.ChkDepThisAsset.Checked = True Then
                Me.cboDepreciationMethod.Items.Remove("No Depreciation Method")
            End If
            If ObjJournal.GetJmlByTipeNPrimaryAsset(Me.txtPrimaryAsset.Text, ClsTransaksiJournal1.EnumTipeJurnal2.AssetDepreciation) > 0 Then
                If Me.ChkDepThisAsset.Checked = True Then
                    Me.ChkDepThisAsset.Enabled = False
                ElseIf Me.ChkDepThisAsset.Checked = False Then


                End If
            End If

            'cek status
            If String.IsNullOrEmpty(Me.TxtStatusEnum.Text) = False Then
                If Me.TxtStatusEnum.Text = ClsMasterAsset.EnumStatusAsset.Dispose Then
                    Me.TxtStatus.Text = "Disposed"
                End If
            Else
                If Me.ChkAcquisitionFinish.Checked = True Then
                    Me.TxtStatus.Text = "Active"
                    Me.TxtStatusEnum.Text = ClsMasterAsset.EnumStatusAsset.Active
                Else : Me.ChkAcquisitionFinish.Checked = False
                    Me.TxtStatus.Text = "On Progress"
                    Me.TxtStatusEnum.Text = ClsMasterAsset.EnumStatusAsset.Active
                End If
            End If




            'jika edit , define value dan mempunyai transaksi depresiasi
            'maka depresiasi life tidak bisa diedit 
            If Me.Baru = False And Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.Define_Value And _
            ObjJournal.GetJmlByTipeNPrimaryAsset(Me.txtPrimaryAsset.Text, ClsTransaksiJournal1.EnumTipeJurnal2.AssetDepreciation) > 0 Then
                Me.TxtDepLife.Enabled = False
            End If

            'jika edit , buy / build dan mempunyai transaksi depresiasi
            'maka salvage value tidak bisa diedit 
            If Me.Baru = False And Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.BuyOrBuild And _
            ObjJournal.GetJmlByTipeNPrimaryAsset(Me.txtPrimaryAsset.Text, ClsTransaksiJournal1.EnumTipeJurnal2.AssetDepreciation) > 0 Then
                Me.TxtSalvageValue.Enabled = False
            Else
                Me.TxtSalvageValue.Enabled = True
            End If

            Me.ChkAcquisitionFinish.Enabled = True
            Me.ChkDepThisAsset.Enabled = True

            If ObjJournal.GetTipeJournal2(Me.txtPrimaryAsset.Text) = ClsTransaksiJournal1.EnumTipeJurnal2.AssetDepreciation Then
                If Me.ChkDepThisAsset.Checked = True Then
                    Me.ChkDepThisAsset.Enabled = False
                End If
                If Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.Define_Value Then
                    Me.TxtDepLife.Enabled = False
                    Me.TxtAkunOnProcess.Enabled = False
                    Me.btnFindAcc_OnProcess.Enabled = False
                ElseIf Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.BuyOrBuild Then
                    Me.TxtAkunOnProcess.Enabled = True
                    Me.btnFindAcc_OnProcess.Enabled = True
                End If
                Me.cboDepreciationMethod.Enabled = False
                Me.BtnFindAccDepAsset.Enabled = False
                Me.btnFindAccAccum.Enabled = False
                Me.TxtAkunDepExpense.Enabled = False
                Me.TxtAkunAccumDep.Enabled = False
            End If

            If ObjJournal.GetTipeJournal2(Me.txtPrimaryAsset.Text) = ClsTransaksiJournal1.EnumTipeJurnal2.AssetAcquistion_Other Then
                Me.btnFindAcc_OnProcess.Enabled = False
                Me.TxtAkunOnProcess.Enabled = False
            End If

            If Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.Define_Value Then
                Me.TxtAkunOnProcess.Enabled = False
                Me.btnFindAcc_OnProcess.Enabled = False
            ElseIf Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.BuyOrBuild Then
                Me.TxtAkunOnProcess.Enabled = True
                Me.btnFindAcc_OnProcess.Enabled = True
                Me.txtAcquiredValue.Enabled = False
                Me.txtAcquiredValue.Text = Me.ObjCekdata.DigitGrouping(0, Me.DefPrec_Amount)
            End If

            If Me.ChkDepThisAsset.Checked = False Then
                Me.cboDepreciationMethod.SelectedIndex = ClsMasterAsset.EnumDepreciationMethod.No_Depreciation_Method
                Me.cboDepreciationMethod.Enabled = False
            End If


            If Me.Baru = False Then
                If ObjJournal.GetJmlByTipeNPrimaryAsset(Me.txtPrimaryAsset.Text, ClsTransaksiJournal1.EnumTipeJurnal2.AssetDepreciation) > 0 Then
                    Me.cboDepreciateFrequency.Enabled = False
                    Me.dtpStart.Enabled = False
                    Me.DtpDepOn.Enabled = False
                End If
            End If

            If Me.Baru = False And Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.BuyOrBuild And _
            Me.ChkAcquisitionFinish.Checked = False Then
                Me.TxtSalvageValue.Enabled = False
            Else
                Me.TxtSalvageValue.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjJournal = Nothing
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCekdata.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.GroupBox1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.cboDepreciationMethod.Items.Remove("No Depreciation Method")
            Me.ChkAcquisitionFinish.Enabled = False
            Me.ChkDepThisAsset.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCekdata.KosongkanControl(Me.Controls, True, True, True)
            Me.ObjCekdata.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.ObjCekdata.KosongkanControl(Me.GroupBox1.Controls, True, True, True)
            Me.ObjCekdata.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.txtPrimaryAsset.Text = -1
            Me.ObjCekdata.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.ObjCekdata.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, True)
            Me.ObjCekdata.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, True)
            Me.TxtQty.Text = Me.ObjCekdata.DigitGrouping(0, Me.DefPrec_Qty)
            Me.TxtDepLife.Text = Me.ObjCekdata.DigitGrouping(1)
            Me.TxtSalvageValue.Text = Me.ObjCekdata.DigitGrouping(0, Me.DefPrec_Amount)
            Me.TxtBookValue.Text = Me.ObjCekdata.DigitGrouping(0, Me.DefPrec_Amount)
            Me.txtAcquiredValue.Text = Me.ObjCekdata.DigitGrouping(0, Me.DefPrec_Amount)
            Me.TxtPrimaryAkun_Accum_Dep.Text = -1
            Me.TxtPrimaryAkun_Asset.Text = -1
            Me.TxtPrimaryAkun_Dep_Expense.Text = -1
            Me.TxtPrimaryAkun_Disposal_Income.Text = -1
            Me.TxtPrimaryAkun_Gain_Loss.Text = -1
            Me.TxtPrimaryAkun_OnProcess.Text = -1
            Me.txtPrimaryGroup.Text = -1
            Me.TxtPrimaryJurnal_SetFinish.Text = -1
            Me.TxtPrimaryUnit.Text = -1
            Me.DtpDepOn.Value = Now
            Me.DtpSetFinish.Value = Now
            Me.dtpStart.Value = Now
            Me.ChkAcquisitionFinish.Checked = True
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.Baru = True
            Me.txtAssetCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAssetCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAssetCode.GotFocus
        Try
            Me.txtAssetCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAssetCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAssetCode.KeyPress
        Me.BlokSimbol(Me.txtAssetCode, e)
    End Sub

    Private Sub txtAssetCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAssetCode.LostFocus
        Try
            Me.txtAssetCode.BackColor = Me.txtPrimaryAsset.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDescription_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.GotFocus
        Try
            Me.txtDescription.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDescription_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDescription.KeyPress
        Me.BlokSimbol(Me.txtDescription, e)
    End Sub

    Private Sub txtDescription_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.LostFocus
        Try
            Me.txtDescription.BackColor = Me.txtPrimaryAsset.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindGroupCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindGroupCode.Click
        Dim FrmCariGroupCode As New FrmMasterAssetGroup_Cari

        Try
            FrmCariGroupCode.BackColorAll = Me.BackColorAll
            If Me.txtPrimaryAsset.Text = -1 And Me.txtGroupCode.Text = "" Then
                FrmCariGroupCode.ShowDialog()
            Else
                FrmCariGroupCode.txtCari.Text = Me.txtGroupCode.Text
                FrmCariGroupCode.ShowDialog()
            End If
            If CDbl(FrmCariGroupCode.PrimaryGroup) > 0 Then
                Me.txtPrimaryGroup.Text = FrmCariGroupCode.PrimaryGroup
                Me.txtGroupCode.Text = Me.ObjAssetGroup.GetGroupCode(FrmCariGroupCode.PrimaryGroup)
                Me.TxtQty.Focus()
            Else
                Me.txtPrimaryGroup.Text = -1
                Me.txtGroupCode.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariGroupCode.Dispose()
        End Try
    End Sub

    Private Sub txtGroupCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGroupCode.GotFocus
        Try
            Me.txtGroupCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGroupCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjAssetGroup.ApakahSudahAda(Me.txtGroupCode.Text) = True Then
                    Me.txtPrimaryGroup.Text = Me.ObjAssetGroup.GetPrimary(Me.txtGroupCode.Text)
                    Me.TxtQty.Focus()
                Else
                    Me.BtnFindGroupCode.PerformClick()
                    Me.TxtQty.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGroupCode.Leave
        Try
            If Me.txtGroupCode.Text <> "" Then
                If Me.ObjAssetGroup.ApakahSudahAda(Me.txtGroupCode.Text) = True Then
                    Me.txtPrimaryGroup.Text = Me.ObjAssetGroup.GetPrimary(Me.txtGroupCode.Text)
                Else
                    Me.BtnFindGroupCode.PerformClick()
                End If
            Else
                Me.txtPrimaryGroup.Text = -1
                Me.txtGroupCode.Text = ""
            End If
            Me.txtGroupCode.BackColor = Me.txtPrimaryAsset.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty.GotFocus
        Me.TxtQty.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtQty.KeyPress
        e.Handled = CumaAngka(Asc(e.KeyChar))
    End Sub

    Private Sub TxtQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty.LostFocus
        Me.TxtQty.BackColor = Me.txtPrimaryAsset.BackColor
        If Me.TxtQty.Text = "" Then
            Me.TxtQty.Text = Me.ObjCekdata.DigitGrouping(0, Me.DefPrec_Qty)
        Else
            Me.TxtQty.Text = Me.ObjCekdata.DigitGrouping(Me.TxtQty.Text, Me.DefPrec_Qty)
        End If
    End Sub

    Private Sub btnFindUnit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindUnit.Click
        Dim FrmCariUnit As New FrmCariMasterUnit
        Try
            FrmCariUnit.BackColorAll = Me.BackColorAll
            If Me.TxtPrimaryUnit.Text = -1 And Me.TxtUnit.Text = "" Then
                FrmCariUnit.ShowDialog()
            Else
                FrmCariUnit.txtCari.Text = Me.TxtUnit.Text
                FrmCariUnit.ShowDialog()
            End If
            If CDbl(FrmCariUnit.PrimaryUnit) > 0 Then
                Me.TxtPrimaryUnit.Text = FrmCariUnit.PrimaryUnit
                Me.TxtUnit.Text = Me.ObjUnit.GetUnitCode(FrmCariUnit.PrimaryUnit)
                Me.txtRemarks.Focus()
            Else
                Me.TxtPrimaryUnit.Text = -1
                Me.TxtUnit.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariUnit.Dispose()
        End Try
    End Sub

    Private Sub TxtUnit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnit.GotFocus
        Me.TxtUnit.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnit.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjUnit.ApakahSudahAda(Me.TxtUnit.Text) = True Then
                    Me.TxtPrimaryUnit.Text = Me.ObjUnit.GetPrimary(Me.TxtUnit.Text)
                    Me.txtRemarks.Focus()
                Else
                    Me.btnFindUnit.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnit.Leave
        Try
            If Me.TxtUnit.Text <> "" Then
                If Me.ObjUnit.ApakahSudahAda(Me.TxtUnit.Text) = True Then
                    Me.TxtPrimaryUnit.Text = Me.ObjUnit.GetPrimary(Me.TxtUnit.Text)
                Else
                    Me.btnFindUnit.PerformClick()
                End If
            Else
                Me.TxtPrimaryUnit.Text = -1
                Me.TxtUnit.Text = ""
            End If
            Me.TxtUnit.BackColor = Me.txtPrimaryAsset.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.GotFocus
        Me.txtRemarks.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtRemarks_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRemarks.KeyPress
        Me.BlokSimbol(Me.txtRemarks, e)
    End Sub

    Private Sub txtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.LostFocus
        Me.txtRemarks.BackColor = Me.TxtPrimaryUnit.BackColor
    End Sub

    Private Sub cboAcquisitionType_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAcquisitionType.SelectedValueChanged
        Dim ObjJournal As New ClsTransaksiJournal1
        Try
            If Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.Define_Value Then
                'akun on proses harus kosong
                Me.btnFindAcc_OnProcess.Enabled = False
                Me.TxtAkunOnProcess.Enabled = False
                Me.TxtAkunOnProcess.Text = ""
                Me.TxtAccDesc_OnProgress.Text = ""
                Me.TxtPrimaryAkun_OnProcess.Text = -1

                'cek finis
                Me.ChkAcquisitionFinish.Checked = True
                Me.ChkAcquisitionFinish.Visible = False
                Me.DtpSetFinish.Visible = False
                Me.DtpSetFinish.Value = "1-1-1950"

                If Me.Baru = False Then
                    If ObjJournal.GetJmlByTipeNPrimaryAsset(Me.txtPrimaryAsset.Text, ClsTransaksiJournal1.EnumTipeJurnal2.AssetDepreciation) > 0 Then
                        Me.txtAcquiredValue.Enabled = False
                    Else
                        Me.txtAcquiredValue.Enabled = True
                    End If
                End If
                Me.txtAcquiredValue.Enabled = True
                Me.TxtSalvageValue.Enabled = True
            ElseIf Me.cboAcquisitionType.SelectedIndex = ClsMasterAsset.EnumAcquisitionType.BuyOrBuild Then
                Me.btnFindAcc_OnProcess.Enabled = True
                Me.TxtAkunOnProcess.Enabled = True
                If Me.Baru = False Then
                    Me.ChkAcquisitionFinish.Enabled = True
                    Me.DtpSetFinish.Enabled = True
                Else
                    Me.ChkAcquisitionFinish.Enabled = False
                    Me.DtpSetFinish.Enabled = False
                End If
                Me.ChkAcquisitionFinish.Checked = False
                Me.DtpSetFinish.Visible = False
                Me.ChkAcquisitionFinish.Visible = True
                Me.DtpSetFinish.Value = Me.dtpStart.Value
                Me.txtAcquiredValue.Text = Me.ObjCekdata.DigitGrouping(0, Me.DefPrec_Amount)
                Me.txtAcquiredValue.Enabled = False
                If Me.TxtStatusEnum.Text = ClsMasterAsset.EnumStatusAsset.On_Progress Then
                    Me.TxtSalvageValue.Text = Me.ObjCekdata.DigitGrouping(0, Me.DefPrec_Amount)
                    Me.TxtSalvageValue.Enabled = False
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjJournal = Nothing
        End Try
    End Sub

    Private Sub ChkAcquisitionFinish_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkAcquisitionFinish.CheckedChanged
        Try
            If Me.ChkAcquisitionFinish.Checked = True Then
                Me.TxtStatus.Text = "Active"
                Me.TxtStatusEnum.Text = ClsMasterAsset.EnumStatusAsset.Active
                Me.DtpSetFinish.Visible = True
                Me.DtpSetFinish.Value = Now
                Me.TxtSalvageValue.Enabled = True
            ElseIf Me.ChkAcquisitionFinish.Checked = False Then
                Me.TxtStatus.Text = "On Progress"
                Me.TxtStatusEnum.Text = ClsMasterAsset.EnumStatusAsset.On_Progress
                Me.DtpSetFinish.Visible = False
                Me.DtpSetFinish.Value = "1-1-1950"
                Me.TxtSalvageValue.Enabled = False
                Me.TxtSalvageValue.Text = Me.ObjCekdata.DigitGrouping(0, Me.DefPrec_Amount)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkDepThisAsset_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkDepThisAsset.CheckedChanged
        Try
            If Me.ChkDepThisAsset.Checked = False Then
                'akun dept asset dan accum dept d kosongkan
                Me.BtnFindAccDepAsset.Enabled = False
                Me.btnFindAccAccum.Enabled = False
                Me.TxtAkunDepExpense.Enabled = False
                Me.TxtAkunAccumDep.Enabled = False
                Me.TxtAccDesc_dept_Expense.Enabled = False
                Me.TxtAccDesc_accum_dept.Enabled = False
                Me.TxtPrimaryAkun_Dep_Expense.Text = -1
                Me.TxtPrimaryAkun_Accum_Dep.Text = -1
                Me.TxtAccDesc_dept_Expense.Text = ""
                Me.TxtAccDesc_accum_dept.Text = ""
                Me.TxtAkunDepExpense.Text = ""
                Me.TxtAkunAccumDep.Text = ""
                'tambah item cbo method
                Me.cboDepreciationMethod.Items.Add("No Depreciation Method")
                Me.cboDepreciationMethod.SelectedIndex = ClsMasterAsset.EnumDepreciationMethod.No_Depreciation_Method
                Me.cboDepreciationMethod.Enabled = False
            Else
                Me.BtnFindAccDepAsset.Enabled = True
                Me.btnFindAccAccum.Enabled = True
                Me.TxtAkunDepExpense.Enabled = True
                Me.TxtAkunAccumDep.Enabled = True
                Me.TxtAccDesc_dept_Expense.Enabled = True
                Me.TxtAccDesc_accum_dept.Enabled = True
                Me.cboDepreciationMethod.Items.Remove("No Depreciation Method")
                Me.cboDepreciationMethod.SelectedIndex = ClsMasterAsset.EnumDepreciationMethod.Straight_Line
                Me.cboDepreciationMethod.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDepLife_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDepLife.GotFocus
        Me.TxtDepLife.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtDepLife_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtDepLife.KeyPress
        e.Handled = CumaAngka(Asc(e.KeyChar))
    End Sub

    Private Sub TxtDepLife_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDepLife.Leave
        Dim ObjJournal As New ClsTransaksiJournal1
        Try
            If Me.TxtDepLife.Text = "" Then
                Me.TxtDepLife.Text = 0
                MsgBox("depreciation life should not be less than 1", MsgBoxStyle.Information, "Information")
                Me.TxtDepLife.Focus()
            Else
                If Me.Baru = True Then
                    If CDbl(Me.TxtDepLife.Text) < 1 Then
                        MsgBox("depreciation life should not be less than 1", MsgBoxStyle.Information, "Information")
                        Me.TxtDepLife.Focus()
                    End If
                Else
                    If ObjJournal.GetJmlDepAsset(Me.txtPrimaryAsset.Text) > 0 Then
                        If CDbl(Me.TxtDepLife.Text) < ObjJournal.GetJmlDepAsset(Me.txtPrimaryAsset.Text) Then
                            MsgBox("depreciation life should not be less than " + _
                            CInt(ObjJournal.GetJmlDepAsset(Me.txtPrimaryAsset.Text)) _
                            , MsgBoxStyle.Information, "Information")
                            Me.TxtDepLife.Focus()
                        End If
                    Else
                        If CDbl(Me.TxtDepLife.Text) < 1 Then
                            MsgBox("depreciation life should not be less than 1", MsgBoxStyle.Information, "Information")
                            Me.TxtDepLife.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjJournal = Nothing
        End Try
    End Sub

    Private Sub TxtDepLife_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDepLife.LostFocus
        Me.TxtDepLife.BackColor = Me.txtPrimaryAsset.BackColor
    End Sub

    Private Sub txtAcquiredValue_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAcquiredValue.Enter
        Me.ObjCekdata.Number_Enter(Me.txtAcquiredValue, sender, e)
    End Sub

    Private Sub txtAcquiredValue_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAcquiredValue.GotFocus
        Me.txtAcquiredValue.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtAcquiredValue_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAcquiredValue.KeyDown
        Me.ObjCekdata.Number_KeyDown(Me.txtAcquiredValue, sender, e)
    End Sub

    Private Sub txtAcquiredValue_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAcquiredValue.KeyUp
        Me.ObjCekdata.Number_KeyUp(Me.txtAcquiredValue, sender, e, False)
    End Sub

    Private Sub txtAcquiredValue_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAcquiredValue.Leave
        Try
            If Me.txtAcquiredValue.Text = "" Then
                Me.txtAcquiredValue.Text = Me.ObjCekdata.DigitGrouping(0, Me.DefPrec_Amount)
            Else
                Me.txtAcquiredValue.Text = Me.ObjCekdata.DigitGrouping(Me.txtAcquiredValue.Text, Me.DefPrec_Amount)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAcquiredValue_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAcquiredValue.LostFocus
        Me.txtAcquiredValue.BackColor = Me.TxtPrimaryUnit.BackColor
    End Sub

    Private Sub TxtSalvageValue_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalvageValue.GotFocus
        Me.TxtSalvageValue.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtSalvageValue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtSalvageValue.KeyPress
        e.Handled = CumaAngka(Asc(e.KeyChar))
    End Sub

    Private Sub TxtSalvageValue_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalvageValue.Leave
        Try
            If Me.TxtSalvageValue.Text = "" Then
                Me.TxtSalvageValue.Text = Me.ObjCekdata.DigitGrouping(0, Me.DefPrec_Amount)
            Else
                Me.TxtSalvageValue.Text = Me.ObjCekdata.DigitGrouping(Me.TxtSalvageValue.Text, Me.DefPrec_Amount)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalvageValue_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalvageValue.LostFocus
        Me.TxtSalvageValue.BackColor = Me.TxtPrimaryUnit.BackColor
    End Sub

    Private Sub cboDepreciateFrequency_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDepreciateFrequency.SelectedValueChanged
        Try
            If Me.cboDepreciateFrequency.SelectedIndex = ClsMasterAsset.EnumFrequency.Monthly Then
                Me.LblLife1.Text = "Month"
                If Me.Baru = True Then
                    Me.dtpStart.Value = CStr(Now.Month) + "-1-" + CStr(Now.Year)
                    Me.DtpDepOn.Value = "1-1-1950"
                End If
                Me.dtpStart.CustomFormat = "MMM yyyy"
                Me.dtpStart.Format = DateTimePickerFormat.Custom
                Me.DtpDepOn.Visible = False
                Me.Label14.Visible = False
            ElseIf Me.cboDepreciateFrequency.SelectedIndex = ClsMasterAsset.EnumFrequency.Yearly Then
                Me.LblLife1.Text = "Year"
                If Me.Baru = True Then
                    Me.dtpStart.Value = "1-" + "1-" + CStr(Now.Year)
                    Me.DtpDepOn.Value = CStr(Now.Month) + "-1" + "-1950"
                End If
                Me.dtpStart.CustomFormat = "yyyy"
                Me.dtpStart.Format = DateTimePickerFormat.Custom
                Me.DtpDepOn.Visible = True
                Me.Label14.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindAcc_OnProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAcc_OnProcess.Click
        Dim FrmCariAkun As New FrmMasterAkun_Cari3
        Try
            FrmCariAkun.BackColorAll = Me.BackColorAll
            FrmCariAkun.txtCari.Text = Me.TxtAkunOnProcess.Text
            FrmCariAkun.AkunAsset = ClsMasterAsset.EnumAkunAsset.OnProcess
            FrmCariAkun.ShowDialog()
            If FrmCariAkun.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAkun_OnProcess.Text = FrmCariAkun.mPrimaryAkun
                Me.TxtAkunOnProcess.Text = Me.ObjAkun.GetAccountCode(FrmCariAkun.mPrimaryAkun)
                Me.TxtAccDesc_OnProgress.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(FrmCariAkun.mPrimaryAkun)
                Me.TxtAkunAsset.Focus()
            Else
                Me.TxtPrimaryAkun_OnProcess.Text = -1
                Me.TxtAkunOnProcess.Text = ""
                Me.TxtAccDesc_OnProgress.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunOnProcess_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunOnProcess.GotFocus
        Me.TxtAkunOnProcess.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtAkunOnProcess_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAkunOnProcess.KeyDown
        Try
            If e.KeyValue = 13 Then

                If Me.ObjAkun.ApakahSudahAda(Me.TxtAkunOnProcess.Text) = True Then
                    Me.TxtPrimaryAkun_OnProcess.Text = Me.ObjAkun.GetPrimary(Me.TxtAkunOnProcess.Text)
                    Me.TxtAccDesc_OnProgress.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkun_OnProcess.Text)
                    Me.TxtAkunAsset.Focus()
                Else
                    Me.btnFindAcc_OnProcess.PerformClick()
                    Me.TxtAkunAsset.Focus()
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunOnProcess_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunOnProcess.Leave
        Try
            If Me.TxtAkunOnProcess.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda(Me.TxtAkunOnProcess.Text) = True Then
                    Me.TxtPrimaryAkun_OnProcess.Text = Me.ObjAkun.GetPrimary(Me.TxtAkunOnProcess.Text)
                    Me.TxtAccDesc_OnProgress.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkun_OnProcess.Text)
                    If Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_OnProcess.Text) <> 4 Then
                        MsgBox("Hanya boleh akun Fixed Asset", MsgBoxStyle.Information, "Information")
                        Me.TxtPrimaryAkun_OnProcess.Text = -1
                        Me.TxtAkunOnProcess.Text = ""
                        Me.TxtAccDesc_OnProgress.Text = ""
                        Me.TxtAkunOnProcess.Focus()
                    End If
                Else
                    Me.btnFindAcc_OnProcess.PerformClick()
                End If
            Else
                Me.TxtPrimaryAkun_OnProcess.Text = -1
                Me.TxtAkunOnProcess.Text = ""
                Me.TxtAccDesc_OnProgress.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunOnProcess_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunOnProcess.LostFocus
        Me.TxtAkunOnProcess.BackColor = Me.TxtPrimaryUnit.BackColor
    End Sub

    Private Sub BtnFindAccAsset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindAccAsset.Click
        Dim FrmCariAkun As New FrmMasterAkun_Cari3
        Try
            FrmCariAkun.BackColorAll = Me.BackColorAll
            FrmCariAkun.txtCari.Text = Me.TxtAkunAsset.Text
            FrmCariAkun.AkunAsset = ClsMasterAsset.EnumAkunAsset.Asset
            FrmCariAkun.ShowDialog()
            If FrmCariAkun.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAkun_Asset.Text = FrmCariAkun.mPrimaryAkun
                Me.TxtAkunAsset.Text = Me.ObjAkun.GetAccountCode(FrmCariAkun.mPrimaryAkun)
                Me.TxtAccDesc_asset.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(FrmCariAkun.mPrimaryAkun)
                If Me.TxtAkunDepExpense.Enabled = True Then
                    Me.TxtAkunDepExpense.Focus()
                Else
                    Me.TxtAkunDisposalIncome.Focus()
                End If
            Else
                Me.TxtPrimaryAkun_Asset.Text = -1
                Me.TxtAkunAsset.Text = ""
                Me.TxtAccDesc_asset.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAsset_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAsset.GotFocus
        Me.TxtAkunAsset.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtAkunAsset_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAkunAsset.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda(Me.TxtAkunAsset.Text) = True Then
                    Me.TxtPrimaryAkun_Asset.Text = Me.ObjAkun.GetPrimary(Me.TxtAkunAsset.Text)
                    Me.TxtAccDesc_asset.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkun_Asset.Text)
                    If Me.TxtAkunDepExpense.Enabled = True Then
                        Me.TxtAkunDepExpense.Focus()
                    Else
                        Me.TxtAkunDisposalIncome.Focus()
                    End If
                Else
                    Me.BtnFindAccAsset.PerformClick()
                    If Me.TxtAkunDepExpense.Enabled = True Then
                        Me.TxtAkunDepExpense.Focus()
                    Else
                        Me.TxtAkunDisposalIncome.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAsset_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAsset.Leave
        Try
            If Me.TxtAkunAsset.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda(Me.TxtAkunAsset.Text) = True Then
                    Me.TxtPrimaryAkun_Asset.Text = Me.ObjAkun.GetPrimary(Me.TxtAkunAsset.Text)
                    Me.TxtAccDesc_asset.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkun_Asset.Text)
                    If Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_Asset.Text) <> 4 Then
                        MsgBox("Hanya boleh akun Fixed Asset", MsgBoxStyle.Information, "Information")
                        Me.TxtPrimaryAkun_Asset.Text = -1
                        Me.TxtAkunAsset.Text = ""
                        Me.TxtAccDesc_asset.Text = ""
                        Me.TxtAkunAsset.Focus()
                    End If
                Else
                    Me.BtnFindAccAsset.PerformClick()
                End If
            Else
                Me.TxtPrimaryAkun_Asset.Text = -1
                Me.TxtAkunAsset.Text = ""
                Me.TxtAccDesc_asset.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAsset_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAsset.LostFocus
        Me.TxtAkunAsset.BackColor = Me.TxtPrimaryUnit.BackColor
    End Sub

    Private Sub BtnFindAccDepAsset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindAccDepAsset.Click
        Dim FrmCariAkun As New FrmMasterAkun_Cari3
        Try
            FrmCariAkun.BackColorAll = Me.BackColorAll
            FrmCariAkun.txtCari.Text = Me.TxtAkunDepExpense.Text
            FrmCariAkun.AkunAsset = ClsMasterAsset.EnumAkunAsset.DeprciationExpense
            FrmCariAkun.ShowDialog()
            If FrmCariAkun.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAkun_Dep_Expense.Text = FrmCariAkun.mPrimaryAkun
                Me.TxtAkunDepExpense.Text = Me.ObjAkun.GetAccountCode(FrmCariAkun.mPrimaryAkun)
                Me.TxtAccDesc_dept_Expense.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(FrmCariAkun.mPrimaryAkun)
                If Me.TxtPrimaryAkun_Accum_Dep.Text <> "" And Me.TxtPrimaryAkun_Accum_Dep.Text <> "-1" Then
                    If Me.ObjAkun.GetPrimary(Me.TxtAkunDepExpense.Text) = Me.TxtPrimaryAkun_Accum_Dep.Text Then
                        MsgBox("Account depreciation can't same with account accum depreciation", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtPrimaryAkun_Dep_Expense.Text = -1
                        Me.TxtAkunDepExpense.Text = ""
                        Me.TxtAccDesc_dept_Expense.Text = ""
                    End If
                End If
                Me.TxtAkunAccumDep.Focus()
            Else
                Me.TxtPrimaryAkun_Dep_Expense.Text = -1
                Me.TxtAkunDepExpense.Text = ""
                Me.TxtAccDesc_dept_Expense.Text = ""
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunDepExpense_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunDepExpense.GotFocus
        Me.TxtAkunDepExpense.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtAkunDepExpense_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAkunDepExpense.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda(Me.TxtAkunDepExpense.Text) = True Then
                    Me.TxtPrimaryAkun_Dep_Expense.Text = Me.ObjAkun.GetPrimary(Me.TxtAkunDepExpense.Text)
                    Me.TxtAccDesc_dept_Expense.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkun_Dep_Expense.Text)
                    If Me.TxtPrimaryAkun_Accum_Dep.Text <> "" And Me.TxtPrimaryAkun_Accum_Dep.Text <> "-1" Then
                        If Me.ObjAkun.GetPrimary(Me.TxtAkunDepExpense.Text) = Me.TxtPrimaryAkun_Accum_Dep.Text Then
                            MsgBox("Account depreciation can't same with account accum depreciation", _
                            MsgBoxStyle.Information, "Information")
                            Me.TxtPrimaryAkun_Dep_Expense.Text = -1
                            Me.TxtAkunDepExpense.Text = ""
                            Me.TxtAccDesc_dept_Expense.Text = ""
                        End If
                    End If
                    Me.TxtAkunAccumDep.Focus()
                Else
                    Me.BtnFindAccDepAsset.PerformClick()
                    Me.TxtAkunAccumDep.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunDepExpense_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunDepExpense.Leave
        Try
            If Me.TxtAkunDepExpense.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda(Me.TxtAkunDepExpense.Text) = True Then
                    Me.TxtPrimaryAkun_Dep_Expense.Text = Me.ObjAkun.GetPrimary(Me.TxtAkunDepExpense.Text)
                    Me.TxtAccDesc_dept_Expense.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkun_Dep_Expense.Text)
                    'cek tipe akun
                    If Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_Dep_Expense.Text) <> 13 And _
                    Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_Dep_Expense.Text) <> 15 Then
                        MsgBox("Hanya boleh akun  Expense dan Other Expense", MsgBoxStyle.Information, "Information")
                        If Me.TxtPrimaryAkun_Accum_Dep.Text <> "" And Me.TxtPrimaryAkun_Accum_Dep.Text <> "-1" Then
                            If Me.ObjAkun.GetPrimary(Me.TxtAkunDepExpense.Text) = Me.TxtPrimaryAkun_Accum_Dep.Text Then
                                MsgBox("Account depreciation can't same with account accum depreciation", _
                                MsgBoxStyle.Information, "Information")
                                Me.TxtPrimaryAkun_Dep_Expense.Text = -1
                                Me.TxtAkunDepExpense.Text = ""
                                Me.TxtAccDesc_dept_Expense.Text = ""
                            End If
                        End If
                        Me.TxtPrimaryAkun_Dep_Expense.Text = -1
                        Me.TxtAkunDepExpense.Text = ""
                        Me.TxtAccDesc_dept_Expense.Text = ""
                        Me.TxtAkunDepExpense.Focus()
                    End If
                Else
                    Me.BtnFindAccDepAsset.PerformClick()
                End If
            Else
                Me.TxtPrimaryAkun_Dep_Expense.Text = -1
                Me.TxtAkunDepExpense.Text = ""
                Me.TxtAccDesc_dept_Expense.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunDepExpense_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunDepExpense.LostFocus
        Me.TxtAkunDepExpense.BackColor = Me.TxtPrimaryUnit.BackColor
    End Sub

    Private Sub TxtAkunAccumDep_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAccumDep.GotFocus
        Me.TxtAkunAccumDep.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtAkunAccumDep_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAkunAccumDep.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda(Me.TxtAkunAccumDep.Text) = True Then
                    Me.TxtPrimaryAkun_Accum_Dep.Text = Me.ObjAkun.GetPrimary(Me.TxtAkunAccumDep.Text)
                    Me.TxtAccDesc_accum_dept.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkun_Accum_Dep.Text)
                    If Me.TxtPrimaryAkun_Dep_Expense.Text <> "" And Me.TxtPrimaryAkun_Dep_Expense.Text <> -1 Then
                        If Me.ObjAkun.GetPrimary(Me.TxtAkunAccumDep.Text) = Me.TxtPrimaryAkun_Dep_Expense.Text Then
                            MsgBox("Account accum depreciation can't same with account depreciation", _
                            MsgBoxStyle.Information, "Information")
                            Me.TxtPrimaryAkun_Accum_Dep.Text = -1
                            Me.TxtAkunAccumDep.Text = ""
                            Me.TxtAccDesc_accum_dept.Text = ""
                        End If
                    End If
                    Me.TxtAkunDisposalIncome.Focus()
                Else
                    Me.btnFindAccAccum.PerformClick()
                    Me.TxtAkunDisposalIncome.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAccumDep_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAccumDep.Leave
        Try
            If Me.TxtAkunAccumDep.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda(Me.TxtAkunAccumDep.Text) = True Then
                    Me.TxtPrimaryAkun_Accum_Dep.Text = Me.ObjAkun.GetPrimary(Me.TxtAkunAccumDep.Text)
                    Me.TxtAccDesc_accum_dept.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkun_Accum_Dep.Text)
                    'cek tipe akun
                    If Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_Accum_Dep.Text) <> 5 Then
                        MsgBox("Hanya boleh akun Accumulation depreciation", MsgBoxStyle.Information, "Information")
                        If Me.TxtPrimaryAkun_Dep_Expense.Text <> "" And Me.TxtPrimaryAkun_Dep_Expense.Text <> -1 Then
                            If Me.ObjAkun.GetPrimary(Me.TxtAkunAccumDep.Text) = Me.TxtPrimaryAkun_Dep_Expense.Text Then
                                MsgBox("Account accum depreciation can't same with account depreciation", _
                                MsgBoxStyle.Information, "Information")
                                Me.TxtPrimaryAkun_Accum_Dep.Text = -1
                                Me.TxtAkunAccumDep.Text = ""
                                Me.TxtAccDesc_accum_dept.Text = ""
                            End If
                        End If
                        Me.TxtPrimaryAkun_Accum_Dep.Text = -1
                        Me.TxtAkunAccumDep.Text = ""
                        Me.TxtAccDesc_accum_dept.Text = ""
                        Me.TxtAkunAccumDep.Focus()
                    End If
                Else
                    Me.btnFindAccAccum.PerformClick()
                End If
            Else
                Me.TxtPrimaryAkun_Accum_Dep.Text = -1
                Me.TxtAkunAccumDep.Text = ""
                Me.TxtAccDesc_accum_dept.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAccumDep_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAccumDep.LostFocus
        Me.TxtAkunAccumDep.BackColor = Me.TxtPrimaryUnit.BackColor
    End Sub

    Private Sub btnFindAccAccum_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindAccAccum.Click
        Dim FrmCariAkun As New FrmMasterAkun_Cari3
        Try
            FrmCariAkun.BackColorAll = Me.BackColorAll
            FrmCariAkun.txtCari.Text = Me.TxtAkunAccumDep.Text
            FrmCariAkun.AkunAsset = ClsMasterAsset.EnumAkunAsset.AccumDepreciation
            FrmCariAkun.ShowDialog()
            If FrmCariAkun.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAkun_Accum_Dep.Text = FrmCariAkun.mPrimaryAkun
                Me.TxtAkunAccumDep.Text = Me.ObjAkun.GetAccountCode(FrmCariAkun.mPrimaryAkun)
                Me.TxtAccDesc_accum_dept.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(FrmCariAkun.mPrimaryAkun)
                If Me.TxtPrimaryAkun_Dep_Expense.Text <> "" And Me.TxtPrimaryAkun_Dep_Expense.Text <> -1 Then
                    If Me.ObjAkun.GetPrimary(Me.TxtAkunAccumDep.Text) = Me.TxtPrimaryAkun_Dep_Expense.Text Then
                        MsgBox("Account accum depreciation can't same with account depreciation", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtPrimaryAkun_Accum_Dep.Text = -1
                        Me.TxtAkunAccumDep.Text = ""
                        Me.TxtAccDesc_accum_dept.Text = ""
                    End If
                End If
                Me.TxtAkunDisposalIncome.Focus()
            Else
                Me.TxtPrimaryAkun_Accum_Dep.Text = -1
                Me.TxtAkunAccumDep.Text = ""
                Me.TxtAccDesc_accum_dept.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunDisposalIncome_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunDisposalIncome.GotFocus
        Me.TxtAkunDisposalIncome.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtAkunDisposalIncome_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAkunDisposalIncome.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda(Me.TxtAkunDisposalIncome.Text) = True Then
                    Me.TxtPrimaryAkun_Disposal_Income.Text = Me.ObjAkun.GetPrimary(Me.TxtAkunDisposalIncome.Text)
                    Me.TxtAccDesc_DisposalIncome.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkun_Disposal_Income.Text)
                    Me.TxtAkunGainLoss.Focus()
                Else
                    Me.BtnFindAccDisposal.PerformClick()
                    Me.TxtAkunGainLoss.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunDisposalIncome_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunDisposalIncome.Leave
        Try
            If Me.TxtAkunDisposalIncome.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda(Me.TxtAkunDisposalIncome.Text) = True Then
                    Me.TxtPrimaryAkun_Disposal_Income.Text = Me.ObjAkun.GetPrimary(Me.TxtAkunDisposalIncome.Text)
                    Me.TxtAccDesc_DisposalIncome.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkun_Disposal_Income.Text)
                    'cek tipe akun
                    If Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_Disposal_Income.Text) <> 13 And _
                    Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_Disposal_Income.Text) <> 15 And _
                    Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_Disposal_Income.Text) <> 11 And _
                    Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_Disposal_Income.Text) <> 14 Then
                        MsgBox("Hanya boleh akun Income, Other Income, Expense, Other Expense", MsgBoxStyle.Information, "Information")
                        Me.TxtPrimaryAkun_Disposal_Income.Text = -1
                        Me.TxtAkunDisposalIncome.Text = ""
                        Me.TxtAccDesc_DisposalIncome.Text = ""
                        Me.TxtAkunDisposalIncome.Focus()
                    End If
                Else
                    Me.BtnFindAccDisposal.PerformClick()
                End If
            Else
                Me.TxtPrimaryAkun_Disposal_Income.Text = -1
                Me.TxtAkunDisposalIncome.Text = ""
                Me.TxtAccDesc_DisposalIncome.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunDisposalIncome_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunDisposalIncome.LostFocus
        Me.TxtAkunDisposalIncome.BackColor = Me.TxtPrimaryUnit.BackColor
    End Sub

    Private Sub BtnFindAccDisposal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFindAccDisposal.Click
        Dim FrmCariAkun As New FrmMasterAkun_Cari3
        Try
            FrmCariAkun.BackColorAll = Me.BackColorAll
            FrmCariAkun.txtCari.Text = Me.TxtAkunDisposalIncome.Text
            FrmCariAkun.AkunAsset = ClsMasterAsset.EnumAkunAsset.DisposalIncome
            FrmCariAkun.ShowDialog()
            If FrmCariAkun.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAkun_Disposal_Income.Text = FrmCariAkun.mPrimaryAkun
                Me.TxtAkunDisposalIncome.Text = Me.ObjAkun.GetAccountCode(FrmCariAkun.mPrimaryAkun)
                Me.TxtAccDesc_DisposalIncome.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(FrmCariAkun.mPrimaryAkun)
                Me.TxtAkunGainLoss.Focus()
            Else
                Me.TxtPrimaryAkun_Disposal_Income.Text = -1
                Me.TxtAkunDisposalIncome.Text = ""
                Me.TxtAccDesc_DisposalIncome.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunGainLoss_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunGainLoss.GotFocus
        Me.TxtAkunGainLoss.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtAkunGainLoss_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAkunGainLoss.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda(Me.TxtAkunGainLoss.Text) = True Then
                    Me.TxtPrimaryAkun_Gain_Loss.Text = Me.ObjAkun.GetPrimary(Me.TxtAkunGainLoss.Text)
                    Me.TxtAccDesc_GainLoss.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkun_Gain_Loss.Text)
                    Me.btnOK.Focus()
                Else
                    Me.BtnFindAccGainLoss.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunGainLoss_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunGainLoss.Leave
        Try
            If Me.TxtAkunGainLoss.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda(Me.TxtAkunGainLoss.Text) = True Then
                    Me.TxtPrimaryAkun_Gain_Loss.Text = Me.ObjAkun.GetPrimary(Me.TxtAkunGainLoss.Text)
                    Me.TxtAccDesc_GainLoss.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkun_Gain_Loss.Text)
                    'cek tipe akun
                    If Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_Gain_Loss.Text) <> 13 And _
                    Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_Gain_Loss.Text) <> 15 And _
                    Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_Gain_Loss.Text) <> 11 And _
                    Me.ObjAkun.GetTipeAkun2(Me.TxtPrimaryAkun_Gain_Loss.Text) <> 14 Then
                        MsgBox("Hanya boleh akun Income, Other Income, Expense, Other Expense", MsgBoxStyle.Information, "Information")
                        Me.TxtPrimaryAkun_Gain_Loss.Text = -1
                        Me.TxtAkunGainLoss.Text = ""
                        Me.TxtAccDesc_GainLoss.Text = ""
                        Me.TxtAkunGainLoss.Focus()
                    End If
                    Me.btnOK.Focus()
                Else
                    Me.BtnFindAccGainLoss.PerformClick()
                End If
            Else
                Me.TxtPrimaryAkun_Gain_Loss.Text = -1
                Me.TxtAkunGainLoss.Text = ""
                Me.TxtAccDesc_GainLoss.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunGainLoss_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunGainLoss.LostFocus
        Me.TxtAkunGainLoss.BackColor = Me.TxtPrimaryAkun_Disposal_Income.BackColor
    End Sub

    Private Sub BtnFindAccGainLoss_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFindAccGainLoss.Click
        Dim FrmCariAkun As New FrmMasterAkun_Cari3
        Try
            FrmCariAkun.BackColorAll = Me.BackColorAll
            FrmCariAkun.txtCari.Text = Me.TxtAkunGainLoss.Text
            FrmCariAkun.AkunAsset = ClsMasterAsset.EnumAkunAsset.GainLossDisposal
            FrmCariAkun.ShowDialog()
            If FrmCariAkun.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAkun_Gain_Loss.Text = FrmCariAkun.mPrimaryAkun
                Me.TxtAkunGainLoss.Text = Me.ObjAkun.GetAccountCode(FrmCariAkun.mPrimaryAkun)
                Me.TxtAccDesc_GainLoss.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(FrmCariAkun.mPrimaryAkun)
                Me.btnOK.Focus()
            Else
                Me.TxtPrimaryAkun_Gain_Loss.Text = -1
                Me.TxtAkunGainLoss.Text = ""
                Me.TxtAccDesc_GainLoss.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'cek data kosong
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjAsset.SimpanBaru(Me.UserCode) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjAsset.SimpanPerubahan(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtAssetCode.Text <> "" And Me.txtAssetCode.Enabled = True Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterAssetCari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryAsset = -1 Then
                If Me.ObjAsset.CariDetail(hj.PrimaryAsset) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                    Me.NonAktifkanDetail()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If Me.txtAssetCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.txtAssetCode.Text <> "" Then
            Me.Baru = False
            Me.AktifkanDetail()
            Me.txtAssetCode.Focus()
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtAssetCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtAssetCode.Text <> "" Then
                hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ObjAsset.Hapus(Me.txtPrimaryAsset.Text, Me.UserCode) = True Then
                        Me.KosongkanField()
                        Me.btnNew.Focus()
                        MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Function CekAdaTransaksi(ByVal vPrimary As Double) As Boolean
        Dim ObjStockInOut As New clsTransaksiStockInOut1
        Dim ObjJurnal As New ClsTransaksiJournal1
        Try

            If vPrimary = -1 Then
                Return False
            End If
            'cek di material usage
            If ObjStockInOut.GetJmlByAsset(vPrimary) > 0 Then
                'ada
                Return True
            End If
            'cek di other , depresiasi , disposal
            If ObjJurnal.GetJmlByPrimaryAsset(vPrimary) > 0 Then
                'ada
                Return True
            End If

            'jika tidak ada
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjStockInOut = Nothing
            ObjJurnal = Nothing
        End Try
    End Function

    Private Sub txtAcquiredValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAcquiredValue.TextChanged

    End Sub
End Class