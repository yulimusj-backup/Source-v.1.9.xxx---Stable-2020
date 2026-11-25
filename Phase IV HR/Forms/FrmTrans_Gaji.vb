Option Explicit On
Option Strict Off
Imports Core_Module, System.Drawing, System.Windows.Forms
Public Class FrmTrans_Gaji 

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ijK As New FrmTransaksiLoan_Payment
    'Keterangan PrimaryQuick untuk posisi edit di DGV Quick
    Public PrimaryQuick As Double = -1
    '------------------------------------------------------
    Public ShowQuick_Edit As Boolean = False
    Public NewQuick_Data As Boolean = False
    Public DeleteQuick_Data As Boolean = False
    Public Baru As Boolean = False
    'Keterangan Primary Gain Acc dan Loss Acc untuk new quick
    Public mPrimaryCashIn As Double = -1
    Public mPrimaryGainAcc As Double = -1
    Public mPrimaryLossAcc As Double = -1
    Public mPrimaryPaymentAkunDefault As String = "-1"
    '--------------------------------------------------------
    'Keterangan ngakal in button simpan
    Public WaitingForm As Boolean = True
    Public TemporaryDataJgnDiHapus As Boolean = False
    '--------------------------------------------------------
    'Untuk Msgbox Error Salary Formula
    Public TBLErrorFormula As New DseTransaksiSalary.TblErrorFormulaDataTable
    Public mPrimarySalaryFormulaRef As Double = -1
    Dim TBLSL_Item_Ref As New DseTransaksiSalary.tblmaster_salary_formula_salarycomponentDataTable 
    Dim DTAVar As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_salarycomponent_variableTableAdapter
    Dim TBLVar_Ref As New DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_variableDataTable
    Dim DTAConst As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_salarycomponent_constraintTableAdapter
    Dim TBLConst_Ref As New DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_constraintDataTable 

    Public Obj As New ClsTransaksiSalary
    Public ObjCD As New Core_Module.ClsCekData
    Dim ObjLicense As New ClsLicense
    Dim ObjItem As New ClsTransaksiSalary_Item
    Dim ObjPayment As New ClsTrans_Payment
    Dim ObjPaymentItem As New ClsTrans_Payment_Item
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjAkun As New AccModule.ClsMasterAkun
    Dim ObjMstSlry As New Core_Module.ClsMasterSalaryFormula
    Dim ObjMstSlry_Item As New Core_Module.ClsMasterSalaryFormula_item
    Dim PrimaryCurr_Default As Double = 1
    Public Def_Prec As Double = Me.ObjCurr.GetDecimal_DefaultCurrency
    Public ObjKrywn As New Core_Module.ClsMasterKaryawan
    Dim ObjKompetensi As New Core_Module.ClsMasterKompetensi
    Dim ObjPostion As New Core_Module.ClsMasterJobPosition
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim Buffer As String = ""
    Dim BufferDari As String = ""
    Dim BufferSampai As String = ""
    Dim BufferCBOPeriod As String = ""
    Dim BufferMetodeTax As String = ""
    Dim BufferPrimaryEmployee As String = ""
    Dim DummyDari As String = ""
    Dim DummySampai As String = ""
    Dim DummyCBOPeriod As String = ""
    Dim DummyMetodeTax As String = ""
    Dim DummyPrimaryEmployee As String = ""
    Dim FrmPaymentData As New FrmTrans_Gaji_Payment
    Dim FrmAtt As New FrmTrans_Gaji_ShowAttendance
    Dim FrmOtherDeduction As New FrmTrans_Gaji_OtherDeduction
    Dim TBLSL_Kary As New DseTransaksiSalary.tblmaster_salary_formula_component_detailemployeeDataTable
    Dim TBLSL_Item As New DseTransaksiSalary.tblmaster_salary_formula_salarycomponentDataTable
    Dim TBLSL_Item_Other As New DseTransaksiSalary.tblmaster_salary_formula_otherdeductionDataTable
    Dim Perct_BiayaJabatan As Decimal = 0.05
    Dim BiayaJabatanMaxValue As Decimal = 6000000
    Dim ObjPTKP As New HR.ClsMasterSettingPajak
    Dim vTBLTarif As New DseTransaksiSalary.tblmaster_tarif_pphDataTable
    Dim vTBLTarifPPHFinal As New DseTransaksiSalary.tblmaster_tarif_pph_final_pesangonDataTable
    Dim vTBLOtherPayment As New DseTransaksiSalary.tbltrans_salary_item_other1DataTable
    Dim vTarifPPhFinal As Decimal = 0
    Dim vPenghasilanBruto As Decimal = 0
    Dim ObjBB As New ClsTrans_Tax_Beginning_Balance
    Dim TBLTaxBB As New DseTrans_Tax_Beginning_Balance.tbltrans_tax_beginning_balance_item1DataTable
    Dim TBLTaxBB2 As New DseTrans_Tax_Beginning_Balance.tbltrans_tax_beginning_balance_item1DataTable

    ''' <summary>
    ''' Untuk Mengambil Nilai Dari Parameter
    ''' </summary>
    ''' <remarks></remarks>
    Dim TBLParam1 As New DseMasterParameter.DataTable1DataTable
    Dim TBLParam2 As New DseMasterParameter.tbltrans_attendanceDataTable
    Dim TBLWeekDay As New DseMasterParameter.tbltrans_attendance1DataTable
    Dim TBLHoliday As New DseMasterParameter.tbltrans_attendance1DataTable
    Dim TblKaryawan As New DseMasterParameter.tblmaster_karyawanDataTable
    Dim ObjParam As New Core_Module.ClsMasterParameterGaji
    Dim TBLStartEnd As New DseTransaksiSalary.tbltrans_salary1DataTable
    Dim DariCancelBtn As Boolean = False
    Dim TBLItemTax As New DseTransaksiSalary.tbltrans_salary_itemtaxDataTable
    Dim TBLItemIterasi As New DseTransaksiSalary.tbltrans_salary_item_iterationDataTable
    Dim TBLItemIterasiPPHFinal As New DseTransaksiSalary.tbltrans_salary_item_iteration1DataTable
    Public DariBtnStop As Boolean = False
    Public DariBtnPause As Boolean = False

    Public ShowSlr_FromQuick As Boolean = False

    'show item tbl prev. income
    Dim TBLPrevIncome As New DseTransaksiSalary.tbltrans_salary_item_prev_income1DataTable
    Dim rowPrevIncome As DseTransaksiSalary.tbltrans_salary_item_prev_income1Row

    Public Sub FungsiLoad_Event()
        Try
            If Me.ShowQuick_Edit = True Then
                If Me.NewQuick_Data = False Then
                    Me.Obj = New ClsTransaksiSalary
                    Me.Baru = True
                    Me.KosongkanField()
                    Me.NonAktifkanDetail()
                    Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
                    Me.FungsiBtnFind()
                    Me.FungsiBtnEdit()
                    Me.BtnStop.Visible = False
                    Me.BtnStop.TabStop = False
                    Me.lblProgress.Visible = False
                    Me.ProgressBar1.Visible = False
                ElseIf Me.NewQuick_Data = True Then
                    'kondisi Btn New Quick
                    'MsgBox("Belum dicoding", MsgBoxStyle.Information, "Informasi")
                    Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
                    If Me.TemporaryDataJgnDiHapus = False Then
                        Me.Obj = New ClsTransaksiSalary
                        Me.Baru = True
                        Me.KosongkanField()
                        Me.NonAktifkanDetail()
                        Me.FungsiBtnNew()
                        Me.DtpDate.Focus()
                    End If
                End If
            Else
                'Kondisi Normal
                Me.Obj = New ClsTransaksiSalary
                Me.Baru = True
                Me.KosongkanField()
                Me.NonAktifkanDetail()
                Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
                Me.BtnStop.Visible = False
                Me.BtnStop.TabStop = False
                Me.lblProgress.Visible = False
                Me.ProgressBar1.Visible = False
            End If

            Me.ijK.vFromSalaryCalc = True

            If Me.ShowSlr_FromQuick = True Then
                Me.Show_Salary(Me.PrimaryQuick)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Show_Salary(ByVal vPrimarySalary As Double)
        Dim ijk As New FrmTrans_Gaji_Cari
        Try
            If Not vPrimarySalary = -1 Then
                If Me.Obj.CariDetail(vPrimarySalary, Me.ObjCD, Me.dgvItem, Me.FrmOtherDeduction.dgvItem, Me.FrmPaymentData.DGVPayment, _
                   Me.Def_Prec, Me, Me.FrmPaymentData, Me.TBLItemIterasi, Me.TBLItemTax, Me.vTBLOtherPayment, Me.vPenghasilanBruto, Me.TBLPrevIncome) = True Then
                    Me.SalinKeField()
                    Me.ObjItem.IsiDgvAttendance(Me.TxtPrimaryEmployee.Text, Me.DtpTo.Value, Me.dtpPeriod.Value, Me.FrmAtt)
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub FrmTransaksiTicketing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ShowIcon = False
            Me.FungsiLoad_Event()
            If Me.UserCode <> "USER" And Me.UserCode <> "ERMIWATI" Then
                Me.llbEditPPh.Visible = False
            End If

            'Me.TxtPeriodCount.Visible = True
            'Me.TxtPeriodEnd.Visible = True
            'Me.TxtPeriodStart.Visible = True
            'Me.TxtPrimaryDept.Visible = True
            'Me.TxtIsResign.Visible = True

            If Me.DTAVar.Connection.ConnectionString <> Core_Module.GetConstring Then
                Me.DTAVar.Connection.ConnectionString = Core_Module.GetConstring
                Me.DTAConst.Connection.ConnectionString = Me.DTAVar.Connection.ConnectionString
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CekData()
        Dim idx As Integer
        Try
            idx = Me.ijK.dgvItem.RowCount
            idx = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnOk()
        Dim hasil As MsgBoxResult
        Try
            'Me.CekData()
            'Exit Sub
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    If Me.NewQuick_Data = True Then
                        hasil = MsgBoxResult.Yes
                    Else
                        hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    End If

                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, Me.FrmAtt.dgvItem, Me.FrmOtherDeduction.dgvItem, _
                        Me.FrmPaymentData.DGVPayment, Me.ijK.TxtDocumentNo.Text, Me, Me.ijK, Me.FrmOtherDeduction, Me.FrmPaymentData, _
                        Me.TBLItemIterasi, Me.TBLItemTax, vTBLOtherPayment, Me.vPenghasilanBruto, Me.TBLPrevIncome) = True Then
                            If Me.NewQuick_Data = True Then
                                If Me.ShowQuick_Edit = False Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                ElseIf Me.ShowQuick_Edit = True Then
                                End If
                            Else
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            End If
                            If Me.ShowQuick_Edit = False Then
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            ElseIf Me.ShowQuick_Edit = True Then
                                If Me.NewQuick_Data = True Then
                                    Me.DariCancelBtn = True
                                    Me.DariBtnStop = False
                                    Me.DariBtnPause = False
                                    Me.Hide()
                                End If
                            End If
                        Else
                            'hitung ulang Loan Installment
                            Me.TxtLoanInstallment.Text = Me.ObjCD.DigitGrouping(Me.ijK.TotalLoanInstallment, Me.ijK.Def_Prec)
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, Me.FrmAtt.dgvItem, Me.FrmOtherDeduction.dgvItem, _
                        Me.FrmPaymentData.DGVPayment, Me, Me.ijK, Me.FrmOtherDeduction, Me.FrmPaymentData, Me.TBLItemIterasi, Me.TBLItemTax, _
                        Me.vTBLOtherPayment, Me.vPenghasilanBruto, Me.TBLPrevIncome) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            If Me.ShowQuick_Edit = False Then
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            ElseIf Me.ShowQuick_Edit = True Then
                                If Me.NewQuick_Data = False Then
                                    Me.DariCancelBtn = True
                                    Me.Close()
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If Me.DariBtnPause = True Then
                    Me.ApakahDataBenar()
                    Me.DariBtnPause = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FungsiBtnOk_TanpaPeringatan() As Boolean
        'Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    'hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    'If hasil = MsgBoxResult.Yes Then
                    If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, Me.FrmAtt.dgvItem, _
                    Me.FrmOtherDeduction.dgvItem, Me.FrmPaymentData.DGVPayment, Me.ijK.TxtDocumentNo.Text, Me, Me.ijK, _
                    Me.FrmOtherDeduction, Me.FrmPaymentData, Me.TBLItemIterasi, Me.TBLItemTax, Me.vTBLOtherPayment, _
                    Me.vPenghasilanBruto, Me.TBLPrevIncome) = True Then
                        'MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        If Me.ShowQuick_Edit = False Then
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        ElseIf Me.ShowQuick_Edit = True Then
                            If Me.NewQuick_Data = True Then
                                Me.DariCancelBtn = True
                                Me.DariBtnStop = False
                                Me.DariBtnPause = False
                                Me.Hide()
                            End If
                        End If
                    Else
                        If Me.ShowQuick_Edit = True Then
                            If Me.NewQuick_Data = True Then
                                Return False
                            End If
                        End If
                    End If
                    'End If
                Else
                    'konfirmasi disimpan atau tidak
                    'hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    'If hasil = MsgBoxResult.Yes Then
                    If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, Me.FrmAtt.dgvItem, Me.FrmOtherDeduction.dgvItem, _
                    Me.FrmPaymentData.DGVPayment, Me, Me.ijK, Me.FrmOtherDeduction, Me.FrmPaymentData, Me.TBLItemIterasi, Me.TBLItemTax, _
                    Me.vTBLOtherPayment, Me.vPenghasilanBruto, Me.TBLPrevIncome) = True Then
                        'MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        If Me.ShowQuick_Edit = False Then
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        ElseIf Me.ShowQuick_Edit = True Then
                            If Me.NewQuick_Data = False Then
                                Me.DariCancelBtn = True
                                Me.Hide()
                            End If
                        End If
                    Else
                        If Me.ShowQuick_Edit = True Then
                            If Me.NewQuick_Data = True Then
                                Return False
                            End If
                        End If
                    End If
                    'End If
                End If
            Else
                If Me.DariBtnPause = True Then
                    Me.ApakahDataBenar()
                    Me.DariBtnPause = False
                End If
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub BtnSimpan_PerformClick()
        Try
            If Me.ShowQuick_Edit = True And Me.NewQuick_Data = True Then
                If Me.WaitingForm = True Then
                    Me.FungsiBtnOk()
                Else
                    If Me.FungsiBtnOk_TanpaPeringatan() = False Then
                        Me.TemporaryDataJgnDiHapus = True
                        Me.WaitingForm = True
                        Me.ShowDialog()
                        Me.TemporaryDataJgnDiHapus = False
                    End If
                End If
            Else
                Me.FungsiBtnOk()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Kontrol Button
    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.BtnSimpan_PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnCancel()
        Dim hasil As MsgBoxResult
        Try
            If Me.ShowQuick_Edit = False Then
                If Not Me.TxtDocumentNo.Text = "" Then
                    'konfirmasi batal atau tidak
                    hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Baru = True Then
                            Me.KosongkanField()
                            Me.TxtDocumentNo.Clear()
                            'LP
                            Me.ijK.KosongkanField()
                            Me.ijK.TxtDocumentNo.Clear()
                        Else
                            Me.Obj.CariDetail(Me.TxtPrimary.Text, Me.ObjCD, _
                            Me.dgvItem, Me.FrmOtherDeduction.dgvItem, Me.FrmPaymentData.DGVPayment, Me.Def_Prec, Me, _
                            Me.FrmPaymentData, Me.TBLItemIterasi, Me.TBLItemTax, Me.vTBLOtherPayment, Me.vPenghasilanBruto, Me.TBLPrevIncome)
                            Me.SalinKeField()
                            'LP
                            Me.ijK.Obj.CariDetail(Me.ijK.TxtPrimaryLoanPayment.Text, Me.ObjCD, _
                            Me.ijK.dgvItem, Me.ijK.Def_Prec, Me.ijK)
                            Me.ijK.SalinKeField()
                        End If
                        Me.ijK.Hide()
                    Else
                        Exit Sub
                    End If
                Else
                    Me.KosongkanField()
                End If

                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            ElseIf Me.ShowQuick_Edit = True Then
                If Me.NewQuick_Data = False Then
                    Me.DariCancelBtn = True
                    Me.Close()
                ElseIf Me.NewQuick_Data = True Then
                    Me.DariCancelBtn = True
                    Me.Close()
                    'MsgBox("Belum dikoding.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.FungsiBtnCancel()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnNew()
        Try
            Me.Baru = True
            Obj.ResetObject(Me.dgvItem, Me.FrmOtherDeduction.dgvItem)
            Me.AktifkanDetail()
            Me.KosongkanField()
            ''Me.DtpDate.Value = Me.Obj.GetTanggalBaru()
            Me.DtpDate.Value = Today
            Me.dtpPeriod.Value = Me.DtpDate.Value
            Me.DtpTo.Value = Me.DtpDate.Value
            Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
            Me.Txt_P_NoForm.Text = Me.Obj.GetNoBaru_BuatNPWP(Me.DtpDate.Value, Me.ObjCD)
            Me.Obj.IsiKeteranganKolomTabPage3(Me)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.DtpDate.Focus()
            'LP
            Me.ijK.Baru = True
            Me.ijK.UserCode = Me.UserCode
            Me.ijK.ShowFromSalary = True
            'Sementara aja---------------------
            If Me.ShowQuick_Edit = True Then
                If Me.NewQuick_Data = True Then
                    Me.ijK.PrimaryCashIn = Me.mPrimaryCashIn
                    Me.ijK.PrimaryGain = Me.mPrimaryGainAcc
                    Me.ijK.PrimaryLoss = Me.mPrimaryLossAcc
                    Me.ijK.BackColorAll = Me.BackColorAll
                End If
            End If
            '----------------------------------
            Me.ijK.Show()
            Me.ijK.Fungsi_Load()
            Me.ijK.Hide()
            'Ohter Deduction
            Me.FrmOtherDeduction.Show()
            Me.FrmOtherDeduction.LoadData()
            Me.FrmOtherDeduction.Hide()
            'Masukan Payment Data
            Me.FrmPaymentData.BackColorAll = Me.BackColorAll
            Me.FrmPaymentData.PrimarySalary = CDbl(Me.TxtPrimary.Text)
            Me.FrmPaymentData.Def_Prec = Me.Def_Prec
            Me.FrmPaymentData.Baru = Me.Baru
            Me.FrmPaymentData.FrmSalary = Me
            'Me.mPrimaryPaymentAkunDefault = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            'ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryCashAccount)
            Me.mPrimaryPaymentAkunDefault = Me.ObjMstSlry.GetPrimaryDefaultAccount(Me.TxtPrimaryEmployee.Text)
            If String.IsNullOrEmpty(Me.mPrimaryPaymentAkunDefault) = True Then
                Me.mPrimaryPaymentAkunDefault = -1
            Else
                Me.BuatPaymentData(Me.Baru)
            End If
            Me.FrmPaymentData.Show()
            Me.FrmPaymentData.LoadData()
            Me.FrmPaymentData.Hide()

            'Me.TxtEmployeeCode1.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub BuatPaymentData(ByVal vBaru As Boolean, Optional ByVal DipanggilDariBtnCalculate As Boolean = False)
        Dim RowDGV As DataGridViewRow
        Dim PaymentAmount As Decimal = 0
        Dim IdentifierStr As String = ""
        Try
            If vBaru = True Then
                Me.FrmPaymentData.DGVPayment.Rows.Clear()
                Me.FrmPaymentData.DGVPayment.Rows.Add()
                RowDGV = Me.FrmPaymentData.DGVPayment.Rows(Me.FrmPaymentData.DGVPayment.Rows.Count - 1)
                RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_PrimaryAkun.Index).Value = Me.mPrimaryPaymentAkunDefault
                RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_AccountCode.Index).Value = Me.ObjAkun.GetAccountCode(Me.mPrimaryPaymentAkunDefault)
                RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_AccountDesc.Index).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.mPrimaryPaymentAkunDefault)
                RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_Remarks.Index).Value = ""
                IdentifierStr = Me.ObjPaymentItem.GetIdentifierUrutan_BuatDGVCellEndEdit( _
                Me.FrmPaymentData.DGVPayment, Me.FrmPaymentData.ColDetailPayment_Identifier.Index, _
                vBaru, Me.TxtPrimary.Text, 0)
                Me.FrmPaymentData.DGVPayment.AllowUserToAddRows = False
                Me.FrmPaymentData.DGVPayment.Item(Me.FrmPaymentData.ColDetailPayment_Identifier.Index, Me.FrmPaymentData.DGVPayment.Rows.Count - 1).Value = IdentifierStr
                RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_IndexRowSalary.Index).Value = "0"
                If Convert.ToDecimal(Me.TxtSalaryReceived.Text) = 0 Then
                    RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_Amount.Index).Value = Me.ObjCD.DigitGrouping("1", Me.Def_Prec)
                    Me.TxtPayment.Text = Me.ObjCD.DigitGrouping("1", Me.Def_Prec)
                Else
                    RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_Amount.Index).Value = Me.TxtSalaryReceived.Text
                    Me.TxtPayment.Text = Me.TxtSalaryReceived.Text
                End If
            Else
                If DipanggilDariBtnCalculate = False Then
                    Me.Obj.BuatPaymentData_Posisi_Edit(Me.FrmPaymentData.DGVPayment, Me.TxtPrimary.Text, Me.FrmPaymentData, Me.ObjAkun, Me.Def_Prec, PaymentAmount)
                    Me.TxtPayment.Text = Me.ObjCD.DigitGrouping(PaymentAmount, Me.Def_Prec)
                Else
                    'Idem Dengan Kondisi New = true
                    Me.FrmPaymentData.DGVPayment.Rows.Clear()
                    Me.FrmPaymentData.DGVPayment.Rows.Add()
                    RowDGV = Me.FrmPaymentData.DGVPayment.Rows(Me.FrmPaymentData.DGVPayment.Rows.Count - 1)
                    RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_PrimaryAkun.Index).Value = Me.mPrimaryPaymentAkunDefault
                    RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_AccountCode.Index).Value = Me.ObjAkun.GetAccountCode(Me.mPrimaryPaymentAkunDefault)
                    RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_AccountDesc.Index).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.mPrimaryPaymentAkunDefault)
                    RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_Remarks.Index).Value = ""
                    IdentifierStr = Me.ObjPaymentItem.GetIdentifierUrutan_BuatDGVCellEndEdit( _
                    Me.FrmPaymentData.DGVPayment, Me.FrmPaymentData.ColDetailPayment_Identifier.Index, _
                    vBaru, Me.TxtPrimary.Text, 0)
                    Me.FrmPaymentData.DGVPayment.AllowUserToAddRows = False
                    Me.FrmPaymentData.DGVPayment.Item(Me.FrmPaymentData.ColDetailPayment_Identifier.Index, Me.FrmPaymentData.DGVPayment.Rows.Count - 1).Value = IdentifierStr
                    RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_IndexRowSalary.Index).Value = "0"
                    If Convert.ToDecimal(Me.TxtSalaryReceived.Text) = 0 Then
                        RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_Amount.Index).Value = Me.ObjCD.DigitGrouping("1", Me.Def_Prec)
                        Me.TxtPayment.Text = Me.ObjCD.DigitGrouping("1", Me.Def_Prec)
                    Else
                        RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_Amount.Index).Value = Me.TxtSalaryReceived.Text
                        Me.TxtPayment.Text = Me.TxtSalaryReceived.Text
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.FungsiBtnNew()
            'Me.TxtEmployeeCode1.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnFind()
        Dim ijk As New FrmTrans_Gaji_Cari
        Try
            If Me.ShowQuick_Edit = False Then
                ijk.BackColorAll = Me.BackColorAll
                ijk.cboCriteria.Items.Clear()
                ijk.txtCari.Text = ""
                ijk.dgvHasil.DataSource = ""
                ijk.ShowIcon = False
                ijk.ShowDialog()
                If Not ijk.mPrimary = -1 Then
                    If Me.Obj.ApakahBolehFindDataSalary(ijk.mPrimary, Me.UserCode) = False Then
                        Exit Sub
                    Else
                        If Me.Obj.CariDetail(ijk.mPrimary, Me.ObjCD, Me.dgvItem, Me.FrmOtherDeduction.dgvItem, Me.FrmPaymentData.DGVPayment, _
                        Me.Def_Prec, Me, Me.FrmPaymentData, Me.TBLItemIterasi, Me.TBLItemTax, Me.vTBLOtherPayment, Me.vPenghasilanBruto, Me.TBLPrevIncome) = True Then
                            Me.SalinKeField()
                            Me.ObjItem.IsiDgvAttendance(Me.TxtPrimaryEmployee.Text, Me.DtpTo.Value, Me.dtpPeriod.Value, Me.FrmAtt)
                            Me.btnEdit.Focus()
                        Else
                            Me.btnFind.Focus()
                        End If
                        Me.btnFind.Focus()
                    End If
                End If
            ElseIf Me.ShowQuick_Edit = True Then
                If Me.NewQuick_Data = False Then
                    If Not Me.PrimaryQuick = "-1" Then
                        If Me.Obj.CariDetail(Me.PrimaryQuick, Me.ObjCD, Me.dgvItem, Me.FrmOtherDeduction.dgvItem, Me.FrmPaymentData.DGVPayment, _
                        Me.Def_Prec, Me, Me.FrmPaymentData, Me.TBLItemIterasi, Me.TBLItemTax, vTBLOtherPayment, Me.vPenghasilanBruto, Me.TBLPrevIncome) = True Then
                            Me.SalinKeField()
                            Me.ObjItem.IsiDgvAttendance(Me.TxtPrimaryEmployee.Text, Me.DtpTo.Value, Me.dtpPeriod.Value, Me.FrmAtt)
                            Me.btnEdit.Focus()
                        Else
                            Me.btnFind.Focus()
                        End If
                        Me.btnFind.Focus()
                    End If
                ElseIf Me.NewQuick_Data = True Then
                    MsgBox("Belum dikoding", MsgBoxStyle.Information, "Informasi")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk.Dispose()
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Try
            Me.FungsiBtnFind()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekTimestamp() As Boolean
        Try
            Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimary.Text)
            If Me.Obj.TimeStamp <> OldTimeStamp Then
                If MsgBox("Other user have modified this data, reload latest data?", _
                MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.Obj.CariDetail(Me.Obj.Primary, ObjCD, Me.dgvItem, Me.FrmOtherDeduction.dgvItem, Me.FrmPaymentData.DGVPayment, _
                    Me.Def_Prec, Me, Me.FrmPaymentData, Me.TBLItemIterasi, Me.TBLItemTax, Me.vTBLOtherPayment, Me.vPenghasilanBruto, Me.TBLPrevIncome)
                    If Me.TxtPrimaryLoan.Text = "-1" Then
                        Me.Obj.LoanInstallment = 0
                    End If
                    Me.SalinKeField()
                Else
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub FungsiBtnEdit()
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtDocumentNo.Text <> "" Then
                If Me.CekTimestamp() = True Then
                    If Me.Obj.ApakahBolehEdit(Me.TxtPrimary.Text, Me.UserCode, True, True) = True Then
                        Me.Baru = False
                        Me.AktifkanDetail()
                        Me.DtpDate.Focus()
                        'LP
                        Me.ijK.Baru = False
                        Me.ijK.UserCode = Me.UserCode
                        Me.ijK.PrimaryLoanPaymentFromSalary = Me.TxtPrimaryLoan.Text
                        Me.ijK.ShowFromSalary = True
                        Me.ijK.vFromSalaryCalc = True

                        Me.ijK.KosongkanField()
                        Me.ijK.Show()
                        Me.ijK.Fungsi_Load()
                        Me.ijK.Hide()
                        'OtherDeduction
                        Me.FrmOtherDeduction.Show()
                        Me.FrmOtherDeduction.LoadData()
                        Me.FrmOtherDeduction.Hide()
                        'Membuat Payment Data
                        Me.FrmPaymentData.BackColorAll = Me.BackColorAll
                        Me.FrmPaymentData.PrimarySalary = CDbl(Me.TxtPrimary.Text)
                        Me.FrmPaymentData.Def_Prec = Me.Def_Prec
                        Me.FrmPaymentData.Baru = Me.Baru
                        Me.FrmPaymentData.FrmSalary = Me
                        Me.BuatPaymentData(Me.Baru)
                        Me.FrmPaymentData.Show()
                        Me.FrmPaymentData.LoadData()
                        Me.FrmPaymentData.Hide()
                    Else
                        If Me.ShowQuick_Edit = True Then
                            Me.Close()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            Me.FungsiBtnEdit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocumentNo.Text <> "" Then
                If Me.CekTimestamp() = True Then
                    If Me.Obj.ApakahBolehHapus(Me.TxtPrimary.Text, Me.UserCode, True, True, Me.dgvItem, Me) = False Then
                        Me.btnFind.Focus()
                    Else
                        'LP
                        Me.ijK.Baru = False
                        Me.ijK.PrimaryLoanPaymentFromSalary = Me.TxtPrimaryLoan.Text
                        Me.ijK.ShowFromSalary = True
                        Me.ijK.KosongkanField()
                        Me.ijK.UserCode = Me.UserCode
                        Me.ijK.Show()
                        Me.ijK.Fungsi_Load()
                        Me.ijK.Hide()
                        'Other Deduction
                        Me.FrmOtherDeduction.Show()
                        Me.FrmOtherDeduction.LoadData()
                        Me.FrmOtherDeduction.Hide()
                        hasil = MsgBox("Delete data?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.Obj.Hapus(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, _
                             Me, Me.ijK, True) = True Then
                                Me.KosongkanField()
                                Me.btnNew.Focus()
                                MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtDocumentNo.Enabled = True Then
                If Me.DariCancelBtn = True Then
                    Me.DariCancelBtn = False
                    Me.WaitingForm = False
                    Exit Sub
                End If
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, _
                            Me.dgvItem, Me.FrmAtt.dgvItem, Me.FrmOtherDeduction.dgvItem, Me.FrmPaymentData.DGVPayment, Me.ijK.TxtDocumentNo.Text, Me, Me.ijK, _
                            Me.FrmOtherDeduction, Me.FrmPaymentData, Me.TBLItemIterasi, Me.TBLItemTax, vTBLOtherPayment, Me.vPenghasilanBruto, Me.TBLPrevIncome) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                If Me.ShowQuick_Edit = False Then
                                    Me.SalinKeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnFind.Focus()
                                ElseIf Me.ShowQuick_Edit = True Then
                                    Me.WaitingForm = False
                                End If
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, Me.FrmAtt.dgvItem, Me.FrmOtherDeduction.dgvItem, _
                            Me.FrmPaymentData.DGVPayment, Me, Me.ijK, Me.FrmOtherDeduction, Me.FrmPaymentData, Me.TBLItemIterasi, Me.TBLItemTax, _
                            vTBLOtherPayment, Me.vPenghasilanBruto, Me.TBLPrevIncome) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                If Me.ShowQuick_Edit = False Then
                                    Me.SalinKeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnFind.Focus()
                                ElseIf Me.ShowQuick_Edit = True Then
                                    Me.WaitingForm = False
                                End If
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

    Private Sub DtpDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.Leave
        Try
            If Me.BtnHitung.Enabled = True Then
                Me.BackgroundWorker1.RunWorkerAsync()
                Me.BackgroundWorkerPPh21.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'GetNoBaru
    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.ValueChanged
        Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Fungsi Standar
    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.BackColorDefault()

            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToDeleteRows = False

            'Me.ChkShowZeroAmount.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()

            'set dgv
            Me.dgvItem.ReadOnly = False
            Me.dgvItem.AllowUserToDeleteRows = False

            'Posisi Edit
            If Me.Baru = False Then
                Me.TxtEmployeeCode1.Enabled = False
                Me.BtnFindEmployeeCode.Enabled = False
                'Fill TBLStartEnd
                'Me.Obj.FillTBLStartEnd(Me.TBLStartEnd, Me.TxtPrimaryEmployee.Text, False, Me.TxtPrimary.Text)
            Else
                'Fill TBLStartEnd
                'Me.Obj.FillTBLStartEnd(Me.TBLStartEnd, Me.TxtPrimaryEmployee.Text, True, -1)
            End If

            If Me.DtpDate.Value.Date.Month <> 12 Then
                Me.TxtIncomeTax.Visible = True
                Me.TxtKurangBayarTax.Visible = False
            Else
                Me.TxtIncomeTax.Visible = False
                Me.TxtKurangBayarTax.Visible = True
            End If
            Me.LblAttendance.Enabled = True
            Me.CboPaymentStatus.Enabled = False
            Me.CboSalary.Enabled = True
            Me.DtpPaymentStatus.Enabled = False
            Me.TxtEmployeeName.Enabled = False
            Me.TxtPosition.Enabled = False
            Me.TxtExpAccDesc.Enabled = False
            Me.TxtIncome.Enabled = False
            Me.TxtDeduction.Enabled = False
            Me.TxtLoanInstallment.Enabled = False
            Me.TxtIncomeTax.Enabled = False
            Me.TxtKurangBayarTax.Enabled = False
            Me.TxtSalaryReceived.Enabled = False
            Me.LinkLabel2.Enabled = True
            Me.LblIncomeTax.Enabled = True
            Me.ChkShowZeroAmount.Enabled = True
            Me.LblOtherDeduction.Enabled = True
            Me.LblPayment.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanFieldPerhitunganPajak()
        Try
            Me.Txt01Pjk_Gaji.Text = "0"
            Me.Txt01Pjk_GajiRutin.Text = "0"
            Me.Txt02PjkTunjanganPPh.Text = "0"
            Me.Txt02Pjk_TunjPPHDisetahunkan.Text = "0"
            Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text = "0"
            Me.Txt02Pjk_TunjPPHNonRutin.Text = "0"
            Me.Txt02Pjk_TunjPPHRutin.Text = "0"
            Me.Txt03Pjk_TunjLainnya.Text = "0"
            Me.Txt03Pjk_TunjLainnyaRutin.Text = "0"
            Me.Txt04PjkHonorarium.Text = "0"
            Me.Txt04PjkHonorariumRutin.Text = "0"
            Me.Txt05PjkPremiAsuransi.Text = "0"
            Me.Txt05PjkPremiAsuransiRutin.Text = "0"
            Me.Txt06PjkNatura.Text = "0"
            Me.Txt06PjkNaturaRutin.Text = "0"
            Me.Txt07PjkTantiem.Text = "0"
            Me.Txt07PjkTantiemRutin.Text = "0"
            Me.Txt08PjkPenghasilanBruto.Text = "0"
            Me.Txt08PjkPenghasilanBrutoRutin.Text = "0"
            Me.Txt09PjkBiayaJabatan.Text = "0"
            Me.Txt09PjkBiayaJabatanRutin.Text = "0"
            Me.Txt10PjkIuranPensiun.Text = "0"
            Me.Txt10PjkIuranPensiunRutin.Text = "0"
            Me.Txt11PjkJmlPengurangan.Text = "0"
            Me.Txt11PjkJmlPenguranganRutin.Text = "0"
            Me.Txt12PjkJmlPenghasilanNetto.Text = "0"
            Me.Txt13PjkJmlHasilMsPjkSblm.Text = "0"
            Me.Txt14PjkJmlHasilNettoUPPh.Text = "0"
            Me.Txt14PjkJmlHasilNettoUPPhRutin.Text = "0"
            Me.Txt15Pjk_PTKP.Text = "0"
            Me.Txt16Pjk_PKPSetahun.Text = "0"
            Me.Txt16Pjk_PKPSetahunRutin.Text = "0"
            Me.Txt17Pjk_PPH21.Text = "0"
            Me.Txt17Pjk_PPH21Rutin.Text = "0"
            Me.Txt18.Text = "0"
            Me.Txt19Pjk_PPH21Terutang.Text = "0"
            Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text = "0"
            Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text = "0"
            Me.Txt19Pjk_PPH21TerutangSetahun.Text = "0"
            Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text = "0"
            Me.Txt20.Text = "0"
            Me.Txt21Pjk_Pesangon.Text = "0"
            Me.Txt21Pjk_PPhFinalPesangon.Text = "0"
            Me.Txt21Pjk_TunjPPhPesangon.Text = "0"

            Me.Txt_P_NoForm.Text = ""
            Me.Txt_P_NPWPPemotong.Text = ""
            Me.Txt_P_NPWPPemotong.Text = ""
            Me.Txt_P_A_NPWP.Text = ""
            Me.Txt_P_NoNIK.Text = ""
            Me.Txt_P_ANama.Text = ""
            Me.Txt_P_Alamat1.Text = ""
            Me.Txt_P_Alamat2.Text = ""
            Me.Txt_P_JenisKelamin.Text = ""
            Me.Txt_P_StatusNPWP.Text = ""
            Me.Txt_P_NamaJabatan.Text = ""
            Me.Txt_P_KaryawanAsing.Text = ""
            Me.Txt_P_KodeNegaraDomisili.Text = ""
            Me.Txt_P_NPWP.Text = ""
            Me.Txt_P_Nama.Text = ""
            Me.TxtKurangBayarTax.Text = "0"
            Me.TxtIncomeTax.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurr.GetDecimal_DefaultCurrency)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.TxtPrimary.Text = "-1"
            Me.TxtPrimaryEmployee.Text = "-1"
            Me.TxtPrimaryDept.Text = "-1"
            Me.TxtPrimaryPosition.Text = "-1"
            Me.TxtPrimarySalaryFormula.Text = "-1"
            Me.TxtPrimaryTaxExpAcc.Text = "-1"
            Me.TxtPrimaryLoan.Text = "-1"
            Me.CboPaymentStatus.SelectedIndex = 0
            Me.CboSalary.SelectedIndex = 0
            Me.DtpPaymentStatus.Value = #1/1/1950#
            Me.dtpPeriod.Value = Today
            Me.BufferDari = Me.dtpPeriod.Value.ToString
            Me.DtpTo.Value = Today
            Me.BufferSampai = Me.DtpTo.Value.ToString
            Me.TxtIncome.Text = "0"
            Me.TxtDeduction.Text = "0"
            Me.TxtLoanInstallment.Text = "0"
            Me.TxtIncomeTax.Text = "0"
            Me.TxtEmployeeExpense.Text = "0"
            Me.TxtOtherPayment.Text = "0"
            Me.TxtCompanyExpense.Text = "0"
            Me.TxtSalaryReceived.Text = "0"
            Me.TxtKurangBayarTax.Text = "0"
            Me.Txt_A_IncomeAdjust.Text = "0"
            Me.Txt_A_PKPValue.Text = "0"
            Me.Txt_A_PositionExpense.Text = "0"
            Me.Txt_A_PositionExpenseAdj.Text = "0"
            Me.Txt_A_PTKPMonthly.Text = "0"
            Me.Txt_A_PPhValue.Text = "0"
            Me.TxtOtherDedution.Text = "0"
            Me.TxtGrossSalary.Text = "0"
            Me.TxtPayment.Text = "0"
            Me.KosongkanFieldPerhitunganPajak()
            Me.cboTaxCalculation.SelectedIndex = 0
            Me.vTBLOtherPayment.Rows.Clear()
            Me.vTarifPPhFinal = 0
            Me.vPenghasilanBruto = 0
            Me.TxtPeriodStart.Text = "1"
            Me.TxtPeriodEnd.Text = "12"
            Me.TxtPeriodCount.Text = 0
            Me.TxtIsResign.Text = "0"
            Me.TxtLosesTaxObligation.Text = "0"
            Me.TxtPrevIncome.Text = "0"
            Me.TxtPrevTaxPaid.Text = "0"

            If Me.Baru = True Then
                'Status paid visible
                Me.LblPaid.Text = "Unpaid"
                Me.DtpPaymentStatus.Value = #1/1/1950#
                Me.DtpPaymentStatus.Visible = False
                Me.CboPaymentStatus.SelectedIndex = 0
                Me.CboPaymentStatus.Visible = False
                Me.LblPaid.Visible = True
                Me.LblPaymentStatus.Visible = False
            Else
                'Status paid visible
                If Me.LblPaid.Text = "Paid" Then
                    Me.CboPaymentStatus.Visible = True
                    Me.DtpPaymentStatus.Visible = True
                    Me.LblPaid.Visible = False
                    Me.LblPaymentStatus.Visible = True
                Else
                    Me.CboPaymentStatus.Visible = False
                    Me.DtpPaymentStatus.Visible = False
                    Me.LblPaid.Visible = True
                    Me.LblPaymentStatus.Visible = False
                End If
            End If

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
            Me.FrmOtherDeduction.dgvItem.Rows.Clear()
            Me.BufferCBOPeriod = "-1"
            Me.DummyCBOPeriod = "-1"
            Me.BufferMetodeTax = "-1"
            Me.DummyMetodeTax = "-1"
            Me.BufferPrimaryEmployee = "-1"
            Me.DummyPrimaryEmployee = "-1"
            Me.ChkShowZeroAmount.Checked = False

            Me.TBLItemIterasi.Rows.Clear()
            Me.TBLItemTax.Rows.Clear()
            Me.TBLPrevIncome.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim RowDGV As DataGridViewRow
        Dim ObjMstSlry As New ClsTransaksiSalary
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.DtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1. cek data yang harus diisi
            KoleksiNamaData.Add("Employee Code")
            KoleksiTextBox.Add(Me.TxtEmployeeCode1)
            'KoleksiNamaData.Add("Tax Account Expense")
            'KoleksiTextBox.Add(Me.TxtExpAcc)

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '2. Fungsi Alt+O Btn Find
            If Me.TxtPrimaryEmployee.Text = -1 Then
                If Me.TxtEmployeeCode1.Text <> "" Then
                    If ObjKrywn.ApakahSudahAda(Me.TxtEmployeeCode1.Text) = False Then
                        Me.BtnFindEmployeeCode.PerformClick()
                    End If
                End If
            End If
            If Me.TxtPrimaryTaxExpAcc.Text = -1 Then
                If Me.TxtExpAcc.Text <> "" Then
                    If ObjAkun.ApakahSudahAda(Me.TxtExpAcc.Text) = False Then
                        Me.BtnFindEmployeeCode.PerformClick()
                    End If
                End If
            End If

            '3. harus ada data item
            If Me.dgvItem.Rows.Count <= 0 Then
                MsgBox("Please input data item.", MsgBoxStyle.Information, "Information")
                Me.dgvItem.Focus()
                Return False
            End If
            'Harus isi
            'If Me.TxtPrimaryTaxExpAcc.Text = "-1" Then
            '    MsgBox("Please choose tax expense account.", MsgBoxStyle.Information, "Information")
            '    Me.TxtExpAcc.Focus()
            '    Return False
            'End If
            If Me.TxtPrimaryEmployee.Text = "-1" Then
                MsgBox("Please input employee code.", MsgBoxStyle.Information, "Information")
                Me.TxtEmployeeCode1.Focus()
                Return False
            End If
            If Me.ObjCD.CekDataKosong(Me.TxtRemarks.Text, False) = False Then
                MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                Me.TxtRemarks.Focus()
                Return False
            End If

            '4. cek perubahan nomor transaksi
            If Me.Baru = False Then
                If Me.Obj.TanggalTransaksi.Month = Me.DtpDate.Value.Month And _
                Me.Obj.TanggalTransaksi.Year = Me.DtpDate.Value.Year Then
                    Me.TxtDocumentNo.Text = Me.Obj.NoTransaksi
                Else
                    Dim NoTransaksiLama As String = Me.Obj.GetNoTransaksi(Me.TxtPrimary.Text)
                    If NoTransaksiLama <> Me.TxtDocumentNo.Text Then
                        If Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = False Then
                            If MsgBox("You just changed transaction date, this will change your transaction " _
                                      + "number and there's a possibility your transaction no cannot be change back." _
                                      + "Continue and accept this change?", _
                                      MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                Me.DtpDate.Value = Me.Obj.TanggalTransaksi
                                Me.TxtDocumentNo.Text = Me.Obj.NoTransaksi
                                Return False
                            End If
                        End If
                    End If
                End If
            End If

            '5. cek data yang salah, double dalam kode
            If Me.Baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = True Then
                    Me.TxtDocumentNo.Text = Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetNoTransaksi(Me.TxtPrimary.Text)
                If TipeCodeLama <> Me.TxtDocumentNo.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = True Then
                        Me.TxtDocumentNo.Text = Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                    End If
                End If
                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.Obj.Primary, ObjCD, Me.dgvItem, Me.FrmOtherDeduction.dgvItem, _
                        Me.FrmPaymentData.DGVPayment, Me.Def_Prec, Me, Me.FrmPaymentData, Me.TBLItemIterasi, Me.TBLItemTax, _
                        vTBLOtherPayment, Me.vPenghasilanBruto, Me.TBLPrevIncome)
                        If Me.TxtPrimaryLoan.Text = "-1" Then
                            Me.Obj.LoanInstallment = 0
                        End If
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            '6. Cek Account Tax
            'Dim Counter As Double = 0
            'Dim TotalCounter As Double = 5
            'Dim Ketemu As Boolean = False
            'While Counter <= TotalCounter - 1
            '    If Counter = 0 Then
            '        If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryTaxExpAcc.Text)) = CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Income) Then
            '            Ketemu = True
            '            Exit While
            '        End If
            '    End If
            '    If Counter = 1 Then
            '        If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryTaxExpAcc.Text)) = CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Income) Then
            '            Ketemu = True
            '            Exit While
            '        End If
            '    End If
            '    If Counter = 2 Then
            '        If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryTaxExpAcc.Text)) = CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Expense) Then
            '            Ketemu = True
            '            Exit While
            '        End If
            '    End If
            '    If Counter = 3 Then
            '        If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryTaxExpAcc.Text)) = CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Other_Expense) Then
            '            Ketemu = True
            '            Exit While
            '        End If
            '    End If
            '    If Counter = 4 Then
            '        If CDbl(Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryTaxExpAcc.Text)) = CDbl(AccModule.ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales) Then
            '            Ketemu = True
            '            Exit While
            '        End If
            '    End If
            '    Counter += 1
            'End While
            'If Ketemu = False Then
            '    MsgBox("Please input account code with type income account or other income account or expense or other expense or COGS.", MsgBoxStyle.Information, "Information")
            '    Me.TxtExpAcc.Focus()
            '    Return False
            'End If

            'Tanggal Transaksi Tidak boleh < dari DtpFrom
            If Me.DtpDate.Value.Date < Me.dtpPeriod.Value.Date Then
                MsgBox("Transaction date can't be less than period date.", MsgBoxStyle.Information, "Information")
                Me.DtpDate.Focus()
                Return False
            End If

            If CDate(Me.DummyDari) <> Me.dtpPeriod.Value.Date Or CDate(Me.DummySampai) <> Me.DtpTo.Value Then
                If Me.dtpPeriod.Value.Date <> CDate(Me.BufferDari) Or Me.DtpTo.Value.Date <> CDate(Me.BufferSampai) Then
                    MsgBox("Please re-calculate salary.", MsgBoxStyle.Information, "Information")
                    Me.BtnHitung.Focus()
                    Return False
                End If
            End If

            If Me.CboSalary.SelectedIndex <> CInt(Me.DummyCBOPeriod) Then
                If Me.CboSalary.SelectedIndex <> CInt(Me.BufferCBOPeriod) Then
                    MsgBox("Please re-calculate salary.", MsgBoxStyle.Information, "Information")
                    Me.BtnHitung.Focus()
                    Return False
                End If
            End If

            'If Me.cboTaxCalculation.SelectedIndex <> CInt(Me.DummyMetodeTax) Then
            '    If Me.cboTaxCalculation.SelectedIndex <> CInt(Me.BufferMetodeTax) Then
            '        MsgBox("Please re-calculate salary.", MsgBoxStyle.Information, "Information")
            '        Me.BtnHitung.Focus()
            '        Return False
            '    End If
            'End If

            If Me.TxtPrimaryEmployee.Text <> CInt(Me.DummyPrimaryEmployee) Then
                If Me.TxtPrimaryEmployee.Text <> CInt(Me.BufferPrimaryEmployee) Then
                    MsgBox("Please re-calculate salary.", MsgBoxStyle.Information, "Information")
                    Me.BtnHitung.Focus()
                    Return False
                End If
            End If

            '1 bulan 1 transaksi salary
            If Me.Baru = True Then
                If Me.cboTaxCalculation.SelectedIndex <> 3 Then
                    If Me.Obj.GetJmlByTanggalDariSampai(Me.ObjCD.GetLasDateOfTheMonth(Me.DtpDate.Value), _
                    Me.ObjCD.GetFirstDateofTheMonth(Me.DtpDate.Value), Me.TxtPrimaryEmployee.Text) > 0 Then
                        MsgBox("There can be only  1 salary calculation with tax calculation for each month.", MsgBoxStyle.Information, "Information")
                        Me.DtpDate.Focus()
                        Return False
                    End If
                End If
            Else
                If Me.cboTaxCalculation.SelectedIndex <> 3 Then
                    If Me.Obj.GetJmlByTanggalDariSampai(Me.ObjCD.GetLasDateOfTheMonth(Me.DtpDate.Value), _
                    Me.ObjCD.GetFirstDateofTheMonth(Me.DtpDate.Value), Me.TxtPrimaryEmployee.Text, _
                    Me.TxtPrimary.Text) > 0 Then
                        MsgBox("There can be only  1 salary calculation with tax calculation for each month.", MsgBoxStyle.Information, "Information")
                        Me.DtpDate.Focus()
                        Return False
                    End If
                End If
            End If

            'Cek amount dengan original value
            If Me.dgvItem.Rows.Count > 0 Then
                For Each RowDGV In Me.dgvItem.Rows
                    If RowDGV.Cells(Me.ColOverride.Index).Value = True Then
                        If CDec(RowDGV.Cells(Me.ColAmount.Index).Value) = CDec(RowDGV.Cells(Me.ColOriginalValue.Index).Value) Then
                            RowDGV.Cells(Me.ColOverride.Index).Value = False
                        End If
                    End If
                Next
            End If


            '8. batasan bulan 12, cek bulan 1-11 metode tax calculation harus sama
            'If ObjMstSlry.ApakahSelama1TahunTaxCalculationSama(Me.DtpDate.Value.Date, Me.cboTaxCalculation.SelectedIndex, Me.TxtPrimaryEmployee.Text, -1) = False Then
            '    If Me.UserCode.ToUpper = "USER" Then
            '        If MsgBox("The tax calculation method is inconsistent. Continue process anyway?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
            '            Me.BtnHitung.Focus()
            '            Exit Function
            '        End If
            '    Else
            '        MsgBox("Please cek tax calculation because you must choose the same tax calculation for around one year. please, recalculate salary ", MsgBoxStyle.Information, "Information")
            '        Me.BtnHitung.Focus()
            '        Exit Function
            '    End If
            'End If


            ''Batasan IGI Payment, Cek Jumlah Payment harus sama dengan Take Home Pay
            'Dim HitungPaymentAmount As Decimal = 0
            'If Me.FrmPaymentData.DGVPayment.Rows.Count > 0 Then
            '    For Each RowDGV In Me.FrmPaymentData.DGVPayment.Rows
            '        HitungPaymentAmount += Convert.ToDecimal(RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_Amount.Index).Value)
            '    Next
            'End If
            'If CDec(Me.TxtSalaryReceived.Text) <> HitungPaymentAmount Then
            '    MsgBox("The following payment isn't suitable from take home pay value : " + vbNewLine + _
            '    "Salary no : " + Me.TxtDocumentNo.Text + ", take home pay : " + Me.TxtSalaryReceived.Text + _
            '    ", payment : " + Me.TxtPayment.Text + ".", MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            ''Batasan untuk IGI Payment, cek tidak boleh ada baris yang belum diinput akun-nya 
            'If Me.FrmPaymentData.DGVPayment.Rows.Count > 0 Then
            '    If Convert.ToDecimal(Me.TxtPayment.Text) <= 0 Then
            '        MsgBox("Please input payment data.", MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            '    For Each RowDGV In Me.FrmPaymentData.DGVPayment.Rows
            '        If RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_PrimaryAkun.Index).Value = "-1" Then
            '            MsgBox("Please input account code at payment data.", MsgBoxStyle.Information, "Information")
            '            Me.FrmPaymentData.NonAktif = True
            '            Me.FrmPaymentData.ShowDialog()
            '            Me.FrmPaymentData.NonAktif = False
            '            Return False
            '        End If
            '    Next
            'End If

            If Me.cboTaxCalculation.SelectedIndex <> 3 Then
                If Format(Me.DtpDate.Value, "MMM yyyy") <> Format(Me.DtpTo.Value.Date, "MMM yyyy") Then
                    'If Me.cboTaxCalculation.SelectedIndex <> 0 Then
                    MsgBox("Salary calculation period must be in the same month as salary end period.", MsgBoxStyle.Information, "Information")
                    Return False
                    'End If
                End If
            End If

            'cek periode pajak
            If Me.cboTaxCalculation.SelectedIndex <> 3 Then
                If Me.ObjKrywn.GetFirstTaxPeriod(Me.TxtPrimaryEmployee.Text).Date > Me.ObjCD.GetFirstDateofTheMonth(Me.DtpTo.Value).Date Then
                    MsgBox("You define first tax period on " + Format(Me.ObjKrywn.GetFirstTaxPeriod(Me.TxtPrimaryEmployee.Text), "MMM yyyy") + _
                    ", every salary calculation with tax must not be date before first tax period.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'sampe sini
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .NoTransaksi = Me.TxtDocumentNo.Text
                .TanggalTransaksi = Me.DtpDate.Value
                .PrimaryEmployee = Me.TxtPrimaryEmployee.Text
                .PrimaryDept = Me.TxtPrimaryDept.Text
                .PrimaryPosition = Me.TxtPrimaryPosition.Text
                .PrimarySalaryFormula = Me.TxtPrimarySalaryFormula.Text
                .StartDate = Me.dtpPeriod.Value
                .EndDate = Me.DtpTo.Value
                If Me.cboTaxCalculation.SelectedIndex = 0 Then
                    .Tax = ClsTransaksiSalary.EnumTax.Gross_Up
                ElseIf Me.cboTaxCalculation.SelectedIndex = 1 Then
                    .Tax = ClsTransaksiSalary.EnumTax.Netto
                ElseIf Me.cboTaxCalculation.SelectedIndex = 2 Then
                    .Tax = ClsTransaksiSalary.EnumTax.Gross
                ElseIf Me.cboTaxCalculation.SelectedIndex = 3 Then
                    .Tax = ClsTransaksiSalary.EnumTax.Manual_Tax_Calculation
                End If
                .PrimaryTaxExpense = Me.TxtPrimaryTaxExpAcc.Text
                .Income = Me.TxtIncome.Text
                .Deduction = Me.TxtDeduction.Text
                .LoanInstallment = Me.TxtLoanInstallment.Text
                .EmployeeExpense = Me.TxtEmployeeExpense.Text
                .CompanyExpense = Me.TxtCompanyExpense.Text
                .IncomeTax = Me.TxtIncomeTax.Text
                .SalaryReceived = Me.TxtSalaryReceived.Text
                If Me.CboPaymentStatus.SelectedIndex = 0 Then
                    .PaymentStatus = ClsTransaksiSalary.EnumPaymentStatus.unpaid
                Else
                    .PaymentStatus = ClsTransaksiSalary.EnumPaymentStatus.paid
                End If
                .PaymentDate = Me.DtpPaymentStatus.Value
                .Remarks = Me.TxtRemarks.Text
                .PrimaryLoanPayment = Me.TxtPrimaryLoan.Text
                .Adj_IncomeAdjust = Me.Txt_A_IncomeAdjust.Text
                .Adj_PKPValue = Me.Txt_A_PKPValue.Text
                .Adj_PositionExpense = Me.Txt_A_PositionExpense.Text
                .Adj_PositionExpenseAdjust = Me.Txt_A_PositionExpenseAdj.Text
                .Adj_PTKPMonthly = Me.Txt_A_PTKPMonthly.Text
                .Adj_PPhValue = Me.Txt_A_PPhValue.Text
                If Me.CboSalary.SelectedIndex = 0 Then
                    .Period_Paid = ClsTransaksiSalary.EnumPeriodPaid.One_Month
                ElseIf Me.CboSalary.SelectedIndex = 1 Then
                    .Period_Paid = ClsTransaksiSalary.EnumPeriodPaid.One_Week
                ElseIf Me.CboSalary.SelectedIndex = 2 Then
                    .Period_Paid = ClsTransaksiSalary.EnumPeriodPaid.Two_Weeks
                End If
                ._01Gaji = Me.Txt01Pjk_Gaji.Text
                ._01GajiRutin = Me.Txt01Pjk_GajiRutin.Text
                ._02TunjanganPPh = Me.Txt02PjkTunjanganPPh.Text
                ._02TunjanganPPhDisetahunkan = Me.Txt02Pjk_TunjPPHDisetahunkan.Text
                ._02TunjanganPPhDisetahunkanRutin = Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text
                ._02TunjanganPPhNonRutin = Me.Txt02Pjk_TunjPPHNonRutin.Text
                ._02TunjanganPPhRutin = Me.Txt02Pjk_TunjPPHRutin.Text
                ._03Tunjanganlainnya = Me.Txt03Pjk_TunjLainnya.Text
                ._03TunjanganlainnyaRutin = Me.Txt03Pjk_TunjLainnyaRutin.Text
                ._04Honorarium = Me.Txt04PjkHonorarium.Text
                ._04HonorariumRutin = Me.Txt04PjkHonorariumRutin.Text
                ._05PremiAsuransi = Me.Txt05PjkPremiAsuransi.Text
                ._05PremiAsuransiRutin = Me.Txt05PjkPremiAsuransiRutin.Text
                ._06Natura = Me.Txt06PjkNatura.Text
                ._06NaturaRutin = Me.Txt06PjkNaturaRutin.Text
                ._07Tantiem = Me.Txt07PjkTantiem.Text
                ._07TantiemRutin = Me.Txt07PjkTantiemRutin.Text
                ._08PenghasilanBruto = Me.Txt08PjkPenghasilanBruto.Text
                ._08PenghasilanBrutoRutin = Me.Txt08PjkPenghasilanBrutoRutin.Text
                ._09BiayaJabatan = Me.Txt09PjkBiayaJabatan.Text
                ._09BiayaJabatanRutin = Me.Txt09PjkBiayaJabatanRutin.Text
                ._10IuranPensiun = Me.Txt10PjkIuranPensiun.Text
                ._10IuranPensiunRutin = Me.Txt10PjkIuranPensiunRutin.Text
                ._11JumlahPengurangan = Me.Txt11PjkJmlPengurangan.Text
                ._11JumlahPenguranganRutin = Me.Txt11PjkJmlPenguranganRutin.Text
                ._12JumlahPenghasilanBruto = Me.Txt12PjkJmlPenghasilanNetto.Text
                ._13JumlahPenghasilanMsPjkSblm = Me.Txt13PjkJmlHasilMsPjkSblm.Text
                ._13JumlahPenghasilanMsPjkSblm_1 = Me.TxtPrevIncome.Text
                ._14JumlahPenghasilanNettoUtkPPh = Me.Txt14PjkJmlHasilNettoUPPh.Text
                ._14JumlahPenghasilanNettoUtkPPhRutin = Me.Txt14PjkJmlHasilNettoUPPhRutin.Text
                ._15PTKP = Me.Txt15Pjk_PTKP.Text
                ._16PKPSetahun = Me.Txt16Pjk_PKPSetahun.Text
                ._16PKPSetahunRutin = Me.Txt16Pjk_PKPSetahunRutin.Text
                ._17PPh21 = Me.Txt17Pjk_PPH21.Text
                ._17PPh21Rutin = Me.Txt17Pjk_PPH21Rutin.Text
                ._18PPh21YgDiPotongMsPjkSblm = Me.Txt18.Text
                ._18PPh21YgDiPotongMsPjkSblm_1 = Me.TxtPrevTaxPaid.Text
                ._19PPhTerhutang = Me.Txt19Pjk_PPH21Terutang.Text
                ._19PPhTerhutangSebulanNonRutin = Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text
                ._19PPhTerhutangSebulanRutin = Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text
                ._19PPhTerhutangSetahun = Me.Txt19Pjk_PPH21TerutangSetahun.Text
                ._19PPhTerhutangSetahunRutin = Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text
                ._20PPhPs21danPPh26Lunas = Me.Txt20.Text
                ._21Pesangon = Me.Txt21Pjk_Pesangon.Text
                ._21PPhFinalPesangon = Me.Txt21Pjk_PPhFinalPesangon.Text
                ._21TunjPPhFinalPesangon = Me.Txt21Pjk_TunjPPhPesangon.Text
                ._21TarifPPhFinalPesangon = Me.vTarifPPhFinal

                .A_NoForm1721A1 = Me.Txt_P_NoForm.Text
                .A_NPWP_Pemotongan = Me.Txt_P_NPWPPemotong.Text
                .A_Nama_Pemotongan = Me.Txt_P_NPWPPemotong.Text
                .A_NPWP = Me.Txt_P_A_NPWP.Text
                .A_NIK = Me.Txt_P_NoNIK.Text
                .A_Nama = Me.Txt_P_ANama.Text
                .A_Alamat1 = Me.Txt_P_Alamat1.Text
                .A_Alamat2 = Me.Txt_P_Alamat2.Text
                If Me.Txt_P_JenisKelamin.Text = "LAKI-LAKI" Then
                    .A_Jenis_Kelamin = ClsTransaksiSalary.EnumJenisKelamin.Pria
                ElseIf Me.Txt_P_JenisKelamin.Text = "PEREMPUAN" Then
                    .A_Jenis_Kelamin = ClsTransaksiSalary.EnumJenisKelamin.Wanita
                End If
                If Me.Txt_P_StatusNPWP.Text = "TK0" Then
                    .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.TK0
                ElseIf Me.Txt_P_StatusNPWP.Text = "TK1" Then
                    .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.TK1
                ElseIf Me.Txt_P_StatusNPWP.Text = "TK2" Then
                    .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.TK2
                ElseIf Me.Txt_P_StatusNPWP.Text = "TK3" Then
                    .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.TK3
                ElseIf Me.Txt_P_StatusNPWP.Text = "K0" Then
                    .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.K0
                ElseIf Me.Txt_P_StatusNPWP.Text = "K1" Then
                    .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.K1
                ElseIf Me.Txt_P_StatusNPWP.Text = "K2" Then
                    .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.K2
                ElseIf Me.Txt_P_StatusNPWP.Text = "K3" Then
                    .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.K3
                End If
                .A_Nama_Jabatan = Me.Txt_P_NamaJabatan.Text
                .A_Karyawan_Asing = Me.Txt_P_KaryawanAsing.Text
                .A_Kode_Negara_Domisili = Me.Txt_P_KodeNegaraDomisili.Text
                .C_NPWP = Me.Txt_P_NPWP.Text
                .C_Nama = Me.Txt_P_Nama.Text
                If Me.Baru = True Then
                    .PrimaryPayment = -1
                Else
                    .PrimaryPayment = Me.ObjPayment.GetPrimaryPayment_ByPrimarySalary(Me.TxtPrimary.Text)
                End If
                .GrossSalary = Me.TxtGrossSalary.Text
                .OtherDeduction = Me.TxtOtherDedution.Text
                If Me.DtpDate.Value.Month <> 12 Then
                    .KurangBayarTax = 0
                Else
                    .KurangBayarTax = Me.TxtKurangBayarTax.Text
                End If

                .PeriodCount = Me.TxtPeriodCount.Text
                .PeriodEnd = Me.TxtPeriodEnd.Text
                .PeriodStart = Me.TxtPeriodStart.Text
                .IsResign = Me.TxtIsResign.Text
                .LosesTaxObligation = Me.TxtLosesTaxObligation.Text
                .PrimaryDept = Me.TxtPrimaryDept.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function HitungTakeHomePay(ByVal vIncome As Decimal, ByVal vDeduction As Decimal, ByVal vIncomeTax As Decimal, ByVal vLoanInstallment As Decimal, _
    ByVal vOtherDeduction As Decimal, ByVal vKurangBayarTax As Decimal, ByVal vEmployeeExpense As Decimal) As Decimal
        Dim Retval As Decimal = 0
        Try
            If Me.cboTaxCalculation.SelectedIndex = 0 Or Me.cboTaxCalculation.SelectedIndex = 1 Then
                Retval = vIncome - vDeduction - vEmployeeExpense - vLoanInstallment - vOtherDeduction
            Else
                If Me.DtpDate.Value.Month <> 12 Then
                    Retval = vIncome - vDeduction - vEmployeeExpense - vIncomeTax - vLoanInstallment - vOtherDeduction
                Else
                    Retval = vIncome - vDeduction - vEmployeeExpense - vKurangBayarTax - vLoanInstallment - vOtherDeduction
                End If
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function HitungGrossSalary(ByVal vIncome As Decimal, ByVal vDeduction As Decimal, ByVal vEmployeeExpense As Decimal) As Decimal
        Dim Retval As Decimal = 0
        Try
            Retval = vIncome - vDeduction - vEmployeeExpense
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeField()
        Dim rowOther As DseTransaksiSalary.tbltrans_salary_item_other1Row
        Dim JmlOther As Decimal = 0

        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.DtpDate.Value = .TanggalTransaksi
                Me.TxtDocumentNo.Text = .NoTransaksi
                Me.TxtPrimaryEmployee.Text = .PrimaryEmployee
                Me.TxtPrimaryDept.Text = .PrimaryDept
                Me.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(.PrimaryEmployee)
                Me.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(.PrimaryEmployee)
                Me.TxtDepartment.Text = Me.ObjKrywn.GetDeptDescription(Me.TxtPrimaryEmployee.Text)
                Me.TxtPrimaryPosition.Text = .PrimaryPosition
                Me.TxtPosition.Text = Me.ObjPostion.GetDesc(Me.TxtPrimaryPosition.Text)
                Me.dtpPeriod.Value = .StartDate
                Me.DtpTo.Value = .EndDate
                If .Tax = ClsTransaksiSalary.EnumTax.Gross_Up Then
                    Me.cboTaxCalculation.SelectedIndex = 0
                ElseIf .Tax = ClsTransaksiSalary.EnumTax.Netto Then
                    Me.cboTaxCalculation.SelectedIndex = 1
                ElseIf .Tax = ClsTransaksiSalary.EnumTax.Gross Then
                    Me.cboTaxCalculation.SelectedIndex = 2
                ElseIf .Tax = ClsTransaksiSalary.EnumTax.Manual_Tax_Calculation Then
                    Me.cboTaxCalculation.SelectedIndex = 3
                End If
                Me.TxtPrimaryTaxExpAcc.Text = .PrimaryTaxExpense
                Me.TxtExpAcc.Text = Me.ObjAkun.GetAccountCode(.PrimaryTaxExpense)
                Me.TxtExpAccDesc.Text = Me.ObjAkun.GetDescription(Me.TxtExpAcc.Text)
                Me.TxtIncome.Text = Me.ObjCD.DigitGrouping(.Income, Me.Def_Prec)
                Me.TxtDeduction.Text = Me.ObjCD.DigitGrouping(.Deduction, Me.Def_Prec)
                Me.TxtLoanInstallment.Text = Me.ObjCD.DigitGrouping(.LoanInstallment, Me.Def_Prec)
                Me.TxtCompanyExpense.Text = Me.ObjCD.DigitGrouping(.CompanyExpense, Me.Def_Prec)
                Me.TxtEmployeeExpense.Text = Me.ObjCD.DigitGrouping(.EmployeeExpense, Me.Def_Prec)
                Me.TxtIncomeTax.Text = Me.ObjCD.DigitGrouping(.IncomeTax, Me.Def_Prec)
                Me.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(.SalaryReceived, Me.Def_Prec)
                If .PaymentStatus = ClsTransaksiSalary.EnumPaymentStatus.unpaid Then
                    Me.CboPaymentStatus.SelectedIndex = 0
                Else
                    Me.CboPaymentStatus.SelectedIndex = 1
                End If
                Me.TxtRemarks.Text = .Remarks
                Me.TxtPrimaryLoan.Text = .PrimaryLoanPayment
                Me.TxtPrimarySalaryFormula.Text = .PrimarySalaryFormula

                Me.Txt_A_IncomeAdjust.Text = .Adj_IncomeAdjust
                Me.Txt_A_PKPValue.Text = .Adj_PKPValue
                Me.Txt_A_PositionExpense.Text = .Adj_PositionExpense
                Me.Txt_A_PositionExpenseAdj.Text = .Adj_PositionExpenseAdjust
                Me.Txt_A_PTKPMonthly.Text = .Adj_PTKPMonthly
                Me.Txt_A_PPhValue.Text = .Adj_PPhValue
                Me.DummyDari = .StartDate
                Me.DummySampai = .EndDate

                If .Period_Paid = ClsTransaksiSalary.EnumPeriodPaid.One_Month Then
                    Me.CboSalary.SelectedIndex = 0
                ElseIf .Period_Paid = ClsTransaksiSalary.EnumPeriodPaid.One_Week Then
                    Me.CboSalary.SelectedIndex = 1
                ElseIf .Period_Paid = ClsTransaksiSalary.EnumPeriodPaid.Two_Weeks Then
                    Me.CboSalary.SelectedIndex = 2
                End If

                '----------------------------------------------
                Me.Txt01Pjk_Gaji.Text = ._01Gaji
                Me.Txt01Pjk_GajiRutin.Text = ._01GajiRutin
                Me.Txt02PjkTunjanganPPh.Text = ._02TunjanganPPh
                Me.Txt02Pjk_TunjPPHDisetahunkan.Text = ._02TunjanganPPhDisetahunkan
                Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text = ._02TunjanganPPhDisetahunkanRutin
                Me.Txt02Pjk_TunjPPHNonRutin.Text = ._02TunjanganPPhNonRutin
                Me.Txt02Pjk_TunjPPHRutin.Text = ._02TunjanganPPhRutin
                Me.Txt03Pjk_TunjLainnya.Text = ._03Tunjanganlainnya
                Me.Txt03Pjk_TunjLainnyaRutin.Text = ._03TunjanganlainnyaRutin
                Me.Txt04PjkHonorarium.Text = ._04Honorarium
                Me.Txt04PjkHonorariumRutin.Text = ._04HonorariumRutin
                Me.Txt05PjkPremiAsuransi.Text = ._05PremiAsuransi
                Me.Txt05PjkPremiAsuransiRutin.Text = ._05PremiAsuransiRutin
                Me.Txt06PjkNatura.Text = ._06Natura
                Me.Txt06PjkNaturaRutin.Text = ._06NaturaRutin
                Me.Txt07PjkTantiem.Text = ._07Tantiem
                Me.Txt07PjkTantiemRutin.Text = ._07TantiemRutin
                Me.Txt08PjkPenghasilanBruto.Text = ._08PenghasilanBruto
                Me.Txt08PjkPenghasilanBrutoRutin.Text = ._08PenghasilanBrutoRutin
                Me.Txt09PjkBiayaJabatan.Text = ._09BiayaJabatan
                Me.Txt09PjkBiayaJabatanRutin.Text = ._09BiayaJabatanRutin
                Me.Txt10PjkIuranPensiun.Text = ._10IuranPensiun
                Me.Txt10PjkIuranPensiunRutin.Text = ._10IuranPensiunRutin
                Me.Txt11PjkJmlPengurangan.Text = ._11JumlahPengurangan
                Me.Txt11PjkJmlPenguranganRutin.Text = ._11JumlahPenguranganRutin
                Me.Txt12PjkJmlPenghasilanNetto.Text = ._12JumlahPenghasilanBruto
                Me.Txt13PjkJmlHasilMsPjkSblm.Text = ._13JumlahPenghasilanMsPjkSblm
                Me.TxtPrevIncome.Text = Me.ObjCD.DigitGrouping(._13JumlahPenghasilanMsPjkSblm_1, 0)
                Me.Txt14PjkJmlHasilNettoUPPh.Text = ._14JumlahPenghasilanNettoUtkPPh
                Me.Txt14PjkJmlHasilNettoUPPhRutin.Text = ._14JumlahPenghasilanNettoUtkPPhRutin
                Me.Txt15Pjk_PTKP.Text = ._15PTKP
                Me.Txt16Pjk_PKPSetahun.Text = ._16PKPSetahun
                Me.Txt16Pjk_PKPSetahunRutin.Text = ._16PKPSetahunRutin
                Me.Txt17Pjk_PPH21.Text = ._17PPh21
                Me.Txt17Pjk_PPH21Rutin.Text = ._17PPh21Rutin
                Me.Txt18.Text = ._18PPh21YgDiPotongMsPjkSblm
                Me.TxtPrevTaxPaid.Text = Me.ObjCD.DigitGrouping(._18PPh21YgDiPotongMsPjkSblm_1, 0)
                Me.Txt19Pjk_PPH21Terutang.Text = ._19PPhTerhutang
                Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text = ._19PPhTerhutangSebulanNonRutin
                Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text = ._19PPhTerhutangSebulanRutin
                Me.Txt19Pjk_PPH21TerutangSetahun.Text = ._19PPhTerhutangSetahun
                Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text = ._19PPhTerhutangSetahunRutin
                Me.Txt20.Text = ._20PPhPs21danPPh26Lunas
                Me.Txt21Pjk_Pesangon.Text = ._21Pesangon
                Me.Txt21Pjk_PPhFinalPesangon.Text = ._21PPhFinalPesangon
                Me.Txt21Pjk_TunjPPhPesangon.Text = ._21TunjPPhFinalPesangon
                Me.vTarifPPhFinal = ._21TarifPPhFinalPesangon

                Me.Txt_P_NoForm.Text = .A_NoForm1721A1
                Me.Txt_P_NPWPPemotong.Text = .A_NPWP_Pemotongan
                Me.Txt_P_NPWPPemotong.Text = .A_Nama_Pemotongan
                Me.Txt_P_A_NPWP.Text = .A_NPWP
                Me.Txt_P_NoNIK.Text = .A_NIK
                Me.Txt_P_ANama.Text = .A_Nama
                Me.Txt_P_Alamat1.Text = .A_Alamat1
                Me.Txt_P_Alamat2.Text = .A_Alamat2
                If .A_Jenis_Kelamin = ClsTransaksiSalary.EnumJenisKelamin.Pria Then
                    Me.Txt_P_JenisKelamin.Text = "LAKI-LAKI"
                ElseIf .A_Jenis_Kelamin = ClsTransaksiSalary.EnumJenisKelamin.Wanita Then
                    Me.Txt_P_JenisKelamin.Text = "PEREMPUAN"
                End If
                If .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.TK0 Then
                    Me.Txt_P_StatusNPWP.Text = "TK0"
                ElseIf .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.TK1 Then
                    Me.Txt_P_StatusNPWP.Text = "TK1"
                ElseIf .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.TK2 Then
                    Me.Txt_P_StatusNPWP.Text = "TK2"
                ElseIf .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.TK3 Then
                    Me.Txt_P_StatusNPWP.Text = "TK3"
                ElseIf .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.K0 Then
                    Me.Txt_P_StatusNPWP.Text = "K0"
                ElseIf .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.K1 Then
                    Me.Txt_P_StatusNPWP.Text = "K1"
                ElseIf .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.K2 Then
                    Me.Txt_P_StatusNPWP.Text = "K2"
                ElseIf .A_Status_NPWP = ClsMasterKaryawan.GolonganPajak.K3 Then
                    Me.Txt_P_StatusNPWP.Text = "K3"
                End If
                Me.Txt_P_NamaJabatan.Text = .A_Nama_Jabatan
                Me.Txt_P_KaryawanAsing.Text = .A_Karyawan_Asing
                Me.Txt_P_KodeNegaraDomisili.Text = .A_Kode_Negara_Domisili
                Me.Txt_P_NPWP.Text = .C_NPWP
                Me.Txt_P_Nama.Text = .C_Nama
                Me.TxtGrossSalary.Text = Me.ObjCD.DigitGrouping(Me.HitungGrossSalary(Me.TxtIncome.Text, Me.TxtDeduction.Text, Me.TxtEmployeeExpense.Text), Me.Def_Prec)
                Me.TxtKurangBayarTax.Text = Me.ObjCD.DigitGrouping(.KurangBayarTax, Me.Def_Prec)
                Me.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(Me.HitungTakeHomePay(Me.TxtIncome.Text, Me.TxtDeduction.Text, _
                Me.TxtIncomeTax.Text, Me.TxtLoanInstallment.Text, Me.TxtOtherDedution.Text, Me.TxtKurangBayarTax.Text, Me.TxtEmployeeExpense.Text), Me.Def_Prec)
                '----------------------------------------------
                Me.BufferCBOPeriod = Me.CboSalary.SelectedIndex
                Me.DummyCBOPeriod = Me.CboSalary.SelectedIndex
                Me.BufferMetodeTax = Me.cboTaxCalculation.SelectedIndex
                Me.DummyMetodeTax = Me.cboTaxCalculation.SelectedIndex
                Me.BufferPrimaryEmployee = Me.TxtPrimaryEmployee.Text
                Me.DummyPrimaryEmployee = Me.TxtPrimaryEmployee.Text

                If Me.DtpDate.Value.Date.Month <> 12 Then
                    Me.TxtIncomeTax.Visible = True
                    Me.TxtKurangBayarTax.Visible = False
                Else
                    Me.TxtIncomeTax.Visible = False
                    Me.TxtKurangBayarTax.Visible = True
                End If
                'Status paid visible
                If Me.LblPaid.Text = "Paid" Then
                    Me.CboPaymentStatus.Visible = True
                    Me.DtpPaymentStatus.Visible = True
                    Me.LblPaid.Visible = False
                    Me.LblPaymentStatus.Visible = True
                Else
                    Me.CboPaymentStatus.Visible = False
                    Me.DtpPaymentStatus.Visible = False
                    Me.LblPaid.Visible = True
                    Me.LblPaymentStatus.Visible = False
                End If

                For Each rowOther In Me.vTBLOtherPayment.Rows
                    JmlOther += rowOther.Amount
                Next
                Me.TxtPeriodCount.Text = .PeriodCount
                Me.TxtPeriodEnd.Text = .PeriodEnd
                Me.TxtPeriodStart.Text = .PeriodStart
                Me.TxtIsResign.Text = .IsResign
                Me.TxtLosesTaxObligation.Text = .LosesTaxObligation
                Me.TxtPrimaryDept.Text = .PrimaryDept

                Me.TxtOtherPayment.Text = Me.ObjCD.DigitGrouping(JmlOther, Me.Def_Prec)
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.Baru = True Then
                Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                If Me.ijK.TxtRemark.Enabled = True Then
                    Me.ijK.TxtDocumentNo.Text = Me.ijK.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                    Me.ijK.dtpDate.Value = Me.DtpDate.Value
                End If
            Else
                If Me.DtpDate.Value.Month = Me.Obj.GetTanggalTransaksi(Me.TxtPrimary.Text).Month And _
                Me.DtpDate.Value.Year = Obj.GetTanggalTransaksi(Me.TxtPrimary.Text).Year Then
                    Me.TxtDocumentNo.Text = Me.Obj.GetNoTransaksi(Me.TxtPrimary.Text)
                    If Me.ijK.TxtRemark.Enabled = True Then
                        Me.ijK.TxtDocumentNo.Text = Me.ijK.Obj.GetNoTransaksi(Me.ijK.TxtPrimaryLoanPayment.Text)
                        Me.ijK.dtpDate.Value = Me.DtpDate.Value
                    End If
                Else
                    If Me.DtpDate.Value.Month <> Me.Obj.TanggalTransaksi.Month Or _
                    Me.DtpDate.Value.Year <> Me.Obj.TanggalTransaksi.Year Then
                        Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                        If Me.ijK.TxtRemark.Enabled = True Then
                            Me.ijK.TxtDocumentNo.Text = Me.ijK.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                            Me.ijK.dtpDate.Value = Me.DtpDate.Value
                        End If
                    Else
                        Me.TxtDocumentNo.Text = Me.Obj.GetNoTransaksi(Me.TxtPrimary.Text)
                        If Me.ijK.TxtRemark.Enabled = True Then
                            Me.ijK.TxtDocumentNo.Text = Me.ijK.Obj.GetNoTransaksi(Me.ijK.TxtPrimaryLoanPayment.Text)
                            Me.ijK.dtpDate.Value = Me.DtpDate.Value
                        End If
                    End If
                End If
            End If
            If Me.DtpDate.Value.Month <> 12 Then
                Me.TxtIncomeTax.Visible = True
                Me.TxtKurangBayarTax.Visible = False
            Else
                Me.TxtIncomeTax.Visible = False
                Me.TxtKurangBayarTax.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtEmployeeCode1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode1.Enter
        Try
            Me.BufferPrimaryEmployee = Me.TxtPrimaryEmployee.Text
            Me.Buffer = Me.TxtEmployeeCode1.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtEmployeeCode1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode1.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtEmployeeCode1.Enabled = True Then
                    If Me.TxtEmployeeCode1.ReadOnly = False Then
                        Me.TxtEmployeeCode1.BackColor = Me.BackColorAll
                    Else
                        Me.TxtEmployeeCode1.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtEmployeeCode1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode1.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtEmployeeCode1.Enabled = True Then
                    If Me.TxtEmployeeCode1.ReadOnly = False Then
                        Me.TxtEmployeeCode1.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtEmployeeCode1.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExpAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExpAcc.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtExpAcc.Enabled = True Then
                    If Me.TxtExpAcc.ReadOnly = False Then
                        Me.TxtExpAcc.BackColor = Me.BackColorAll
                    Else
                        Me.TxtExpAcc.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExpAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExpAcc.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtExpAcc.Enabled = True Then
                    If Me.TxtExpAcc.ReadOnly = False Then
                        Me.TxtExpAcc.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtExpAcc.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.Enabled = True Then
                    If Me.TxtRemarks.ReadOnly = False Then
                        Me.TxtRemarks.BackColor = Me.BackColorAll
                    Else
                        Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.Enabled = True Then
                    If Me.TxtRemarks.ReadOnly = False Then
                        Me.TxtRemarks.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BersihkanTBL(ByVal vTBLParam1 As DseMasterParameter.DataTable1DataTable, _
    ByVal vTBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, _
    ByVal vTBLErrorFormula As DseTransaksiSalary.TblErrorFormulaDataTable)
        Try
            vTBLParam1.Rows.Clear()
            vTBLParam2.Rows.Clear()
            TBLWeekDay.Rows.Clear()
            TBLHoliday.Rows.Clear()
            TblKaryawan.Rows.Clear()
            vTBLErrorFormula.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Employee Code
    Public Sub FungsiBtnFindEmployee()
        Dim i As New Core_Module.FrmMasterKaryawanCari
        i.BackColorAll = Me.BackColorAll
        i.ShowSalary = True
        i.USERCode = Me.UserCode
        Dim Primary_Lama As Double = -1
        Try
            If Me.NewQuick_Data = False Then
                Primary_Lama = Me.TxtPrimaryEmployee.Text
                If Me.TxtEmployeeCode1.Text <> "" Then
                    i.txtCari.Text = Me.TxtEmployeeCode1.Text
                End If

                i.ShowSalary = True
                i.ShowDialog()
                If i.PrimaryEmploye > 0 Then
                    Me.TxtPrimaryEmployee.Text = i.PrimaryEmploye
                    Me.TxtPrimaryDept.Text = Me.ObjKrywn.GetPrimaryDept(Me.TxtPrimaryEmployee.Text) 
                    Me.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtPosition.Text = Me.ObjPostion.GetDesc(Me.ObjKrywn.GetJobByPrimary(Me.TxtPrimaryEmployee.Text))
                    Me.TxtPrimaryPosition.Text = Me.ObjKrywn.GetPrimaryJobPosition(Me.TxtPrimaryEmployee.Text)
                    Me.TxtDepartment.Text = Me.ObjKrywn.GetDeptDescription(Me.TxtPrimaryEmployee.Text)
                    Me.TxtPrimarySalaryFormula.Text = Me.Obj.GetPrimaryMasterSalaryFormula(i.PrimaryEmploye)
                    Me.CboSalary.SelectedIndex = Me.Obj.MenentukanNilaiAwal(ClsTransaksiSalary.EnumNilaiAwal.Periode_Hitung, Me.TxtPrimaryEmployee.Text)
                    Me.cboTaxCalculation.SelectedIndex = Me.Obj.MenentukanNilaiAwal(ClsTransaksiSalary.EnumNilaiAwal.Tax_Calculation, Me.TxtPrimaryEmployee.Text)
                    If Me.cboTaxCalculation.SelectedIndex = -1 Then
                        Me.cboTaxCalculation.SelectedIndex = Me.Obj.GetTaxCalculationDefault(i.PrimaryEmploye, Me.TxtPrimarySalaryFormula.Text)
                    End If
                    Me.Obj.IsiKeteranganKolomTabPage3(Me)
                    Me.dtpPeriod.Focus()
                    If Me.dgvItem.Rows.Count > 0 Then
                        Me.dgvItem.Rows.Clear()

                        Me.TxtIncome.Text = "0"
                        Me.TxtDeduction.Text = "0"
                        Me.TxtLoanInstallment.Text = "0"
                        Me.TxtIncomeTax.Text = "0"
                        Me.TxtEmployeeExpense.Text = "0"
                        Me.TxtCompanyExpense.Text = "0"
                        Me.TxtSalaryReceived.Text = "0"
                        Me.TxtKurangBayarTax.Text = "0"
                        Me.Txt_A_IncomeAdjust.Text = "0"
                        Me.Txt_A_PKPValue.Text = "0"
                        Me.Txt_A_PositionExpense.Text = "0"
                        Me.Txt_A_PositionExpenseAdj.Text = "0"
                        Me.Txt_A_PTKPMonthly.Text = "0"
                        Me.Txt_A_PPhValue.Text = "0"
                        Me.TxtOtherDedution.Text = "0"
                        Me.TxtGrossSalary.Text = "0"
                        Me.TxtPayment.Text = "0"
                    End If
                Else
                    Me.TxtPrimaryEmployee.Text = "-1"
                    Me.TxtPrimaryDept.Text = "-1"
                    Me.TxtPrimarySalaryFormula.Text = "-1"
                    Me.CboSalary.SelectedIndex = 0
                    Me.cboTaxCalculation.SelectedIndex = 0
                    Me.TxtEmployeeCode1.Focus()
                    Me.TxtEmployeeName.Clear()
                    Me.TxtPosition.Clear()
                End If
            ElseIf Me.NewQuick_Data = True Then
                'buat new quick
                If Me.TxtPrimaryEmployee.Text <> "-1" Then
                    Me.TxtPrimaryDept.Text = Me.ObjKrywn.GetPrimaryDept(Me.TxtPrimaryEmployee.Text) 
                    Me.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtPosition.Text = Me.ObjPostion.GetCode(Me.ObjKrywn.GetJobByPrimary(Me.TxtPrimaryEmployee.Text))
                    Me.TxtPrimaryPosition.Text = Me.ObjKrywn.GetPrimaryJobPosition(Me.TxtPrimaryEmployee.Text)
                    Me.TxtPrimarySalaryFormula.Text = Me.Obj.GetPrimaryMasterSalaryFormula(Me.TxtPrimaryEmployee.Text)
                    Me.CboSalary.SelectedIndex = Me.Obj.MenentukanNilaiAwal(ClsTransaksiSalary.EnumNilaiAwal.Periode_Hitung, Me.TxtPrimaryEmployee.Text)
                    Me.cboTaxCalculation.SelectedIndex = Me.Obj.MenentukanNilaiAwal(ClsTransaksiSalary.EnumNilaiAwal.Tax_Calculation, Me.TxtPrimaryEmployee.Text)
                    If Me.cboTaxCalculation.SelectedIndex = -1 Then
                        Me.cboTaxCalculation.SelectedIndex = Me.Obj.GetTaxCalculationDefault(Me.TxtPrimaryEmployee.Text, Me.TxtPrimarySalaryFormula.Text)
                    End If
                    Me.Obj.IsiKeteranganKolomTabPage3(Me)
                    Me.dtpPeriod.Focus()
                Else
                    Me.TxtPrimaryEmployee.Text = "-1"
                    Me.TxtPrimaryDept.Text = "-1"
                    Me.TxtPrimarySalaryFormula.Text = "-1"
                    Me.CboSalary.SelectedIndex = 0
                    Me.cboTaxCalculation.SelectedIndex = 0
                    Me.TxtEmployeeCode1.Focus()
                    Me.TxtEmployeeName.Clear()
                    Me.TxtPosition.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnFindEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindEmployeeCode.Click
        Try
            Me.FungsiBtnFindEmployee()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtEmployeeCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtEmployeeCode1.KeyDown
        Dim objCode As New ClsMasterKaryawan
        Dim i As New FrmMasterKaryawanCari
        i.BackColorAll = Me.BackColorAll
        i.ShowSalary = True
        i.USERCode = Me.UserCode
        Dim Primary_Lama As Double = -1
        Try
            If e.KeyValue = 13 Then
                If Me.TxtEmployeeCode1.Text <> "" Then
                    Primary_Lama = Me.TxtPrimaryEmployee.Text
                    If Me.ObjKrywn.ApakahSudahAda_ShowSalary(Me.TxtEmployeeCode1.Text, Me.UserCode) = True Then
                        Me.TxtPrimaryEmployee.Text = Me.ObjKrywn.GetPrimaryByCode(Me.TxtEmployeeCode1.Text)
                        Me.TxtPrimaryDept.Text = Me.ObjKrywn.GetPrimaryDept(Me.TxtPrimaryEmployee.Text)
                        Me.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                        Me.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                        Me.TxtPosition.Text = Me.ObjPostion.GetCode(Me.ObjKrywn.GetJobByPrimary(Me.TxtPrimaryEmployee.Text))
                        Me.TxtPrimarySalaryFormula.Text = Me.Obj.GetPrimaryMasterSalaryFormula(Me.TxtPrimaryEmployee.Text)
                        Me.CboSalary.SelectedIndex = Me.Obj.MenentukanNilaiAwal(ClsTransaksiSalary.EnumNilaiAwal.Periode_Hitung, Me.TxtPrimaryEmployee.Text)
                        Me.cboTaxCalculation.SelectedIndex = Me.Obj.MenentukanNilaiAwal(ClsTransaksiSalary.EnumNilaiAwal.Tax_Calculation, Me.TxtPrimaryEmployee.Text)
                        If Me.cboTaxCalculation.SelectedIndex = -1 Then
                            Me.cboTaxCalculation.SelectedIndex = Me.Obj.GetTaxCalculationDefault(Me.TxtPrimaryEmployee.Text, Me.TxtPrimarySalaryFormula.Text)
                        End If
                        Me.Obj.IsiKeteranganKolomTabPage3(Me)
                        Me.dtpPeriod.Focus()
                    Else
                        Me.BtnFindEmployeeCode.PerformClick()
                    End If
                Else
                    Me.BtnFindEmployeeCode.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtEmployeeCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode1.Leave
        Dim i As New FrmMasterKaryawanCari
        Dim objCode As New ClsMasterKaryawan
        i.BackColorAll = Me.BackColorAll
        i.ShowSalary = True
        i.USERCode = Me.UserCode
        Dim Primary_Lama As Double = -1

        Try
            If Me.Buffer = Me.TxtEmployeeCode1.Text Then
                Exit Sub
            End If
            If Me.TxtEmployeeCode1.Text <> "" Then
                If Me.ObjKrywn.ApakahSudahAda_ShowSalary(Me.TxtEmployeeCode1.Text, Me.UserCode) = True Then
                    Primary_Lama = Me.TxtPrimaryEmployee.Text
                    Me.TxtPrimaryEmployee.Text = Me.ObjKrywn.GetPrimaryByCode(Me.TxtEmployeeCode1.Text)
                    Me.TxtPrimaryDept.Text = Me.ObjKrywn.GetPrimaryDept(Me.TxtPrimaryEmployee.Text)
                    Me.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtPosition.Text = Me.ObjPostion.GetCode(Me.ObjKrywn.GetJobByPrimary(Me.TxtPrimaryEmployee.Text))
                    Me.TxtPrimarySalaryFormula.Text = Me.Obj.GetPrimaryMasterSalaryFormula(Me.TxtPrimaryEmployee.Text)
                    Me.CboSalary.SelectedIndex = Me.Obj.MenentukanNilaiAwal(ClsTransaksiSalary.EnumNilaiAwal.Periode_Hitung, Me.TxtPrimaryEmployee.Text)
                    Me.cboTaxCalculation.SelectedIndex = Me.Obj.MenentukanNilaiAwal(ClsTransaksiSalary.EnumNilaiAwal.Tax_Calculation, Me.TxtPrimaryEmployee.Text)
                    If Me.cboTaxCalculation.SelectedIndex = -1 Then
                        Me.cboTaxCalculation.SelectedIndex = Me.Obj.GetTaxCalculationDefault(Me.TxtPrimaryEmployee.Text, Me.TxtPrimarySalaryFormula.Text)
                    End If
                    Me.Obj.IsiKeteranganKolomTabPage3(Me)
                    Me.TxtPrimaryPosition.Text = Me.ObjKrywn.GetPrimaryJobPosition(Me.TxtPrimaryEmployee.Text)
                Else
                    Primary_Lama = Me.TxtPrimaryEmployee.Text
                    If Me.TxtEmployeeCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtEmployeeCode1.Text
                    End If

                    i.ShowSalary = True
                    i.ShowDialog()
                    If i.PrimaryEmploye > 0 Then
                        Me.TxtPrimaryEmployee.Text = i.PrimaryEmploye
                        Me.TxtPrimaryDept.Text = Me.ObjKrywn.GetPrimaryDept(Me.TxtPrimaryEmployee.Text)
                        Me.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                        Me.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                        Me.TxtPosition.Text = Me.ObjPostion.GetCode(Me.ObjKrywn.GetJobByPrimary(Me.TxtPrimaryEmployee.Text))
                        Me.TxtPrimaryPosition.Text = Me.ObjKrywn.GetPrimaryJobPosition(Me.TxtPrimaryEmployee.Text)
                        Me.TxtPrimarySalaryFormula.Text = Me.Obj.GetPrimaryMasterSalaryFormula(Me.TxtPrimaryEmployee.Text)
                        Me.CboSalary.SelectedIndex = Me.Obj.MenentukanNilaiAwal(ClsTransaksiSalary.EnumNilaiAwal.Periode_Hitung, Me.TxtPrimaryEmployee.Text)
                        Me.cboTaxCalculation.SelectedIndex = Me.Obj.MenentukanNilaiAwal(ClsTransaksiSalary.EnumNilaiAwal.Tax_Calculation, Me.TxtPrimaryEmployee.Text)
                        If Me.cboTaxCalculation.SelectedIndex = -1 Then
                            Me.cboTaxCalculation.SelectedIndex = Me.Obj.GetTaxCalculationDefault(Me.TxtPrimaryEmployee.Text, Me.TxtPrimarySalaryFormula.Text)
                        End If
                        Me.Obj.IsiKeteranganKolomTabPage3(Me)
                    Else
                        Me.TxtPrimaryEmployee.Text = "-1"
                        Me.TxtPrimaryDept.Text = "-1"
                        Me.TxtPrimarySalaryFormula.Text = "-1"
                        Me.TxtPrimaryPosition.Text = "-1"
                        Me.CboSalary.SelectedIndex = 0
                        Me.cboTaxCalculation.SelectedIndex = 0
                        Me.TxtEmployeeCode1.Clear()
                        Me.TxtEmployeeName.Clear()
                        Me.TxtPosition.Clear()
                    End If
                End If
            Else
                Me.TxtEmployeeCode1.Clear()
                Me.TxtPrimaryEmployee.Text = "-1"
                Me.TxtPrimaryDept.Text = "-1"
                Me.TxtPrimarySalaryFormula.Text = "-1"
                Me.TxtPrimaryPosition.Text = "-1"
                Me.CboSalary.SelectedIndex = 0
                Me.cboTaxCalculation.SelectedIndex = 0
                Me.TxtEmployeeName.Clear()
                Me.TxtPosition.Clear()
            End If

            If Me.Buffer <> Me.TxtEmployeeCode1.Text Or Me.TxtEmployeeCode1.Text = "" Then
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.dgvItem.Rows.Clear()

                    Me.TxtIncome.Text = "0"
                    Me.TxtDeduction.Text = "0"
                    Me.TxtLoanInstallment.Text = "0"
                    Me.TxtIncomeTax.Text = "0"
                    Me.TxtEmployeeExpense.Text = "0"
                    Me.TxtCompanyExpense.Text = "0"
                    Me.TxtSalaryReceived.Text = "0"
                    Me.TxtKurangBayarTax.Text = "0"
                    Me.Txt_A_IncomeAdjust.Text = "0"
                    Me.Txt_A_PKPValue.Text = "0"
                    Me.Txt_A_PositionExpense.Text = "0"
                    Me.Txt_A_PositionExpenseAdj.Text = "0"
                    Me.Txt_A_PTKPMonthly.Text = "0"
                    Me.Txt_A_PPhValue.Text = "0"
                    Me.TxtOtherDedution.Text = "0"
                    Me.TxtGrossSalary.Text = "0"
                    Me.TxtPayment.Text = "0"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Public Sub FungsiBtnFindTaxAcc()
        Dim i As New AccModule.FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.NewQuick_Data = False Then
                If Me.TxtExpAcc.Text <> "" Then
                    i.txtCari.Text = Me.TxtExpAcc.Text
                End If

                i.ShowDialog()
                If i.mPrimaryAkun > 0 Then
                    Me.TxtPrimaryTaxExpAcc.Text = i.mPrimaryAkun
                    Me.TxtExpAcc.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryTaxExpAcc.Text)
                    Me.TxtExpAccDesc.Text = Me.ObjAkun.GetDescription(Me.TxtExpAcc.Text)
                    Me.BtnHitung.Focus()
                Else
                    Me.TxtExpAcc.Clear()
                    Me.TxtPrimaryTaxExpAcc.Text = "-1"
                    Me.TxtExpAcc.Focus()
                    Me.TxtExpAccDesc.Clear()
                End If
            ElseIf Me.NewQuick_Data = True Then
                If Me.TxtPrimaryTaxExpAcc.Text <> "-1" Then
                    Me.TxtExpAcc.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryTaxExpAcc.Text)
                    Me.TxtExpAccDesc.Text = Me.ObjAkun.GetDescription(Me.TxtExpAcc.Text)
                    Me.BtnHitung.Focus()
                Else
                    Me.TxtExpAcc.Clear()
                    Me.TxtPrimaryTaxExpAcc.Text = "-1"
                    Me.TxtExpAcc.Focus()
                    Me.TxtExpAccDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnFindTaxAcc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindTaxAcc.Click
        Try
            Me.FungsiBtnFindTaxAcc()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExpAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtExpAcc.KeyDown
        Dim objCode As New AccModule.ClsMasterAkun
        Dim i As New AccModule.FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtExpAcc.Text <> "" Then
                    If objCode.ApakahSudahAda_Aktif_NonHeader(Me.TxtExpAcc.Text) = True Then
                        Me.TxtPrimaryTaxExpAcc.Text = objCode.GetPrimary(Me.TxtExpAcc.Text)
                        Me.TxtExpAcc.Text = objCode.GetAccountCode(Me.TxtPrimaryTaxExpAcc.Text)
                        Me.TxtExpAccDesc.Text = objCode.GetDescription(Me.TxtExpAcc.Text)
                        Me.BtnHitung.Focus()
                    Else
                        Me.BtnFindTaxAcc.PerformClick()
                    End If
                Else
                    Me.BtnFindTaxAcc.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtExpAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtExpAcc.Leave
        Dim i As New AccModule.FrmMasterAkun_Cari
        Dim objCode As New AccModule.ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtExpAcc.Text <> "" Then
                If objCode.ApakahSudahAda_Aktif_NonHeader(Me.TxtExpAcc.Text) = True Then
                    Me.TxtPrimaryTaxExpAcc.Text = objCode.GetPrimary(Me.TxtExpAcc.Text)
                    Me.TxtExpAcc.Text = objCode.GetAccountCode(Me.TxtPrimaryTaxExpAcc.Text)
                    Me.TxtExpAccDesc.Text = objCode.GetDescription(Me.TxtExpAcc.Text)
                Else
                    If Me.TxtExpAcc.Text <> "" Then
                        i.txtCari.Text = Me.TxtExpAcc.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryTaxExpAcc.Text = i.mPrimaryAkun
                        Me.TxtExpAcc.Text = objCode.GetAccountCode(Me.TxtPrimaryTaxExpAcc.Text)
                        Me.TxtExpAccDesc.Text = objCode.GetDescription(Me.TxtExpAcc.Text)
                    Else
                        Me.TxtExpAcc.Clear()
                        Me.TxtPrimaryTaxExpAcc.Text = "-1"
                        Me.TxtExpAccDesc.Clear()
                    End If
                End If
            Else
                Me.TxtExpAcc.Clear()
                Me.TxtPrimaryTaxExpAcc.Text = "-1"
                Me.TxtExpAccDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        'Dim ij As New FrmTrans_Gaji_Cetak
        Dim ij As New FrmTrans_Gaji_Cetak_IGI_Options
        Dim jk As New FrmTrans_Gaji_Cetak

        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocumentNo.Text <> "" Then
                'ij.mPrimary = Me.TxtPrimary.Text
                'ij.mDef_Prec = Me.Def_Prec
                'ij.ShowDialog()
                'ij.Dispose()

                If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.SalaryCalculation, ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    jk.mPrimary = Me.TxtPrimary.Text
                    jk.UseCustom = True
                    jk.ShowDialog()
                Else
                    ij.mPrimarySalary = Me.TxtPrimary.Text
                    ij.mQuickSalary_Form = False
                    ij.mTanggalTrans_Dari = Me.dtpPeriod.Value.Date
                    ij.mTanggalTrans_Sampai = Me.DtpTo.Value.Date
                    ij.ShowDialog()
                    ij.Dispose()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            jk.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If Me.dgvItem.Item(Me.ColOverride.Index, e.RowIndex).Value = True Then
                If Me.dgvItem.Item(Me.ColOriginalValue.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value Then
                    If e.ColumnIndex = Me.ColAmount.Index Or e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        If e.ColumnIndex = Me.ColAmount.Index Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf Me.dgvItem.Item(Me.ColOriginalValue.Index, e.RowIndex).Value <> Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value Then
                    If e.ColumnIndex = Me.ColAmount.Index Or _
                    e.ColumnIndex = Me.ColOverrideReason.Index Or _
                    e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        If e.ColumnIndex = Me.ColAmount.Index Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                        End If
                    Else
                        e.Cancel = True
                    End If
                End If
            Else
                If e.ColumnIndex = Me.ColOverride.Index Or e.ColumnIndex = Me.ColRemarks.Index Then
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
        Dim FrmExplain As New FrmTrans_Gaji_ShowExplaination
        Dim RowDGV As DataGridViewRow
        Try
            If Me.TxtEmployeeCode1.Text = "-1" Then
                Exit Sub
            End If
            If Me.dgvItem.Rows.Count <= 0 Then
                Exit Sub
            End If
            If e.RowIndex < 0 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColExplain.Index Then
                FrmExplain.TxtFormula.Text = Me.dgvItem.Item(Me.ColExplaination.Index, e.RowIndex).Value
                If Me.dgvItem.Item(Me.ColOverride.Index, e.RowIndex).Value = True Then
                    FrmExplain.TxtFormula.Text += vbCrLf + "Override Value = " + Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value.ToString
                    FrmExplain.TxtFormula.Text += vbCrLf + "Reason = " + Me.dgvItem.Item(Me.ColOverrideReason.Index, e.RowIndex).Value.ToString
                End If
                FrmExplain.ShowDialog()
            End If
            If e.ColumnIndex = Me.ColOverride.Index Then
                Me.Buffer = Me.dgvItem.Item(Me.ColOverride.Index, e.RowIndex).Value
                Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColOriginalValue.Index, e.RowIndex).Value
                Me.dgvItem.Item(Me.ColOverrideReason.Index, e.RowIndex).Value = ""
                Me.dgvItem.Item(Me.ColOverride.Index, e.RowIndex).Value = False
            End If
            'Checklist kalo false 
            If e.ColumnIndex = Me.ColOverride.Index Then
                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = False Then
                    'hitung total
                    Me.HitungTotal()
                    'Update Payment
                    If Me.FrmPaymentData.DGVPayment.Rows.Count = 1 Then
                        RowDGV = Me.FrmPaymentData.DGVPayment.Rows(0)
                        RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_Amount.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtSalaryReceived.Text, Me.Def_Prec)
                        Me.TxtPayment.Text = Me.ObjCD.DigitGrouping(Me.TxtSalaryReceived.Text, Me.Def_Prec)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmExplain.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Dim RowDGV As DataGridViewRow
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColAmount.Index Then
                        If Me.ObjCD.CekNonNegatif(Row.Cells(Me.ColAmount.Index).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        Else
                            'digit grouping
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.Def_Prec)
                            'hitung ulang header
                            Me.HitungTotal()
                            'Update Payment
                            If Me.FrmPaymentData.DGVPayment.Rows.Count = 1 Then
                                RowDGV = Me.FrmPaymentData.DGVPayment.Rows(0)
                                RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_Amount.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtSalaryReceived.Text, Me.Def_Prec)
                                Me.TxtPayment.Text = Me.ObjCD.DigitGrouping(Me.TxtSalaryReceived.Text, Me.Def_Prec)
                            End If
                            '---------------
                        End If
                    End If
                    If e.ColumnIndex = Me.ColOverrideReason.Index Then
                        If Me.ObjCD.CekDataKosong(Row.Cells(Me.ColOverrideReason.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If Me.ObjCD.CekDataKosong(Row.Cells(Me.ColRemarks.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HitungLoanPayment(ByVal vTBLLoan_OnGoing As DseTransaksiLoanPayment.DataTable2DataTable)
        Dim RowLoan_Going As DseTransaksiLoanPayment.DataTable2Row
        Dim RowView As Data.DataRowView
        Dim RowDGV As DataGridViewRow
        Dim PaymentAmount As Decimal = 0
        Dim Filter As String = ""
        Dim PrimaryCashAcc As String
        Dim PrimaryGainAcc As String
        Dim PrimaryLossAcc As String
        Try
            Me.ijK.KosongkanField()
            Me.ijK.FungsiBtnNew()
            Me.ijK.dtpDate.Value = Me.DtpDate.Value
            Me.ijK.TxtDocumentNo.Text = Me.ijK.Obj.GetNoBaru(Me.ijK.dtpDate.Value, Me.ObjCD)
            'Clear Object
            Me.ijK.TxtPrimaryEmployee.Text = "-1"
            Me.ijK.TxtEmployeeCode1.Text = ""
            Me.ijK.TxtEmployeeName.Text = ""
            PrimaryCashAcc = ""
            PrimaryGainAcc = ""
            PrimaryLossAcc = ""
            Me.ijK.CboCashAccount.Text = ""
            Me.ijK.TxtDescCash.Text = ""
            Me.ijK.TxtGainAccount.Text = ""
            Me.ijK.TxtDescGain.Text = ""
            Me.ijK.TxtLossAccount.Text = ""
            Me.ijK.TxtDescLoss.Text = ""
            Me.ijK.TxtRemark.Text = ""
            'Fill data
            Me.ijK.Obj.IsiTBLLoanPayment_OnGoing(Me.ijK.TBLLoan, Me.TxtPrimaryEmployee.Text, True, _
            Me.ijK.TxtPrimaryLoanPayment.Text, Me.ijK.dgvItem, Me.ijK)
            Me.ijK.Obj.IsiCBOLoanPayment(Me.ijK.TBLLoan, Me.ObjCD, Me.ijK.Def_Prec, Me.ijK, Me.ijK.TBLCBOItem)
            Me.ijK.TxtDescCash.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.ijK.TxtPrimaryCashAcc.Text)
            'Input Data Item
            Filter = "[Payment Amount] = " + PaymentAmount.ToString + " AND [Date Schedule] <= #" + Format(Me.DtpTo.Value, "MM/dd/yyyy") + "#"
            vTBLLoan_OnGoing.DefaultView.RowFilter = Filter
            vTBLLoan_OnGoing.DefaultView.Sort = "[Loan No], [Installment Period]"
            If vTBLLoan_OnGoing.DefaultView.Count > 0 Then
                Me.ijK.dgvItem.Rows.Clear()
                For Each RowView In vTBLLoan_OnGoing.DefaultView
                    RowLoan_Going = RowView.Row
                    Me.ijK.dgvItem.Rows.Add()
                    RowDGV = Me.ijK.dgvItem.Rows(Me.ijK.dgvItem.Rows.Count - 1)
                    With RowDGV
                        .Cells(Me.ijK.ColNo.Index).Value = ""
                        .Cells(Me.ijK.ColPrimaryLoan.Index).Value = RowLoan_Going.Primary
                        .Cells(Me.ijK.ColIdentifier.Index).Value = RowLoan_Going.Identifier
                        .Cells(Me.ijK.ColLoanNo.Index).Value = Me.ijK.Obj.GetLoanNumber(.Cells(Me.ijK.ColPrimaryLoan.Index).Value)
                        .Cells(Me.ijK.ColInstallementAt.Index).Value = RowLoan_Going.Installment_Period
                        .Cells(Me.ijK.ColBase.Index).Value = Me.ObjCD.DigitGrouping(RowLoan_Going.Base, Me.ijK.Def_Prec)
                        .Cells(Me.ijK.ColInterest.Index).Value = Me.ObjCD.DigitGrouping(RowLoan_Going.Interest, Me.ijK.Def_Prec)
                        .Cells(Me.ijK.ColTotalInstallment.Index).Value = Me.ObjCD.DigitGrouping(RowLoan_Going.Base + RowLoan_Going.Interest, Me.ijK.Def_Prec)
                        .Cells(Me.ijK.ColPayment.Index).Value = Me.ObjCD.DigitGrouping(RowLoan_Going.Base + RowLoan_Going.Interest, Me.ijK.Def_Prec)
                        .Cells(Me.ijK.ColRemarks.Index).Value = ""
                        .Cells(Me.ijK.ColPrimaryRev.Index).Value = RowLoan_Going.Primary_Receivable_Account
                    End With
                    'Update Nilai Payment
                    RowLoan_Going.Payment_Amount = RowLoan_Going.Base + RowLoan_Going.Interest
                    RowLoan_Going.Update = 1
                Next
                'set number
                Me.ObjCD.SetNumber(Me.ijK.dgvItem)
            End If
            vTBLLoan_OnGoing.DefaultView.RowFilter = ""
            vTBLLoan_OnGoing.DefaultView.Sort = ""
            'Input Header + CBO Loan Posisi Akhir
            Me.ijK.TxtPrimaryEmployee.Text = Me.TxtPrimaryEmployee.Text
            Me.ijK.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
            Me.ijK.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
            Me.ijK.Obj.IsiCBOLoanPayment(Me.ijK.TBLLoan, Me.ObjCD, Me.ijK.Def_Prec, Me.ijK, Me.ijK.TBLCBOItem)
            If Me.NewQuick_Data = False Then
                PrimaryCashAcc = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryCashInAccount)
                PrimaryGainAcc = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryGainAccount)
                PrimaryLossAcc = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSetting.EnumNamaParameter.Default_LoanPayment_PrimaryLossAccount)
                If String.IsNullOrEmpty(PrimaryCashAcc) = False Then
                    Me.ijK.CboCashAccount.Text = Me.ObjAkun.GetAccountCode(PrimaryCashAcc)
                    Me.ijK.TxtDescCash.Text = Me.ObjAkun.GetDescription(Me.ObjAkun.GetAccountCode(PrimaryCashAcc))
                End If
                If String.IsNullOrEmpty(PrimaryGainAcc) = False Then
                    Me.ijK.TxtGainAccount.Text = Me.ObjAkun.GetAccountCode(PrimaryGainAcc)
                    Me.ijK.TxtDescGain.Text = Me.ObjAkun.GetDescription(Me.ObjAkun.GetAccountCode(PrimaryGainAcc))
                End If
                If String.IsNullOrEmpty(PrimaryLossAcc) = False Then
                    Me.ijK.TxtLossAccount.Text = Me.ObjAkun.GetAccountCode(PrimaryLossAcc)
                    Me.ijK.TxtDescLoss.Text = Me.ObjAkun.GetDescription(Me.ObjAkun.GetAccountCode(PrimaryLossAcc))
                End If
            Else
                'untuk New Quick Data
                PrimaryCashAcc = Me.mPrimaryCashIn
                PrimaryGainAcc = Me.mPrimaryGainAcc
                PrimaryLossAcc = Me.mPrimaryLossAcc
                If String.IsNullOrEmpty(PrimaryCashAcc) = False Then
                    Me.ijK.CboCashAccount.Text = Me.ObjAkun.GetAccountCode(PrimaryCashAcc)
                    Me.ijK.TxtDescCash.Text = Me.ObjAkun.GetDescription(Me.ObjAkun.GetAccountCode(PrimaryCashAcc))
                End If
                If String.IsNullOrEmpty(PrimaryGainAcc) = False Then
                    Me.ijK.TxtGainAccount.Text = Me.ObjAkun.GetAccountCode(PrimaryGainAcc)
                    Me.ijK.TxtDescGain.Text = Me.ObjAkun.GetDescription(Me.ObjAkun.GetAccountCode(PrimaryGainAcc))
                End If
                If String.IsNullOrEmpty(PrimaryLossAcc) = False Then
                    Me.ijK.TxtLossAccount.Text = Me.ObjAkun.GetAccountCode(PrimaryLossAcc)
                    Me.ijK.TxtDescLoss.Text = Me.ObjAkun.GetDescription(Me.ObjAkun.GetAccountCode(PrimaryLossAcc))
                End If
            End If
            Me.ijK.TxtRemark.Text = ""
            'Total Loan
            Me.ijK.TotalLoanInstallment = 0
            For Each RowDGV In Me.ijK.dgvItem.Rows
                Me.ijK.TotalLoanInstallment += CDec(RowDGV.Cells(Me.ijK.ColPayment.Index).Value)
            Next
            Me.TxtLoanInstallment.Text = Me.ObjCD.DigitGrouping(Me.ijK.TotalLoanInstallment, Me.ijK.Def_Prec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Public Function ApakahDataitemBenar() As Boolean
        Try
            If Me.TxtPrimaryEmployee.Text = "-1" Then
                MsgBox("Please input employee code.", MsgBoxStyle.Information, "Information")
                Me.TxtEmployeeCode1.Focus()
                Return False
            End If
            If Me.dtpPeriod.Value > Me.DtpTo.Value Then
                MsgBox("Period cannot larger than period to.", MsgBoxStyle.Information, "Information")
                Me.dtpPeriod.Focus()
                Return False
            End If
            'Cek apakah ada yang tumpang tindih
            If Me.Obj.FillTBLStartEnd(Me.TxtPrimaryEmployee.Text, Me.DtpTo.Value, Me.dtpPeriod.Value, Me.Baru, Me.TxtPrimary.Text) = False Then
                Me.BtnHitung.Focus()
                Return False
            End If
            'Cek Berkaitan dengan attendance 
            If Me.Obj.CekValidPeriodEmp_HubunganAttendance(Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, Me.DtpTo.Value) = False Then
                Me.dtpPeriod.Focus()
                Return False
            End If

            'sampai sini
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub CekCobaCobaData_Item()
        Dim rowDGV As DataGridViewRow
        Try
            'autosize dgv
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item overhead cost
            Me.dgvItem.Rows.Clear()
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
            With rowDGV
                .Cells(Me.ColNo.Index).Value = ""
                .Cells(Me.ColSlipDesc.Index).Value = "Trail-1"
                .Cells(Me.ColComponentName.Index).Value = "Percobaan Pertama"
                .Cells(Me.ColTipeKomponen.Index).Value = "Income"
                .Cells(Me.ColTaxElement.Index).Value = "1"
                .Cells(Me.ColExpenseAccount.Index).Value = "5.00.01"
                .Cells(Me.ColAmount.Index).Value = "5,000,000"
                .Cells(Me.ColOverride.Index).Value = False
                .Cells(Me.ColOverrideReason.Index).Value = "-"
                .Cells(Me.ColOriginalValue.Index).Value = "5,000,000"
                .Cells(Me.ColRemarks.Index).Value = "-"
                .Cells(Me.ColPrimarySlipDesc.Index).Value = "-1"
                .Cells(Me.ColIdentifierKomponen.Index).Value = "1-1"
                .Cells(Me.ColPrimaryAkunExpense.Index).Value = "-1"
                .Cells(Me.ColExplaination.Index).Value = "--"
            End With

            'set number
            ObjCD.SetNumber(Me.dgvItem)

            'autosize dgv
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Me.TxtIncome.Text = "5,000,000"
            Me.TxtSalaryReceived.Text = "5,000,000"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FormatListVarConst_OtherDeduction(ByVal vPrimaryOtherDeduction As Double, _
    ByVal TBLVar As DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_variableDataTable, _
    ByVal TBLConst As DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_constraintDataTable) As String
        Dim RowVar As DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_variableRow
        Dim RowConst As DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_constraintRow
        Dim Retval As String = ""
        Dim Filter As String = ""
        Dim DV As Data.DataRowView
        Dim Counter As Integer = 0
        Dim Comparer As String = ""
        Try
            Filter = "[Primary Transaksi] = " + vPrimaryOtherDeduction.ToString + ""
            TBLVar.DefaultView.RowFilter = Filter
            If TBLVar.DefaultView.Count > 0 Then
                Counter = 0
                For Each DV In TBLVar.DefaultView
                    RowVar = DV.Row
                    If Counter = 0 Then
                        Retval = "Where" + vbCrLf + vbTab + RowVar.Variable.ToString + " = " + Me.ObjParam.GetDescByPrimary(RowVar.Primary_Parameter).ToString
                    Else
                        Retval += vbCrLf + vbTab + RowVar.Variable.ToString + " = " + Me.ObjParam.GetDescByPrimary(RowVar.Primary_Parameter).ToString
                    End If
                    Counter += 1
                Next
            End If
            TBLConst.DefaultView.RowFilter = Filter
            If TBLConst.DefaultView.Count > 0 Then
                Counter = 0
                For Each DV In TBLConst.DefaultView
                    RowConst = DV.Row
                    If RowConst.Compare = ClsMasterParameterGaji.EnumComparer.SamaDengan Then
                        Comparer = "="
                    ElseIf RowConst.Compare = ClsMasterParameterGaji.EnumComparer.LebihBesar Then
                        Comparer = ">"
                    ElseIf RowConst.Compare = ClsMasterParameterGaji.EnumComparer.LebihKecil Then
                        Comparer = "<"
                    ElseIf RowConst.Compare = ClsMasterParameterGaji.EnumComparer.LebihBesar_SamaDengan Then
                        Comparer = ">="
                    ElseIf RowConst.Compare = ClsMasterParameterGaji.EnumComparer.LebihKecil_SamaDengan Then
                        Comparer = "<="
                    ElseIf RowConst.Compare = ClsMasterParameterGaji.EnumComparer.TidakSamaDengan Then
                        Comparer = "<>"
                    End If
                    If Counter = 0 Then
                        Retval = "Constraint" + vbCrLf + vbTab + Me.ObjParam.GetCodeByPrimary(RowConst.Primary_Parameter) + " " + Comparer.ToString + " " + RowConst.Ref_Value.ToString
                    Else
                        Retval += vbCrLf + vbTab + Me.ObjParam.GetCodeByPrimary(RowConst.Primary_Parameter) + " " + Comparer.ToString + " " + RowConst.Ref_Value.ToString
                    End If
                    Counter += 1
                Next
            End If
            TBLVar.DefaultView.RowFilter = ""
            TBLConst.DefaultView.RowFilter = ""
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function FormatListVarConst(ByVal vIdentifierMasterSalaryFormula As String, _
    ByVal TBLVar As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_variableDataTable, _
    ByVal TBLConst As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_constraintDataTable) As String
        Dim RowVar As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_variableRow
        Dim RowConst As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_constraintRow
        Dim Retval As String = ""
        Dim Filter As String = ""
        Dim DV As Data.DataRowView
        Dim Counter As Integer = 0
        Dim Comparer As String = ""
        Dim TipeVarConst As Core_Module.ClsMasterSalaryFormula_item.EnumTipeConstraitVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.CustomValue

        Try
            'Retval += vbCrLf

            Filter = "[Identifier Salary Component] = '" + vIdentifierMasterSalaryFormula.ToString + "'"
            TBLVar.DefaultView.RowFilter = Filter
            If TBLVar.DefaultView.Count > 0 Then
                Counter = 0
                For Each DV In TBLVar.DefaultView
                    RowVar = DV.Row
                    TipeVarConst = RowVar.Type2
                    If TipeVarConst = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                        If Counter = 0 Then
                            Retval = "Where " + vbCrLf + vbTab + RowVar.Variable.ToString + " = " + Me.ObjParam.GetDescByPrimary(RowVar.Primary_Parameter).ToString
                        Else
                            Retval += vbCrLf + vbTab + RowVar.Variable.ToString + " = " + Me.ObjParam.GetDescByPrimary(RowVar.Primary_Parameter).ToString
                        End If
                    ElseIf TipeVarConst = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                        If Counter = 0 Then
                            Retval = "Where " + vbCrLf + vbTab + RowVar.Variable.ToString + " = " + Me.ObjKompetensi.GetDesc(RowVar.Primary_Parameter).ToString
                        Else
                            Retval += vbCrLf + vbTab + RowVar.Variable.ToString + " = " + Me.ObjKompetensi.GetDesc(RowVar.Primary_Parameter).ToString
                        End If
                    Else
                        MsgBox("belum")
                    End If
                    Counter += 1
                Next
            End If

            'Retval += vbCrLf
            TBLConst.DefaultView.RowFilter = Filter
            If TBLConst.DefaultView.Count > 0 Then
                Counter = 0
                For Each DV In TBLConst.DefaultView
                    RowConst = DV.Row
                    If RowConst.Compare = ClsMasterParameterGaji.EnumComparer.SamaDengan Then
                        Comparer = "="
                    ElseIf RowConst.Compare = ClsMasterParameterGaji.EnumComparer.LebihBesar Then
                        Comparer = ">"
                    ElseIf RowConst.Compare = ClsMasterParameterGaji.EnumComparer.LebihKecil Then
                        Comparer = "<"
                    ElseIf RowConst.Compare = ClsMasterParameterGaji.EnumComparer.LebihBesar_SamaDengan Then
                        Comparer = ">="
                    ElseIf RowConst.Compare = ClsMasterParameterGaji.EnumComparer.LebihKecil_SamaDengan Then
                        Comparer = "<="
                    ElseIf RowConst.Compare = ClsMasterParameterGaji.EnumComparer.TidakSamaDengan Then
                        Comparer = "<>"
                    End If
                    TipeVarConst = RowConst.Type2
                    If TipeVarConst = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                        If Counter = 0 Then
                            If Retval = "" Then
                                Retval = "Constraint " + vbCrLf + vbTab + Me.ObjParam.GetCodeByPrimary(RowConst.Primary_Parameter) + " " + Comparer.ToString + " " + RowConst.Ref_Value.ToString
                            Else
                                Retval += vbCrLf + "Constraint " + vbCrLf + vbTab + Me.ObjParam.GetCodeByPrimary(RowConst.Primary_Parameter) + " " + Comparer.ToString + " " + RowConst.Ref_Value.ToString
                            End If
                        Else
                            Retval += vbCrLf + vbTab + Me.ObjParam.GetCodeByPrimary(RowConst.Primary_Parameter) + " " + Comparer.ToString + " " + RowConst.Ref_Value.ToString
                        End If
                    ElseIf TipeVarConst = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                        If Counter = 0 Then
                            If Retval = "" Then
                                Retval = "Constraint " + vbCrLf + vbTab + Me.ObjKompetensi.GetDesc(RowConst.Primary_Parameter) + " " + Comparer.ToString + " " + _
                            Me.ObjKompetensi.SetFormatDataCustom(RowConst.Ref_Value.ToString, RowConst.Primary_Parameter, Me.ObjCD)
                            Else
                                Retval += vbCrLf + "Constraint " + vbCrLf + vbTab + Me.ObjKompetensi.GetDesc(RowConst.Primary_Parameter) + " " + Comparer.ToString + " " + _
                            Me.ObjKompetensi.SetFormatDataCustom(RowConst.Ref_Value.ToString, RowConst.Primary_Parameter, Me.ObjCD)
                            End If
                        Else
                            Retval += vbCrLf + vbTab + Me.ObjKompetensi.GetDesc(RowConst.Primary_Parameter) + " " + Comparer.ToString + " " + _
                            Me.ObjKompetensi.SetFormatDataCustom(RowConst.Ref_Value.ToString, RowConst.Primary_Parameter, Me.ObjCD)
                        End If
                    Else
                        MsgBox("belum")
                    End If
                    Counter += 1
                Next
            End If
            TBLVar.DefaultView.RowFilter = ""
            TBLConst.DefaultView.RowFilter = ""

            Return Retval + vbCrLf
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub CalculateHitungItemSalaryFormula_OtherDeduction(ByVal FrmOther As FrmTrans_Gaji_OtherDeduction)
        Dim RowDGV As DataGridViewRow
        'Dim RowSL_Kary As DseTransaksiSalary.tblmaster_salary_formula_component_detailemployeeRow
        Dim RowSL_Item_Other As DseTransaksiSalary.tblmaster_salary_formula_otherdeductionRow
        Dim Param_Primary_SalaryComp As Double = -1
        'Dim CboDgv As DataGridViewComboBoxCell
        Dim TS As TimeSpan = Me.DtpTo.Value - Me.dtpPeriod.Value
        Dim Counter As Integer = 0
        Dim TglDaily As Date = #1/1/1950#
        Dim TotalAmount As Decimal
        Dim TotalRounding As Decimal
        Dim Explanation As String
        Dim DTAVar As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_otherdeduction_variableTableAdapter
        Dim TBLVar As New DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_variableDataTable
        Dim DTAConst As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_otherdeduction_constraintTableAdapter
        Dim TBLConst As New DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_constraintDataTable
        Try
            DTAVar.Connection.ConnectionString = Core_Module.GetConstring
            DTAConst.Connection.ConnectionString = Core_Module.GetConstring

            FrmOther.dgvItem.Rows.Clear()

            Me.Obj.FillSalaryComponent_Item_OtherDeduction(Me.TBLSL_Item_Other, Me.TxtPrimaryEmployee.Text)
            DTAVar.Fill(TBLVar)
            DTAConst.Fill(TBLConst)
            If Me.TBLSL_Item_Other.Rows.Count > 0 Then
                For Each RowSL_Item_Other In Me.TBLSL_Item_Other.Rows
                    Explanation = ""
                    TotalAmount = 0
                    FrmOther.dgvItem.Rows.Add()
                    RowDGV = FrmOther.dgvItem.Rows(FrmOther.dgvItem.Rows.Count - 1)
                    RowDGV.Cells(FrmOther.ColNo.Index).Value = ""
                    RowDGV.Cells(FrmOther.ColPrimarySlipDesc.Index).Value = RowSL_Item_Other.Primary_Slip_Desc
                    RowDGV.Cells(FrmOther.ColSlipDesc.Index).Value = RowSL_Item_Other.SlipDesc
                    RowDGV.Cells(FrmOther.ColComponentName.Index).Value = RowSL_Item_Other.Description
                    'CboDgv = RowDGV.Cells(Me.ColTipeKomponen.Index)
                    'RowDGV.Cells(Me.ColTipeKomponen.Index).Value = CboDgv.Items.Item(RowSL_Item_Other.Type)
                    RowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Other Deduction"
                    'CboDgv = RowDGV.Cells(Me.ColTaxElement.Index)
                    'RowDGV.Cells(Me.ColTaxElement.Index).Value = CboDgv.Items.Item(RowSL_Item_Other.Tax_Component)
                    RowDGV.Cells(Me.ColTaxElement.Index).Value = "GAJI/PENSIUN ATAU THT/JHT"
                    RowDGV.Cells(Me.ColPrimaryAkunExpense.Index).Value = RowSL_Item_Other.Primary_Referred_Account
                    RowDGV.Cells(Me.ColExpenseAccount.Index).Value = Me.ObjAkun.GetAccountCode(RowSL_Item_Other.Primary_Referred_Account)
                    If RowSL_Item_Other.Calculation_Period = 0 Then
                        'Daily
                        Counter = 0
                        TglDaily = Me.dtpPeriod.Value.Date
                        While Counter <= TS.TotalDays
                            If Me.Obj.IsConstraint_Daily_OtherDeduction(RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, _
                            Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, TBLConst, Me.ObjKrywn, Me.ObjKompetensi) = False Then
                                'False
                                If Explanation = "" Then
                                    Explanation = RowSL_Item_Other.Description.ToString
                                    If IsNumeric(RowSL_Item_Other.Formula) = True Then
                                        Explanation += vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec) + ""
                                    Else
                                        Explanation += vbCrLf + "Formula = " + RowSL_Item_Other.Formula.ToString + ""
                                    End If
                                    Explanation += vbCrLf + Me.FormatListVarConst_OtherDeduction(RowSL_Item_Other.Primary, TBLVar, TBLConst)
                                    Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + "0"
                                Else
                                    Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + "0"
                                End If
                                'BarisTerakhir
                                If Counter = TS.TotalDays Then
                                    TotalRounding = 0
                                    Explanation += vbCrLf + "Total" + vbTab + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                    If RowSL_Item_Other.Rounding_Type = 0 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                                    ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                                    ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                                    ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                                    End If
                                    If RowSL_Item_Other.Rounding_Type = 0 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                        TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                    ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                                    ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                                    ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                                    End If
                                    If IsNumeric(RowSL_Item_Other.Monthly_Max_Value) = True Then
                                        Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Max_Value, Me.Def_Prec)
                                    Else
                                        Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item_Other.Monthly_Max_Value
                                    End If
                                    If IsNumeric(RowSL_Item_Other.Monthly_Min_Value) = True Then
                                        Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Min_Value, Me.Def_Prec)
                                    Else
                                        Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item_Other.Monthly_Min_Value
                                    End If
                                    If IsNumeric(RowSL_Item_Other.Yearly_Max_Value) = True Then
                                        Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Max_Value, Me.Def_Prec)
                                    Else
                                        Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item_Other.Yearly_Max_Value
                                    End If
                                    If IsNumeric(RowSL_Item_Other.Yearly_Min_Value) = True Then
                                        Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                                    Else
                                        Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item_Other.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                                    End If
                                    Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                                End If
                            Else
                                'True
                                If Explanation = "" Then
                                    Explanation = RowSL_Item_Other.Description.ToString
                                    If IsNumeric(RowSL_Item_Other.Formula) = True Then
                                        Explanation += vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec) + ""
                                        If RowSL_Item_Other.Rounding_Type = 0 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item_Other.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item_Other.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item_Other.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        End If
                                        Explanation += vbCrLf + Me.FormatListVarConst_OtherDeduction(RowSL_Item_Other.Primary, TBLVar, TBLConst)
                                        If RowSL_Item_Other.Rounding_Type = 0 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item_Other.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item_Other.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item_Other.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        End If
                                    Else
                                        Explanation += vbCrLf + "Formula = " + RowSL_Item_Other.Formula.ToString + ""
                                        If RowSL_Item_Other.Rounding_Type = 0 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        End If
                                        Explanation += vbCrLf + Me.FormatListVarConst_OtherDeduction(RowSL_Item_Other.Primary, TBLVar, TBLConst)
                                        If RowSL_Item_Other.Rounding_Type = 0 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        End If
                                    End If
                                Else
                                    If IsNumeric(RowSL_Item_Other.Formula) = True Then
                                        If RowSL_Item_Other.Rounding_Type = 0 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec)
                                            TotalAmount += Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item_Other.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item_Other.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item_Other.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item_Other.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item_Other.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item_Other.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        End If
                                    Else
                                        If RowSL_Item_Other.Rounding_Type = 0 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        End If
                                    End If
                                End If
                                'BarisTerakhir
                                If Counter = TS.TotalDays Then
                                    TotalRounding = 0
                                    Explanation += vbCrLf + "Total" + vbTab + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                    If RowSL_Item_Other.Rounding_Type = 0 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                                    ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                                    ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                                    ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                                    End If
                                    If RowSL_Item_Other.Rounding_Type = 0 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                        TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                    ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                                    ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                                    ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                                    End If
                                    If IsNumeric(RowSL_Item_Other.Monthly_Max_Value) = True Then
                                        Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Max_Value, Me.Def_Prec)
                                    Else
                                        Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item_Other.Monthly_Max_Value
                                    End If
                                    If IsNumeric(RowSL_Item_Other.Monthly_Min_Value) = True Then
                                        Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Min_Value, Me.Def_Prec)
                                    Else
                                        Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item_Other.Monthly_Min_Value
                                    End If
                                    If IsNumeric(RowSL_Item_Other.Yearly_Max_Value) = True Then
                                        Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Max_Value, Me.Def_Prec)
                                    Else
                                        Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item_Other.Yearly_Max_Value
                                    End If
                                    If IsNumeric(RowSL_Item_Other.Yearly_Min_Value) = True Then
                                        Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                                    Else
                                        Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item_Other.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                                    End If
                                    TotalRounding = Me.Obj.EndResultValue_Final_OtherDeduction(True, RowSL_Item_Other.Yearly_Max_Value, RowSL_Item_Other.Yearly_Min_Value, RowSL_Item_Other.Monthly_Max_Value, RowSL_Item_Other.Monthly_Min_Value, _
                                    Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, RowSL_Item_Other.Calculate_On, TotalRounding, RowSL_Item_Other.Primary, _
                                    Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                    Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn)
                                    Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                                End If
                            End If
                            TglDaily = TglDaily.AddDays(1)
                            Counter += 1
                        End While
                    ElseIf RowSL_Item_Other.Calculation_Period = 1 Then
                        'Monthly
                        TglDaily = Me.dtpPeriod.Value.Date
                        If Me.Obj.IsConstraint_OtherDeduction(RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, _
                        Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, TBLConst, Me.ObjKompetensi, Me.ObjKrywn) = False Then
                            'False
                            Explanation = RowSL_Item_Other.Description.ToString
                            If IsNumeric(RowSL_Item_Other.Formula) = True Then
                                Explanation += vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec) + ""
                            Else
                                Explanation += vbCrLf + "Formula = " + RowSL_Item_Other.Formula.ToString + ""
                            End If
                            Explanation += vbCrLf + Me.FormatListVarConst_OtherDeduction(RowSL_Item_Other.Primary, TBLVar, TBLConst)
                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + "0"
                            'BarisTerakhir
                            TotalRounding = 0
                            Explanation += vbCrLf + "Total" + vbTab + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                            If RowSL_Item_Other.Rounding_Type = 0 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                            ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                            ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                            ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                            End If
                            If RowSL_Item_Other.Rounding_Type = 0 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                            ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                            ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                            ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                            End If
                            If IsNumeric(RowSL_Item_Other.Monthly_Max_Value) = True Then
                                Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Max_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item_Other.Monthly_Max_Value
                            End If
                            If IsNumeric(RowSL_Item_Other.Monthly_Min_Value) = True Then
                                Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Min_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item_Other.Monthly_Min_Value
                            End If
                            If IsNumeric(RowSL_Item_Other.Yearly_Max_Value) = True Then
                                Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Max_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item_Other.Yearly_Max_Value
                            End If
                            If IsNumeric(RowSL_Item_Other.Yearly_Min_Value) = True Then
                                Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                            Else
                                Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item_Other.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                            End If
                            Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                        Else
                            'True
                            Explanation = RowSL_Item_Other.Description.ToString
                            If IsNumeric(RowSL_Item_Other.Formula) = True Then
                                Explanation += vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec) + ""
                                TotalAmount += Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec)
                                Explanation += vbCrLf + Me.FormatListVarConst_OtherDeduction(RowSL_Item_Other.Primary, TBLVar, TBLConst)
                                Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Formula = " + RowSL_Item_Other.Formula.ToString + ""
                                TotalAmount += Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormulaMonthlyYearly_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, _
                                Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec)
                                Explanation += vbCrLf + Me.FormatListVarConst_OtherDeduction(RowSL_Item_Other.Primary, TBLVar, TBLConst)
                                Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormulaMonthlyYearly_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, _
                                Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec)
                            End If
                            'BarisTerakhir
                            TotalRounding = 0
                            Explanation += vbCrLf + "Total" + vbTab + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                            If RowSL_Item_Other.Rounding_Type = 0 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                            ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                            ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                            ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                            End If
                            If RowSL_Item_Other.Rounding_Type = 0 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                            ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD).ToString, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                            ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD).ToString, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                            ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD).ToString, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                            End If
                            If IsNumeric(RowSL_Item_Other.Monthly_Max_Value) = True Then
                                Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Max_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item_Other.Monthly_Max_Value
                            End If
                            If IsNumeric(RowSL_Item_Other.Monthly_Min_Value) = True Then
                                Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Min_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item_Other.Monthly_Min_Value
                            End If
                            If IsNumeric(RowSL_Item_Other.Yearly_Max_Value) = True Then
                                Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Max_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item_Other.Yearly_Max_Value
                            End If
                            If IsNumeric(RowSL_Item_Other.Yearly_Min_Value) = True Then
                                Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                            Else
                                Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item_Other.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                            End If
                            TotalRounding = Me.Obj.EndResultValue_Final_OtherDeduction(False, RowSL_Item_Other.Yearly_Max_Value, RowSL_Item_Other.Yearly_Min_Value, RowSL_Item_Other.Monthly_Max_Value, RowSL_Item_Other.Monthly_Min_Value, _
                            Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, RowSL_Item_Other.Calculate_On, TotalRounding, RowSL_Item_Other.Primary, _
                            Me.dtpPeriod.Value, Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn)
                            Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                        End If
                    ElseIf RowSL_Item_Other.Calculation_Period = 2 Then
                        If RowSL_Item_Other.Calculate_On = Me.EnumBln_CalculateOn(Me.DtpDate.Value.Date) Then
                            'Yearly
                            TglDaily = Me.dtpPeriod.Value.Date
                            If Me.Obj.IsConstraint_OtherDeduction(RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, _
                            Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, TBLConst, Me.ObjKompetensi, Me.ObjKrywn) = False Then
                                'False
                                Explanation = RowSL_Item_Other.Description.ToString
                                If IsNumeric(RowSL_Item_Other.Formula) = True Then
                                    Explanation += vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec) + ""
                                Else
                                    Explanation += vbCrLf + "Formula = " + RowSL_Item_Other.Formula.ToString + ""
                                End If
                                Explanation += vbCrLf + Me.FormatListVarConst_OtherDeduction(RowSL_Item_Other.Primary, TBLVar, TBLConst)
                                Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + "0"
                                'BarisTerakhir
                                TotalRounding = 0
                                Explanation += vbCrLf + "Total" + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                If RowSL_Item_Other.Rounding_Type = 0 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                                ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                                ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                                ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                                End If
                                If RowSL_Item_Other.Rounding_Type = 0 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                    TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                    TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                                ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                    TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                                ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                    TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                                End If
                                If IsNumeric(RowSL_Item_Other.Monthly_Max_Value) = True Then
                                    Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Max_Value, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item_Other.Monthly_Max_Value
                                End If
                                If IsNumeric(RowSL_Item_Other.Monthly_Min_Value) = True Then
                                    Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Min_Value, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item_Other.Monthly_Min_Value
                                End If
                                If IsNumeric(RowSL_Item_Other.Yearly_Max_Value) = True Then
                                    Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Max_Value, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item_Other.Yearly_Max_Value
                                End If
                                If IsNumeric(RowSL_Item_Other.Yearly_Min_Value) = True Then
                                    Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                                Else
                                    Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item_Other.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                                End If
                                Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                            Else
                                'True
                                Explanation = RowSL_Item_Other.Description.ToString
                                If IsNumeric(RowSL_Item_Other.Formula) = True Then
                                    Explanation += vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec) + ""
                                    TotalAmount += Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec)
                                    Explanation += vbCrLf + Me.FormatListVarConst_OtherDeduction(RowSL_Item_Other.Primary, TBLVar, TBLConst)
                                    Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Formula = " + RowSL_Item_Other.Formula.ToString + ""
                                    TotalAmount += Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormulaMonthlyYearly_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, _
                                    Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec)
                                    Explanation += vbCrLf + Me.FormatListVarConst_OtherDeduction(RowSL_Item_Other.Primary, TBLVar, TBLConst)
                                    Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormulaMonthlyYearly_OtherDeduction(RowSL_Item_Other.Formula, RowSL_Item_Other.Primary, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, _
                                    Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn), Me.Def_Prec)
                                End If
                                'BarisTerakhir
                                TotalRounding = 0
                                Explanation += vbCrLf + "Total" + vbTab + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                If RowSL_Item_Other.Rounding_Type = 0 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                                ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                                ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                                ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                                End If
                                If RowSL_Item_Other.Rounding_Type = 0 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                    TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                    TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                                ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                    TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                                ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                    TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                                End If
                                If IsNumeric(RowSL_Item_Other.Monthly_Max_Value) = True Then
                                    Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Max_Value, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item_Other.Monthly_Max_Value
                                End If
                                If IsNumeric(RowSL_Item_Other.Monthly_Min_Value) = True Then
                                    Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Min_Value, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item_Other.Monthly_Min_Value
                                End If
                                If IsNumeric(RowSL_Item_Other.Yearly_Max_Value) = True Then
                                    Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Max_Value, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item_Other.Yearly_Max_Value
                                End If
                                If IsNumeric(RowSL_Item_Other.Yearly_Min_Value) = True Then
                                    Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                                Else
                                    Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item_Other.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                                End If
                                TotalRounding = Me.Obj.EndResultValue_Final_OtherDeduction(False, RowSL_Item_Other.Yearly_Max_Value, RowSL_Item_Other.Yearly_Min_Value, RowSL_Item_Other.Monthly_Max_Value, RowSL_Item_Other.Monthly_Min_Value, _
                                Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, RowSL_Item_Other.Calculate_On, TotalRounding, RowSL_Item_Other.Primary, _
                                Me.dtpPeriod.Value, Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn)
                                Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                            End If
                        Else
                            'False
                            Explanation = RowSL_Item_Other.Description.ToString
                            If IsNumeric(RowSL_Item_Other.Formula) = True Then
                                Explanation += vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Formula, Me.Def_Prec) + ""
                            Else
                                Explanation += vbCrLf + "Formula = " + RowSL_Item_Other.Formula.ToString + ""
                            End If
                            Explanation += vbCrLf + Me.FormatListVarConst_OtherDeduction(RowSL_Item_Other.Primary, TBLVar, TBLConst)
                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + "0"
                            'BarisTerakhir
                            TotalRounding = 0
                            Explanation += vbCrLf + "Total" + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                            If RowSL_Item_Other.Rounding_Type = 0 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                            ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                            ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                            ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item_Other.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                            End If
                            If RowSL_Item_Other.Rounding_Type = 0 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                            ElseIf RowSL_Item_Other.Rounding_Type = 1 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                            ElseIf RowSL_Item_Other.Rounding_Type = 2 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                            ElseIf RowSL_Item_Other.Rounding_Type = 3 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item_Other.Rounding_Unit, Me.ObjCD)
                            End If
                            If IsNumeric(RowSL_Item_Other.Monthly_Max_Value) = True Then
                                Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Max_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item_Other.Monthly_Max_Value
                            End If
                            If IsNumeric(RowSL_Item_Other.Monthly_Min_Value) = True Then
                                Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Monthly_Min_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item_Other.Monthly_Min_Value
                            End If
                            If IsNumeric(RowSL_Item_Other.Yearly_Max_Value) = True Then
                                Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Max_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item_Other.Yearly_Max_Value
                            End If
                            If IsNumeric(RowSL_Item_Other.Yearly_Min_Value) = True Then
                                Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item_Other.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                            Else
                                Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item_Other.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item_Other.Calculate_On).ToString + "/1/1950#"), "MMM")
                            End If
                            Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                        End If
                    End If
                    RowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                    RowDGV.Cells(Me.ColExplaination.Index).Value = Explanation
                    RowDGV.Cells(Me.ColOriginalValue.Index).Value = RowDGV.Cells(Me.ColAmount.Index).Value
                    RowDGV.Cells(Me.ColIdentifierKomponen.Index).Value = RowSL_Item_Other.Primary
                    RowDGV.Cells(Me.ColOverride.Index).Value = False
                    RowDGV.Cells(Me.ColOverrideReason.Index).Value = ""
                    RowDGV.Cells(Me.ColRemarks.Index).Value = ""
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAVar.Dispose()
            TBLVar.Dispose()
            DTAConst.Dispose()
            TBLConst.Dispose()
        End Try
    End Sub

    Public Sub CalculateHitungItemSalaryFormula()
        Dim RowDGV As DataGridViewRow
        Dim RowSL_Kary As DseTransaksiSalary.tblmaster_salary_formula_component_detailemployeeRow
        Dim RowSL_Item As DseTransaksiSalary.tblmaster_salary_formula_salarycomponentRow
        Dim Param_Primary_SalaryComp As Double = -1
        Dim CboDgv As DataGridViewComboBoxCell
        Dim TS As TimeSpan = Me.DtpTo.Value - Me.dtpPeriod.Value
        Dim Counter As Integer = 0
        Dim TglDaily As Date = #1/1/1950#
        Dim TotalAmount As Decimal
        Dim TotalRounding As Decimal
        Dim Explanation As String
        Dim TBLVar As New DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_variableDataTable
        Dim TBLConst As New DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_constraintDataTable

        Try
            'DTAVar.Connection.ConnectionString = Core_Module.GetConstring
            'DTAConst.Connection.ConnectionString = Core_Module.GetConstring

            Me.dgvItem.Rows.Clear()
            If Me.TBLSL_Kary.Rows.Count > 0 Then
                RowSL_Kary = Me.TBLSL_Kary.Rows(0)
                Param_Primary_SalaryComp = RowSL_Kary.Primary_Transaksi
                Me.dgvItem.Rows.Add()
                RowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                RowDGV.Cells(Me.ColNo.Index).Value = ""
                RowDGV.Cells(Me.ColPrimarySlipDesc.Index).Value = RowSL_Kary.Primary_Slip_Desc
                RowDGV.Cells(Me.ColSlipDesc.Index).Value = RowSL_Kary.Slip_Desc
                RowDGV.Cells(Me.ColTipeKomponen.Index).Value = "Income"
                RowDGV.Cells(Me.ColComponentName.Index).Value = "BASIC SALARY"
                CboDgv = RowDGV.Cells(Me.ColTaxElement.Index)
                RowDGV.Cells(Me.ColTaxElement.Index).Value = CboDgv.Items.Item(RowSL_Kary.Tax_Component)
                RowDGV.Cells(Me.ColPrimaryAkunExpense.Index).Value = RowSL_Kary.Primary_Expense_Account
                RowDGV.Cells(Me.ColExpenseAccount.Index).Value = Me.ObjAkun.GetAccountCode(RowSL_Kary.Primary_Expense_Account)
                RowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(RowSL_Kary.Basic_Salary, Me.Def_Prec)
                RowDGV.Cells(Me.ColOriginalValue.Index).Value = Me.ObjCD.DigitGrouping(RowSL_Kary.Basic_Salary, Me.Def_Prec)
                RowDGV.Cells(Me.ColOverride.Index).Value = False
                RowDGV.Cells(Me.ColOverrideReason.Index).Value = ""
                RowDGV.Cells(Me.ColRemarks.Index).Value = ""
                RowDGV.Cells(Me.ColIdentifierKomponen.Index).Value = ""
                RowDGV.Cells(Me.ColExplaination.Index).Value = "BASIC SALARY" + vbCrLf + "Nominal" + vbTab + "         =" + Me.ObjCD.DigitGrouping(RowDGV.Cells(Me.ColAmount.Index).Value, Me.Def_Prec).ToString
                RowDGV.Cells(Me.ColIsRutin.Index).Value = 1
                RowDGV.Cells(Me.ColPrimaryAkunPayment.Index).Value = -1
            End If

            If Me.mPrimarySalaryFormulaRef = Param_Primary_SalaryComp And Me.mPrimarySalaryFormulaRef <> -1 Then
                Me.CopyData(Me.TBLVar_Ref, TBLVar)
                Me.CopyData(Me.TBLSL_Item_Ref, Me.TBLSL_Item)
                Me.CopyData(Me.TBLConst_Ref, TBLConst)
            Else
                Me.Obj.FillSalaryComponent_Item(Me.TBLSL_Item, Param_Primary_SalaryComp)
                DTAVar.FillBy_Primary(TBLVar, Param_Primary_SalaryComp)
                DTAConst.FillBy_Primary(TBLConst, Param_Primary_SalaryComp)
                Me.mPrimarySalaryFormulaRef = Param_Primary_SalaryComp
                Me.CopyData(TBLVar, Me.TBLVar_Ref)
                Me.CopyData(Me.TBLSL_Item, Me.TBLSL_Item_Ref)
                Me.CopyData(TBLConst, Me.TBLConst_Ref)
            End If

            If Me.TBLSL_Item.Rows.Count > 0 Then
                For Each RowSL_Item In Me.TBLSL_Item.Rows
                    Explanation = ""
                    TotalAmount = 0
                    Me.dgvItem.Rows.Add()
                    RowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    RowDGV.Cells(Me.ColNo.Index).Value = ""
                    RowDGV.Cells(Me.ColIsRutin.Index).Value = RowSL_Item.rutin_disetahunkan
                    RowDGV.Cells(Me.ColPrimarySlipDesc.Index).Value = RowSL_Item.Primary_Slip_Desc
                    RowDGV.Cells(Me.ColSlipDesc.Index).Value = RowSL_Item.Slip_Desc
                    RowDGV.Cells(Me.ColComponentName.Index).Value = RowSL_Item.Description
                    CboDgv = RowDGV.Cells(Me.ColTipeKomponen.Index)
                    RowDGV.Cells(Me.ColTipeKomponen.Index).Value = CboDgv.Items.Item(RowSL_Item.Type)
                    CboDgv = RowDGV.Cells(Me.ColTaxElement.Index)
                    RowDGV.Cells(Me.ColTaxElement.Index).Value = CboDgv.Items.Item(RowSL_Item.Tax_Component)
                    RowDGV.Cells(Me.ColPrimaryAkunExpense.Index).Value = RowSL_Item.Primary_Expense_Account
                    RowDGV.Cells(Me.ColExpenseAccount.Index).Value = Me.ObjAkun.GetAccountCode(RowSL_Item.Primary_Expense_Account)
                    RowDGV.Cells(Me.ColPrimaryAkunPayment.Index).Value = RowSL_Item.primary_payment_account
                    If RowSL_Item.Calculation_Period = 0 Then
                        'Daily
                        Counter = 0
                        TglDaily = Me.dtpPeriod.Value.Date
                        While Counter <= TS.TotalDays
                            If Me.Obj.IsConstraint_Daily(RowSL_Item.Identifier, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, _
                            Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, TBLConst, Me.ObjKrywn, Me.ObjKompetensi) = False Then
                                'False
                                If Explanation = "" Then
                                    Explanation = RowSL_Item.Description.ToString
                                    If IsNumeric(RowSL_Item.Formula) = True Then
                                        Explanation += vbCrLf + vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec) + ""
                                    Else
                                        Explanation += vbCrLf + vbCrLf + "Formula = " + RowSL_Item.Formula.ToString + ""
                                    End If
                                    Explanation += vbCrLf + Me.FormatListVarConst(RowSL_Item.Identifier, TBLVar, TBLConst)
                                    Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + "0"
                                Else
                                    Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + "0"
                                End If
                                'BarisTerakhir
                                If Counter = TS.TotalDays Then
                                    TotalRounding = 0
                                    Explanation += vbCrLf + "Total" + vbTab + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                    If RowSL_Item.Rounding_Type = 0 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                                    ElseIf RowSL_Item.Rounding_Type = 1 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                                    ElseIf RowSL_Item.Rounding_Type = 2 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                                    ElseIf RowSL_Item.Rounding_Type = 3 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                                    End If
                                    If RowSL_Item.Rounding_Type = 0 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                        TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                    ElseIf RowSL_Item.Rounding_Type = 1 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD)
                                    ElseIf RowSL_Item.Rounding_Type = 2 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD)
                                    ElseIf RowSL_Item.Rounding_Type = 3 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD)
                                    End If
                                    If IsNumeric(RowSL_Item.Monthly_Max_Value) = True Then
                                        Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Max_Value, Me.Def_Prec)
                                    Else
                                        Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item.Monthly_Max_Value
                                    End If
                                    If IsNumeric(RowSL_Item.Monthly_Min_Value) = True Then
                                        Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Min_Value, Me.Def_Prec)
                                    Else
                                        Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item.Monthly_Min_Value
                                    End If
                                    If IsNumeric(RowSL_Item.Yearly_Max_Value) = True Then
                                        Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Max_Value, Me.Def_Prec)
                                    Else
                                        Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item.Yearly_Max_Value
                                    End If
                                    If IsNumeric(RowSL_Item.Yearly_Min_Value) = True Then
                                        Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                                    Else
                                        Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                                    End If
                                    Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                                End If
                            Else
                                'True
                                If Explanation = "" Then
                                    Explanation = RowSL_Item.Description.ToString
                                    If IsNumeric(RowSL_Item.Formula) = True Then
                                        Explanation += vbCrLf + vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec) + ""
                                        If RowSL_Item.Rounding_Type = 0 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec)
                                        ElseIf RowSL_Item.Rounding_Type = 1 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        ElseIf RowSL_Item.Rounding_Type = 2 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        ElseIf RowSL_Item.Rounding_Type = 3 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        End If
                                        Explanation += vbCrLf + Me.FormatListVarConst(RowSL_Item.Identifier, TBLVar, TBLConst)
                                        If RowSL_Item.Rounding_Type = 0 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec)
                                        ElseIf RowSL_Item.Rounding_Type = 1 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        ElseIf RowSL_Item.Rounding_Type = 2 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        ElseIf RowSL_Item.Rounding_Type = 3 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        End If
                                    Else
                                        Explanation += vbCrLf + vbCrLf + "Formula = " + RowSL_Item.Formula.ToString + ""
                                        If RowSL_Item.Rounding_Type = 0 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec)
                                        ElseIf RowSL_Item.Rounding_Type = 1 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        ElseIf RowSL_Item.Rounding_Type = 2 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        ElseIf RowSL_Item.Rounding_Type = 3 Then
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        End If
                                        Explanation += vbCrLf + Me.FormatListVarConst(RowSL_Item.Identifier, TBLVar, TBLConst)
                                        If RowSL_Item.Rounding_Type = 0 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec)
                                        ElseIf RowSL_Item.Rounding_Type = 1 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        ElseIf RowSL_Item.Rounding_Type = 2 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        ElseIf RowSL_Item.Rounding_Type = 3 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        End If
                                    End If
                                Else
                                    If IsNumeric(RowSL_Item.Formula) = True Then
                                        If RowSL_Item.Rounding_Type = 0 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec)
                                            TotalAmount += Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec)
                                        ElseIf RowSL_Item.Rounding_Type = 1 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        ElseIf RowSL_Item.Rounding_Type = 2 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        ElseIf RowSL_Item.Rounding_Type = 3 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(RowSL_Item.Formula, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        End If
                                    Else
                                        If RowSL_Item.Rounding_Type = 0 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec)
                                        ElseIf RowSL_Item.Rounding_Type = 1 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        ElseIf RowSL_Item.Rounding_Type = 2 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        ElseIf RowSL_Item.Rounding_Type = 3 Then
                                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                            TotalAmount += Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(Me.Obj.TranslateFormula_Daily(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                            Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        End If
                                    End If
                                End If

                                'BarisTerakhir
                                If Counter = TS.TotalDays Then
                                    TotalRounding = 0
                                    Explanation += vbCrLf + "Total" + vbTab + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                    If RowSL_Item.Rounding_Type = 0 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                                    ElseIf RowSL_Item.Rounding_Type = 1 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                                    ElseIf RowSL_Item.Rounding_Type = 2 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                                    ElseIf RowSL_Item.Rounding_Type = 3 Then
                                        Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                                    End If
                                    If RowSL_Item.Rounding_Type = 0 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                        TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                    ElseIf RowSL_Item.Rounding_Type = 1 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                        TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD)
                                    ElseIf RowSL_Item.Rounding_Type = 2 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                        TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD)
                                    ElseIf RowSL_Item.Rounding_Type = 3 Then
                                        Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                        TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD)
                                    End If
                                    If IsNumeric(RowSL_Item.Monthly_Max_Value) = True Then
                                        Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Max_Value, Me.Def_Prec)
                                    Else
                                        Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item.Monthly_Max_Value
                                    End If
                                    If IsNumeric(RowSL_Item.Monthly_Min_Value) = True Then
                                        Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Min_Value, Me.Def_Prec)
                                    Else
                                        Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item.Monthly_Min_Value
                                    End If
                                    If IsNumeric(RowSL_Item.Yearly_Max_Value) = True Then
                                        Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Max_Value, Me.Def_Prec)
                                    Else
                                        Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item.Yearly_Max_Value
                                    End If
                                    If IsNumeric(RowSL_Item.Yearly_Min_Value) = True Then
                                        Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                                    Else
                                        Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                                    End If

                                    TotalRounding = Me.Obj.EndResultValue_Final(True, RowSL_Item.Yearly_Max_Value, RowSL_Item.Yearly_Min_Value, RowSL_Item.Monthly_Max_Value, RowSL_Item.Monthly_Min_Value, _
                                    Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, RowSL_Item.Calculate_On, TotalRounding, RowSL_Item.Identifier, _
                                    RowSL_Item.Primary_Transaksi, Me.dtpPeriod.Value + TimeSpan.FromDays(Counter), _
                                    Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn)
                                    Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                                End If
                            End If
                            TglDaily = TglDaily.AddDays(1)
                            Counter += 1
                        End While
                    ElseIf RowSL_Item.Calculation_Period = 1 Then
                        'Monthly
                        TglDaily = Me.dtpPeriod.Value.Date
                        If Me.Obj.IsConstraint(RowSL_Item.Identifier, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, _
                        Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, TBLConst, Me.ObjKompetensi, Me.ObjKrywn) = False Then
                            'False
                            Explanation = RowSL_Item.Description.ToString
                            If IsNumeric(RowSL_Item.Formula) = True Then
                                Explanation += vbCrLf + vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec) + ""
                            Else
                                Explanation += vbCrLf + vbCrLf + "Formula = " + RowSL_Item.Formula.ToString + ""
                            End If
                            Explanation += vbCrLf + Me.FormatListVarConst(RowSL_Item.Identifier, TBLVar, TBLConst)
                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + "0"
                            'BarisTerakhir
                            TotalRounding = 0
                            Explanation += vbCrLf + "Total" + vbTab + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                            If RowSL_Item.Rounding_Type = 0 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                            ElseIf RowSL_Item.Rounding_Type = 1 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                            ElseIf RowSL_Item.Rounding_Type = 2 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                            ElseIf RowSL_Item.Rounding_Type = 3 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                            End If
                            If RowSL_Item.Rounding_Type = 0 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                            ElseIf RowSL_Item.Rounding_Type = 1 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD)
                            ElseIf RowSL_Item.Rounding_Type = 2 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD)
                            ElseIf RowSL_Item.Rounding_Type = 3 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD)
                            End If
                            If IsNumeric(RowSL_Item.Monthly_Max_Value) = True Then
                                Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Max_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item.Monthly_Max_Value
                            End If
                            If IsNumeric(RowSL_Item.Monthly_Min_Value) = True Then
                                Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Min_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item.Monthly_Min_Value
                            End If
                            If IsNumeric(RowSL_Item.Yearly_Max_Value) = True Then
                                Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Max_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item.Yearly_Max_Value
                            End If
                            If IsNumeric(RowSL_Item.Yearly_Min_Value) = True Then
                                Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                            Else
                                Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                            End If
                            Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                        Else
                            'True
                            Explanation = RowSL_Item.Description.ToString
                            If IsNumeric(RowSL_Item.Formula) = True Then
                                Explanation += vbCrLf + vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec) + ""
                                TotalAmount += Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec)
                                Explanation += vbCrLf + Me.FormatListVarConst(RowSL_Item.Identifier, TBLVar, TBLConst)
                                Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + vbCrLf + "Formula = " + RowSL_Item.Formula.ToString + ""
                                TotalAmount += Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormula(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, _
                                Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec)
                                Explanation += vbCrLf + Me.FormatListVarConst(RowSL_Item.Identifier, TBLVar, TBLConst)
                                Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormula(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, _
                                Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec)
                            End If
                            'BarisTerakhir
                            TotalRounding = 0
                            Explanation += vbCrLf + "Total" + vbTab + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                            If RowSL_Item.Rounding_Type = 0 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                            ElseIf RowSL_Item.Rounding_Type = 1 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                            ElseIf RowSL_Item.Rounding_Type = 2 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                            ElseIf RowSL_Item.Rounding_Type = 3 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                            End If
                            If RowSL_Item.Rounding_Type = 0 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                            ElseIf RowSL_Item.Rounding_Type = 1 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD).ToString, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD)
                            ElseIf RowSL_Item.Rounding_Type = 2 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD).ToString, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD)
                            ElseIf RowSL_Item.Rounding_Type = 3 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD).ToString, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD)
                            End If
                            If IsNumeric(RowSL_Item.Monthly_Max_Value) = True Then
                                Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Max_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item.Monthly_Max_Value
                            End If
                            If IsNumeric(RowSL_Item.Monthly_Min_Value) = True Then
                                Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Min_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item.Monthly_Min_Value
                            End If
                            If IsNumeric(RowSL_Item.Yearly_Max_Value) = True Then
                                Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Max_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item.Yearly_Max_Value
                            End If
                            If IsNumeric(RowSL_Item.Yearly_Min_Value) = True Then
                                Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                            Else
                                Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                            End If
                            TotalRounding = Me.Obj.EndResultValue_Final(False, RowSL_Item.Yearly_Max_Value, RowSL_Item.Yearly_Min_Value, RowSL_Item.Monthly_Max_Value, RowSL_Item.Monthly_Min_Value, _
                            Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, RowSL_Item.Calculate_On, TotalRounding, RowSL_Item.Identifier, _
                            RowSL_Item.Primary_Transaksi, Me.dtpPeriod.Value, Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn)
                            Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                        End If
                    ElseIf RowSL_Item.Calculation_Period = 2 Then
                        If RowSL_Item.Calculate_On = Me.EnumBln_CalculateOn(Me.DtpDate.Value.Date) Then
                            'Yearly
                            TglDaily = Me.dtpPeriod.Value.Date
                            If Me.Obj.IsConstraint(RowSL_Item.Identifier, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, _
                            Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, TBLConst, Me.ObjKompetensi, Me.ObjKrywn) = False Then
                                'False
                                Explanation = RowSL_Item.Description.ToString
                                If IsNumeric(RowSL_Item.Formula) = True Then
                                    Explanation += vbCrLf + vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec) + ""
                                Else
                                    Explanation += vbCrLf + vbCrLf + "Formula = " + RowSL_Item.Formula.ToString + ""
                                End If
                                Explanation += vbCrLf + Me.FormatListVarConst(RowSL_Item.Identifier, TBLVar, TBLConst)
                                Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + "0"
                                'BarisTerakhir
                                TotalRounding = 0
                                Explanation += vbCrLf + "Total" + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                If RowSL_Item.Rounding_Type = 0 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                                ElseIf RowSL_Item.Rounding_Type = 1 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                                ElseIf RowSL_Item.Rounding_Type = 2 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                                ElseIf RowSL_Item.Rounding_Type = 3 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                                End If
                                If RowSL_Item.Rounding_Type = 0 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                    TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                ElseIf RowSL_Item.Rounding_Type = 1 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                    TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD)
                                ElseIf RowSL_Item.Rounding_Type = 2 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                    TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD)
                                ElseIf RowSL_Item.Rounding_Type = 3 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                    TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD)
                                End If
                                If IsNumeric(RowSL_Item.Monthly_Max_Value) = True Then
                                    Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Max_Value, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item.Monthly_Max_Value
                                End If
                                If IsNumeric(RowSL_Item.Monthly_Min_Value) = True Then
                                    Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Min_Value, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item.Monthly_Min_Value
                                End If
                                If IsNumeric(RowSL_Item.Yearly_Max_Value) = True Then
                                    Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Max_Value, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item.Yearly_Max_Value
                                End If
                                If IsNumeric(RowSL_Item.Yearly_Min_Value) = True Then
                                    Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                                Else
                                    Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                                End If
                                Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                            Else
                                'True
                                Explanation = RowSL_Item.Description.ToString
                                If IsNumeric(RowSL_Item.Formula) = True Then
                                    Explanation += vbCrLf + vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec) + ""
                                    TotalAmount += Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec)
                                    Explanation += vbCrLf + Me.FormatListVarConst(RowSL_Item.Identifier, TBLVar, TBLConst)
                                    Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + vbCrLf + "Formula = " + RowSL_Item.Formula.ToString + ""
                                    TotalAmount += Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormula(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, _
                                    Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec)
                                    Explanation += vbCrLf + Me.FormatListVarConst(RowSL_Item.Identifier, TBLVar, TBLConst)
                                    Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + Me.ObjCD.DigitGrouping(Me.Obj.TranslateFormula(RowSL_Item.Formula, RowSL_Item.Primary_Transaksi, Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, _
                                    Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, RowSL_Item.Identifier, Me.ObjKrywn), Me.Def_Prec)
                                End If
                                'BarisTerakhir
                                TotalRounding = 0
                                Explanation += vbCrLf + "Total" + vbTab + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                If RowSL_Item.Rounding_Type = 0 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                                ElseIf RowSL_Item.Rounding_Type = 1 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                                ElseIf RowSL_Item.Rounding_Type = 2 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                                ElseIf RowSL_Item.Rounding_Type = 3 Then
                                    Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                                End If
                                If RowSL_Item.Rounding_Type = 0 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                    TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                ElseIf RowSL_Item.Rounding_Type = 1 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                    TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD)
                                ElseIf RowSL_Item.Rounding_Type = 2 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                    TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD)
                                ElseIf RowSL_Item.Rounding_Type = 3 Then
                                    Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                    TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD)
                                End If
                                If IsNumeric(RowSL_Item.Monthly_Max_Value) = True Then
                                    Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Max_Value, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item.Monthly_Max_Value
                                End If
                                If IsNumeric(RowSL_Item.Monthly_Min_Value) = True Then
                                    Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Min_Value, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item.Monthly_Min_Value
                                End If
                                If IsNumeric(RowSL_Item.Yearly_Max_Value) = True Then
                                    Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Max_Value, Me.Def_Prec)
                                Else
                                    Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item.Yearly_Max_Value
                                End If
                                If IsNumeric(RowSL_Item.Yearly_Min_Value) = True Then
                                    Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Min_Value, Me.Def_Prec) + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                                Else
                                    Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item.Yearly_Min_Value + "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                                End If
                                TotalRounding = Me.Obj.EndResultValue_Final(False, RowSL_Item.Yearly_Max_Value, RowSL_Item.Yearly_Min_Value, RowSL_Item.Monthly_Max_Value, RowSL_Item.Monthly_Min_Value, _
                                Me.TxtPrimaryEmployee.Text, Me.dtpPeriod.Value, RowSL_Item.Calculate_On, TotalRounding, RowSL_Item.Identifier, _
                                RowSL_Item.Primary_Transaksi, Me.dtpPeriod.Value, Me.DtpTo.Value, Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.ObjParam, Me.ObjKrywn)
                                Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                            End If
                        Else
                            'Jika bukan dalam bulan calculate on maka ikut in kondisi isconstraint = false 
                            Explanation = RowSL_Item.Description.ToString
                            If IsNumeric(RowSL_Item.Formula) = True Then
                                Explanation += vbCrLf + vbCrLf + "Formula = " + Me.ObjCD.DigitGrouping(RowSL_Item.Formula, Me.Def_Prec) + ""
                            Else
                                Explanation += vbCrLf + vbCrLf + "Formula = " + RowSL_Item.Formula.ToString + ""
                            End If
                            Explanation += vbCrLf + Me.FormatListVarConst(RowSL_Item.Identifier, TBLVar, TBLConst)
                            Explanation += vbCrLf + Format(TglDaily, "dd-MMM-yyyy").ToString + vbTab + "0"
                            'BarisTerakhir
                            TotalRounding = 0
                            Explanation += vbCrLf + "Total" + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                            If RowSL_Item.Rounding_Type = 0 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = no rounding"
                            ElseIf RowSL_Item.Rounding_Type = 1 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding down"
                            ElseIf RowSL_Item.Rounding_Type = 2 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = rounding middle point"
                            ElseIf RowSL_Item.Rounding_Type = 3 Then
                                Explanation += vbCrLf + "Rounding Unit = " + RowSL_Item.Rounding_Unit.ToString + "; " + "Rounding Type = round up"
                            End If
                            If RowSL_Item.Rounding_Type = 0 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                                TotalRounding = Me.ObjCD.DigitGrouping(TotalAmount, Me.Def_Prec)
                            ElseIf RowSL_Item.Rounding_Type = 1 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Down)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Down, RowSL_Item.Rounding_Unit, Me.ObjCD)
                            ElseIf RowSL_Item.Rounding_Type = 2 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Middlepoint, RowSL_Item.Rounding_Unit, Me.ObjCD)
                            ElseIf RowSL_Item.Rounding_Type = 3 Then
                                Explanation += vbCrLf + "Total Rounding = " + Me.ObjCD.DigitGrouping(Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD), Me.Def_Prec, ClsCekData.EnumTipeRounding.Up)
                                TotalRounding = Me.ObjCD.GetRoundingValue(TotalAmount, Me.Def_Prec, ClsCekData.EnumTipeRounding.Up, RowSL_Item.Rounding_Unit, Me.ObjCD)
                            End If
                            If IsNumeric(RowSL_Item.Monthly_Max_Value) = True Then
                                Explanation += vbCrLf + "Monthly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Max_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Monthly Max Value = " + RowSL_Item.Monthly_Max_Value
                            End If
                            If IsNumeric(RowSL_Item.Monthly_Min_Value) = True Then
                                Explanation += vbCrLf + "Monthly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Monthly_Min_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Monthly Min Value = " + RowSL_Item.Monthly_Min_Value
                            End If
                            If IsNumeric(RowSL_Item.Yearly_Max_Value) = True Then
                                Explanation += vbCrLf + "Yearly Max Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Max_Value, Me.Def_Prec)
                            Else
                                Explanation += vbCrLf + "Yearly Max Value = " + RowSL_Item.Yearly_Max_Value
                            End If
                            If IsNumeric(RowSL_Item.Yearly_Min_Value) = True Then
                                Explanation += vbCrLf + "Yearly Min Value = " + Me.ObjCD.DigitGrouping(RowSL_Item.Yearly_Min_Value, Me.Def_Prec) + _
                                "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                            Else
                                Explanation += vbCrLf + "Yearly Min Value = " + RowSL_Item.Yearly_Min_Value + _
                                "; " + "Calculate On = " + Format(CDate("#" + Me.ObjMstSlry_Item.ConversionMonth(RowSL_Item.Calculate_On).ToString + "/1/1950#"), "MMM")
                            End If
                            Explanation += vbCrLf + "End Result = " + Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                        End If
                    End If
                    RowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(TotalRounding, Me.Def_Prec)
                    RowDGV.Cells(Me.ColExplaination.Index).Value = Explanation
                    RowDGV.Cells(Me.ColOriginalValue.Index).Value = RowDGV.Cells(Me.ColAmount.Index).Value
                    RowDGV.Cells(Me.ColIdentifierKomponen.Index).Value = RowSL_Item.Identifier
                    RowDGV.Cells(Me.ColOverride.Index).Value = False
                    RowDGV.Cells(Me.ColOverrideReason.Index).Value = ""
                    RowDGV.Cells(Me.ColRemarks.Index).Value = ""
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'DTAVar.Dispose()
            TBLVar.Rows.Clear()
            TBLVar.Dispose()
            'DTAConst.Dispose()
            TBLConst.Rows.Clear()
            TBLConst.Dispose()
            TBLSL_Item.Rows.Clear()
        End Try
    End Sub

    Private Function EnumBln_CalculateOn(ByVal vTanggalTransaksi As Date) As Integer
        Dim EnumBln As Core_Module.ClsMasterSalaryFormula_item.EnumCalculateOn
        Dim Bulan As Integer = vTanggalTransaksi.Date.Month
        Try
            If Bulan = 1 Then
                EnumBln = ClsMasterSalaryFormula_item.EnumCalculateOn.Jan
            ElseIf Bulan = 2 Then
                EnumBln = ClsMasterSalaryFormula_item.EnumCalculateOn.Feb
            ElseIf Bulan = 3 Then
                EnumBln = ClsMasterSalaryFormula_item.EnumCalculateOn.Mar
            ElseIf Bulan = 4 Then
                EnumBln = ClsMasterSalaryFormula_item.EnumCalculateOn.Apr
            ElseIf Bulan = 5 Then
                EnumBln = ClsMasterSalaryFormula_item.EnumCalculateOn.May
            ElseIf Bulan = 6 Then
                EnumBln = ClsMasterSalaryFormula_item.EnumCalculateOn.June
            ElseIf Bulan = 7 Then
                EnumBln = ClsMasterSalaryFormula_item.EnumCalculateOn.Jul
            ElseIf Bulan = 8 Then
                EnumBln = ClsMasterSalaryFormula_item.EnumCalculateOn.Agt
            ElseIf Bulan = 9 Then
                EnumBln = ClsMasterSalaryFormula_item.EnumCalculateOn.Sept
            ElseIf Bulan = 10 Then
                EnumBln = ClsMasterSalaryFormula_item.EnumCalculateOn.Oct
            ElseIf Bulan = 11 Then
                EnumBln = ClsMasterSalaryFormula_item.EnumCalculateOn.Nov
            ElseIf Bulan = 12 Then
                EnumBln = ClsMasterSalaryFormula_item.EnumCalculateOn.Dec
            End If
            Return EnumBln
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub SetNumber(ByVal dgv As DataGridView, Optional ByVal vSetAutosize As Boolean = True, _
    Optional ByVal BerikutYgdHide As Boolean = True)
        Dim NomorBaru As Integer = 1
        Dim Putih As Boolean = False
        Dim Row As DataGridViewRow
        Try
            If dgv.RowCount = 0 Then
                Exit Sub
            End If
            If vSetAutosize = True Then
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            End If
            dgv.DefaultCellStyle.BackColor = Color.White
            If BerikutYgdHide = False Then
                For i As Integer = 1 To dgv.Rows.Count
                    Row = Me.dgvItem.Rows(i - 1)
                    If dgv.Rows(i - 1).Visible = True Then
                        dgv.Item(0, i - 1).Value = NomorBaru
                        NomorBaru += 1
                        If Putih = False Then
                            Row.DefaultCellStyle.BackColor = Color.White
                            Putih = True
                        Else
                            Row.DefaultCellStyle.BackColor = Me.BackColorAll
                            Putih = False
                        End If
                    End If
                Next
            Else
                For i As Integer = 1 To dgv.Rows.Count
                    Row = Me.dgvItem.Rows(i - 1)
                    dgv.Item(0, i - 1).Value = i
                    If Putih = False Then
                        Row.DefaultCellStyle.BackColor = Color.White
                        Putih = True
                    Else
                        Row.DefaultCellStyle.BackColor = Me.BackColorAll
                        Putih = False
                    End If
                Next
            End If
            If vSetAutosize = True Then
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungOverride(ByVal vOtherDeduction As Boolean, ByVal vPeriodFrom As Date, _
    ByVal vPeriodTo As Date)
        Dim rowDGV As DataGridViewRow
        Dim DTAHitung As New DseTransaksiSalary_OverrideTableAdapters.tblhitung1TableAdapter
        'Dim DTAReason As New DseTransaksiSalary_OverrideTableAdapters.tblhitung2TableAdapter
        Dim TBLHitung As New DseTransaksiSalary_Override.tblhitung1DataTable
        'Dim TBLReason As New DseTransaksiSalary_Override.tblhitung2DataTable
        Dim ParamStr As String = ""
        Dim ParamDec As Decimal = 0
        Dim DataFound As Boolean = False

        Try
            DTAHitung.Connection.ConnectionString = Core_Module.GetConstring
            'DTAReason.Connection.ConnectionString = Core_Module.GetConstring
            'load data override berdsarkan tanggal dari sampai, primary karyawan, other deduction atau bukan, untuk amount harus sum
            If vOtherDeduction = True Then
                DTAHitung.FillBy_HitungByOthers(TBLHitung, Me.DtpTo.Value.Date, Me.dtpPeriod.Value.Date)
                'DTAReason.FillBy_ReasonByOthers(TBLReason)
                'Dari Other Deduction
                If Me.FrmOtherDeduction.dgvItem.Rows.Count > 0 Then
                    For id As Integer = 0 To Me.FrmOtherDeduction.dgvItem.Rows.Count - 1
                        rowDGV = Me.FrmOtherDeduction.dgvItem.Rows(id)
                        ParamDec = Me.GetDecimalAmount(TBLHitung, CDbl(rowDGV.Cells(FrmOtherDeduction.ColIdentifierKomponen.Index).Value), "-1-1", True, _
                        Me.TxtPrimaryEmployee.Text, Me.DtpTo.Value.Date, Me.dtpPeriod.Value.Date, ParamStr, DataFound)
                        'ParamStr = Me.GetStringReason(TBLHitung, CDbl(rowDGV.Cells(FrmOtherDeduction.ColIdentifierKomponen.Index).Value), "-1-1", True, _
                        'Me.TxtPrimaryEmployee.Text, Me.DtpTo.Value.Date, Me.dtpPeriod.Value.Date)
                        If ParamDec <> 0 Then
                            rowDGV.Cells(FrmOtherDeduction.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(ParamDec, Me.Def_Prec)
                            rowDGV.Cells(FrmOtherDeduction.ColOverrideReason.Index).Value = ParamStr
                            rowDGV.Cells(FrmOtherDeduction.ColOverride.Index).Value = True
                        End If
                    Next
                End If
            Else
                DTAHitung.FillBy_HitungBySalaryFormula1(TBLHitung, Me.DtpTo.Value.Date, Me.dtpPeriod.Value.Date, Me.TxtPrimaryEmployee.Text)
                'DTAReason.FillBy_ReasonBySalaryFormula(TBLReason)
                'Dari Salary Formula
                If Me.dgvItem.Rows.Count > 0 Then
                    For id As Integer = 0 To Me.dgvItem.Rows.Count - 1
                        rowDGV = Me.dgvItem.Rows(id)
                        ParamDec = Me.GetDecimalAmount(TBLHitung, -1, rowDGV.Cells(Me.ColIdentifierKomponen.Index).Value, False, _
                        Me.TxtPrimaryEmployee.Text, Me.DtpTo.Value.Date, Me.dtpPeriod.Value.Date, ParamStr, DataFound)
                        'ParamStr = Me.GetStringReason(TBLHitung, -1, rowDGV.Cells(Me.ColIdentifierKomponen.Index).Value, False, _
                        'Me.TxtPrimaryEmployee.Text, Me.DtpTo.Value.Date, Me.dtpPeriod.Value.Date)

                        If ParamDec <> 0 Or DataFound = True Then
                            rowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(ParamDec, Me.Def_Prec)
                            rowDGV.Cells(Me.ColOverrideReason.Index).Value = ParamStr
                            rowDGV.Cells(Me.ColOverride.Index).Value = True
                        End If
                    Next
                End If
            End If
            'untuk setiap komponen yang ada dalam dgv, cari komponen yang bersesuaian identifiernya, lalu override
            'masukkan override amount + override reason + check list
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAHitung.Dispose()
            'DTAReason.Dispose()
            TBLHitung.Dispose()
            'TBLReason.Dispose()
        End Try
    End Sub

    Private Function GetDecimalAmount(ByVal TBL As DseTransaksiSalary_Override.tblhitung1DataTable, _
ByVal vPrimaryFormula As Double, ByVal vIdentifier_Formula As String, ByVal vOtherDeduction As Boolean, _
ByVal vPrimaryEmployee As Double, ByVal vPeriodEnd As Date, ByVal vPeriodStart As Date, ByRef vReason As String, ByRef DataFound As Boolean) As Decimal
        Dim Row As DseTransaksiSalary_Override.tblhitung1Row
        Dim Retval As Decimal = 0
        Dim Filter As String = ""
        Dim RowView As DataRowView

        Try
            vReason = ""
            DataFound = False

            If vOtherDeduction = False Then
                'Other Deduction
                Filter = "[Identifier Formula] = '" + vIdentifier_Formula.ToString + "' AND [Primary Employee] = " + vPrimaryEmployee.ToString _
                + " AND [Date] <= #" + Format(vPeriodEnd, "MM/dd/yyyy").ToString + "# AND [Date] >= #" + Format(vPeriodStart, "MM/dd/yyyy").ToString + "#"
            Else
                'Salary Formula
                Filter = "[Primary Formula] = " + vPrimaryFormula.ToString + " AND [Primary Employee] = " + vPrimaryEmployee.ToString _
                 + " AND [Date] <= #" + Format(vPeriodEnd, "MM/dd/yyyy").ToString + "# AND [Date] >= #" + Format(vPeriodStart, "MM/dd/yyyy").ToString + "#"
            End If
            TBL.DefaultView.RowFilter = Filter
            If TBL.DefaultView.Count > 0 Then
                DataFound = True

                For Each RowView In TBL.DefaultView
                    Row = RowView.Row
                    Retval += Row.Override_Amount

                    If String.IsNullOrEmpty(Row.Override_Reason) = False Then
                        If String.IsNullOrEmpty(vReason) = True Then
                            vReason = Row.Override_Reason
                        Else
                            vReason += ", " + Row.Override_Reason
                        End If
                    End If
                Next
                If vReason.Length > 250 Then
                    vReason = vReason.Substring(0, 250)
                    If vReason.EndsWith(", ") = True Or vReason.EndsWith(",") = True Then
                        vReason = vReason.Remove(vReason.IndexOf(","))
                    End If
                End If
            End If
            TBL.DefaultView.RowFilter = ""
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Function GetStringReason(ByVal TBL As DseTransaksiSalary_Override.tblhitung1DataTable, _
    'ByVal vPrimaryFormula As Double, ByVal vIdentifier_Formula As String, ByVal vOtherDeduction As Boolean, _
    'ByVal vPrimaryEmployee As Double, ByVal vPeriodEnd As Date, ByVal vPeriodStart As Date) As String
    '    Dim Row As DseTransaksiSalary_Override.tblhitung1Row
    '    Dim Retval As String = ""
    '    Dim Filter As String = ""
    '    Dim RowView As DataRowView
    '    Try
    '        If vOtherDeduction = False Then
    '            'Other Deduction
    '            Filter = "[Identifier Formula] = '" + vIdentifier_Formula.ToString + "' AND [Primary Employee] = " + vPrimaryEmployee.ToString _
    '             + " AND [Date] <= #" + Format(vPeriodEnd, "MM/dd/yyyy").ToString + "# AND [Date] >= #" + Format(vPeriodStart, "MM/dd/yyyy").ToString + "#"
    '        Else
    '            'Salary Formula
    '            Filter = "[Primary Formula] = " + vPrimaryFormula.ToString + " AND [Primary Employee] = " + vPrimaryEmployee.ToString _
    '             + " AND [Date] <= #" + Format(vPeriodEnd, "MM/dd/yyyy").ToString + "# AND [Date] >= #" + Format(vPeriodStart, "MM/dd/yyyy").ToString + "#"
    '        End If
    '        TBL.DefaultView.RowFilter = Filter
    '        If TBL.DefaultView.Count > 0 Then
    '            For Each RowView In TBL.DefaultView
    '                Row = RowView.Row
    '                If String.IsNullOrEmpty(Row.Override_Reason) = False Then
    '                    If String.IsNullOrEmpty(Retval) = True Then
    '                        Retval = Row.Override_Reason
    '                    Else
    '                        Retval += ", " + Row.Override_Reason
    '                    End If
    '                End If
    '            Next
    '            If Retval.Length > 25 Then
    '                Retval = Retval.Substring(0, 25)
    '                If Retval.EndsWith(", ") = True Or Retval.EndsWith(",") = True Then
    '                    Retval = Retval.Remove(Retval.IndexOf(","))
    '                End If
    '            End If
    '        End If
    '        TBL.DefaultView.RowFilter = ""
    '        Return Retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    End Try
    'End Function

    Private Function GetPrimaryMax(ByVal vTBLItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iterationDataTable) As Double
        Dim rowItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iterationRow
        Dim RetVal As Double = Double.MinValue

        Try
            For Each rowItemIterasi In vTBLItemIterasi.Rows
                If RetVal < rowItemIterasi.primary Then
                    RetVal = rowItemIterasi.primary
                End If
            Next

            If RetVal = Double.MinValue Then
                Return 0
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax(ByVal vTBLItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iteration1DataTable) As Double
        Dim rowItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iteration1Row
        Dim RetVal As Double = Double.MinValue

        Try
            For Each rowItemIterasi In vTBLItemIterasi.Rows
                If RetVal < rowItemIterasi.primary Then
                    RetVal = rowItemIterasi.primary
                End If
            Next

            If RetVal = Double.MinValue Then
                Return 0
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTipeElemenPajak(ByVal vParameter As ClsMasterSalaryFormula_item.EnumTax2) As ClsTransaksiSalary_Item.EnumTaxElement
        Try
            If vParameter = ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun Then
                Return ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT
            ElseIf vParameter = ClsMasterSalaryFormula_item.EnumTax2.honorarium Then
                Return ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya
            ElseIf vParameter = ClsMasterSalaryFormula_item.EnumTax2.iuran_pensiun Then
                Return ClsTransaksiSalary_Item.EnumTaxElement.No10_IuranPensiun_THT_JHT
            ElseIf vParameter = ClsMasterSalaryFormula_item.EnumTax2.no_tax Then
                Return ClsTransaksiSalary_Item.EnumTaxElement.NoTax
            ElseIf vParameter = ClsMasterSalaryFormula_item.EnumTax2.penerimaan_dalam_bentuk_natura Then
                Return ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH
            ElseIf vParameter = ClsMasterSalaryFormula_item.EnumTax2.premi_asuransi Then
                Return ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja
            ElseIf vParameter = ClsMasterSalaryFormula_item.EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr Then
                Return ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi
            ElseIf vParameter = ClsMasterSalaryFormula_item.EnumTax2.tunjangan_lainnya Then
                Return ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb
            ElseIf vParameter = ClsMasterSalaryFormula_item.EnumTax2.pesangon Then
                Return ClsTransaksiSalary_Item.EnumTaxElement.No21_Pesangon
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetAmountMonth(ByVal vAmount As Decimal, ByVal TipeSalaryKomponen As FrmMasterSalaryFormula_Component.EnumType, _
    ByVal TipeElemenPajak As ClsTransaksiSalary_Item.EnumTaxElement) As Decimal
        Try
            If TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT Or _
            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb Or _
            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya Or _
            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja Or _
            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH Or _
            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi Or _
            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No21_Pesangon Then
                If TipeSalaryKomponen = FrmMasterSalaryFormula_Component.EnumType.Income Or _
                TipeSalaryKomponen = FrmMasterSalaryFormula_Component.EnumType.CompanyExpense Then
                    Return vAmount
                ElseIf TipeSalaryKomponen = FrmMasterSalaryFormula_Component.EnumType.Deduction Or _
                TipeSalaryKomponen = FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense Then
                    Return vAmount * -1
                End If
            ElseIf TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No10_IuranPensiun_THT_JHT Then

                Return vAmount
                'If TipeSalaryKomponen = FrmMasterSalaryFormula_Component.EnumType.Income Or _
                'TipeSalaryKomponen = FrmMasterSalaryFormula_Component.EnumType.CompanyExpense Then
                '    Return vAmount * -1
                'ElseIf TipeSalaryKomponen = FrmMasterSalaryFormula_Component.EnumType.Deduction Or _
                'TipeSalaryKomponen = FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense Then
                '    Return vAmount
                'End If
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsikanItemOtherPayment()
        Dim DTA1 As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_other2TableAdapter
        Dim TBL1 As New DseTransaksiSalary.tbltrans_salary_item_other2DataTable
        Dim row1 As DseTransaksiSalary.tbltrans_salary_item_other2Row
        Dim rowOther As DseTransaksiSalary.tbltrans_salary_item_other1Row
        Dim Amount As Decimal = 0

        Try
            Me.vTBLOtherPayment.Rows.Clear()
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Fill(TBL1, Me.TxtPrimaryEmployee.Text, Me.DtpTo.Value.Date, Me.dtpPeriod.Value.Date)
            For Each row1 In TBL1.Rows
                rowOther = Me.vTBLOtherPayment.Newtbltrans_salary_item_other1Row
                With rowOther
                    .Amount = row1.Override_Amount
                    Amount += row1.Override_Amount
                    .Description = row1.Description
                    .Document_No = row1.Document_No
                    .Identifier_Salary_Component = row1.Identifier_Formula
                    .Is_Rutin = row1.rutin_disetahunkan
                    .Primary = Me.vTBLOtherPayment.Rows.Count + 1
                    .Primary_Ref_Doc = row1.Primary_Ref_Doc
                    .Primary_Transaksi = Me.TxtPrimary.Text
                    .Remarks = row1.Override_Reason
                    .Tax_Component = row1.Tax_Component
                    .Type = row1.Type
                    ._Date = row1._Date
                End With
                Me.vTBLOtherPayment.Rows.Add(rowOther)
            Next

            Me.TxtOtherPayment.Text = Me.ObjCD.DigitGrouping(Amount, Me.ObjCurr.GetDecimal_DefaultCurrency)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Private Sub IsikanItemOtherPayment_EndOfYear_Resign(ByVal vStartDate As Date, ByVal vEndDate As Date, _
    ByVal vTBLOtherPayment1 As DseTransaksiSalary.tbltrans_salary_item_other1DataTable)
        Dim DTA1 As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_other2TableAdapter
        Dim TBL1 As New DseTransaksiSalary.tbltrans_salary_item_other2DataTable
        Dim row1 As DseTransaksiSalary.tbltrans_salary_item_other2Row
        Dim rowOther As DseTransaksiSalary.tbltrans_salary_item_other1Row
        Dim Amount As Decimal = 0
        Dim Dummy As Double = -1

        Try
            If vEndDate = #1/1/1950# Then
                'karyawan ybs bukan resign, tapi hanya perhitungan akhir tahun biasa saj
                vEndDate = Me.dtpPeriod.Value.Date - TimeSpan.FromDays(1)
            ElseIf vEndDate.Date >= Me.dtpPeriod.Value.Date Then
                vEndDate = Me.dtpPeriod.Value.Date - TimeSpan.FromDays(1)
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If vStartDate.Year < Me.dtpPeriod.Value.Year Then
                vStartDate = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpPeriod.Value)
            End If
            Dummy = Me.Obj.GetPrimary_ByPeriod(vStartDate, Me.TxtPrimaryEmployee.Text)
            If Dummy = -1 Then
                vTBLOtherPayment1.Rows.Clear()
                'MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Exit Sub
            Else
                vStartDate = Me.Obj.GetStartDate(Dummy)
            End If

            vTBLOtherPayment1.Rows.Clear()
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Fill(TBL1, Me.TxtPrimaryEmployee.Text, vEndDate, vStartDate)
            For Each row1 In TBL1.Rows
                rowOther = vTBLOtherPayment1.Newtbltrans_salary_item_other1Row
                With rowOther
                    .Amount = row1.Override_Amount
                    'Amount += row1.Override_Amount
                    .Description = row1.Description
                    .Document_No = row1.Document_No
                    .Identifier_Salary_Component = row1.Identifier_Formula
                    .Is_Rutin = row1.rutin_disetahunkan
                    .Primary = vTBLOtherPayment1.Rows.Count + 1
                    .Primary_Ref_Doc = row1.Primary_Ref_Doc
                    .Primary_Transaksi = Me.TxtPrimary.Text
                    .Remarks = row1.Override_Reason
                    .Tax_Component = row1.Tax_Component
                    .Type = row1.Type
                    ._Date = row1._Date
                End With
                vTBLOtherPayment1.Rows.Add(rowOther)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Private Function GetAmount_OtherPayment(ByVal vIdentifierSalaryFormula As String) As Decimal
        Dim RetVal As Decimal = 0
        Dim row As DseTransaksiSalary.tbltrans_salary_item_other1Row

        Try
            Me.vTBLOtherPayment.DefaultView.RowFilter = "[Identifier Salary Component] = '" + vIdentifierSalaryFormula + "'"
            For Each DV As DataRowView In Me.vTBLOtherPayment.DefaultView
                row = DV.Row
                RetVal += row.Amount
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahAdaDalamDGV(ByVal vIdentifier_OtherPayment As String) As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColIdentifierKomponen.Index).Value = vIdentifier_OtherPayment Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub HitungPajak1(Optional ByVal vPPhRutin As Decimal = 0, Optional ByVal vPPhNonRutin As Decimal = 0)
        Dim rowDGV As DataGridViewRow
        Dim rowItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Dim rowItemIterasi, rowItemIterasiSebelum As DseTransaksiSalary.tbltrans_salary_item_iterationRow
        Dim CboDgv As DataGridViewComboBoxCell
        Dim TipeSalaryKomponen As FrmMasterSalaryFormula_Component.EnumType = FrmMasterSalaryFormula_Component.EnumType.Income
        Dim TipeElemenPajak As ClsTransaksiSalary_Item.EnumTaxElement
        Dim DTATarif As New DseTransaksiSalaryTableAdapters.tblmaster_tarif_pphTableAdapter
        Dim JmlIterasi As Double = 0
        Dim Selisih1, Selisih2 As Decimal
        Dim UseNPWP As ClsMasterKaryawan.EnumIsNPWP = ClsMasterKaryawan.EnumIsNPWP.Yes
        Dim FaktorKaliDisetahunkan As Integer = 12
        Dim StartWorkingDate As Date = Today
        Dim ResignDate As Date = #1/1/1950#
        Dim vResign As Boolean = False
        Dim vTBLOtherPayment1 As New DseTransaksiSalary.tbltrans_salary_item_other1DataTable
        Dim NPWPCategory As ClsMasterKaryawan.GolonganPajak = ClsMasterKaryawan.GolonganPajak.K0
        Dim TaxCategoryDesc As String = ""
        Dim RowTaxBB As DseTrans_Tax_Beginning_Balance.tbltrans_tax_beginning_balance_item1Row
        Dim Amount As Decimal = 0
        Dim Remarks1 As String = ""
        Dim LosesTaxObligation As Integer = 0
        Dim LanjutIterasi As Boolean = True

        Try
            ResignDate = Me.ObjKrywn.GetResignDate(Me.TxtPrimaryEmployee.Text)
            LosesTaxObligation = Me.ObjKrywn.GetLosesTaxObligation(Me.TxtPrimaryEmployee.Text)

            Me.TBLTaxBB.Rows.Clear()
            Me.TBLTaxBB2.Rows.Clear()

            Me.vPenghasilanBruto = 0
            Me.IsikanItemOtherPayment()

            StartWorkingDate = Me.ObjKrywn.GetFirstTaxPeriod(Me.TxtPrimaryEmployee.Text)
            If ResignDate >= Me.dtpPeriod.Value.Date And ResignDate <= Me.DtpTo.Value Then
                vResign = True
            Else
                vResign = False
            End If
            If vResign = True Then
                Me.TxtIsResign.Text = "1"
            Else
                Me.TxtIsResign.Text = "0"
            End If
            Me.TxtLosesTaxObligation.Text = LosesTaxObligation

            If StartWorkingDate.Date.Year < Me.ObjCD.GetFirstDateofTheMonth(Me.DtpTo.Value.Date).Date.Year Then
                If ResignDate = #1/1/1950# Then
                    FaktorKaliDisetahunkan = 12

                    Me.TxtPeriodStart.Text = "1"
                    Me.TxtPeriodEnd.Text = "12"
                    Me.TxtPeriodCount.Text = "12"
                Else
                    If ResignDate.Date <= Me.DtpTo.Value.Date And ResignDate >= Me.dtpPeriod.Value.Date Then
                        FaktorKaliDisetahunkan = ResignDate.Month

                        Me.TxtPeriodStart.Text = "1"
                        Me.TxtPeriodEnd.Text = ResignDate.Month
                        Me.TxtPeriodCount.Text = ResignDate.Month
                    Else
                        'MsgBox("Logic error: no code here.", MsgBoxStyle.Information, "Information")
                        FaktorKaliDisetahunkan = 12
                        Me.TxtPeriodStart.Text = "1"
                        Me.TxtPeriodEnd.Text = "12"
                        Me.TxtPeriodCount.Text = "12"
                        Me.TxtIsResign.Text = "0"
                        'Exit Sub


                        'Me.TxtPeriodStart.Text = "1"
                        'Me.TxtPeriodEnd.Text = "12"
                        'Me.TxtPeriodCount.Text = "12"
                        'MsgBox(Me.TxtPeriodStart.Text + " - " + Me.TxtPeriodEnd.Text + " - " + Me.TxtPeriodCount.Text)
                    End If
                End If
            ElseIf StartWorkingDate.Date.Year = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpTo.Value.Date).Date.Year Then
                If ResignDate = #1/1/1950# Then
                    If StartWorkingDate.Month = 1 Then
                        FaktorKaliDisetahunkan = 12
                    ElseIf StartWorkingDate.Month = 2 Then
                        FaktorKaliDisetahunkan = 11
                    ElseIf StartWorkingDate.Month = 3 Then
                        FaktorKaliDisetahunkan = 10
                    ElseIf StartWorkingDate.Month = 4 Then
                        FaktorKaliDisetahunkan = 9
                    ElseIf StartWorkingDate.Month = 5 Then
                        FaktorKaliDisetahunkan = 8
                    ElseIf StartWorkingDate.Month = 6 Then
                        FaktorKaliDisetahunkan = 7
                    ElseIf StartWorkingDate.Month = 7 Then
                        FaktorKaliDisetahunkan = 6
                    ElseIf StartWorkingDate.Month = 8 Then
                        FaktorKaliDisetahunkan = 5
                    ElseIf StartWorkingDate.Month = 9 Then
                        FaktorKaliDisetahunkan = 4
                    ElseIf StartWorkingDate.Month = 10 Then
                        FaktorKaliDisetahunkan = 3
                    ElseIf StartWorkingDate.Month = 11 Then
                        FaktorKaliDisetahunkan = 2
                    ElseIf StartWorkingDate.Month = 12 Then
                        FaktorKaliDisetahunkan = 1
                    End If

                    Me.TxtPeriodStart.Text = StartWorkingDate.Month
                    Me.TxtPeriodEnd.Text = 12
                    Me.TxtPeriodCount.Text = FaktorKaliDisetahunkan
                Else
                    If ResignDate.Date <= Me.DtpTo.Value.Date And ResignDate >= Me.dtpPeriod.Value.Date Then
                        FaktorKaliDisetahunkan = ResignDate.Month - StartWorkingDate.Month + 1

                        Me.TxtPeriodStart.Text = StartWorkingDate.Month
                        Me.TxtPeriodEnd.Text = ResignDate.Month
                        Me.TxtPeriodCount.Text = FaktorKaliDisetahunkan
                    Else
                        'MsgBox("Logic error: no code here.", MsgBoxStyle.Information, "Information")
                        'Me.TxtPeriodStart.Text = "1"
                        'Me.TxtPeriodEnd.Text = "12"
                        'Me.TxtPeriodCount.Text = "12"
                        'Me.TxtIsResign.Text = "0"
                        'Exit Sub

                        If StartWorkingDate.Month = 1 Then
                            FaktorKaliDisetahunkan = 12
                        ElseIf StartWorkingDate.Month = 2 Then
                            FaktorKaliDisetahunkan = 11
                        ElseIf StartWorkingDate.Month = 3 Then
                            FaktorKaliDisetahunkan = 10
                        ElseIf StartWorkingDate.Month = 4 Then
                            FaktorKaliDisetahunkan = 9
                        ElseIf StartWorkingDate.Month = 5 Then
                            FaktorKaliDisetahunkan = 8
                        ElseIf StartWorkingDate.Month = 6 Then
                            FaktorKaliDisetahunkan = 7
                        ElseIf StartWorkingDate.Month = 7 Then
                            FaktorKaliDisetahunkan = 6
                        ElseIf StartWorkingDate.Month = 8 Then
                            FaktorKaliDisetahunkan = 5
                        ElseIf StartWorkingDate.Month = 9 Then
                            FaktorKaliDisetahunkan = 4
                        ElseIf StartWorkingDate.Month = 10 Then
                            FaktorKaliDisetahunkan = 3
                        ElseIf StartWorkingDate.Month = 11 Then
                            FaktorKaliDisetahunkan = 2
                        ElseIf StartWorkingDate.Month = 12 Then
                            FaktorKaliDisetahunkan = 1
                        End If

                        Me.TxtPeriodStart.Text = StartWorkingDate.Month
                        Me.TxtPeriodEnd.Text = 12
                        Me.TxtPeriodCount.Text = FaktorKaliDisetahunkan
                    End If
                End If
            Else
                FaktorKaliDisetahunkan = 12

                Me.TxtPeriodStart.Text = "1"
                Me.TxtPeriodEnd.Text = "12"
                Me.TxtPeriodCount.Text = "12"
            End If

            'isikan prev income
            '1. hanya diisikan ketika resign atau bulan 12
            '2. hanya diisikan untuk karyawan yang bekerja mulai dari pertengahan tahun pada tahun ybs
            '3. prev. income tidak boleh direfer oleh transaksi salary lainnya
            Me.TBLPrevIncome.Rows.Clear()
            Me.TxtPrevIncome.Text = Me.ObjCD.DigitGrouping(0, 0)
            Me.TxtPrevTaxPaid.Text = Me.ObjCD.DigitGrouping(0, 0)
            If Me.TxtIsResign.Text = "1" Or Me.ObjCD.GetFirstDateofTheMonth(Me.DtpTo.Value.Date).Date.Month = 12 Then
                If StartWorkingDate.Year >= Me.DtpTo.Value.Date.Year Then
                    Dim DTA1 As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_prev_income2TableAdapter
                    Dim TBL1 As New DseTransaksiSalary.tbltrans_salary_item_prev_income2DataTable
                    Dim row1 As DseTransaksiSalary.tbltrans_salary_item_prev_income2Row

                    DTA1.Fill(TBL1, Me.TxtPrimaryEmployee.Text, Me.ObjCD.GetFisrtDateOfTheYear(StartWorkingDate), Me.ObjCD.GetLasDateOfTheMonth(StartWorkingDate))
                    If Me.Baru = False Then
                        Dim DTA2 As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item_prev_income1TableAdapter
                        Dim TBL2 As New DseTransaksiSalary.tbltrans_salary_item_prev_income1DataTable
                        Dim row2 As DseTransaksiSalary.tbltrans_salary_item_prev_income1Row

                        DTA2.Fill(TBL2, Me.TxtPrimary.Text)
                        For Each row2 In TBL2.Rows
                            row1 = TBL1.Newtbltrans_salary_item_prev_income2Row

                            row1.a1_form_no = row2.a1_form_no
                            row1.employer = row2.employer
                            row1.end_period = row2.end_period
                            row1.income_tax_paid = row2.income_tax_paid
                            row1.nett_income = row2.nett_income
                            row1.primary = row2.primary
                            row1.primary_employee = row2.primary_employee 
                            row1.remarks = row2.remarks
                            row1.start_period = row2.start_period
                            row1._date = row2._date
                            row1.Jml_Salary = 0

                            TBL1.Rows.Add(row1)
                        Next

                        DTA2.Dispose()
                        TBL2.Dispose()
                    End If
                    For Each row1 In TBL1.Rows
                        If row1.Jml_Salary = 0 Then
                            Me.rowPrevIncome = Me.TBLPrevIncome.Newtbltrans_salary_item_prev_income1Row

                            Me.rowPrevIncome.a1_form_no = row1.a1_form_no
                            Me.rowPrevIncome.employer = row1.employer
                            Me.rowPrevIncome.end_period = row1.end_period
                            Me.rowPrevIncome.income_tax_paid = row1.income_tax_paid
                            Me.rowPrevIncome.nett_income = row1.nett_income
                            Me.rowPrevIncome.primary = row1.primary
                            Me.rowPrevIncome.primary_employee = row1.primary_employee
                            Me.rowPrevIncome.primary_salary = Me.TxtPrimary.Text
                            Me.rowPrevIncome.remarks = row1.remarks
                            Me.rowPrevIncome.start_period = row1.start_period
                            Me.rowPrevIncome._date = row1._date

                            Me.TBLPrevIncome.Rows.Add(Me.rowPrevIncome)
                            Me.TxtPrevIncome.Text = Me.ObjCD.DigitGrouping(Me.rowPrevIncome.nett_income, 0)
                            Me.TxtPrevTaxPaid.Text = Me.ObjCD.DigitGrouping(Me.rowPrevIncome.income_tax_paid, 0)

                            Exit For
                        End If
                    Next

                    DTA1.Dispose()
                    TBL1.Dispose()
                End If
            End If

            DTATarif.Connection.ConnectionString = Core_Module.GetConstring
            DTATarif.Fill(Me.vTBLTarif)

            'kosongkan item tax detail dan item iterasi
            Me.TBLItemTax.Clear()
            Me.TBLItemIterasi.Clear()

            'kosongkan semua text box yang terkait dengan pajak penghasilan
            Me.KosongkanFieldPerhitunganPajak()
            Me.Obj.IsiKeteranganKolomTabPage3(Me)

            If Me.cboTaxCalculation.SelectedIndex = 3 Then
                'manual tax calculation
                Exit Sub
            End If

            'masukkan data item tax detail
            'hati2, untuk akhir tahun, harus memperhitungkan balance payroll dan memperhitungkan komponen gaji bulan2 sebelumnya
            For Each rowDGV In Me.dgvItem.Rows
                If CDec(rowDGV.Cells(Me.ColAmount.Index).Value) + _
                            Me.GetAmount_OtherPayment(rowDGV.Cells(Me.ColIdentifierKomponen.Index).Value) <> 0 Then
                    CboDgv = rowDGV.Cells(Me.ColTaxElement.Index)
                    TipeElemenPajak = Me.GetTipeElemenPajak(CboDgv.Items.IndexOf(rowDGV.Cells(Me.ColTaxElement.Index).Value))
                    If TipeElemenPajak <> ClsTransaksiSalary_Item.EnumTaxElement.NoTax And TipeElemenPajak <> ClsTransaksiSalary_Item.EnumTaxElement.No21_Pesangon Then
                        rowItemTax = Me.TBLItemTax.Newtbltrans_salary_itemtaxRow()
                        With rowItemTax
                            CboDgv = rowDGV.Cells(Me.ColTipeKomponen.Index)
                            TipeSalaryKomponen = CboDgv.Items.IndexOf(rowDGV.Cells(Me.ColTipeKomponen.Index).Value)
                            CboDgv = rowDGV.Cells(Me.ColTaxElement.Index)
                            .tipe_tax = Me.GetTipeElemenPajak(CboDgv.Items.IndexOf(rowDGV.Cells(Me.ColTaxElement.Index).Value))
                            TipeElemenPajak = .tipe_tax
                            .amount_month = Me.GetAmountMonth(Me.ObjCD.RoundDown1(rowDGV.Cells(Me.ColAmount.Index).Value, 0), TipeSalaryKomponen, TipeElemenPajak) + _
                            Me.GetAmount_OtherPayment(rowDGV.Cells(Me.ColIdentifierKomponen.Index).Value)

                            If TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT Or _
                            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH Or _
                            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb Or _
                            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya Or _
                            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja Or _
                            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH Or _
                            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi Then
                                Me.vPenghasilanBruto += .amount_month
                            End If
                            'If rowDGV.Cells(Me.ColIdentifierKomponen.Index).Value = "1_21" Or rowDGV.Cells(Me.ColIdentifierKomponen.Index).Value = "1_48" Or _
                            'rowDGV.Cells(Me.ColIdentifierKomponen.Index).Value = "1_22" Then
                            '.amount_month = 0
                            'End If

                            .identifier_salary_component = rowDGV.Cells(Me.ColIdentifierKomponen.Index).Value
                            .primary = Me.GetPrimaryMax(Me.TBLItemTax) + 1
                            .primary_transaksi = Me.TxtPrimary.Text
                            .is_rutin = rowDGV.Cells(Me.ColIsRutin.Index).Value
                            If Me.DtpTo.Value.Month = 12 Or vResign = True Then
                                If .is_rutin = 1 Then
                                    .amount_gabungan_disetahunkan = .amount_month
                                    .remarks_gabungan = ""
                                    .amount_rutin_disetahunkan = .amount_month
                                    .remarks_rutin = ""
                                Else
                                    .amount_gabungan_disetahunkan = .amount_month
                                    .remarks_gabungan = ""
                                    .amount_rutin_disetahunkan = 0
                                    .remarks_rutin = ""
                                End If
                            Else
                                If .is_rutin = 1 Then
                                    .amount_gabungan_disetahunkan = FaktorKaliDisetahunkan * .amount_month
                                    .remarks_gabungan = FaktorKaliDisetahunkan.ToString + " * " + Me.ObjCD.DigitGrouping(.amount_month, 0)
                                    .amount_rutin_disetahunkan = FaktorKaliDisetahunkan * .amount_month
                                    .remarks_rutin = ""
                                Else
                                    .amount_gabungan_disetahunkan = .amount_month
                                    .remarks_gabungan = ""
                                    .amount_rutin_disetahunkan = 0
                                    .remarks_rutin = ""
                                End If
                            End If
                        End With
                        Me.TBLItemTax.Rows.Add(rowItemTax)
                    End If
                End If
            Next

            'tambahkan item tax dengan other payment, untuk bulan ini, tidak memperhitungkan resign
            Me.TambahkanDgnItemOtherPayment(Me.vTBLOtherPayment, FaktorKaliDisetahunkan, vResign)

            If Me.DtpTo.Value.Month = 12 Or vResign = True Then
                If Me.TxtPrimaryEmployee.Text = "-1" Then
                    vTBLOtherPayment1.Rows.Clear()
                Else
                    Me.IsikanItemOtherPayment_EndOfYear_Resign(StartWorkingDate, ResignDate, vTBLOtherPayment1)
                End If
                Me.TambahkanDgnItemOtherPayment_Resign_EndOfYear(vTBLOtherPayment1, FaktorKaliDisetahunkan)

                Dim DTAAdd As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item1TableAdapter
                Dim TBLAdd As New DseTransaksiSalary.tbltrans_salary_item1DataTable
                Dim rowAdd As DseTransaksiSalary.tbltrans_salary_item1Row

                '1. masukkan item tax detail untuk bulan 1-11
                DTAAdd.Fill(TBLAdd, Me.ObjCD.GetFirstDateofTheMonth(Me.DtpTo.Value) - TimeSpan.FromDays(1), _
                Me.ObjCD.GetFisrtDateOfTheYear(Me.DtpTo.Value), _
                Me.TxtPrimaryEmployee.Text)
                For Each rowAdd In TBLAdd.Rows
                    TipeElemenPajak = Me.GetTipeElemenPajak(rowAdd.Tax_Component)
                    If TipeElemenPajak <> ClsTransaksiSalary_Item.EnumTaxElement.NoTax And TipeElemenPajak <> ClsTransaksiSalary_Item.EnumTaxElement.No21_Pesangon Then
                        '#sinkron2018_pphcls
                        If TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT Or _
                        TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb Or _
                        TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya Or _
                        TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja Or _
                        TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH Or _
                        TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi Then
                            If rowAdd.Type = FrmMasterSalaryFormula_Component.EnumType.Income Or _
                            rowAdd.Type = FrmMasterSalaryFormula_Component.EnumType.CompanyExpense Then
                                rowAdd.Amount = rowAdd.Amount
                            ElseIf rowAdd.Type = FrmMasterSalaryFormula_Component.EnumType.Deduction Or _
                            rowAdd.Type = FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense Then
                                rowAdd.Amount = rowAdd.Amount * (-1)
                            End If
                        ElseIf TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No10_IuranPensiun_THT_JHT Then
                            If rowAdd.Type = FrmMasterSalaryFormula_Component.EnumType.Income Or _
                            rowAdd.Type = FrmMasterSalaryFormula_Component.EnumType.CompanyExpense Then
                                rowAdd.Amount = rowAdd.Amount * (-1)
                            ElseIf rowAdd.Type = FrmMasterSalaryFormula_Component.EnumType.Deduction Or _
                            rowAdd.Type = FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense Then
                                rowAdd.Amount = rowAdd.Amount
                            End If
                        End If


                        'cari apakah ada yang identifier salary komponennya sama
                        Me.TBLItemTax.DefaultView.RowFilter = "[identifier salary component] = '" + rowAdd.Identifier_Salary_Component + "'"
                        If Me.TBLItemTax.DefaultView.Count = 0 Then
                            'jika tidak ada yang sama, maka tambahkan baris baru
                            rowItemTax = Me.TBLItemTax.Newtbltrans_salary_itemtaxRow()
                            With rowItemTax
                                .tipe_tax = CInt(TipeElemenPajak)
                                .amount_month = 0
                                .identifier_salary_component = rowAdd.Identifier_Salary_Component
                                .primary = Me.GetPrimaryMax(Me.TBLItemTax) + 1
                                .primary_transaksi = Me.TxtPrimary.Text
                                .is_rutin = rowAdd.Is_Rutin
                                If .is_rutin = 1 Then
                                    .amount_gabungan_disetahunkan = rowAdd.Amount
                                    .remarks_gabungan = ""
                                    .amount_rutin_disetahunkan = rowAdd.Amount
                                    .remarks_rutin = ""
                                Else
                                    .amount_gabungan_disetahunkan = rowAdd.Amount
                                    .remarks_gabungan = ""
                                    .amount_rutin_disetahunkan = 0
                                    .remarks_rutin = ""
                                End If
                            End With
                            Me.TBLItemTax.Rows.Add(rowItemTax)
                        Else
                            'kalo ada, maka hanya tambahkan amount saja
                            rowItemTax = Me.TBLItemTax.DefaultView(0).Row
                            With rowItemTax
                                If .is_rutin = 1 Then
                                    .amount_gabungan_disetahunkan += rowAdd.Amount
                                    .amount_rutin_disetahunkan += rowAdd.Amount
                                Else
                                    .amount_gabungan_disetahunkan += rowAdd.Amount
                                    .amount_rutin_disetahunkan += 0
                                End If
                            End With
                        End If
                        'Else
                        'MsgBox("")
                    End If
                Next

                '2. masukkan item tax detail untuk begining balance tax

                '3. masukkan item terkait dengan form potongan pajak tahunan

                DTAAdd.Dispose()
                TBLAdd.Dispose()
            End If

            'hitung detail pajak untuk iterasi 0
            UseNPWP = Me.ObjKrywn.GetIsNPWP(Me.TxtPrimaryEmployee.Text)
            NPWPCategory = Me.ObjKrywn.GetNPWPCategory(Me.TxtPrimaryEmployee.Text)
            TaxCategoryDesc = Me.ObjPTKP.GetDesc(NPWPCategory)
            Txt15Pjk_PTKP.Text = Me.ObjPTKP.GetPTKPPerTahun(NPWPCategory)

            If Me.DtpTo.Value.Month = 12 Or vResign = True Then
                If Format(StartWorkingDate, "MMM yyyy") <> Format(ResignDate, "MMM yyyy") Then
                    If ResignDate = #1/1/1950# Then
                        ResignDate = Me.DtpTo.Value
                    End If
                    '#sinkron2018_pphcls
                    If StartWorkingDate.Year = ResignDate.Year Then 'tahun pertama
                        Txt18.Text = Me.Obj.GetPPHTerutang_Rutin_NonRutin(Me.TxtPrimaryEmployee.Text, StartWorkingDate, _
                        ObjCD.GetFirstDateofTheMonth(ResignDate) - TimeSpan.FromDays(1))
                        Me.ObjBB.GetValue(Me.TBLTaxBB, Me.ObjCD.GetFirstDateofTheMonth(StartWorkingDate), _
                        Me.ObjCD.GetLasDateOfTheMonth(ObjCD.GetFirstDateofTheMonth(ResignDate) - TimeSpan.FromDays(1)), Me.TxtPrimaryEmployee.Text)
                        'MsgBox("Start: " + Format(Me.ObjCD.GetFirstDateofTheMonth(StartWorkingDate), "MMM yyyy") + vbCrLf + "End: " + Format(Me.ObjCD.GetLasDateOfTheMonth(ObjCD.GetFirstDateofTheMonth(ResignDate) - TimeSpan.FromDays(1)), "MMM yyyy"))
                        'Me.ObjBB.GetValue(Me.TBLTaxBB2, Me.ObjCD.GetFisrtDateOfTheYear(StartWorkingDate), _
                        'Me.ObjCD.GetLasDateOfTheMonth(Me.ObjCD.GetFirstDateofTheMonth(StartWorkingDate) - TimeSpan.FromDays(1)), Me.TxtPrimaryEmployee.Text) (kemungkinan bagian ini obselete, karena menggunakan table prev income)
                        'MsgBox("Start: " + Format(Me.ObjCD.GetFisrtDateOfTheYear(StartWorkingDate), "MMM yyyy") + vbCrLf + "End: " + Format(Me.ObjCD.GetLasDateOfTheMonth(Me.ObjCD.GetFirstDateofTheMonth(StartWorkingDate) - TimeSpan.FromDays(1)), "MMM yyyy"))
                    Else
                        If Me.DtpTo.Value.Month = 12 Then
                            'bulan 12
                            Txt18.Text = Me.Obj.GetPPHTerutang_Rutin_NonRutin(Me.TxtPrimaryEmployee.Text, ObjCD.GetFisrtDateOfTheYear(Me.DtpTo.Value), _
                            CDate("11/30/" + Format(Me.DtpTo.Value.Date, "yyyy")))
                            'Me.ObjBB.GetValue(Me.TBLTaxBB, ObjCD.GetFisrtDateOfTheYear(ResignDate), _
                            'ObjCD.GetFirstDateofTheMonth(ResignDate) - TimeSpan.FromDays(1), Me.TxtPrimaryEmployee.Text) (kemungkinan bagian ini obselete, karena menggunakan table prev income)
                        Else
                            'resign =true
                            Txt18.Text = Me.Obj.GetPPHTerutang_Rutin_NonRutin(Me.TxtPrimaryEmployee.Text, ObjCD.GetFisrtDateOfTheYear(ResignDate), _
                            ObjCD.GetFirstDateofTheMonth(ResignDate) - TimeSpan.FromDays(1))
                            'Me.ObjBB.GetValue(Me.TBLTaxBB, ObjCD.GetFisrtDateOfTheYear(ResignDate), _
                            'ObjCD.GetFirstDateofTheMonth(ResignDate) - TimeSpan.FromDays(1), Me.TxtPrimaryEmployee.Text) (kemungkinan bagian ini obselete, karena menggunakan table prev income)
                        End If
                        Debug.Print("PPh sudah disetorkan Jan - Nov:" + Me.ObjCD.DigitGrouping(Txt18.Text, 0))
                        Debug.Print("Start working date:" + Format(StartWorkingDate, "dd MMM yyyy"))
                        Debug.Print("Start resign date:" + Format(ResignDate, "dd MMM yyyy"))
                    End If

                    'tambahkan dengan tunjangan yang tidak dihitung dengan SMESS, gaji dihitung menggunakan SMESS
                    'tapi pajaknya tidak. kemudian diakhir tahun, mau hitung semua menggunakan SMESS
                    'contoh kasus infinity ketika mulai implementasi pph ditahun 2018 (kemungkinan bagian ini obselete, karena menggunakan table prev income)
                    Me.Txt13PjkJmlHasilMsPjkSblm.Text = 0
                    'If Me.TBLTaxBB.Rows.Count > 0 Then
                    '    For Each RowTaxBB In Me.TBLTaxBB
                    '        '#sinkron2018_pphcls
                    '        Txt18.Text = CDec(Txt18.Text) + RowTaxBB.Tunjangan_PPh
                    '    Next
                    'End If
                    'If Me.TBLTaxBB2.Rows.Count > 0 Then
                    '    For Each RowTaxBB In Me.TBLTaxBB2
                    '        '#sinkron2018_pphcls
                    '        Txt18.Text = CDec(Txt18.Text) + RowTaxBB.Tunjangan_PPh

                    '        'ini hanya jalan untuk sebelum start work periode, di tahun pertama
                    '        If RowTaxBB.Gaji_Pensiun > 0 Then
                    '            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No13_PenghasilanNettoMasaSebelum
                    '            TipeSalaryKomponen = FrmMasterSalaryFormula_Component.EnumType.Income
                    '            Amount = RowTaxBB.Gaji_Pensiun
                    '            Remarks1 = Format(RowTaxBB.Period, "MMM yyyy")

                    '            Me.Txt13PjkJmlHasilMsPjkSblm.Text = CDec(Me.Txt13PjkJmlHasilMsPjkSblm.Text) + Amount

                    '            'cari apakah ada yang identifier salary komponennya sama
                    '            rowItemTax = Me.TBLItemTax.Newtbltrans_salary_itemtaxRow()
                    '            With rowItemTax
                    '                .tipe_tax = CInt(TipeElemenPajak)
                    '                .amount_month = 0
                    '                .identifier_salary_component = ""
                    '                .primary = Me.GetPrimaryMax(Me.TBLItemTax) + 1
                    '                .primary_transaksi = Me.TxtPrimary.Text
                    '                .is_rutin = 1
                    '                .amount_gabungan_disetahunkan = Amount
                    '                .remarks_gabungan = Remarks1
                    '                .amount_rutin_disetahunkan = Amount
                    '                .remarks_rutin = ""
                    '            End With
                    '            Me.TBLItemTax.Rows.Add(rowItemTax)
                    '        End If
                    '    Next
                    'End If

                    'Txt18.Text = Me.Obj.GetPPHTerutang_Rutin_NonRutin(Me.TxtPrimaryEmployee.Text, StartWorkingDate, _
                    'ObjCD.GetFirstDateofTheMonth(ResignDate) - TimeSpan.FromDays(1))
                Else
                    Txt18.Text = 0
                End If
            End If

            If vPPhRutin + vPPhNonRutin <> 0 Then
                If Me.DtpTo.Value.Month <> 12 And vResign = False Then
                    Me.Obj.HitungPajak1_Aktual(UseNPWP, Me.vTBLTarif, FaktorKaliDisetahunkan, _
                    Me.TBLItemTax, Me.Perct_BiayaJabatan, Me.BiayaJabatanMaxValue, Me.ObjCD, NPWPCategory, TaxCategoryDesc, _
                    Me.TxtPrimary.Text, Me.TxtPrimaryEmployee.Text, Me.DtpDate.Value, Me.TxtDocumentNo.Text, FaktorKaliDisetahunkan * vPPhRutin, _
                    vPPhNonRutin + FaktorKaliDisetahunkan * vPPhRutin, Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text, Me.Txt02Pjk_TunjPPHDisetahunkan.Text, Me.Txt02Pjk_TunjPPHRutin.Text, _
                    Me.Txt02Pjk_TunjPPHNonRutin.Text, Me.Txt02PjkTunjanganPPh.Text, Me.Txt01Pjk_GajiRutin.Text, Me.Txt01Pjk_Gaji.Text, _
                    Me.Txt03Pjk_TunjLainnya.Text, Me.Txt03Pjk_TunjLainnyaRutin.Text, Me.Txt04PjkHonorarium.Text, Me.Txt04PjkHonorariumRutin.Text, _
                    Me.Txt05PjkPremiAsuransi.Text, Me.Txt05PjkPremiAsuransiRutin.Text, Me.Txt06PjkNatura.Text, Me.Txt06PjkNaturaRutin.Text, _
                    Me.Txt07PjkTantiem.Text, Me.Txt07PjkTantiemRutin.Text, Me.Txt08PjkPenghasilanBruto.Text, Me.Txt08PjkPenghasilanBrutoRutin.Text, _
                    Me.Txt09PjkBiayaJabatan.Text, Me.Txt09PjkBiayaJabatanRutin.Text, Me.Txt10PjkIuranPensiun.Text, Me.Txt10PjkIuranPensiunRutin.Text, _
                    Me.Txt11PjkJmlPengurangan.Text, Me.Txt11PjkJmlPenguranganRutin.Text, Me.Txt12PjkJmlPenghasilanNetto.Text, Me.Txt13PjkJmlHasilMsPjkSblm.Text, _
                    Me.Txt14PjkJmlHasilNettoUPPh.Text, Me.Txt14PjkJmlHasilNettoUPPhRutin.Text, Me.Txt15Pjk_PTKP.Text, Me.Txt16Pjk_PKPSetahun.Text, _
                    Me.Txt16Pjk_PKPSetahunRutin.Text, Me.Txt17Pjk_PPH21.Text, Me.Txt17Pjk_PPH21Rutin.Text, Me.Txt18.Text, Me.Txt19Pjk_PPH21TerutangSetahun.Text, _
                    Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text, Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text, Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text, _
                    Me.Txt19Pjk_PPH21Terutang.Text, Me.TxtIncomeTax.Text, Me.TxtKurangBayarTax.Text, Me.Txt20.Text)
                    rowItemIterasi = Me.TBLItemIterasi.Newtbltrans_salary_item_iterationRow
                    Me.Txt19Pjk_PPH21TerutangSetahun.Text = ((FaktorKaliDisetahunkan * vPPhRutin) + vPPhNonRutin)
                    Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text = (FaktorKaliDisetahunkan * vPPhRutin)
                    Me.Txt19Pjk_PPH21Terutang.Text = (vPPhRutin + vPPhNonRutin)
                    Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text = vPPhNonRutin
                    Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text = vPPhRutin

                    Me.Txt02Pjk_TunjPPHDisetahunkan.Text = Me.Txt19Pjk_PPH21TerutangSetahun.Text
                    Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text = Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text
                    Me.Txt02Pjk_TunjPPHNonRutin.Text = Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text
                    Me.Txt02Pjk_TunjPPHRutin.Text = Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text
                    Me.Txt02PjkTunjanganPPh.Text = Me.Txt19Pjk_PPH21Terutang.Text

                    'End With
                    With rowItemIterasi
                        .PPH_Terutang_Non_Rutin = CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) - CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text)
                        .PPH_Terutang_Rutin = CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text)
                        .primary = Me.GetPrimaryMax(Me.TBLItemIterasi)
                        .primary_transaksi = Me.TxtPrimary.Text
                        .Tunjangan_PPH_Rutin = CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text)
                        ._Tunjangan_PPH_Non_Rutin = CDec(Me.Txt02Pjk_TunjPPHDisetahunkan.Text) - CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text)
                    End With
                    Me.TBLItemIterasi.Rows.Add(rowItemIterasi)
                Else
                    MsgBox("belum coding")
                End If
            Else
                If Me.DtpTo.Value.Month = 12 Or vResign = True Then
                    Me.Obj.HitungPajak1_Aktual_EndOfYear(0, 0, UseNPWP, Me.vTBLTarif, 0, 0, StartWorkingDate, ResignDate, FaktorKaliDisetahunkan, _
                    TBLItemTax, Me.Perct_BiayaJabatan, Me.BiayaJabatanMaxValue, Me.ObjCD, NPWPCategory, TaxCategoryDesc, Me.TxtPrimary.Text, _
                    Me.TxtPrimaryEmployee.Text, Me.DtpDate.Value, Me.TxtDocumentNo.Text, Me.DtpTo.Value, Txt18.Text, _
                    Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text, Me.Txt02Pjk_TunjPPHDisetahunkan.Text, Me.Txt02Pjk_TunjPPHRutin.Text, _
                    Me.Txt02Pjk_TunjPPHNonRutin.Text, Me.Txt02PjkTunjanganPPh.Text, Me.Txt01Pjk_GajiRutin.Text, Me.Txt01Pjk_Gaji.Text, _
                    Me.Txt03Pjk_TunjLainnya.Text, Me.Txt03Pjk_TunjLainnyaRutin.Text, Me.Txt04PjkHonorarium.Text, Me.Txt04PjkHonorariumRutin.Text, _
                    Me.Txt05PjkPremiAsuransi.Text, Me.Txt05PjkPremiAsuransiRutin.Text, Me.Txt06PjkNatura.Text, Me.Txt06PjkNaturaRutin.Text, _
                    Me.Txt07PjkTantiem.Text, Me.Txt07PjkTantiemRutin.Text, Me.Txt08PjkPenghasilanBruto.Text, Me.Txt08PjkPenghasilanBrutoRutin.Text, _
                    Me.Txt09PjkBiayaJabatan.Text, Me.Txt09PjkBiayaJabatanRutin.Text, Me.Txt10PjkIuranPensiun.Text, Me.Txt10PjkIuranPensiunRutin.Text, _
                    Me.Txt11PjkJmlPengurangan.Text, Me.Txt11PjkJmlPenguranganRutin.Text, Me.Txt12PjkJmlPenghasilanNetto.Text, Me.Txt13PjkJmlHasilMsPjkSblm.Text, _
                    Me.Txt14PjkJmlHasilNettoUPPh.Text, Me.Txt14PjkJmlHasilNettoUPPhRutin.Text, Me.Txt15Pjk_PTKP.Text, Me.Txt16Pjk_PKPSetahun.Text, _
                    Me.Txt16Pjk_PKPSetahunRutin.Text, Me.Txt17Pjk_PPH21.Text, Me.Txt17Pjk_PPH21Rutin.Text, Me.Txt18.Text, Me.Txt19Pjk_PPH21TerutangSetahun.Text, _
                    Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text, Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text, Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text, _
                    Me.Txt19Pjk_PPH21Terutang.Text, Me.TxtIncomeTax.Text, Me.TxtKurangBayarTax.Text, Me.Txt20.Text, Me.TxtLosesTaxObligation.Text, _
                    Me.TxtPeriodEnd.Text, Me.TxtPeriodCount.Text, vResign, Me.DtpDate.Value.Date, Me.TxtPeriodStart.Text, Me.TxtPrevIncome.Text, Me.TxtPrevTaxPaid.Text, _
                    Me.TBLPrevIncome)
                    rowItemIterasi = Me.TBLItemIterasi.Newtbltrans_salary_item_iterationRow
                    With rowItemIterasi
                        .PPH_Terutang_Non_Rutin = CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) - CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text)
                        .PPH_Terutang_Rutin = CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text)
                        .primary = Me.GetPrimaryMax(Me.TBLItemIterasi)
                        .primary_transaksi = Me.TxtPrimary.Text
                        .Tunjangan_PPH_Rutin = CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text)
                        ._Tunjangan_PPH_Non_Rutin = CDec(Me.Txt02Pjk_TunjPPHDisetahunkan.Text) - CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text)
                    End With
                    Me.TBLItemIterasi.Rows.Add(rowItemIterasi)
                Else
                    Me.Obj.HitungPajak1_Aktual(UseNPWP, Me.vTBLTarif, FaktorKaliDisetahunkan, _
                    Me.TBLItemTax, Me.Perct_BiayaJabatan, Me.BiayaJabatanMaxValue, Me.ObjCD, NPWPCategory, TaxCategoryDesc, _
                    Me.TxtPrimary.Text, Me.TxtPrimaryEmployee.Text, Me.DtpDate.Value, Me.TxtDocumentNo.Text, 0, _
                    0, Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text, Me.Txt02Pjk_TunjPPHDisetahunkan.Text, Me.Txt02Pjk_TunjPPHRutin.Text, _
                    Me.Txt02Pjk_TunjPPHNonRutin.Text, Me.Txt02PjkTunjanganPPh.Text, Me.Txt01Pjk_GajiRutin.Text, Me.Txt01Pjk_Gaji.Text, _
                    Me.Txt03Pjk_TunjLainnya.Text, Me.Txt03Pjk_TunjLainnyaRutin.Text, Me.Txt04PjkHonorarium.Text, Me.Txt04PjkHonorariumRutin.Text, _
                    Me.Txt05PjkPremiAsuransi.Text, Me.Txt05PjkPremiAsuransiRutin.Text, Me.Txt06PjkNatura.Text, Me.Txt06PjkNaturaRutin.Text, _
                    Me.Txt07PjkTantiem.Text, Me.Txt07PjkTantiemRutin.Text, Me.Txt08PjkPenghasilanBruto.Text, Me.Txt08PjkPenghasilanBrutoRutin.Text, _
                    Me.Txt09PjkBiayaJabatan.Text, Me.Txt09PjkBiayaJabatanRutin.Text, Me.Txt10PjkIuranPensiun.Text, Me.Txt10PjkIuranPensiunRutin.Text, _
                    Me.Txt11PjkJmlPengurangan.Text, Me.Txt11PjkJmlPenguranganRutin.Text, Me.Txt12PjkJmlPenghasilanNetto.Text, Me.Txt13PjkJmlHasilMsPjkSblm.Text, _
                    Me.Txt14PjkJmlHasilNettoUPPh.Text, Me.Txt14PjkJmlHasilNettoUPPhRutin.Text, Me.Txt15Pjk_PTKP.Text, Me.Txt16Pjk_PKPSetahun.Text, _
                    Me.Txt16Pjk_PKPSetahunRutin.Text, Me.Txt17Pjk_PPH21.Text, Me.Txt17Pjk_PPH21Rutin.Text, Me.Txt18.Text, Me.Txt19Pjk_PPH21TerutangSetahun.Text, _
                    Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text, Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text, Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text, _
                    Me.Txt19Pjk_PPH21Terutang.Text, Me.TxtIncomeTax.Text, Me.TxtKurangBayarTax.Text, Me.Txt20.Text)

                    rowItemIterasi = Me.TBLItemIterasi.Newtbltrans_salary_item_iterationRow
                    With rowItemIterasi
                        .PPH_Terutang_Non_Rutin = CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) - CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text)
                        .PPH_Terutang_Rutin = CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text)
                        .primary = Me.GetPrimaryMax(Me.TBLItemIterasi)
                        .primary_transaksi = Me.TxtPrimary.Text
                        .Tunjangan_PPH_Rutin = CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text)
                        ._Tunjangan_PPH_Non_Rutin = CDec(Me.Txt02Pjk_TunjPPHDisetahunkan.Text) - CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text)
                    End With
                    Me.TBLItemIterasi.Rows.Add(rowItemIterasi)
                End If
            End If

            ''sementara u/gross up akhir tahun s.d iterasi 01 saja
            'If Me.DtpTo.Value.Month = 12 Then
            '    Exit Sub
            'End If

            'jika gross up, maka hitung pajak untuk iterasi selanjutnya sampai dengan tunjangan PPH = PPH terhutang
            If Me.cboTaxCalculation.SelectedIndex = 0 Then
                'While CDec(Me.Txt02PjkTunjanganPPh.Text) <> CDec(Me.Txt19Pjk_PPH21Terutang.Text)
                'While CDec(Me.Txt02Pjk_TunjPPHDisetahunkan.Text) <> CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) Or _
                'CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text) <> CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text)
                While LanjutIterasi = True
                    If JmlIterasi > 10000 Then
                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    JmlIterasi += 1
                    Debug.Print("----------------------------------------------------")
                    Debug.Print("Start Iterasi: " + Me.ObjCD.DigitGrouping(JmlIterasi))

                    'If JmlIterasi >= 1 Then
                    '    Exit Sub
                    'End If
                    'JmlIterasi += 1
                    'If JmlIterasi = 1 Then
                    '    Selisih1 = 0
                    '    Selisih2 = 0
                    'Else
                    rowItemIterasiSebelum = Me.TBLItemIterasi.Rows(Me.TBLItemIterasi.Rows.Count - 1)
                    If Me.TxtPeriodEnd.Text <> 12 And vResign = True And LosesTaxObligation = 1 Then
                        Selisih1 = Math.Round(((rowItemIterasiSebelum.PPH_Terutang_Non_Rutin + rowItemIterasiSebelum.PPH_Terutang_Rutin) / 12) * CDec(Me.TxtPeriodCount.Text), 0, MidpointRounding.AwayFromZero) - rowItemIterasiSebelum._Tunjangan_PPH_Non_Rutin - rowItemIterasiSebelum.Tunjangan_PPH_Rutin
                        Selisih2 = Math.Round((rowItemIterasiSebelum.PPH_Terutang_Rutin / 12) * CDec(Me.TxtPeriodCount.Text), 0, MidpointRounding.AwayFromZero) - rowItemIterasiSebelum.Tunjangan_PPH_Rutin
                    Else
                        Selisih1 = rowItemIterasiSebelum.PPH_Terutang_Non_Rutin + rowItemIterasiSebelum.PPH_Terutang_Rutin - rowItemIterasiSebelum._Tunjangan_PPH_Non_Rutin - rowItemIterasiSebelum.Tunjangan_PPH_Rutin
                        Selisih2 = rowItemIterasiSebelum.PPH_Terutang_Rutin - rowItemIterasiSebelum.Tunjangan_PPH_Rutin
                    End If

                    If Selisih1 = 1 Then
                        Selisih1 = 0
                    End If
                    If Selisih2 = 1 Then
                        Selisih2 = 0
                    End If
                    Debug.Print("Selisih 1: " + Me.ObjCD.DigitGrouping(Selisih1, 0) + " -  Selisih 2: " + Me.ObjCD.DigitGrouping(Selisih2, 0))

                    'End If
                    If Me.DtpTo.Value.Month <> 12 And vResign = False Then
                        Me.Obj.HitungPajak1_Aktual(UseNPWP, Me.vTBLTarif, FaktorKaliDisetahunkan, _
                        Me.TBLItemTax, Me.Perct_BiayaJabatan, Me.BiayaJabatanMaxValue, Me.ObjCD, NPWPCategory, TaxCategoryDesc, _
                        Me.TxtPrimary.Text, Me.TxtPrimaryEmployee.Text, Me.DtpDate.Value, Me.TxtDocumentNo.Text, CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text) + Me.ObjCD.RoundDown1(Selisih2 / 2, 0), _
                        CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) + Me.ObjCD.RoundDown1(Selisih1 / 2, 0), Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text, Me.Txt02Pjk_TunjPPHDisetahunkan.Text, Me.Txt02Pjk_TunjPPHRutin.Text, _
                        Me.Txt02Pjk_TunjPPHNonRutin.Text, Me.Txt02PjkTunjanganPPh.Text, Me.Txt01Pjk_GajiRutin.Text, Me.Txt01Pjk_Gaji.Text, _
                        Me.Txt03Pjk_TunjLainnya.Text, Me.Txt03Pjk_TunjLainnyaRutin.Text, Me.Txt04PjkHonorarium.Text, Me.Txt04PjkHonorariumRutin.Text, _
                        Me.Txt05PjkPremiAsuransi.Text, Me.Txt05PjkPremiAsuransiRutin.Text, Me.Txt06PjkNatura.Text, Me.Txt06PjkNaturaRutin.Text, _
                        Me.Txt07PjkTantiem.Text, Me.Txt07PjkTantiemRutin.Text, Me.Txt08PjkPenghasilanBruto.Text, Me.Txt08PjkPenghasilanBrutoRutin.Text, _
                        Me.Txt09PjkBiayaJabatan.Text, Me.Txt09PjkBiayaJabatanRutin.Text, Me.Txt10PjkIuranPensiun.Text, Me.Txt10PjkIuranPensiunRutin.Text, _
                        Me.Txt11PjkJmlPengurangan.Text, Me.Txt11PjkJmlPenguranganRutin.Text, Me.Txt12PjkJmlPenghasilanNetto.Text, Me.Txt13PjkJmlHasilMsPjkSblm.Text, _
                        Me.Txt14PjkJmlHasilNettoUPPh.Text, Me.Txt14PjkJmlHasilNettoUPPhRutin.Text, Me.Txt15Pjk_PTKP.Text, Me.Txt16Pjk_PKPSetahun.Text, _
                        Me.Txt16Pjk_PKPSetahunRutin.Text, Me.Txt17Pjk_PPH21.Text, Me.Txt17Pjk_PPH21Rutin.Text, Me.Txt18.Text, Me.Txt19Pjk_PPH21TerutangSetahun.Text, _
                        Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text, Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text, Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text, _
                        Me.Txt19Pjk_PPH21Terutang.Text, Me.TxtIncomeTax.Text, Me.TxtKurangBayarTax.Text, Me.Txt20.Text)

                        rowItemIterasi = Me.TBLItemIterasi.Newtbltrans_salary_item_iterationRow

                        With rowItemIterasi
                            .PPH_Terutang_Non_Rutin = CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) - CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text)
                            .PPH_Terutang_Rutin = CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text)
                            .primary = Me.GetPrimaryMax(Me.TBLItemIterasi)
                            .primary_transaksi = Me.TxtPrimary.Text
                            .Tunjangan_PPH_Rutin = CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text)
                            ._Tunjangan_PPH_Non_Rutin = CDec(Me.Txt02Pjk_TunjPPHDisetahunkan.Text) - CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text)
                        End With
                        Me.TBLItemIterasi.Rows.Add(rowItemIterasi)
                    Else
                        If Me.TxtPeriodEnd.Text <> 12 And vResign = True And LosesTaxObligation = 1 Then
                            Debug.Print("Tunj PPH Rutin (Setahun): " + Me.ObjCD.DigitGrouping(Math.Round((CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text) / 12) * CDec(Me.TxtPeriodCount.Text), 0, MidpointRounding.AwayFromZero) + Me.ObjCD.RoundDown1(Selisih2 / 2, 0), 0))
                            Debug.Print("Tunj PPH Gabungan (Setahun): " + Me.ObjCD.DigitGrouping(Math.Round((CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) / 12) * CDec(Me.TxtPeriodCount.Text), 0, MidpointRounding.AwayFromZero) + Me.ObjCD.RoundDown1(Selisih1 / 2, 0), 0))

                            Me.Obj.HitungPajak1_Aktual_EndOfYear(CDec(Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text) + Me.ObjCD.RoundDown1(Selisih2 / 2, 0), _
                            CDec(Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text) + Me.ObjCD.RoundDown1(Selisih1 / 2, 0), _
                            UseNPWP, Me.vTBLTarif, _
                            Math.Round((CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text) / 12) * CDec(Me.TxtPeriodCount.Text), 0, MidpointRounding.AwayFromZero) + Me.ObjCD.RoundDown1(Selisih2 / 2, 0), _
                            Math.Round((CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) / 12) * CDec(Me.TxtPeriodCount.Text), 0, MidpointRounding.AwayFromZero) + Me.ObjCD.RoundDown1(Selisih1 / 2, 0), StartWorkingDate, ResignDate, FaktorKaliDisetahunkan, _
                            TBLItemTax, Me.Perct_BiayaJabatan, Me.BiayaJabatanMaxValue, Me.ObjCD, NPWPCategory, TaxCategoryDesc, Me.TxtPrimary.Text, _
                            Me.TxtPrimaryEmployee.Text, Me.DtpDate.Value, Me.TxtDocumentNo.Text, Me.DtpTo.Value, Txt18.Text, _
                            Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text, Me.Txt02Pjk_TunjPPHDisetahunkan.Text, Me.Txt02Pjk_TunjPPHRutin.Text, _
                            Me.Txt02Pjk_TunjPPHNonRutin.Text, Me.Txt02PjkTunjanganPPh.Text, Me.Txt01Pjk_GajiRutin.Text, Me.Txt01Pjk_Gaji.Text, _
                            Me.Txt03Pjk_TunjLainnya.Text, Me.Txt03Pjk_TunjLainnyaRutin.Text, Me.Txt04PjkHonorarium.Text, Me.Txt04PjkHonorariumRutin.Text, _
                            Me.Txt05PjkPremiAsuransi.Text, Me.Txt05PjkPremiAsuransiRutin.Text, Me.Txt06PjkNatura.Text, Me.Txt06PjkNaturaRutin.Text, _
                            Me.Txt07PjkTantiem.Text, Me.Txt07PjkTantiemRutin.Text, Me.Txt08PjkPenghasilanBruto.Text, Me.Txt08PjkPenghasilanBrutoRutin.Text, _
                            Me.Txt09PjkBiayaJabatan.Text, Me.Txt09PjkBiayaJabatanRutin.Text, Me.Txt10PjkIuranPensiun.Text, Me.Txt10PjkIuranPensiunRutin.Text, _
                            Me.Txt11PjkJmlPengurangan.Text, Me.Txt11PjkJmlPenguranganRutin.Text, Me.Txt12PjkJmlPenghasilanNetto.Text, Me.Txt13PjkJmlHasilMsPjkSblm.Text, _
                            Me.Txt14PjkJmlHasilNettoUPPh.Text, Me.Txt14PjkJmlHasilNettoUPPhRutin.Text, Me.Txt15Pjk_PTKP.Text, Me.Txt16Pjk_PKPSetahun.Text, _
                            Me.Txt16Pjk_PKPSetahunRutin.Text, Me.Txt17Pjk_PPH21.Text, Me.Txt17Pjk_PPH21Rutin.Text, Me.Txt18.Text, Me.Txt19Pjk_PPH21TerutangSetahun.Text, _
                            Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text, Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text, Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text, _
                            Me.Txt19Pjk_PPH21Terutang.Text, Me.TxtIncomeTax.Text, Me.TxtKurangBayarTax.Text, Me.Txt20.Text, Me.TxtLosesTaxObligation.Text, _
                            Me.TxtPeriodEnd.Text, Me.TxtPeriodCount.Text, vResign, Me.DtpDate.Value.Date, Me.TxtPeriodStart.Text, Me.TxtPrevIncome.Text, Me.TxtPrevTaxPaid.Text, _
                            Me.TBLPrevIncome)

                            Debug.Print("TxtIncomeTax: " + Me.ObjCD.DigitGrouping(Me.TxtIncomeTax.Text, 0))
                        Else
                            Debug.Print("Tunj PPH Rutin (Setahun): " + Me.ObjCD.DigitGrouping(CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text) + Me.ObjCD.RoundDown1(Selisih2 / 2, 0), 0))
                            Debug.Print("Tunj PPH Gabungan (Setahun): " + Me.ObjCD.DigitGrouping(CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) + Me.ObjCD.RoundDown1(Selisih1 / 2, 0), 0))

                            Me.Obj.HitungPajak1_Aktual_EndOfYear(CDec(Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text) + Me.ObjCD.RoundDown1(Selisih2 / 2, 0), _
                            CDec(Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text) + Me.ObjCD.RoundDown1(Selisih1 / 2, 0), _
                            UseNPWP, Me.vTBLTarif, _
                            CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text) + Me.ObjCD.RoundDown1(Selisih2 / 2, 0), _
                            CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) + Me.ObjCD.RoundDown1(Selisih1 / 2, 0), StartWorkingDate, ResignDate, FaktorKaliDisetahunkan, _
                            TBLItemTax, Me.Perct_BiayaJabatan, Me.BiayaJabatanMaxValue, Me.ObjCD, NPWPCategory, TaxCategoryDesc, Me.TxtPrimary.Text, _
                            Me.TxtPrimaryEmployee.Text, Me.DtpDate.Value, Me.TxtDocumentNo.Text, Me.DtpTo.Value, Txt18.Text, _
                            Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text, Me.Txt02Pjk_TunjPPHDisetahunkan.Text, Me.Txt02Pjk_TunjPPHRutin.Text, _
                            Me.Txt02Pjk_TunjPPHNonRutin.Text, Me.Txt02PjkTunjanganPPh.Text, Me.Txt01Pjk_GajiRutin.Text, Me.Txt01Pjk_Gaji.Text, _
                            Me.Txt03Pjk_TunjLainnya.Text, Me.Txt03Pjk_TunjLainnyaRutin.Text, Me.Txt04PjkHonorarium.Text, Me.Txt04PjkHonorariumRutin.Text, _
                            Me.Txt05PjkPremiAsuransi.Text, Me.Txt05PjkPremiAsuransiRutin.Text, Me.Txt06PjkNatura.Text, Me.Txt06PjkNaturaRutin.Text, _
                            Me.Txt07PjkTantiem.Text, Me.Txt07PjkTantiemRutin.Text, Me.Txt08PjkPenghasilanBruto.Text, Me.Txt08PjkPenghasilanBrutoRutin.Text, _
                            Me.Txt09PjkBiayaJabatan.Text, Me.Txt09PjkBiayaJabatanRutin.Text, Me.Txt10PjkIuranPensiun.Text, Me.Txt10PjkIuranPensiunRutin.Text, _
                            Me.Txt11PjkJmlPengurangan.Text, Me.Txt11PjkJmlPenguranganRutin.Text, Me.Txt12PjkJmlPenghasilanNetto.Text, Me.Txt13PjkJmlHasilMsPjkSblm.Text, _
                            Me.Txt14PjkJmlHasilNettoUPPh.Text, Me.Txt14PjkJmlHasilNettoUPPhRutin.Text, Me.Txt15Pjk_PTKP.Text, Me.Txt16Pjk_PKPSetahun.Text, _
                            Me.Txt16Pjk_PKPSetahunRutin.Text, Me.Txt17Pjk_PPH21.Text, Me.Txt17Pjk_PPH21Rutin.Text, Me.Txt18.Text, Me.Txt19Pjk_PPH21TerutangSetahun.Text, _
                            Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text, Me.Txt19Pjk_PPH21TerutangSebulanNonRutin.Text, Me.Txt19Pjk_PPH21TerutangSebulanRutin.Text, _
                            Me.Txt19Pjk_PPH21Terutang.Text, Me.TxtIncomeTax.Text, Me.TxtKurangBayarTax.Text, Me.Txt20.Text, Me.TxtLosesTaxObligation.Text, _
                            Me.TxtPeriodEnd.Text, Me.TxtPeriodCount.Text, vResign, Me.DtpDate.Value.Date, Me.TxtPeriodStart.Text, Me.TxtPrevIncome.Text, Me.TxtPrevTaxPaid.Text, _
                            Me.TBLPrevIncome)
                        End If

                        rowItemIterasi = Me.TBLItemIterasi.Newtbltrans_salary_item_iterationRow
                        With rowItemIterasi
                            .PPH_Terutang_Non_Rutin = CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) - CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text)
                            .PPH_Terutang_Rutin = CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text)
                            .primary = Me.GetPrimaryMax(Me.TBLItemIterasi)
                            .primary_transaksi = Me.TxtPrimary.Text
                            .Tunjangan_PPH_Rutin = CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text)
                            ._Tunjangan_PPH_Non_Rutin = CDec(Me.Txt02Pjk_TunjPPHDisetahunkan.Text) - CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text)
                        End With
                        Me.TBLItemIterasi.Rows.Add(rowItemIterasi)
                    End If


                    LanjutIterasi = False
                    If Me.TxtPeriodEnd.Text <> 12 And vResign = True And LosesTaxObligation = 1 Then
                        Debug.Print("Tunj. PPH Disetahunkan:" + Me.ObjCD.DigitGrouping(Me.Txt02Pjk_TunjPPHDisetahunkan.Text, 0) + " vs " + ObjCD.DigitGrouping(Math.Round((CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) / 12) * CDec(Me.TxtPeriodCount.Text), 0, MidpointRounding.AwayFromZero), 0))
                        Debug.Print("Tunj. PPH Rutin Disetahunkan:" + Me.ObjCD.DigitGrouping(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text, 0) + " vs " + ObjCD.DigitGrouping(Math.Round((CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) / 12) * CDec(Me.TxtPeriodCount.Text), 0, MidpointRounding.AwayFromZero), 0))

                        If CDec(Me.Txt02Pjk_TunjPPHDisetahunkan.Text) <> Math.Round((CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) / 12) * CDec(Me.TxtPeriodCount.Text), 0, MidpointRounding.AwayFromZero) Or _
                        CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text) <> Math.Round((CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text) / 12) * CDec(Me.TxtPeriodCount.Text), 0, MidpointRounding.AwayFromZero) Then
                            LanjutIterasi = True
                        End If
                    Else
                        If CDec(Me.Txt02Pjk_TunjPPHDisetahunkan.Text) <> CDec(Me.Txt19Pjk_PPH21TerutangSetahun.Text) Or _
                        CDec(Me.Txt02Pjk_TunjPPHDisetahunkanRutin.Text) <> CDec(Me.Txt19Pjk_PPH21TerutangSetahunRutin.Text) Then
                            LanjutIterasi = True
                        End If
                    End If
                End While
            End If

            Debug.Print("Tunj Lain Rutin: " + Me.ObjCD.DigitGrouping(CDec(Me.Txt03Pjk_TunjLainnyaRutin.Text), 0))
            Debug.Print("Tunj Lain: " + Me.ObjCD.DigitGrouping(CDec(Me.Txt03Pjk_TunjLainnya.Text), 0))
            Debug.Print("TxtIncomeTax: " + Me.ObjCD.DigitGrouping(Me.TxtIncomeTax.Text, 0))
            Me.HitungPPhFinalPesangon(Me.TxtPeriodEnd.Text, vResign, Me.TxtLosesTaxObligation.Text, Me.TxtPeriodCount.Text)
            Debug.Print("TxtIncomeTax: " + Me.ObjCD.DigitGrouping(Me.TxtIncomeTax.Text, 0))
            'MsgBox("Tunj. PPh Non Rutin: " + Me.Txt02Pjk_TunjPPHNonRutin.Text)
            'MsgBox("Tunj. PPh Rutin: " + Me.Txt02Pjk_TunjPPHRutin.Text)
            '#sinkron2018_pphcls
            'MsgBox("ganti disini untuk perbaikan data cmsi: resign atau akhir tahun")
            ' yang diganti:
            '1. kolom header: tunj. pph, tunj. pph rutin, tunj. pph non rutin
            '2. item tax amount bulannya untuk item tunjangan pph rutin dan non rutin pada nilai bulanannya (nilai tahunnya sudah bener)
            '3. otomatis nilai penghasil bruto dibawah ini akan menjadi bener
            If Me.cboTaxCalculation.SelectedIndex = 0 Then
                If Me.DtpTo.Value.Month <> 12 And vResign = False Then  
                    Me.Txt02PjkTunjanganPPh.Text = CDec(Me.Txt02Pjk_TunjPPHRutin.Text) + CDec(Txt02Pjk_TunjPPHNonRutin.Text)
                    Me.vPenghasilanBruto += CDec(Me.Txt02Pjk_TunjPPHNonRutin.Text) + CDec(Me.Txt02Pjk_TunjPPHRutin.Text) '- CDec(Me.Txt21Pjk_TunjPPhPesangon.Text) '- CDec(Me.Txt18.Text) '#sinkron2018_pphcls
                Else
                    Debug.Print("TxtKurangBayar: " + Me.ObjCD.DigitGrouping(Me.TxtKurangBayarTax.Text, 0))
                    Me.Txt02Pjk_TunjPPHNonRutin.Text = 0
                    For Each rowItemTax In Me.TBLItemTax.Rows
                        If rowItemTax.tipe_tax = ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH Then
                            If rowItemTax.is_rutin = 0 Then
                                rowItemTax.amount_month = 0
                            End If
                        End If
                    Next
                    Me.Txt02Pjk_TunjPPHRutin.Text = TxtKurangBayarTax.Text
                    For Each rowItemTax In Me.TBLItemTax.Rows
                        If rowItemTax.tipe_tax = ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH Then
                            If rowItemTax.is_rutin = 1 Then
                                Debug.Print("rowitemTax.amount: " + Me.ObjCD.DigitGrouping(rowItemTax.amount_month, 0))
                                rowItemTax.amount_month = CDec(Me.Txt02Pjk_TunjPPHRutin.Text) - CDec(Me.Txt21Pjk_TunjPPhPesangon.Text)
                                Debug.Print("rowitemTax.amount: " + Me.ObjCD.DigitGrouping(rowItemTax.amount_month, 0))
                            End If
                        End If
                    Next
                    Me.Txt02PjkTunjanganPPh.Text = CDec(Me.Txt02Pjk_TunjPPHRutin.Text) + CDec(Txt02Pjk_TunjPPHNonRutin.Text)
                    Me.vPenghasilanBruto += CDec(Me.Txt02Pjk_TunjPPHNonRutin.Text) + CDec(Me.Txt02Pjk_TunjPPHRutin.Text) - CDec(Me.Txt21Pjk_TunjPPhPesangon.Text) '- CDec(Me.Txt18.Text) '#sinkron2018_pphcls
                End If
                For Each rowItemTax In Me.TBLItemTax.Rows
                    If rowItemTax.tipe_tax = ClsTransaksiSalary_Item.EnumTaxElement.No8_Jumlah_Penghasilan_Bruto Then
                        rowItemTax.amount_month = Me.vPenghasilanBruto
                    End If
                Next
            End If

            Me.TxtKurangBayarTax.Text = Me.ObjCD.DigitGrouping(Me.TxtKurangBayarTax.Text, 0)
            Debug.Print("TxtKurangBayar: " + Me.ObjCD.DigitGrouping(Me.TxtKurangBayarTax.Text, 0))
            Me.TxtIncomeTax.Text = Me.ObjCD.DigitGrouping(Me.TxtIncomeTax.Text, 0)
            Debug.Print("TxtIncomeTax: " + Me.ObjCD.DigitGrouping(Me.TxtIncomeTax.Text, 0))

            ''#sinkron2018_pphcls
            If Me.DtpTo.Value.Month = 12 Or vResign = True Then
                If Me.Baru = True Then
                    Me.Txt_P_NoForm.Text = Me.Obj.GetNoBaru_BuatNPWP(Me.DtpDate.Value, Me.ObjCD)
                Else
                    If Me.DtpDate.Value.Month = Me.Obj.GetTanggalTransaksi(Me.TxtPrimary.Text).Month And _
                    Me.DtpDate.Value.Year = Obj.GetTanggalTransaksi(Me.TxtPrimary.Text).Year Then
                        Me.Txt_P_NoForm.Text = Me.Obj.GetNoForm1721A1(Me.TxtPrimary.Text)
                        Me.Txt_P_NoForm.Text = Me.Txt_P_NoForm.Text.Trim
                        If Me.Txt_P_NoForm.Text = "" Then
                            Me.Txt_P_NoForm.Text = Me.Obj.GetNoBaru_BuatNPWP(Me.DtpDate.Value, Me.ObjCD)
                        End If
                    Else
                        If Me.DtpDate.Value.Month <> Me.Obj.TanggalTransaksi.Month Or _
                        Me.DtpDate.Value.Year <> Me.Obj.TanggalTransaksi.Year Then
                            Me.Txt_P_NoForm.Text = Me.Obj.GetNoBaru_BuatNPWP(Me.DtpDate.Value, Me.ObjCD)
                        Else
                            Me.Txt_P_NoForm.Text = Me.Obj.GetNoForm1721A1(Me.TxtPrimary.Text)
                            Me.Txt_P_NoForm.Text = Me.Txt_P_NoForm.Text.Trim
                            If Me.Txt_P_NoForm.Text = "" Then
                                Me.Txt_P_NoForm.Text = Me.Obj.GetNoBaru_BuatNPWP(Me.DtpDate.Value, Me.ObjCD)
                            End If
                        End If
                    End If
                End If
            End If

            Me.TxtPrevTaxPaid.Text = Me.ObjCD.DigitGrouping(Me.TxtPrevTaxPaid.Text, 0)
            Me.TxtPrevIncome.Text = Me.ObjCD.DigitGrouping(Me.TxtPrevIncome.Text, 0)

            'MsgBox(Me.Txt20.Text)
            'MsgBox(Me.TxtKurangBayarTax.Text)
            'MsgBox(Me.ObjCD.DigitGrouping(Me.vPenghasilanBruto, 0))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vTBLOtherPayment1.Dispose()
        End Try
    End Sub

    Private Function GetPrimaryMax(ByVal vTBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable) As Double
        Dim rowItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Dim RetVal As Double = Double.MinValue

        Try
            For Each rowItemTax In vTBLItemTax.Rows
                If RetVal < rowItemTax.primary Then
                    RetVal = rowItemTax.primary
                End If
            Next

            If RetVal = Double.MinValue Then
                Return 0
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TambahkanDgnItemOtherPayment(ByVal vTBLOtherPayment2 As DseTransaksiSalary.tbltrans_salary_item_other1DataTable, _
    ByVal FaktorKaliDisetahunkan As Integer, ByVal vResign As Boolean)
        Dim TipeSalaryKomponen As FrmMasterSalaryFormula_Component.EnumType = FrmMasterSalaryFormula_Component.EnumType.Income
        Dim TipeElemenPajak As ClsTransaksiSalary_Item.EnumTaxElement
        Dim rowItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Dim rowOther As DseTransaksiSalary.tbltrans_salary_item_other1Row

        Try
            If vTBLOtherPayment2.Rows.Count > 0 Then
                For Each rowOther In vTBLOtherPayment2.Rows
                    If Me.ApakahAdaDalamDGV(rowOther.Identifier_Salary_Component) = False Then
                        TipeElemenPajak = Me.GetTipeElemenPajak(rowOther.Tax_Component)
                        If TipeElemenPajak <> ClsTransaksiSalary_Item.EnumTaxElement.NoTax And TipeElemenPajak <> ClsTransaksiSalary_Item.EnumTaxElement.No21_Pesangon Then
                            rowItemTax = Me.TBLItemTax.Newtbltrans_salary_itemtaxRow()
                            With rowItemTax
                                TipeSalaryKomponen = rowOther.Type
                                .tipe_tax = Me.GetTipeElemenPajak(rowOther.Tax_Component)
                                TipeElemenPajak = .tipe_tax
                                .amount_month = rowOther.Amount
                                If TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT Or _
                                TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH Or _
                                TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb Or _
                                TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya Or _
                                TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja Or _
                                TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH Or _
                                TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi Then
                                    Me.vPenghasilanBruto += .amount_month
                                End If
                                .identifier_salary_component = rowOther.Identifier_Salary_Component
                                .primary = Me.GetPrimaryMax(Me.TBLItemTax) + 1
                                .primary_transaksi = Me.TxtPrimary.Text
                                .is_rutin = rowOther.Is_Rutin
                                If Me.DtpTo.Value.Month = 12 Or vResign = True Then
                                    If .is_rutin = 1 Then
                                        .amount_gabungan_disetahunkan = .amount_month
                                        .remarks_gabungan = ""
                                        .amount_rutin_disetahunkan = .amount_month
                                        .remarks_rutin = ""
                                    Else
                                        .amount_gabungan_disetahunkan = .amount_month
                                        .remarks_gabungan = ""
                                        .amount_rutin_disetahunkan = 0
                                        .remarks_rutin = ""
                                    End If
                                Else
                                    If .is_rutin = 1 Then
                                        .amount_gabungan_disetahunkan = FaktorKaliDisetahunkan * .amount_month
                                        .remarks_gabungan = FaktorKaliDisetahunkan.ToString + " * " + Me.ObjCD.DigitGrouping(.amount_month, 0)
                                        .amount_rutin_disetahunkan = FaktorKaliDisetahunkan * .amount_month
                                        .remarks_rutin = ""
                                    Else
                                        .amount_gabungan_disetahunkan = .amount_month
                                        .remarks_gabungan = ""
                                        .amount_rutin_disetahunkan = 0
                                        .remarks_rutin = ""
                                    End If
                                End If
                            End With
                            Me.TBLItemTax.Rows.Add(rowItemTax)
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanDgnItemOtherPayment_Resign_EndOfYear(ByVal vTBLOtherPayment2 As DseTransaksiSalary.tbltrans_salary_item_other1DataTable, _
    ByVal FaktorKaliDisetahunkan As Integer)
        Dim TipeSalaryKomponen As FrmMasterSalaryFormula_Component.EnumType = FrmMasterSalaryFormula_Component.EnumType.Income
        Dim TipeElemenPajak As ClsTransaksiSalary_Item.EnumTaxElement
        Dim rowItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Dim rowOther As DseTransaksiSalary.tbltrans_salary_item_other1Row

        Try
            If vTBLOtherPayment2.Rows.Count > 0 Then
                For Each rowOther In vTBLOtherPayment2.Rows
                    TipeElemenPajak = Me.GetTipeElemenPajak(rowOther.Tax_Component)
                    TBLItemTax.DefaultView.RowFilter = "[identifier salary component] = '" + rowOther.Identifier_Salary_Component + "'"
                    If TBLItemTax.DefaultView.Count = 0 Then
                        rowItemTax = Me.TBLItemTax.Newtbltrans_salary_itemtaxRow()
                        With rowItemTax
                            TipeSalaryKomponen = rowOther.Type
                            .tipe_tax = Me.GetTipeElemenPajak(rowOther.Tax_Component)
                            TipeElemenPajak = .tipe_tax
                            .amount_month = 0
                            If TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT Or _
                            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH Or _
                            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb Or _
                            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya Or _
                            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja Or _
                            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH Or _
                            TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi Then
                                Me.vPenghasilanBruto += 0
                            End If
                            .identifier_salary_component = rowOther.Identifier_Salary_Component
                            .primary = Me.GetPrimaryMax(Me.TBLItemTax) + 1
                            .primary_transaksi = Me.TxtPrimary.Text
                            .is_rutin = rowOther.Is_Rutin
                            If .is_rutin = 1 Then
                                .amount_gabungan_disetahunkan = rowOther.Amount
                                .remarks_gabungan = ""
                                .amount_rutin_disetahunkan = rowOther.Amount
                                .remarks_rutin = ""
                            Else
                                .amount_gabungan_disetahunkan = rowOther.Amount
                                .remarks_gabungan = ""
                                .amount_rutin_disetahunkan = 0
                                .remarks_rutin = ""
                            End If
                        End With
                        Me.TBLItemTax.Rows.Add(rowItemTax)
                    Else
                        rowItemTax = Me.TBLItemTax.DefaultView(0).Row
                        With rowItemTax
                            If .is_rutin = 1 Then
                                .amount_gabungan_disetahunkan += rowOther.Amount
                                .amount_rutin_disetahunkan += rowOther.Amount
                            Else
                                .amount_gabungan_disetahunkan += rowOther.Amount
                                .amount_rutin_disetahunkan += 0
                            End If
                        End With
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungPPhFinalPesangon(ByVal vEndPeriod As Integer, _
    ByVal vIsResign As Boolean, ByVal vLosesTaxObligation As Integer, ByVal vPeriodCount As Integer)
        Dim rowDGV As DataGridViewRow
        Dim rowItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Dim rowItemIterasi, rowItemIterasiSebelum As DseTransaksiSalary.tbltrans_salary_item_iteration1Row
        Dim CboDgv As DataGridViewComboBoxCell
        Dim TipeSalaryKomponen As FrmMasterSalaryFormula_Component.EnumType = FrmMasterSalaryFormula_Component.EnumType.Income
        Dim TipeElemenPajak As ClsTransaksiSalary_Item.EnumTaxElement
        Dim DTATarif As New DseTransaksiSalaryTableAdapters.tblmaster_tarif_pph_final_pesangonTableAdapter
        Dim JmlIterasi As Double = 0
        Dim Selisih1 As Decimal
        Dim UseNPWP As ClsMasterKaryawan.EnumIsNPWP = ClsMasterKaryawan.EnumIsNPWP.Yes
        Dim JmlPesangon As Decimal = 0

        Try
            DTATarif.Connection.ConnectionString = Core_Module.GetConstring
            DTATarif.Fill(Me.vTBLTarifPPHFinal)

            'kosongkan item tax detail dan item iterasi
            'Me.TBLItemTax.Clear()
            Me.TBLItemIterasiPPHFinal.Clear()

            'kosongkan semua text box yang terkait dengan pajak penghasilan
            'Me.KosongkanFieldPerhitunganPajak()
            'Me.Obj.IsiKeteranganKolomTabPage3(Me)

            If Me.cboTaxCalculation.SelectedIndex = 3 Then
                'manual tax calculation
                Exit Sub
            End If

            'masukkan data item tax detail
            'hati2, untuk akhir tahun, harus memperhitungkan balance payroll dan memperhitungkan komponen gaji bulan2 sebelumnya
            For Each rowDGV In Me.dgvItem.Rows
                If CDec(rowDGV.Cells(Me.ColAmount.Index).Value) <> 0 Then
                    CboDgv = rowDGV.Cells(Me.ColTaxElement.Index)
                    TipeElemenPajak = Me.GetTipeElemenPajak(CboDgv.Items.IndexOf(rowDGV.Cells(Me.ColTaxElement.Index).Value))
                    If TipeElemenPajak = ClsTransaksiSalary_Item.EnumTaxElement.No21_Pesangon Then
                        rowItemTax = Me.TBLItemTax.Newtbltrans_salary_itemtaxRow()
                        With rowItemTax
                            CboDgv = rowDGV.Cells(Me.ColTipeKomponen.Index)
                            TipeSalaryKomponen = CboDgv.Items.IndexOf(rowDGV.Cells(Me.ColTipeKomponen.Index).Value)
                            CboDgv = rowDGV.Cells(Me.ColTaxElement.Index)
                            .tipe_tax = Me.GetTipeElemenPajak(CboDgv.Items.IndexOf(rowDGV.Cells(Me.ColTaxElement.Index).Value))
                            TipeElemenPajak = .tipe_tax
                            .amount_month = Me.GetAmountMonth(Me.ObjCD.RoundDown1(rowDGV.Cells(Me.ColAmount.Index).Value, 0), TipeSalaryKomponen, TipeElemenPajak) + _
                            Me.GetAmount_OtherPayment(rowDGV.Cells(Me.ColIdentifierKomponen.Index).Value)
                            JmlPesangon += .amount_month
                            .identifier_salary_component = rowDGV.Cells(Me.ColIdentifierKomponen.Index).Value
                            .primary = Me.GetPrimaryMax(Me.TBLItemTax) + 1
                            .primary_transaksi = Me.TxtPrimary.Text
                            .is_rutin = rowDGV.Cells(Me.ColIsRutin.Index).Value
                            .amount_gabungan_disetahunkan = 0
                            .remarks_gabungan = ""
                            .amount_rutin_disetahunkan = 0
                            .remarks_rutin = ""
                        End With
                        Me.TBLItemTax.Rows.Add(rowItemTax)
                    End If
                End If
            Next

            'hitung detail pajak untuk iterasi 0
            UseNPWP = Me.ObjKrywn.GetIsNPWP(Me.TxtPrimaryEmployee.Text)
            Me.Obj.HitungPPhFinalPesangon_Aktual(0, UseNPWP, Me.vTBLTarifPPHFinal, JmlPesangon, Me.TxtPrimary.Text, Me.TxtPrimaryEmployee.Text, _
            Me.DtpDate.Value, Me.TxtDocumentNo.Text, Me.TBLItemTax, Me.ObjCD, Me.DtpTo.Value, Me.Txt21Pjk_Pesangon.Text, _
            Me.Txt21Pjk_TunjPPhPesangon.Text, Me.Txt21Pjk_PPhFinalPesangon.Text, Me.vTarifPPhFinal, Me.Txt19Pjk_PPH21Terutang.Text, _
            Me.Txt18.Text, Me.TxtIncomeTax.Text, Me.TxtKurangBayarTax.Text, Me.Txt20.Text, vEndPeriod, vIsResign, vLosesTaxObligation, vPeriodCount, Me.TxtPrevTaxPaid.Text)
            rowItemIterasi = Me.TBLItemIterasiPPHFinal.Newtbltrans_salary_item_iteration1Row
            With rowItemIterasi
                .PPH_Terutang = CDec(Me.Txt21Pjk_PPhFinalPesangon.Text)
                .Tunjangan_PPH = CDec(Me.Txt21Pjk_TunjPPhPesangon.Text)
                .primary = Me.GetPrimaryMax(Me.TBLItemIterasiPPHFinal)
                .primary_transaksi = Me.TxtPrimary.Text
            End With
            Me.TBLItemIterasiPPHFinal.Rows.Add(rowItemIterasi)

            'jika gross up, maka hitung pajak untuk iterasi selanjutnya sampai dengan tunjangan PPH = PPH terhutang
            Debug.Print(Me.ObjCD.DigitGrouping(Me.TxtIncomeTax.Text, 0))
            If Me.cboTaxCalculation.SelectedIndex = 0 Then
                While CDec(Me.Txt21Pjk_PPhFinalPesangon.Text) <> CDec(Me.Txt21Pjk_TunjPPhPesangon.Text)
                    If JmlIterasi > 10000 Then
                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    'If JmlIterasi >= 1 Then
                    '    Exit Sub
                    'End If
                    JmlIterasi += 1
                    rowItemIterasiSebelum = Me.TBLItemIterasiPPHFinal.Rows(Me.TBLItemIterasiPPHFinal.Rows.Count - 1)
                    Selisih1 = rowItemIterasiSebelum.PPH_Terutang - rowItemIterasiSebelum.Tunjangan_PPH
                    If Selisih1 = 1 Then
                        Selisih1 = 0
                    End If

                    Me.Obj.HitungPPhFinalPesangon_Aktual(CDec(Me.Txt21Pjk_PPhFinalPesangon.Text) + Me.ObjCD.RoundDown1(Selisih1 / 2, 0), _
                    UseNPWP, Me.vTBLTarifPPHFinal, JmlPesangon, Me.TxtPrimary.Text, Me.TxtPrimaryEmployee.Text, _
                    Me.DtpDate.Value, Me.TxtDocumentNo.Text, Me.TBLItemTax, Me.ObjCD, Me.DtpTo.Value, Me.Txt21Pjk_Pesangon.Text, _
                    Me.Txt21Pjk_TunjPPhPesangon.Text, Me.Txt21Pjk_PPhFinalPesangon.Text, Me.vTarifPPhFinal, Me.Txt19Pjk_PPH21Terutang.Text, _
                    Me.Txt18.Text, Me.TxtIncomeTax.Text, Me.TxtKurangBayarTax.Text, Me.Txt20.Text, vEndPeriod, vIsResign, vLosesTaxObligation, vPeriodCount, Me.TxtPrevTaxPaid.Text)

                    rowItemIterasi = Me.TBLItemIterasiPPHFinal.Newtbltrans_salary_item_iteration1Row
                    With rowItemIterasi
                        .PPH_Terutang = CDec(Me.Txt21Pjk_PPhFinalPesangon.Text)
                        .Tunjangan_PPH = CDec(Me.Txt21Pjk_TunjPPhPesangon.Text)
                        .primary = Me.GetPrimaryMax(Me.TBLItemIterasiPPHFinal)
                        .primary_transaksi = Me.TxtPrimary.Text
                    End With
                    Me.TBLItemIterasiPPHFinal.Rows.Add(rowItemIterasi)
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnHitung()
        Dim Hasil As MsgBoxResult
        Dim RowDGV As DataGridViewRow
        'Dim StartTIme As Date

        Try
            If Me.ShowQuick_Edit = True Then
                If Me.NewQuick_Data = True Then
                    Hasil = MsgBoxResult.Yes
                Else
                    Hasil = MsgBox("Do you want to calculate transaction salary ?", MsgBoxStyle.YesNo, "Confirm")
                End If
            Else
                Hasil = MsgBox("Do you want to calculate transaction salary ?", MsgBoxStyle.YesNo, "Confirm")
            End If

            If Hasil = MsgBoxResult.Yes Then
                If Me.ApakahDataitemBenar() = True Then
                    If Me.Baru = True Then
                        'Hitung Loan Payment
                        Me.HitungLoanPayment(Me.ijK.TBLLoan)
                    End If
                    'Hitung Data Item berdasarkan salary formula
                    Me.BersihkanTBL(Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.TBLErrorFormula)
                    'StartTIme = Now
                    Me.ObjParam.FillDataOnly(Me.TBLParam1, Me.TBLParam2, Me.TBLWeekDay, Me.TBLHoliday, Me.TblKaryawan, Me.dtpPeriod.Value.Date, Me.DtpTo.Value.Date, Me.TxtPrimaryEmployee.Text)
                    'MsgBox("A: " + Me.ObjCD.DigitGrouping((StartTIme - Now).TotalSeconds, 1) + " s")
                    Me.Obj.FillSalaryComponent_DetailEmployee(Me.TBLSL_Kary, Me.TxtPrimaryEmployee.Text)
                    'Beres in Format Explaination perbaris
                    'StartTIme = Now
                    Me.CalculateHitungItemSalaryFormula()
                    'MsgBox("B: " + Me.ObjCD.DigitGrouping((StartTIme - Now).TotalSeconds, 1) + " s")
                    Me.HitungOverride(False, Me.dtpPeriod.Value, Me.DtpTo.Value)
                    'DGV item & format explaination other deduction perbaris
                    Me.CalculateHitungItemSalaryFormula_OtherDeduction(Me.FrmOtherDeduction)
                    Me.HitungOverride(True, Me.dtpPeriod.Value, Me.DtpTo.Value)

                    'Hitung Total Diheader
                    Me.HitungTotal()
                    '--------------------------------------------------------------------------
                    'Payment Data
                    Me.mPrimaryPaymentAkunDefault = Me.ObjMstSlry.GetPrimaryDefaultAccount(Me.TxtPrimaryEmployee.Text)
                    Me.BuatPaymentData(Me.Baru, True)
                    '--------------------------------------------------------------------------
                    Me.ObjItem.IsiDgvAttendance(Me.TxtPrimaryEmployee.Text, Me.DtpTo.Value, Me.dtpPeriod.Value, Me.FrmAtt)
                    Me.DummyDari = Me.dtpPeriod.Value.Date.ToString
                    Me.DummySampai = Me.DtpTo.Value.Date.ToString
                    Me.DummyCBOPeriod = Me.CboSalary.SelectedIndex.ToString
                    Me.DummyMetodeTax = Me.cboTaxCalculation.SelectedIndex.ToString
                    Me.DummyPrimaryEmployee = Me.TxtPrimaryEmployee.Text
                    'Hide Row Item Salary
                    If Me.ChkShowZeroAmount.Checked = False Then
                        If Me.dgvItem.Rows.Count > 0 Then
                            For Each RowDGV In Me.dgvItem.Rows
                                If CDec(RowDGV.Cells(Me.ColAmount.Index).Value) = 0 Then
                                    Me.dgvItem.Rows(RowDGV.Index).Visible = False
                                End If
                            Next
                        End If
                        Me.SetNumber(Me.dgvItem, True, False)
                    Else
                        If Me.dgvItem.Rows.Count > 0 Then
                            For Each RowDGV In Me.dgvItem.Rows
                                Me.dgvItem.Rows(RowDGV.Index).Visible = True
                            Next
                        End If
                        Me.SetNumber(Me.dgvItem, True, True)
                    End If
                    'Hide Row Item Other Deduction
                    If Me.ChkShowZeroAmount.Checked = False Then
                        If FrmOtherDeduction.dgvItem.Rows.Count > 0 Then
                            For Each RowDGV In FrmOtherDeduction.dgvItem.Rows
                                If CDec(RowDGV.Cells(FrmOtherDeduction.ColAmount.Index).Value) = 0 Then
                                    FrmOtherDeduction.dgvItem.Rows(RowDGV.Index).Visible = False
                                End If
                            Next
                        End If
                        Me.SetNumber(FrmOtherDeduction.dgvItem, True, False)
                    Else
                        If FrmOtherDeduction.dgvItem.Rows.Count > 0 Then
                            For Each RowDGV In FrmOtherDeduction.dgvItem.Rows
                                FrmOtherDeduction.dgvItem.Rows(RowDGV.Index).Visible = True
                            Next
                        End If
                        Me.SetNumber(FrmOtherDeduction.dgvItem, True, True)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungTotal()
        Dim rowDGV As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell

        Try
            Me.TxtIncome.Text = "0"
            Me.TxtDeduction.Text = "0"
            Me.TxtIncomeTax.Text = "0"
            Me.TxtSalaryReceived.Text = "0"
            Me.TxtCompanyExpense.Text = "0"
            Me.TxtEmployeeExpense.Text = "0"

            If Me.dgvItem.Rows.Count > 0 Then
                For Each rowDGV In Me.dgvItem.Rows
                    CboDgv = rowDGV.Cells(Me.ColTipeKomponen.Index)
                    If CboDgv.Items.IndexOf(rowDGV.Cells(Me.ColTipeKomponen.Index).Value) = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income Then
                        Me.TxtIncome.Text += CDec(rowDGV.Cells(Me.ColAmount.Index).Value)
                    ElseIf CboDgv.Items.IndexOf(rowDGV.Cells(Me.ColTipeKomponen.Index).Value) = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Deduction Then
                        Me.TxtDeduction.Text += CDec(rowDGV.Cells(Me.ColAmount.Index).Value)
                    ElseIf CboDgv.Items.IndexOf(rowDGV.Cells(Me.ColTipeKomponen.Index).Value) = Core_Module.FrmMasterSalaryFormula_Component.EnumType.CompanyExpense Then
                        Me.TxtCompanyExpense.Text += CDec(rowDGV.Cells(Me.ColAmount.Index).Value)
                    ElseIf CboDgv.Items.IndexOf(rowDGV.Cells(Me.ColTipeKomponen.Index).Value) = Core_Module.FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense Then
                        Me.TxtEmployeeExpense.Text += CDec(rowDGV.Cells(Me.ColAmount.Index).Value)
                    End If
                Next
            End If

            Me.TxtIncome.Text = Me.ObjCD.DigitGrouping(Me.TxtIncome.Text, Me.Def_Prec)
            Me.TxtDeduction.Text = Me.ObjCD.DigitGrouping(Me.TxtDeduction.Text, Me.Def_Prec)
            Me.TxtCompanyExpense.Text = Me.ObjCD.DigitGrouping(Me.TxtCompanyExpense.Text, Me.Def_Prec)
            Me.TxtEmployeeExpense.Text = Me.ObjCD.DigitGrouping(Me.TxtEmployeeExpense.Text, Me.Def_Prec)

            Me.TxtGrossSalary.Text = Me.ObjCD.DigitGrouping(Me.HitungGrossSalary(CDec(Me.TxtIncome.Text), _
            CDec(Me.TxtDeduction.Text), CDec(Me.TxtEmployeeExpense.Text)), Me.Def_Prec)
            Me.TxtOtherDedution.Text = "0"
            If Me.FrmOtherDeduction.dgvItem.Rows.Count > 0 Then
                For Each rowDGV In Me.FrmOtherDeduction.dgvItem.Rows
                    Me.TxtOtherDedution.Text += CDec(rowDGV.Cells(Me.FrmOtherDeduction.ColAmount.Index).Value)
                Next
            End If
            Me.TxtOtherDedution.Text = Me.ObjCD.DigitGrouping(Me.TxtOtherDedution.Text, Me.Def_Prec)
            '------------------Diganti dengan koding fitur terbaru------------------
            'Hitung Income Tax
            Me.TxtPeriodStart.Text = "1"
            Me.TxtPeriodEnd.Text = "12"
            Me.TxtPeriodCount.Text = "12"
            Me.TxtIsResign.Text = "0"

            Me.HitungPajak1()
            'hitung salary received
            Me.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(Me.HitungTakeHomePay(CDec(Me.TxtIncome.Text), CDec(Me.TxtDeduction.Text), _
            CDec(Me.TxtIncomeTax.Text), CDec(Me.TxtLoanInstallment.Text), CDec(Me.TxtOtherDedution.Text), CDec(Me.TxtKurangBayarTax.Text), _
            CDec(Me.TxtEmployeeExpense.Text)), Me.Def_Prec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RecalculatePajak()
        Try
            Me.HitungPajak1()
            'If Me.ObjKrywn.GetIsNPWP(Me.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.EnumIsNPWP.No Then
            Debug.Print("TxtIncomeTax: " + Me.ObjCD.DigitGrouping(Me.TxtIncomeTax.Text, 0))
            Me.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(Me.HitungTakeHomePay(CDec(Me.TxtIncome.Text), CDec(Me.TxtDeduction.Text), _
            CDec(Me.TxtIncomeTax.Text), CDec(Me.TxtLoanInstallment.Text), CDec(Me.TxtOtherDedution.Text), CDec(Me.TxtKurangBayarTax.Text), _
            CDec(Me.TxtEmployeeExpense.Text)), Me.Def_Prec)
            '    Me.TxtIncomeTax.Text = Me.ObjCD.DigitGrouping(CDec(Me.Txt19Pjk_PPH21Terutang.Text), Me.Def_Prec)
            '    Me.TxtKurangBayarTax.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtKurangBayarTax.Text), Me.Def_Prec)
            'Else
            '    'Salary Received
            '    Me.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(Me.HitungTakeHomePay(CDec(Me.TxtIncome.Text), CDec(Me.TxtDeduction.Text), _
            '    CDec(Me.TxtIncomeTax.Text), CDec(Me.TxtLoanInstallment.Text), CDec(Me.TxtOtherDedution.Text), CDec(Me.TxtKurangBayarTax.Text), _
            '    CDec(Me.TxtEmployeeExpense.Text)), Me.Def_Prec)
            '    Me.TxtIncomeTax.Text = Me.ObjCD.DigitGrouping(CDec(Me.Txt19Pjk_PPH21Terutang.Text), Me.Def_Prec)
            '    Me.TxtKurangBayarTax.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtKurangBayarTax.Text), Me.Def_Prec)
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Public Sub MasukanTxtIncomeTax()
    '    Try
    '        If Me.cboTaxCalculation.SelectedIndex = ClsTransaksiSalary.EnumTax.Gross_Up Then
    '            If Me.ApakahAkhirBulan = True Then
    '                Me.Obj.ClearObjectTxtPPh21(Me)
    '                Me.Obj.HitungIncomeTax_GrossUp_Bulan12(Me)
    '            Else
    '                Me.Obj.ClearObjectTxtPPh21(Me)
    '                Me.Obj.HitungIncomeTax_GrossUp_DariAwalSampe11(Me)
    '            End If
    '        End If
    '        If Me.cboTaxCalculation.SelectedIndex = ClsTransaksiSalary.EnumTax.Netto Or _
    '        Me.cboTaxCalculation.SelectedIndex = ClsTransaksiSalary.EnumTax.Gross Then
    '            If Me.ApakahAkhirBulan = True Then
    '                Me.Obj.ClearObjectTxtPPh21(Me)
    '                Me.Obj.HitungIncomeTax_GrossNNetto_Bulan12(Me)
    '            Else
    '                Me.Obj.ClearObjectTxtPPh21(Me)
    '                Me.Obj.HitungIncomeTax_GrossNNetto_DariAwalSampe11(Me)
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Public Function ApakahAkhirBulan() As Boolean
    '    Try
    '        If Me.DtpDate.Value.Date.Month = 12 Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function




    Private Sub BtnHitung_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnHitung.Click
        Dim ObjMstSlry As New ClsTransaksiSalary
        Try
            '8. batasan bulan 12, cek bulan 1-11 metode tax calculation harus sama
            'If ObjMstSlry.ApakahSelama1TahunTaxCalculationSama(Me.DtpDate.Value.Date, Me.cboTaxCalculation.SelectedIndex, Me.TxtPrimaryEmployee.Text, -1) = False Then
            '    MsgBox("Please cek tax calculation because you must choose the same tax calculation for around one year. ", MsgBoxStyle.Information, "Information")
            '    Me.cboTaxCalculation.Focus()
            '    Exit Sub
            'End If

            Me.FungsiBtnHitung()
            If NewQuick_Data = False Then
                My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Beep)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim ObjLoanPay As New ClsTransaksiLoanPayment
        Try
            If Me.TxtRemarks.Enabled = False Then
                If Me.TxtDocumentNo.Text <> "" Then
                    Me.ijK.ReadOnlyFromSalary = True
                Else
                    Exit Sub
                End If
                'LP
                Me.ijK.Baru = False
                Me.ijK.UserCode = Me.UserCode
                Me.ijK.PrimaryLoanPaymentFromSalary = Me.TxtPrimaryLoan.Text
                Me.ijK.ShowFromLinkLabel = False
                Me.ijK.ShowFromSalary = True
                Me.ijK.KosongkanField()
                Me.ijK.Show()
                Me.ijK.Fungsi_Load()
                Me.ijK.Hide()
                If Me.TxtPrimaryLoan.Text <> "-1" Then
                    Me.ijK.ShowFromLinkLabel = True
                    Me.ijK.BackColorAll = Me.BackColorAll
                    If Me.TxtRemarks.Enabled = False Then
                        Me.ijK.ReadOnlyFromSalary = True
                    End If
                    ijK.Enabled = True
                    Me.ijK.ShowDialog()
                    Me.ijK.ShowFromLinkLabel = False
                End If
            Else
                Me.ijK.ReadOnlyFromSalary = False
                If Me.Baru = True Then
                    If Me.ijK.dgvItem.Rows.Count > 0 Then
                        Me.ijK.ShowFromLinkLabel = True
                        Me.ijK.BackColorAll = Me.BackColorAll
                        If Me.TxtRemarks.Enabled = False Then
                            If Me.TxtDocumentNo.Text <> "" Then
                                Me.ijK.ReadOnlyFromSalary = True
                            Else
                                Exit Sub
                            End If
                        End If
                        Me.ijK.ShowDialog()
                        Me.TxtLoanInstallment.Text = Me.ObjCD.DigitGrouping(Me.ijK.TotalLoanInstallment, Me.ijK.Def_Prec)
                        'Hitung Salary Received
                        Me.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(Me.HitungTakeHomePay(CDec(Me.TxtIncome.Text), CDec(Me.TxtDeduction.Text), _
                        CDec(Me.TxtIncomeTax.Text), CDec(Me.TxtLoanInstallment.Text), CDec(Me.TxtOtherDedution.Text), CDec(Me.TxtKurangBayarTax.Text), _
                        CDec(Me.TxtEmployeeExpense.Text)), Me.Def_Prec)
                        'If Me.cboTaxCalculation.SelectedIndex = 0 Then
                        '    'Gross Up
                        '    Me.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtIncome.Text) - CDec(Me.TxtDeduction.Text) - CDec(Me.TxtLoanInstallment.Text), Me.Def_Prec)
                        'Else
                        '    'Netto
                        '    Me.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtIncome.Text) - CDec(Me.TxtDeduction.Text) - CDec(Me.TxtLoanInstallment.Text) - CDec(Me.TxtIncomeTax.Text), Me.Def_Prec)
                        'End If
                        Me.ijK.ShowFromLinkLabel = False
                    End If
                End If
                If Me.Baru = False Then
                    If Me.ijK.TxtDocumentNo.Text = "" Then
                        Me.TxtPrimaryLoan.Text = "-1"
                    End If
                    If Me.TxtPrimaryLoan.Text <> "-1" Then
                        Me.ijK.ShowFromLinkLabel = True
                        Me.ijK.BackColorAll = Me.BackColorAll
                        If Me.TxtRemarks.Enabled = False Then
                            Me.ijK.ReadOnlyFromSalary = True
                        End If
                        Me.ijK.ShowDialog()
                        Me.TxtLoanInstallment.Text = Me.ObjCD.DigitGrouping(Me.ijK.TotalLoanInstallment, Me.ijK.Def_Prec)
                        'Hitung Salary Received
                        Me.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(Me.HitungTakeHomePay(CDec(Me.TxtIncome.Text), CDec(Me.TxtDeduction.Text), _
                        CDec(Me.TxtIncomeTax.Text), CDec(Me.TxtLoanInstallment.Text), CDec(Me.TxtOtherDedution.Text), CDec(Me.TxtKurangBayarTax.Text), _
                        CDec(Me.TxtEmployeeExpense.Text)), Me.Def_Prec)
                        'If Me.cboTaxCalculation.SelectedIndex = 0 Then
                        '    'Gross Up
                        '    Me.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtIncome.Text) - CDec(Me.TxtDeduction.Text) - CDec(Me.TxtLoanInstallment.Text), Me.Def_Prec)
                        'Else
                        '    'Netto
                        '    Me.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtIncome.Text) - CDec(Me.TxtDeduction.Text) - CDec(Me.TxtLoanInstallment.Text) - CDec(Me.TxtIncomeTax.Text), Me.Def_Prec)
                        'End If
                        Me.ijK.ShowFromLinkLabel = False
                        'If Me.Baru = False Then
                        '    If Me.ijK.dgvItem.Rows.Count <= 0 Then
                        '        Me.TxtPrimaryLoan.Text = ObjLoanPay.GetPrimary(Me.ijK.TxtDocumentNo.Text)
                        '    End If
                        'End If
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblIncomeTax_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblIncomeTax.LinkClicked
        'Dim k As New FrmTrans_Gaji_IncomeTax
        'Try
        '    k.ik = Me
        '    k.ShowDialog()

        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
        Dim k As New FrmTrans_Gaji_IncomeTax2
        Dim ResignDate As Date = #1/1/1950#
        Dim StartWorkingPeriod As Date = #1/1/1950#
        Dim TBLItemTaxD As New DseTransaksiSalary.tbltrans_salary_itemtaxDataTable
        Dim rowItemTax, rowItemTaxD As DseTransaksiSalary.tbltrans_salary_itemtaxRow

        Try
            If Me.dgvItem.RowCount > 0 Then
                ResignDate = Me.ObjKrywn.GetResignDate(Me.TxtPrimaryEmployee.Text)

                For Each rowItemTax In Me.TBLItemTax.Rows
                    rowItemTaxD = TBLItemTaxD.Newtbltrans_salary_itemtaxRow
                    For ij As Integer = 0 To TBLItemTaxD.Columns.Count - 1
                        rowItemTaxD.Item(ij) = rowItemTax.Item(ij)
                    Next
                    TBLItemTaxD.Rows.Add(rowItemTaxD)
                Next
                k.vTBLItemTax = TBLItemTaxD

                k.vEmployeeCode = Me.TxtEmployeeCode1.Text
                k.vTransactionDate = Me.DtpDate.Value
                k.vEmployeeName = Me.TxtEmployeeName.Text
                k.vNIK_Kary = Me.Txt_P_NoNIK.Text
                k.vNPWP = Me.Txt_P_A_NPWP.Text
                k.vStatusNPWP = Me.Txt_P_StatusNPWP.Text
                k.vKodeNegaraDomisili = Me.Txt_P_KodeNegaraDomisili.Text
                k.vJenisKelamin = Me.Txt_P_JenisKelamin.Text
                k.vAlamatPajak1 = Me.Txt_P_Alamat1.Text
                k.vAlamatPajak2 = Me.Txt_P_Alamat2.Text
                k.vPrimaryEmployee = Me.TxtPrimaryEmployee.Text
                k.vDepartement = Me.TxtDepartment.Text
                k.vPosition = Me.TxtPosition.Text
                k.vIsResign = Me.TxtIsResign.Text

                If Me.DtpTo.Value.Month = 12 Or (ResignDate.Date >= Me.dtpPeriod.Value.Date And ResignDate.Date <= Me.DtpTo.Value.Date) Then
                    k.vEndOfYear = True
                    StartWorkingPeriod = Me.ObjKrywn.GetFirstTaxPeriod(Me.TxtPrimaryEmployee.Text)
                    If StartWorkingPeriod.Year < Me.DtpTo.Value.Date.Year Then
                        StartWorkingPeriod = Me.ObjCD.GetFisrtDateOfTheYear(Me.DtpTo.Value)
                    End If
                    k.vStartPeriod = StartWorkingPeriod
                    If ResignDate = #1/1/1950# Then
                        ResignDate = Me.DtpTo.Value.Date
                    End If
                End If
                If Me.DtpTo.Value.Month = 12 Then
                    k.vEndPeriod = Me.DtpTo.Value
                End If
                If k.vIsResign = 1 Then
                    k.vEndPeriod = ResignDate
                End If
                k.ShowDialog()

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLItemTaxD.Rows.Clear()
            TBLItemTaxD.Dispose()
        End Try
    End Sub

    Private Sub dtpPeriod_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriod.Enter
        Try
            Me.BufferDari = Me.dtpPeriod.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpTo_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpTo.Enter
        Try
            Me.BufferSampai = Me.DtpTo.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpPeriod_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriod.Leave
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpTo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpTo.Leave
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Function PengecekanDTPStartEndDanPosisiRangeKetikaHitung(ByVal vFrom As Date, ByVal vTo As Date, _
    'ByVal vTBLStartDate As DseTransaksiSalary.tbltrans_salary1DataTable, Optional ByVal ApakahCekBtnHitung As Boolean = False) As Boolean
    '    Dim rowStartDate As DseTransaksiSalary.tbltrans_salary1Row
    '    Dim Filter As String = ""
    '    Try
    '        If ApakahCekBtnHitung = False Then
    '            vTBLStartDate.DefaultView.RowFilter = ""
    '            Filter = "[Start Date] <= #" + Format(vFrom.Date, "MM/dd/yyyy") + "# AND [End Date] >= #" + Format(vTo.Date, "MM/dd/yyyy") + "#"
    '            vTBLStartDate.DefaultView.RowFilter = Filter
    '            If vTBLStartDate.DefaultView.Count > 0 Then
    '                Return False
    '            End If
    '        Else
    '            vTBLStartDate.DefaultView.RowFilter = ""
    '            Filter = "[Start Date] <= #" + Format(vFrom.Date, "MM/dd/yyyy") + "# AND [End Date] >= #" + Format(vTo.Date, "MM/dd/yyyy") + "#"
    '            While vTo.Date < vFrom.Date

    '                vTo.Date.AddDays(-1)
    '            End While
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function


    Private Sub LblAttendance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAttendance.LinkClicked
        Dim j As New FrmTrans_Gaji_ShowAttendance
        Try
            If Me.TxtDocumentNo.Text <> "" Then
                j.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
                Me.ObjItem.IsiDgvAttendance(Me.TxtPrimaryEmployee.Text, Me.DtpTo.Value, Me.dtpPeriod.Value, j)
                j.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            j.Dispose()
        End Try
    End Sub

    Private Sub cboTaxCalculation_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTaxCalculation.Enter
        Try
            Me.BufferMetodeTax = Me.cboTaxCalculation.SelectedIndex.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RecalcuateTakeHomePay()
        Dim RowDGV As DataGridViewRow
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.RecalculatePajak()
                Me.TxtSalaryReceived.Text = Me.ObjCD.DigitGrouping(Me.HitungTakeHomePay(CDec(Me.TxtIncome.Text), CDec(Me.TxtDeduction.Text), _
                CDec(Me.TxtIncomeTax.Text), CDec(Me.TxtLoanInstallment.Text), CDec(Me.TxtOtherDedution.Text), CDec(Me.TxtKurangBayarTax.Text), _
                CDec(Me.TxtEmployeeExpense.Text)), Me.Def_Prec)
                'Update Payment
                If Me.FrmPaymentData.DGVPayment.Rows.Count = 1 Then
                    RowDGV = Me.FrmPaymentData.DGVPayment.Rows(0)
                    RowDGV.Cells(Me.FrmPaymentData.ColDetailPayment_Amount.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtSalaryReceived.Text, Me.Def_Prec)
                    Me.TxtPayment.Text = Me.ObjCD.DigitGrouping(Me.TxtSalaryReceived.Text, Me.Def_Prec)
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboTaxCalculation_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTaxCalculation.SelectedIndexChanged
        Try
                Me.RecalcuateTakeHomePay()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboSalary_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboSalary.Enter
        Try
            Me.BufferCBOPeriod = Me.CboSalary.SelectedIndex.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowZeroAmount_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkShowZeroAmount.CheckedChanged
        Dim RowDGV As DataGridViewRow
        Try
            If Me.ChkShowZeroAmount.Checked = False Then
                If Me.dgvItem.Rows.Count > 0 Then
                    For Each RowDGV In Me.dgvItem.Rows
                        If CDec(RowDGV.Cells(Me.ColAmount.Index).Value) = 0 Then
                            Me.dgvItem.Rows(RowDGV.Index).Visible = False
                        End If
                    Next
                End If
                Me.SetNumber(Me.dgvItem, True, False)
            Else
                If Me.dgvItem.Rows.Count > 0 Then
                    For Each RowDGV In Me.dgvItem.Rows
                        Me.dgvItem.Rows(RowDGV.Index).Visible = True
                    Next
                End If
                Me.SetNumber(Me.dgvItem, True, True)
            End If
            If Me.ChkShowZeroAmount.Checked = False Then
                If FrmOtherDeduction.dgvItem.Rows.Count > 0 Then
                    For Each RowDGV In FrmOtherDeduction.dgvItem.Rows
                        If CDec(RowDGV.Cells(FrmOtherDeduction.ColAmount.Index).Value) = 0 Then
                            FrmOtherDeduction.dgvItem.Rows(RowDGV.Index).Visible = False
                        End If
                    Next
                End If
                Me.SetNumber(FrmOtherDeduction.dgvItem, True, False)
            Else
                If FrmOtherDeduction.dgvItem.Rows.Count > 0 Then
                    For Each RowDGV In FrmOtherDeduction.dgvItem.Rows
                        FrmOtherDeduction.dgvItem.Rows(RowDGV.Index).Visible = True
                    Next
                End If
                Me.SetNumber(FrmOtherDeduction.dgvItem, True, True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblOtherDeduction_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblOtherDeduction.LinkClicked
        Try
            Me.FrmOtherDeduction.BackColorAll = Me.BackColorAll
            Me.FrmOtherDeduction.ijGaji = Me
            Me.FrmOtherDeduction.ijPay = Me.FrmPaymentData
            If Me.TxtRemarks.Enabled = False Then
                Me.FrmOtherDeduction.NonAktif = False
            Else
                Me.FrmOtherDeduction.NonAktif = True
            End If
            If Me.Baru = False Then
                If Me.TxtRemarks.Enabled = False Then
                    Me.FrmOtherDeduction.DipanggilKetika_PosisiNonAktif = True
                Else
                    Me.FrmOtherDeduction.DipanggilKetika_PosisiNonAktif = False
                End If
            End If
            Me.FrmOtherDeduction.ShowDialog()
            'Hitung ulang payement data
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorkerPPh21_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorkerPPh21.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorkerPPh21_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorkerPPh21.RunWorkerCompleted
        Try
            If Me.Baru = True Then
                Me.Txt_P_NoForm.Text = Me.Obj.GetNoBaru_BuatNPWP(Me.DtpDate.Value, Me.ObjCD)
            Else
                If Me.DtpDate.Value.Month = Me.Obj.GetTanggalTransaksi(Me.TxtPrimary.Text).Month And _
                Me.DtpDate.Value.Year = Obj.GetTanggalTransaksi(Me.TxtPrimary.Text).Year Then
                    Me.Txt_P_NoForm.Text = Me.Obj.GetNoForm1721A1(Me.TxtPrimary.Text)
                Else
                    If Me.DtpDate.Value.Month <> Me.Obj.TanggalTransaksi.Month Or _
                    Me.DtpDate.Value.Year <> Me.Obj.TanggalTransaksi.Year Then
                        Me.Txt_P_NoForm.Text = Me.Obj.GetNoBaru_BuatNPWP(Me.DtpDate.Value, Me.ObjCD)
                    Else
                        Me.Txt_P_NoForm.Text = Me.Obj.GetNoForm1721A1(Me.TxtPrimary.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPay.Click
        Dim FrmPay As New FrmTransaksiPayment
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtDocumentNo.Text <> "" Then
                If Me.CekTimestamp() = True Then
                    If Me.Obj.GetPaymentStatus(Me.TxtPrimary.Text) = ClsTransaksiSalary.EnumPaymentStatus.unpaid Then
                        'New data payment
                        FrmPay.ControlBox = False
                        FrmPay.mDipanggilDariCalulateSalary = True
                        FrmPay.mPrimarySalary = Me.TxtPrimary.Text
                        FrmPay.mTanggalSalary = Me.DtpDate.Value
                        FrmPay.UserCode = Me.UserCode
                        FrmPay.ShowDialog()
                    Else
                        'Edit data payment
                        FrmPay.ControlBox = False
                        FrmPay.mDipanggilDariCalulateSalary = True
                        FrmPay.mPrimarySalary = Me.TxtPrimary.Text
                        FrmPay.mTanggalSalary = Me.DtpDate.Value
                        FrmPay.UserCode = Me.UserCode
                        FrmPay.mPrimaryPaymentDariSalary = Me.Obj.GetPrimaryPayment(Me.TxtPrimary.Text)
                        FrmPay.ShowDialog()
                    End If
                    'Pengecekan Status Paid dan UnPaid
                    If Me.Obj.GetPaymentStatus(Me.TxtPrimary.Text) = ClsTransaksiSalary.EnumPaymentStatus.paid Then
                        Me.DtpPaymentStatus.Value = Me.Obj.GetPaymentDate(Me.TxtPrimary.Text)
                        Me.CboPaymentStatus.SelectedIndex = ClsTransaksiSalary.EnumPaymentStatus.paid
                        Me.LblPaid.Text = "Paid"
                    Else
                        Me.DtpPaymentStatus.Value = #1/1/1950#
                        Me.CboPaymentStatus.SelectedIndex = ClsTransaksiSalary.EnumPaymentStatus.unpaid
                        Me.LblPaid.Text = "Unpaid"
                    End If
                    'Status paid visible
                    If Me.LblPaid.Text = "Paid" Then
                        Me.CboPaymentStatus.Visible = True
                        Me.DtpPaymentStatus.Visible = True
                        Me.LblPaid.Visible = False
                        Me.LblPaymentStatus.Visible = True
                    Else
                        Me.CboPaymentStatus.Visible = False
                        Me.DtpPaymentStatus.Visible = False
                        Me.LblPaid.Visible = True
                        Me.LblPaymentStatus.Visible = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmPay.Dispose()
        End Try
    End Sub

    Private Sub LblPayment_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblPayment.LinkClicked
        Dim AmountPayment As Decimal = 0
        Try
            Me.FrmPaymentData.BackColorAll = Me.BackColorAll
            Me.FrmPaymentData.Def_Prec = Me.Def_Prec
            Me.FrmPaymentData.PrimarySalary = CDbl(Me.TxtPrimary.Text)
            Me.FrmPaymentData.Baru = Me.Baru
            Me.FrmPaymentData.FrmSalary = Me
            Me.FrmPaymentData.DGVPayment.AllowUserToAddRows = True
            If Me.TxtRemarks.Enabled = False Then
                Me.FrmPaymentData.NonAktif = False
                Me.FrmPaymentData.DGVPayment.Rows.Clear()
                Me.FrmPaymentData.DGVPayment.AllowUserToAddRows = False
                Me.Obj.BuatPaymentData_Posisi_Edit(Me.FrmPaymentData.DGVPayment, Me.TxtPrimary.Text, Me.FrmPaymentData, _
                Me.ObjAkun, Me.Def_Prec, AmountPayment)
            Else
                Me.FrmPaymentData.NonAktif = True
            End If
            Me.FrmPaymentData.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnHapus_Quick()
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtDocumentNo.Text <> "" Then
                If Me.CekTimestamp() = True Then
                    If Me.Obj.ApakahBolehHapus(Me.TxtPrimary.Text, Me.UserCode, True, True, Me.dgvItem, Me) = False Then
                        Me.btnFind.Focus()
                        Exit Sub
                    Else
                        'LP
                        Me.ijK.Baru = False
                        Me.ijK.PrimaryLoanPaymentFromSalary = Me.TxtPrimaryLoan.Text
                        Me.ijK.ShowFromSalary = True
                        Me.ijK.KosongkanField()
                        Me.ijK.UserCode = Me.UserCode
                        Me.ijK.Show()
                        Me.ijK.Fungsi_Load()
                        Me.ijK.Hide()
                        'Other Deduction
                        Me.FrmOtherDeduction.Show()
                        Me.FrmOtherDeduction.LoadData()
                        Me.FrmOtherDeduction.Hide()
                        Me.FrmOtherDeduction.Location = New Point((Me.FrmOtherDeduction.Width * -1) - 1, (Me.FrmOtherDeduction.Height * -1) - 1)

                        If Me.DeleteQuick_Data = True Then
                            If Me.Obj.Hapus(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, _
                            Me, Me.ijK, True) = True Then
                                Me.KosongkanField()
                                Me.btnNew.Focus()
                            End If
                        ElseIf Me.DeleteQuick_Data = False Then
                            hasil = MsgBox("Delete data?", MsgBoxStyle.YesNo, "Confirm")
                            If hasil = MsgBoxResult.Yes Then
                                If Me.Obj.Hapus(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, _
                                 Me, Me.ijK, True) = True Then
                                    Me.KosongkanField()
                                    Me.btnNew.Focus()
                                    MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnStop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnStop.Click
        Try
            If MsgBox("Stop creating calculate salary?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DariBtnStop = True
                Me.NonAktifkanDetail()
                Me.DariCancelBtn = True
                Me.DariBtnPause = False
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbOtherPayment_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbOtherPayment.Click
        Dim ij As New FrmTrans_Gaji_OtherPayment

        Try
            If Me.vTBLOtherPayment.Rows.Count > 0 Then
                ij.vTbl = Me.vTBLOtherPayment
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub llbEditPPh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbEditPPh.Click
        Dim PPhRutin, PPhNonRutin As Decimal
        Dim ij As New FrmTrans_Gaji_EditPPh

        Try
            PPhRutin = 0
            PPhNonRutin = 0
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.HitungPajak1(ij.TxtPPhRutin.Text, ij.TxtPPHNonRutin.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Public Sub CopyData(ByVal vTBLSumber As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_variableDataTable, _
    ByVal vTBLTarget As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_variableDataTable)
        Dim rowTarget As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_variableRow
        Dim rowSumber As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_variableRow

        Try
            vTBLTarget.Rows.Clear()

            For Each rowSumber In vTBLSumber.Rows
                rowTarget = vTBLTarget.Newtblmaster_salary_formula_salarycomponent_variableRow

                For ij As Integer = 0 To vTBLSumber.Columns.Count - 1
                    rowTarget.Item(ij) = rowSumber.Item(ij)
                Next

                vTBLTarget.Rows.Add(rowTarget)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub CopyData(ByVal vTBLSumber As DseTransaksiSalary.tblmaster_salary_formula_salarycomponentDataTable, _
  ByVal vTBLTarget As DseTransaksiSalary.tblmaster_salary_formula_salarycomponentDataTable)
        Dim rowTarget As DseTransaksiSalary.tblmaster_salary_formula_salarycomponentRow
        Dim rowSumber As DseTransaksiSalary.tblmaster_salary_formula_salarycomponentRow

        Try
            vTBLTarget.Rows.Clear()

            For Each rowSumber In vTBLSumber.Rows
                rowTarget = vTBLTarget.Newtblmaster_salary_formula_salarycomponentRow

                For ij As Integer = 0 To vTBLSumber.Columns.Count - 1
                    rowTarget.Item(ij) = rowSumber.Item(ij)
                Next

                vTBLTarget.Rows.Add(rowTarget)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub CopyData(ByVal vTBLSumber As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_constraintDataTable, _
  ByVal vTBLTarget As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_constraintDataTable)
        Dim rowTarget As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_constraintRow
        Dim rowSumber As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_constraintRow

        Try
            vTBLTarget.Rows.Clear()

            For Each rowSumber In vTBLSumber.Rows
                rowTarget = vTBLTarget.Newtblmaster_salary_formula_salarycomponent_constraintRow

                For ij As Integer = 0 To vTBLSumber.Columns.Count - 1
                    rowTarget.Item(ij) = rowSumber.Item(ij)
                Next

                vTBLTarget.Rows.Add(rowTarget)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbPrevIncome_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbPrevIncome.LinkClicked
        Try
            If CDec(Me.TxtPrevTaxPaid.Text) <> 0 Or CDec(Me.TxtPrevIncome.Text) <> 0 Then
                If Me.TBLPrevIncome.Rows.Count > 0 Then
                    Dim ij As New FrmTrans_Gaji_MasaSblm
                    rowPrevIncome = TBLPrevIncome.Rows(0)
                    ij.vPrimaryLoad = Me.rowPrevIncome.primary
                    ij.ShowDialog()
                    ij = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class