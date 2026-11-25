Option Explicit On
Option Strict Off

Public Class FrmMasterSalaryFormula_OtherDeduction
    Public UserCode As String = ""
    Public UserLama As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCurr As New clsMasterCurrency
    Public DefPrec As Integer = Me.ObjCurr.GetDecimal_DefaultCurrency()
    Dim Baru As Boolean = True
    Public ObjDed As New ClsTrans_OtherDeduction
    Dim ObjSlip As New ClsMasterSalaryFormula_SlipDesc
    Dim ObjAkun As New ClsMasterAkun1
    Public ObjEmp As New ClsMasterKaryawan
    Public ObjDept As New ClsMasterDepartement
    Public ObjJob As New ClsMasterJobPosition
    Dim TipeAkun As ClsMasterAkun1.EnumTipeAkun = ClsMasterAkun1.EnumTipeAkun.Liabilities
    Dim TipeAkun2 As ClsMasterAkun1.EnumTipeAkun = ClsMasterAkun1.EnumTipeAkun.Expense

    Public TblVar As New DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_variableDataTable
    Public TblCons As New DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_constraintDataTable
    Dim buffer As String = ""
    Dim RowIdx As Integer = 0
    Dim ObjCd As New ClsCekData
    Dim sudah As Boolean = False

    Private Function BlokSimbol(ByVal obj As TextBox, ByVal e As Windows.Forms.KeyPressEventArgs) As Boolean
        Try
            If e.KeyChar = "'" Or _
            e.KeyChar = "%" Or _
            e.KeyChar = "*" Or _
            e.KeyChar = "#" Then
                obj.ReadOnly = True
                obj.BackColor = Me.BackColorAll
            Else
                obj.ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndexItem(ByVal vData As String, ByVal vDgv As DataGridView, _
    ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In vDgv.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(e.ColumnIndex).Value) = False Then
                    If rowDGV.Cells(e.ColumnIndex).Value.ToString.ToUpper = vData.ToUpper Then
                        If rowDGV.Index <> e.RowIndex Then
                            Return rowDGV.Index
                        End If
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub SalinKeComponent(ByVal Frm As FrmMasterSalaryFormula_Component2)
        Try
            Frm.DefPrec = Me.DefPrec
            Frm.TxtFormula.Text = Me.TxtFormula.Text
            Frm.CboRoundingType.SelectedIndex = Me.CboRoundingType.SelectedIndex
            Frm.CboRoundingUnit.Text = Me.TxtRoundingUnit.Text
            Frm.TxtYearlyMinValue.Text = Me.TxtYearlyMinValue.Text
            Frm.TxtYearlyMaxValue.Text = Me.TxtYearlyMaxValue.Text
            Frm.TxtMonthlyMinValue.Text = Me.TxtMonthlyMinValue.Text
            Frm.TxtMonthlyMaxValue.Text = Me.TxtMonthlyMaxValue.Text
            Frm.DtpYear.Value = GetFirstDateofTheMonth(Me.DtpYearlyDate.Value)
            Frm.mType = Me.cboCalculatePeriod.SelectedIndex
            Frm.mTblVar = Me.TblVar
            Frm.mTblConst = Me.TblCons
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinDariComponent(ByVal Frm As FrmMasterSalaryFormula_Component2)
        Try
            Me.TxtFormula.Text = Frm.TxtFormula.Text
            Me.CboRoundingType.SelectedIndex = Frm.CboRoundingType.SelectedIndex
            Me.TxtRoundingUnit.Text = Frm.CboRoundingUnit.Text
            Me.TxtYearlyMinValue.Text = Frm.TxtYearlyMinValue.Text
            Me.TxtYearlyMaxValue.Text = Frm.TxtYearlyMaxValue.Text
            Me.TxtMonthlyMinValue.Text = Frm.TxtMonthlyMinValue.Text
            Me.TxtMonthlyMaxValue.Text = Frm.TxtMonthlyMaxValue.Text
            Me.DtpYearlyDate.Value = GetFirstDateofTheMonth(Frm.DtpYear.Value)
            Me.cboCalculateOn.SelectedIndex = Me.DtpYearlyDate.Value.Month - 1
            Me.TblVar = Frm.mTblVar
            Me.TblCons = Frm.mTblConst
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim RowVar As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_variableRow
        Dim RowCons As DseTrans_OtherDeduction.tblmaster_salary_formula_otherdeduction_constraintRow
        Try
            'tidak boleh kosong
            KoleksiNamaData.Add("Deduction Code")
            KoleksiTextBox.Add(Me.TxtCode)
            KoleksiNamaData.Add("Deduction Description")
            KoleksiTextBox.Add(Me.TxtDescription)
            KoleksiNamaData.Add("Salary Slip Code")
            KoleksiTextBox.Add(Me.TxtSlipCode)
            KoleksiNamaData.Add("Referred Account")
            KoleksiTextBox.Add(Me.TxtRefferedCode)

            If CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'tidak boleh double
            If Me.Baru = True Then
                If Me.ObjDed.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                    Me.TxtCode.Focus()
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.ObjDed.GetCodeByPrimary(Me.TxtPrimaryOther.Text)
                If CodeLama <> Me.TxtCode.Text Then
                    If Me.ObjDed.ApakahSudahAda(Me.TxtCode.Text) = True Then
                        MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                        Me.TxtCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses(edit)
                Dim OldTimeStamp As Date = Me.ObjDed.GetTimeStampByPrimary(Me.TxtPrimaryOther.Text)
                If Me.ObjDed.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjDed.CariDetail(Me.TxtPrimaryOther.Text, Me.TblVar, Me.TblCons, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'If Me.ObjCd.CekNumerik(Me.TxtFormula.Text) = True Then
            '    If Me.TxtFormula.Text <= 0 Then
            '        MsgBox("Please insert salary formula", MsgBoxStyle.Information, "Information")
            '        Me.BtnEditFormula.Focus()
            '        Return False
            '    End If
            'Else
            If Me.TxtFormula.Text.Trim(" ") = "" Then
                MsgBox("Please insert salary formula", MsgBoxStyle.Information, "Information")
                Me.BtnEditFormula.Focus()
                Return False
            End If
            'End If

            If Me.cboCalculatePeriod.Text = "" Then
                MsgBox("Please select calculate period", MsgBoxStyle.Information, "Information")
                Me.cboCalculatePeriod.Focus()
                Return False
            End If

            If Me.cboCalculateOn.Text = "" Then
                MsgBox("Please select calculate on", MsgBoxStyle.Information, "Information")
                Me.cboCalculateOn.Focus()
                Return False
            End If

            'minimal harus ada 1 item
            If Me.DgvItem.RowCount <= 1 Then
                MsgBox("Please input employee item", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek calculation period di formula harus sama dengan di form other deduction
            If Me.TblVar.Rows.Count > 0 Then
                For Each RowVar In Me.TblVar.Rows
                    If RowVar.Type <> 3 Then
                        If RowVar.Type <> Me.cboCalculatePeriod.SelectedIndex Then
                            MsgBox("Calculation period is different from salary component variable type, please do adjustment.", MsgBoxStyle.Information, "Information")
                            Me.cboCalculatePeriod.Focus()
                            Return False
                        End If
                    End If
                Next
            End If
            If Me.TblCons.Rows.Count > 0 Then
                For Each RowCons In TblCons.Rows
                    If RowCons.Type <> 3 Then
                        If RowCons.Type <> Me.cboCalculatePeriod.SelectedIndex Then
                            MsgBox("Calculation period is different from salary component constraint type, please do adjustment.", MsgBoxStyle.Information, "Information")
                            Me.cboCalculatePeriod.Focus()
                            Return False
                        End If
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub AktifkanDetail()

        Try
            KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            If Me.Baru = True Then
                Me.cboCalculateOn.SelectedIndex = 0
                Me.cboCalculatePeriod.SelectedIndex = 0
            Else
                'If Me.ObjDed.ApakahBolehEdit(Me.UserCode, Me.TxtPrimaryOther.Text, False, False) = False Then
                '    KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
                '    Me.btnOK.Enabled = True
                '    Me.btnCancel.Enabled = True
                'End If
            End If
            Me.DgvItem.ReadOnly = False
            CancelButton = Me.btnCancel
            Me.DgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)

            Me.DgvItem.ReadOnly = True
            CancelButton = Me.btnExit
            Me.DgvItem.AllowUserToAddRows = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongKanDetail()
        Try
            KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimaryOther.Text = -1
            Me.TxtPrimaryReffered.Text = -1
            Me.TxtPrimarySlip.Text = -1
            Me.TxtFormula.Text = 0
            Me.DgvItem.Rows.Clear()
            Me.cboCalculateOn.SelectedIndex = 0
            Me.cboCalculatePeriod.SelectedIndex = 0
            Me.CboRoundingType.SelectedIndex = 0
            Me.TxtRoundingUnit.Text = 1
            Me.TxtYearlyMinValue.Text = Me.ObjCd.DigitGrouping(0, Me.DefPrec)
            Me.TxtYearlyMaxValue.Text = Me.ObjCd.DigitGrouping(0, Me.DefPrec)
            Me.TxtMonthlyMaxValue.Text = Me.ObjCd.DigitGrouping(0, Me.DefPrec)
            Me.TxtMonthlyMinValue.Text = Me.ObjCd.DigitGrouping(0, Me.DefPrec)
            Me.DtpYearlyDate.Value = GetFirstDateofTheMonth(Now)

            Me.TblVar.Rows.Clear()
            Me.TblCons.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Try
            With Me.ObjDed
                Me.TxtPrimaryOther.Text = .Primary
                Me.TxtCode.Text = .Code
                Me.TxtDescription.Text = .Desc
                Me.TxtPrimarySlip.Text = .PrimarySlip
                Me.TxtSlipCode.Text = Me.ObjSlip.GetCodeByPrimary(.PrimarySlip)
                Me.TxtSlipDesc.Text = Me.ObjSlip.GetDescByPrimary(.PrimarySlip)
                Me.TxtPrimaryReffered.Text = .PrimaryRefAcc
                Me.TxtRefferedCode.Text = Me.ObjAkun.GetAccountCode(.PrimaryRefAcc)
                Me.TxtRefferedDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(.PrimaryRefAcc)
                Me.TxtFormula.Text = .Formula
                Me.cboCalculatePeriod.SelectedIndex = .CalPeriod
                Me.cboCalculateOn.SelectedIndex = .CalOn
                Me.TxtYearlyMaxValue.Text = .YearlyMaxVal
                Me.TxtMonthlyMaxValue.Text = .MonthlyMaxValue
                Me.TxtRoundingUnit.Text = CInt(.RoundingUnit)
                Me.CboRoundingType.SelectedIndex = .RoundingType
                Me.TxtYearlyMinValue.Text = .YearlyMinVal
                Me.TxtMonthlyMinValue.Text = .MonthlyMinVal
                Me.DtpYearlyDate.Value = .YearlyDate
                Me.UserLama = .CreatedBy

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeObject()
        Try
            With Me.ObjDed
                .Primary = Me.TxtPrimaryOther.Text
                .Code = Me.TxtCode.Text
                .Desc = Me.TxtDescription.Text
                .PrimarySlip = Me.TxtPrimarySlip.Text
                .PrimaryRefAcc = Me.TxtPrimaryReffered.Text
                .Formula = Me.TxtFormula.Text
                .CalPeriod = Me.cboCalculatePeriod.SelectedIndex
                .CalOn = Me.cboCalculateOn.SelectedIndex
                .YearlyMaxVal = Me.TxtYearlyMaxValue.Text
                .MonthlyMaxValue = Me.TxtMonthlyMaxValue.Text
                .RoundingUnit = Me.TxtRoundingUnit.Text
                .RoundingType = Me.CboRoundingType.SelectedIndex
                .YearlyMinVal = Me.TxtYearlyMinValue.Text
                .MonthlyMinVal = Me.TxtMonthlyMinValue.Text
                .YearlyDate = Me.DtpYearlyDate.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEditFormula.Click
        Dim ij As New Core_Module.FrmMasterSalaryFormula_Component2
        Try
            Me.SalinKeComponent(ij)
            ij.BackColorAll = Me.BackColorAll
            ij.Buffer_CalculatePeriod = Me.cboCalculatePeriod.Text
            ij.ShowDialog()
            If ij.salin = True Then
                Me.SalinDariComponent(ij)
            End If

            Me.SetCBOCalculatePeriod()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub FrmMasterSalaryFormula_OtherDeduction_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObject()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjDed.SimpanBaru(Me.UserCode, Me) = True Then
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
                            If Me.ObjDed.SimpanPerubahan(Me.UserCode, Me) = True Then
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

    Private Sub FrmMasterSalaryFormula_OtherDeduction_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.NonAktifkanDetail()
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.Baru = True
            Me.AktifkanDetail()
            Me.KosongKanDetail()
            Me.TxtCode.Focus()
            If Me.cboCalculatePeriod.SelectedIndex = 0 Then
                Me.cboCalculateOn.Enabled = False
            End If
            Me.tslCreatedBy.Text = "Created By: " + UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            Me.SetCBOCalculatePeriod()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindSlip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindSlip.Click
        Dim frm As New FrmMasterSalaryFormula_SlipDesc_Cari
        Try
            frm.txtCari.Text = Me.TxtSlipCode.Text
            frm.BackColorAll = Me.BackColorAll
            frm.ShowDialog()
            If frm.Primary > 0 Then
                Me.TxtPrimarySlip.Text = frm.Primary
                Me.TxtSlipCode.Text = Me.ObjSlip.GetCodeByPrimary(frm.Primary)
                Me.TxtSlipDesc.Text = Me.ObjSlip.GetDescByPrimary(frm.Primary)
            Else
                Me.TxtPrimarySlip.Text = -1
                Me.TxtSlipCode.Text = ""
                Me.TxtSlipDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm.Dispose()
        End Try
    End Sub

    Private Sub TxtSlipCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSlipCode.GotFocus
        Me.TxtSlipCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtSlipCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSlipCode.KeyDown
        Dim primary As Double = -1
        Try
            If e.KeyValue = 13 Then
                If Me.ObjSlip.ApakahSudahAda(Me.TxtSlipCode.Text) Then
                    primary = Me.ObjSlip.GetPrimaryByCode(Me.TxtSlipCode.Text)
                    Me.TxtPrimarySlip.Text = primary
                    Me.TxtSlipCode.Text = Me.ObjSlip.GetCodeByPrimary(primary)
                    Me.TxtSlipDesc.Text = Me.ObjSlip.GetDescByPrimary(primary)
                    Me.cboCalculatePeriod.Focus()
                Else
                    Me.BtnFindSlip.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSlipCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSlipCode.Leave
        Dim primary As Double = -1
        Try
            If Me.TxtSlipCode.Text <> "" Then
                If Me.ObjSlip.ApakahSudahAda(Me.TxtSlipCode.Text) Then
                    primary = Me.ObjSlip.GetPrimaryByCode(Me.TxtSlipCode.Text)
                    Me.TxtPrimarySlip.Text = primary
                    Me.TxtSlipCode.Text = Me.ObjSlip.GetCodeByPrimary(primary)
                    Me.TxtSlipDesc.Text = Me.ObjSlip.GetDescByPrimary(primary)
                Else
                    Me.BtnFindSlip.PerformClick()
                End If
            Else
                Me.TxtPrimarySlip.Text = -1
                Me.TxtSlipCode.Text = ""
                Me.TxtSlipDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindReffered_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindReffered.Click
        Dim frm As New FrmMasterAkun_Cari1
        Dim hasil As DialogResult
        Dim TipeAkunHasil As ClsMasterAkun1.EnumTipeAkun

        Try
            frm.BackColorAll = Me.BackColorAll
            frm.txtCari.Text = Me.TxtRefferedCode.Text
            frm.ShowDialog()
            If frm.mPrimaryAkun > -1 Then
                'cek header non aktif
                'cek tipe akun, jika bukan liabilitas kasih peringatan
                TipeAkunHasil = Me.ObjAkun.GetTipeAkun(frm.mPrimaryAkun)
                If TipeAkunHasil <> Me.TipeAkun And _
                TipeAkunHasil <> Me.TipeAkun2 Then
                    hasil = MsgBox("Usually this data uses liabilities or expense type account, do you want to input this account?", _
                    MsgBoxStyle.YesNo, "Confirm")
                    If hasil = Windows.Forms.DialogResult.Yes Then
                        Me.TxtPrimaryReffered.Text = frm.mPrimaryAkun
                        Me.TxtRefferedCode.Text = Me.ObjAkun.GetAccountCode(frm.mPrimaryAkun)
                        Me.TxtRefferedDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(frm.mPrimaryAkun)
                        If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtRefferedCode.Text) = False Then
                            MsgBox("Please input active and non header account", MsgBoxStyle.Information, "Information")
                            Me.TxtPrimaryReffered.Text = -1
                            Me.TxtRefferedCode.Text = ""
                            Me.TxtRefferedDesc.Text = ""
                        End If
                    Else
                        Me.TxtPrimaryReffered.Text = -1
                        Me.TxtRefferedCode.Text = ""
                        Me.TxtRefferedDesc.Text = ""
                    End If
                Else
                    Me.TxtPrimaryReffered.Text = frm.mPrimaryAkun
                    Me.TxtRefferedCode.Text = Me.ObjAkun.GetAccountCode(frm.mPrimaryAkun)
                    Me.TxtRefferedDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(frm.mPrimaryAkun)
                End If
            Else
                Me.TxtPrimaryReffered.Text = -1
                Me.TxtRefferedCode.Text = ""
                Me.TxtRefferedDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRefferedCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRefferedCode.GotFocus
        Me.TxtRefferedCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtRefferedCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRefferedCode.KeyDown
        Dim primary As Double = -1
        Dim hasil As MsgBoxResult
        Dim TipeAkunHasil As ClsMasterAkun1.EnumTipeAkun

        Try
            If e.KeyValue = 13 Then
                'active non header
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtRefferedCode.Text) = True Then
                    primary = Me.ObjAkun.GetPrimary(Me.TxtRefferedCode.Text)
                    'cek tipe akun
                    TipeAkunHasil = Me.ObjAkun.GetTipeAkun(primary)

                    If TipeAkunHasil <> Me.TipeAkun And _
                    TipeAkunHasil <> Me.TipeAkun2 Then
                        hasil = MsgBox("Usually this data uses liabilities or expense type account, do you want to input this account?", _
                        MsgBoxStyle.YesNo, "Confirm")

                        If hasil = Windows.Forms.DialogResult.Yes Then
                            Me.TxtPrimaryReffered.Text = primary
                            Me.TxtRefferedCode.Text = Me.ObjAkun.GetAccountCode(primary)
                            Me.TxtRefferedDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(primary)
                        Else
                            Me.TxtPrimaryReffered.Text = -1
                            Me.TxtRefferedCode.Text = ""
                            Me.TxtRefferedDesc.Text = ""
                        End If
                    Else
                        Me.TxtPrimaryReffered.Text = primary
                        Me.TxtRefferedCode.Text = Me.ObjAkun.GetAccountCode(primary)
                        Me.TxtRefferedDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(primary)
                    End If
                Else
                    Me.BtnFindReffered.PerformClick()
                End If

                Me.BtnEditFormula.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRefferedCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRefferedCode.Leave
        Dim primary As Double = -1
        Dim hasil As DialogResult
        Dim TipeAkunHasil As ClsMasterAkun1.EnumTipeAkun
        Try
            If Me.TxtRefferedCode.Text <> "" Then
                'active non header
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtRefferedCode.Text) = True Then
                    primary = Me.ObjAkun.GetPrimary(Me.TxtRefferedCode.Text)
                    'cek tipe akun
                    TipeAkunHasil = Me.ObjAkun.GetTipeAkun(primary)
                    If TipeAkunHasil <> Me.TipeAkun And _
                    TipeAkunHasil <> Me.TipeAkun2 Then
                        hasil = MsgBox("Usually this data uses liabilities or expense type account, do you want to input this account?", _
                        MsgBoxStyle.YesNo, "Confirm")
                        If hasil = Windows.Forms.DialogResult.Yes Then
                            Me.TxtPrimaryReffered.Text = primary
                            Me.TxtRefferedCode.Text = Me.ObjAkun.GetAccountCode(primary)
                            Me.TxtRefferedDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(primary)
                        Else
                            Me.TxtPrimaryReffered.Text = -1
                            Me.TxtRefferedCode.Text = ""
                            Me.TxtRefferedDesc.Text = ""
                        End If
                    Else
                        Me.TxtPrimaryReffered.Text = primary
                        Me.TxtRefferedCode.Text = Me.ObjAkun.GetAccountCode(primary)
                        Me.TxtRefferedDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(primary)
                    End If
                Else
                    Me.BtnFindReffered.PerformClick()
                End If
            Else
                Me.TxtPrimaryReffered.Text = -1
                Me.TxtRefferedCode.Text = ""
                Me.TxtRefferedDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColRemarks.Index Then
                If Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value <> "" Then
                    'jika edit cek calculate salary
                    If Me.Baru = False Then
                        If Me.ObjDed.CekCalculate(Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value, Me.TxtPrimaryOther.Text) = True Then
                            MsgBox("This data can not be changed because it is already used in calculate salary", MsgBoxStyle.Information, "Information")
                            e.Cancel = True
                        Else
                            Me.buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.ColEmployeeCode.Index Then
                'jika edit cek calculate salary
                If Me.Baru = False Then
                    If Me.ObjDed.CekCalculate(Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value, Me.TxtPrimaryOther.Text) = True Then
                        MsgBox("This data can not be changed because it is already used in calculate salary", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                    Else
                        Me.buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    Me.buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Dim primary As Double = -1
        Try
            If e.ColumnIndex = Me.ColEmployeeCode.Index Then
                If String.IsNullOrEmpty(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    'hapus
                    If Me.DgvItem.RowCount > 1 Then
                        Me.RowIdx = e.RowIndex
                        Me.BgwHapus.RunWorkerAsync()
                        Me.BgwHapus.Dispose()
                    End If
                Else
                    If Me.ObjEmp.ApakahSudahAda(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        primary = Me.ObjEmp.GetPrimaryByCode(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = primary
                        Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(primary)
                        Me.DgvItem.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(primary)
                        Me.DgvItem.Item(Me.ColNIK.Index, e.RowIndex).Value = Me.ObjEmp.GetNIK(primary)
                        Me.DgvItem.Item(Me.ColDept.Index, e.RowIndex).Value = Me.ObjDept.GetDeptDesc(Me.ObjEmp.GetDeptByPrimary(primary))
                        Me.DgvItem.Item(Me.ColPosition.Index, e.RowIndex).Value = Me.ObjJob.GetDesc(Me.ObjEmp.GetJobByPrimary(primary))
                    Else
                        Dim frm As New FrmMasterKaryawanCari
                        frm.BackColorAll = Me.BackColorAll
                        frm.txtCari.Text = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        frm.ShowDialog()
                        If frm.PrimaryEmploye > 0 Then
                            Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = frm.PrimaryEmploye
                            Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(frm.PrimaryEmploye)
                            Me.DgvItem.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(frm.PrimaryEmploye)
                            Me.DgvItem.Item(Me.ColNIK.Index, e.RowIndex).Value = Me.ObjEmp.GetNIK(frm.PrimaryEmploye)
                            Me.DgvItem.Item(Me.ColDept.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(Me.ObjEmp.GetDeptByPrimary(frm.PrimaryEmploye))
                            Me.DgvItem.Item(Me.ColPosition.Index, e.RowIndex).Value = Me.ObjJob.GetCode(Me.ObjEmp.GetJobByPrimary(frm.PrimaryEmploye))
                        Else
                            'hapus
                            If Me.DgvItem.RowCount > 1 Then
                                Me.RowIdx = e.RowIndex
                                Me.BgwHapus.RunWorkerAsync()
                                Me.BgwHapus.Dispose()
                            End If
                        End If
                    End If
                    'cek tidak boleah ada yang sama
                    If Me.GetIndexItem(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.DgvItem, e) > -1 Then
                        MsgBox("Employee code has been used, please change the employee code.", MsgBoxStyle.Information, "Information")
                        If Me.DgvItem.RowCount > 1 Then
                            If Me.buffer = "" Then
                                Me.RowIdx = e.RowIndex
                                Me.BgwHapus.RunWorkerAsync()
                                Me.BgwHapus.Dispose()
                            Else
                                primary = Me.ObjEmp.GetPrimaryByCode(Me.buffer)
                                Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = primary
                                Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(primary)
                                Me.DgvItem.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(primary)
                                Me.DgvItem.Item(Me.ColNIK.Index, e.RowIndex).Value = Me.ObjEmp.GetNIK(primary)
                                Me.DgvItem.Item(Me.ColDept.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(Me.ObjEmp.GetDeptByPrimary(primary))
                                Me.DgvItem.Item(Me.ColPosition.Index, e.RowIndex).Value = Me.ObjJob.GetCode(Me.ObjEmp.GetJobByPrimary(primary))
                            End If
                        End If
                    End If
                End If
            End If
            If e.ColumnIndex = Me.ColRemarks.Index Then
                If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                    MsgBox("Can't input symbol ' , % , * , #", MsgBoxStyle.Information, "Information")
                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwHapus_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwHapus.DoWork
        Threading.Thread.Sleep(1)
    End Sub

    Private Sub BgwHapus_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwHapus.RunWorkerCompleted
        Me.DgvItem.Rows.RemoveAt(RowIdx)
    End Sub

    Private Sub DgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItem.RowsAdded
        Dim RowDgv As DataGridViewRow
        Try
            For Each RowDgv In Me.DgvItem.Rows
                RowDgv.Cells(Me.ColNo.Index).Value = RowDgv.Index + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvItem.RowsRemoved
        Dim RowDgv As DataGridViewRow
        Try
            For Each RowDgv In Me.DgvItem.Rows
                RowDgv.Cells(Me.ColNo.Index).Value = RowDgv.Index + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Dim RowDgv As DataGridViewRow
        Try
            For Each RowDgv In Me.DgvItem.Rows
                RowDgv.Cells(Me.ColNo.Index).Value = RowDgv.Index + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvItem.UserDeletedRow
        Dim RowDgv As DataGridViewRow
        Try
            For Each RowDgv In Me.DgvItem.Rows
                RowDgv.Cells(Me.ColNo.Index).Value = RowDgv.Index + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'cek data kosong
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObject()
                hasil = MsgBox("Save data ?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjDed.SimpanBaru(Me.UserCode, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                        End If
                    Else
                        If Me.ObjDed.SimpanPerubahan(Me.UserCode, Me) Then
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
            If Me.TxtCode.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongKanDetail()
                    Else
                        Me.NonAktifkanDetail()
                        Me.ObjDed.CariDetail(Me.TxtPrimaryOther.Text, Me.TblVar, Me.TblCons, Me)
                        Me.SalinKeField()
                        Me.cboCalculateOn.Enabled = False
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongKanDetail()
                Me.NonAktifkanDetail()
            End If
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            Me.Baru = False
            If Me.TxtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtCode.Text <> "" Then
                Me.ObjDed.CariDetail(Me.TxtPrimaryOther.Text, Me.TblVar, Me.TblCons, Me)
                'If Me.ObjDed.ApakahBolehEdit(Me.UserCode, Me.TxtPrimaryOther.Text) = True Then
                Me.AktifkanDetail()
                Me.TxtCode.Focus()
                'End If
            End If
            If Me.cboCalculatePeriod.SelectedIndex = 0 Or Me.cboCalculatePeriod.SelectedIndex = 1 Then
                Me.cboCalculateOn.Enabled = False
            End If

            Me.SetCBOCalculatePeriod()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCode.Text <> "" Then
                hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ObjDed.Hapus(Me.TxtPrimaryOther.Text, Me.UserCode) = True Then
                        Me.KosongKanDetail()
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
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterSalaryFormula_OtherDeduction_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.primary = -1 Then
                Me.ObjDed.CariDetail(hj.primary, Me.TblVar, Me.TblCons, Me)
                Me.SalinKeField()
                Me.btnEdit.Focus()
                Me.NonAktifkanDetail()

                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub TxtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.GotFocus
        Me.TxtCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCode.KeyPress
        Me.BlokSimbol(Me.TxtCode, e)
    End Sub

    Private Sub TxtCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.LostFocus
        Me.TxtCode.BackColor = Color.White
    End Sub

    Private Sub TxtDescription_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDescription.GotFocus
        Me.TxtDescription.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtDescription_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtDescription.KeyPress
        Me.BlokSimbol(Me.TxtDescription, e)
    End Sub

    Private Sub TxtDescription_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDescription.LostFocus
        Me.TxtDescription.BackColor = Color.White
    End Sub

    Private Sub TxtSlipCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSlipCode.LostFocus
        Me.TxtSlipCode.BackColor = Color.White
    End Sub

    Private Sub TxtRefferedCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRefferedCode.LostFocus
        Me.TxtRefferedCode.BackColor = Color.White
    End Sub

    Private Sub cboCalculatePeriod_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCalculatePeriod.Enter
        Try
            Me.buffer = Me.cboCalculatePeriod.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCBOCalculatePeriod()
        Try
            If Me.btnOK.Enabled = True Then
                If Me.TblVar.Rows.Count > 0 Or Me.TblCons.Rows.Count > 0 Then
                    Me.cboCalculatePeriod.Enabled = False
                Else
                    Me.cboCalculatePeriod.Enabled = True
                End If
            Else
                Me.cboCalculatePeriod.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCalculatePeriod_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCalculatePeriod.Leave
        Try
            'tidak bisa ganti tipe period, jika sudah ada item constraint atau item variable
            If Me.btnOK.Enabled = True Then
                If Me.buffer <> Me.cboCalculatePeriod.Text Then
                    If Me.TblVar.Rows.Count > 0 Or Me.TblCons.Rows.Count > 0 Then
                        MsgBox("This data have constrain and variable in it's formula, can't change calculation period", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCalculatePeriod_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCalculatePeriod.SelectedIndexChanged
        Try
            If Me.cboCalculatePeriod.SelectedIndex <> ClsTrans_OtherDeduction.EnumCalPeriod.Yearly Then
                Me.cboCalculateOn.SelectedIndex = 0
                Me.cboCalculateOn.Enabled = False
            Else
                Me.cboCalculateOn.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCalculateOn_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCalculateOn.SelectedIndexChanged
        Try
            Me.DtpYearlyDate.Value = (Me.cboCalculateOn.SelectedIndex + 1).ToString + "/1/" + Now.Year.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvItem.UserDeletingRow
        Try
            If Me.Baru = False Then
                If Me.ObjDed.CekCalculate(Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.Row.Index).Value, Me.TxtPrimaryOther.Text) = True Then
                    MsgBox("This data can not be deleted because it is already used in calculate salary", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellContentClick

    End Sub
End Class