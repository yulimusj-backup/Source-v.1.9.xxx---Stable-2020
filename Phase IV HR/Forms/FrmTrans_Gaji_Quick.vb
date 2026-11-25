Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms
Public Class FrmTrans_Gaji_Quick
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjQuick As New ClsTransaksiSalary_Quick
    Dim FrmSalary As New FrmTrans_Gaji
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim Buffer As String = ""
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

    Public Enum EnumLink As Integer
        Print
        Pay
        Delete
        Email
    End Enum

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahBolehPrint() As Boolean
        Dim RowDgv As DataGridViewRow
        Dim Ketemu As Boolean = False
        Try
            If Me.DtpPeriod.Value.Date > Me.DtpTo.Value Then
                MsgBox("Period to cannot less than period from.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.DGVSalaryCalculation.Rows.Count <= 0 Then
                MsgBox("Please input data item.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.DGVSalaryCalculation.Rows.Count > 0 Then
                For Each RowDgv In Me.DGVSalaryCalculation.Rows
                    If RowDgv.Cells(Me.ColChkPrint.Index).Value = True Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    MsgBox("Please checked coloumn print.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Dim FrmCetak As New FrmTrans_Gaji_Cetak_IGI_Options
        Dim Mindate As Date = Date.MaxValue
        Dim maxdate As Date = Date.MinValue
        Dim RowDgv As DataGridViewRow

        Try
            'Me.PrintDialog1.ShowDialog()
            'Me.PrintDialog1.PrinterSettings.PrinterName
            'Dim rpt As New CrpLap_Cuti
            'Dim crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
            'crv.ReportSource = rpt
            'rpt.PrintOptions.PrinterName = Me.PrintDialog1.PrinterSettings.PrinterName
            'rpt.PrintToPrinter(1, True, 1, 100)
            '---
            If Me.ApakahBolehPrint = True Then
                If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.SalaryCalculation, ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    Dim jk As New FrmTrans_Gaji_Cetak
                    jk.UseCustom = True
                    jk.QuickSalary = True
                    jk.vListPrimarySalary.Clear()
                    jk.vStartDate = Me.DtpPeriod.Value
                    jk.vEndDate = Me.DtpTo.Value

                    For Each RowDgv In Me.DGVSalaryCalculation.Rows
                        If RowDgv.Cells(Me.ColChkPrint.Index).Value = True Then
                            If jk.vListPrimarySalary.Contains(CDbl(RowDgv.Cells(Me.ColPrimaryTransaksi.Index).Value)) = False Then
                                jk.vListPrimarySalary.Add(CDbl(RowDgv.Cells(Me.ColPrimaryTransaksi.Index).Value))
                            End If
                        End If
                    Next
                    jk.ShowDialog()
                    jk = Nothing
                Else
                    FrmCetak.mQuickSalary_Form = True
                    FrmCetak.mTanggalTrans_Dari = Me.DtpPeriod.Value.Date
                    FrmCetak.mTanggalTrans_Sampai = Me.DtpTo.Value.Date
                    If Me.DGVSalaryCalculation.Rows.Count > 0 Then
                        For Each RowDgv In Me.DGVSalaryCalculation.Rows
                            If RowDgv.Cells(Me.ColChkPrint.Index).Value = True Then
                                My.Application.ChangeCulture("id-ID")
                                My.Application.ChangeUICulture("id-ID")
                                If Mindate.Date > CDate(RowDgv.Cells(Me.ColStartDate.Index).Value) Then
                                    Mindate = CDate(RowDgv.Cells(Me.ColStartDate.Index).Value)
                                End If
                                My.Application.ChangeCulture("en-US")
                                My.Application.ChangeUICulture("en-US")
                            End If
                        Next
                        For Each RowDgv In Me.DGVSalaryCalculation.Rows
                            If RowDgv.Cells(Me.ColChkPrint.Index).Value = True Then
                                My.Application.ChangeCulture("id-ID")
                                My.Application.ChangeUICulture("id-ID")
                                If maxdate < CDate(RowDgv.Cells(Me.ColEndDate.Index).Value) Then
                                    maxdate = CDate(RowDgv.Cells(Me.ColEndDate.Index).Value)
                                End If
                                My.Application.ChangeCulture("en-US")
                                My.Application.ChangeUICulture("en-US")
                            End If
                        Next
                        For Each RowDgv In Me.DGVSalaryCalculation.Rows
                            If RowDgv.Cells(Me.ColChkPrint.Index).Value = True Then
                                If FrmCetak.mListPrimarySalary.Contains(CDbl(RowDgv.Cells(Me.ColPrimaryTransaksi.Index).Value)) = False Then
                                    FrmCetak.mListPrimarySalary.Add(CDbl(RowDgv.Cells(Me.ColPrimaryTransaksi.Index).Value))
                                End If
                            End If
                        Next
                    End If
                    FrmCetak.mStartDate_Min = Mindate
                    FrmCetak.mEndDate_Max = maxdate
                    FrmCetak.mDgvsalaryQuick = Me.DGVSalaryCalculation
                    FrmCetak.mIndexColoumn_StartDate = Me.ColStartDate.Index
                    FrmCetak.mIndexColoumn_EndDate = Me.ColEndDate.Index
                    FrmCetak.mIndexColoumn_PrimarySalary = Me.ColPrimaryTransaksi.Index
                    FrmCetak.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCetak.Dispose()
        End Try
    End Sub

    Private Sub BtnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShow.Click
        Try
            While Me.BackgroundWorker1.IsBusy = True
                Threading.Thread.Sleep(10)
            End While
            While Me.BackgroundWorker2.IsBusy = True
                Threading.Thread.Sleep(10)
            End While
            Me.ObjQuick.IsiDGVQuick(Me.DtpTo.Value, Me.DtpPeriod.Value, Me.ObjCD, Me)

            ObjSettingUser.SetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Salary, ClsTransaksiSettingUser.EnumNamaParameter.HR_Salary_Quick_Type_Filter, Me.CboTypeDateFilter.SelectedIndex, Me.UserCode)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_Quick_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CboTypeDateFilter.Items.Add("Transaction Date")
            Me.CboTypeDateFilter.Items.Add("Period")
            Me.CboTypeDateFilter.SelectedIndex = 0

            Me.DGVSalaryCalculation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            'Me.LblAllPaid.Left -= 50
            Me.LblAllPrint.Left -= 50
            Me.LblAllDelete.Left -= 20
            Me.LblAllPrint.Text = "Print: Check All"
            'Me.LblAllPaid.Text = "Paid: Check All"
            Me.LblAllDelete.Text = "Delete: Check All"
            Me.LblAllEmail.Text = "Email: Check All"

            Dim UserSetting As String = ObjSettingUser.GetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Salary, ClsTransaksiSettingUser.EnumNamaParameter.HR_Salary_Quick_Type_Filter, Me.UserCode)
            If UserSetting = "1" Then
                Me.CboTypeDateFilter.SelectedIndex = 1
            Else
                Me.CboTypeDateFilter.SelectedIndex = 0
            End If

            Me.ColEmployeeCode.Visible = True
            Me.ColDocumentNo.Frozen = True
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVSalaryCalculation_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVSalaryCalculation.CellClick
        Dim FrmSalary As New FrmTrans_Gaji
        Dim RowDGV As DataGridViewRow
        Dim ObjSalary As New ClsTransaksiSalary
        Try
            If e.RowIndex <> -1 Then
                RowDGV = Me.DGVSalaryCalculation.Rows(e.RowIndex)
            Else
                Exit Sub
            End If
            If Me.DGVSalaryCalculation.Rows.Count > 0 Then
                If e.ColumnIndex = Me.ColDocumentNo.Index Then
                    FrmSalary.UserCode = Me.UserCode
                    FrmSalary.BackColorAll = Me.BackColorAll
                    FrmSalary.PrimaryQuick = RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value
                    FrmSalary.ShowSlr_FromQuick = True
                    FrmSalary.ShowDialog()
                End If
                If e.ColumnIndex = Me.ColEdit.Index Then
                    FrmSalary.ShowQuick_Edit = True
                    FrmSalary.NewQuick_Data = False
                    FrmSalary.PrimaryQuick = RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value
                    FrmSalary.UserCode = Me.UserCode
                    FrmSalary.BackColorAll = Me.BackColorAll
                    FrmSalary.MinimizeBox = False
                    If ObjSalary.GetJmlByNoTransaksi(ObjSalary.GetNoTransaksi(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value.ToString)) > 0 Then
                        FrmSalary.ShowDialog()
                        Me.BtnShow.PerformClick()
                    Else
                        MsgBox("Data has been deleted.", MsgBoxStyle.Information, "Information")
                        Me.DGVSalaryCalculation.Rows.RemoveAt(RowDGV.Index)
                        Me.ObjCD.SetNumber(Me.DGVSalaryCalculation)
                        Exit Sub
                    End If
                End If
                If e.ColumnIndex = Me.ColChkPrint.Index Then
                    If Me.DGVSalaryCalculation.Item(e.ColumnIndex, e.RowIndex).Value = False Then
                        Me.DGVSalaryCalculation.Item(e.ColumnIndex, e.RowIndex).Value = True
                    Else
                        Me.DGVSalaryCalculation.Item(e.ColumnIndex, e.RowIndex).Value = False
                    End If
                    Me.KontrolLinkLabel(EnumLink.Print, e.RowIndex)
                End If
                If e.ColumnIndex = Me.ColPay.Index Then
                    If Me.DGVSalaryCalculation.Item(e.ColumnIndex, e.RowIndex).Value = False Then
                        Me.DGVSalaryCalculation.Item(e.ColumnIndex, e.RowIndex).Value = True
                    Else
                        Me.DGVSalaryCalculation.Item(e.ColumnIndex, e.RowIndex).Value = False
                    End If
                    Me.KontrolLinkLabel(EnumLink.Pay, e.RowIndex)
                End If
                If e.ColumnIndex = Me.ColDelete.Index Then
                    If Me.DGVSalaryCalculation.Item(e.ColumnIndex, e.RowIndex).Value = False Then
                        Me.DGVSalaryCalculation.Item(e.ColumnIndex, e.RowIndex).Value = True
                    Else
                        Me.DGVSalaryCalculation.Item(e.ColumnIndex, e.RowIndex).Value = False
                    End If
                    Me.KontrolLinkLabel(EnumLink.Delete, e.RowIndex)
                End If
                If e.ColumnIndex = Me.ColChkEmail.Index Then
                    If Me.DGVSalaryCalculation.Item(e.ColumnIndex, e.RowIndex).Value = False Then
                        Me.DGVSalaryCalculation.Item(e.ColumnIndex, e.RowIndex).Value = True
                    Else
                        Me.DGVSalaryCalculation.Item(e.ColumnIndex, e.RowIndex).Value = False
                    End If
                    Me.KontrolLinkLabel(EnumLink.Email, e.RowIndex)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmSalary.Dispose()
        End Try
    End Sub

    Public Sub KontrolLinkLabel(ByVal vEnumLink As EnumLink, ByVal vIndex As Integer)
        Dim ij As Integer = 0
        Dim RowDGV As DataGridViewRow
        Dim Ketemu As Boolean = False
        Try
            If vEnumLink = EnumLink.Print Then
                While ij <= Me.DGVSalaryCalculation.Rows.Count - 1
                    RowDGV = Me.DGVSalaryCalculation.Rows(ij)
                    If vIndex <> ij Then
                        If RowDGV.Cells(Me.ColChkPrint.Index).Value = False Then
                            Ketemu = True
                            Exit While
                        End If
                    End If
                    ij += 1
                End While
                If Ketemu = False Then
                    Me.LblAllPrint.Text = "UnChecked All Print"
                Else
                    Me.LblAllPrint.Text = "Checked All Print"
                End If
            End If
            If vEnumLink = EnumLink.Pay Then
                While ij <= Me.DGVSalaryCalculation.Rows.Count - 1
                    RowDGV = Me.DGVSalaryCalculation.Rows(ij)
                    If vIndex <> ij Then
                        If RowDGV.Cells(Me.ColPay.Index).Value = False Then
                            Ketemu = True
                            Exit While
                        End If
                    End If
                    ij += 1
                End While
                If Ketemu = False Then
                    Me.LblAllPaid.Text = "Unchecked All Paid"
                Else
                    Me.LblAllPaid.Text = "Checked All Paid"
                End If
            End If
            If vEnumLink = EnumLink.Delete Then
                While ij <= Me.DGVSalaryCalculation.Rows.Count - 1
                    RowDGV = Me.DGVSalaryCalculation.Rows(ij)
                    If vIndex <> ij Then
                        If RowDGV.Cells(Me.ColDelete.Index).Value = False Then
                            Ketemu = True
                            Exit While
                        End If
                    End If
                    ij += 1
                End While
                If Ketemu = False Then
                    Me.LblAllDelete.Text = "UnChecked All Delete"
                Else
                    Me.LblAllDelete.Text = "Checked All Delete"
                End If
            End If
            If vEnumLink = EnumLink.Email Then
                While ij <= Me.DGVSalaryCalculation.Rows.Count - 1
                    RowDGV = Me.DGVSalaryCalculation.Rows(ij)
                    If vIndex <> ij Then
                        If RowDGV.Cells(Me.ColChkEmail.Index).Value = False Then
                            Ketemu = True
                            Exit While
                        End If
                    End If
                    ij += 1
                End While
                If Ketemu = False Then
                    Me.LblAllEmail.Text = "Email: Uncheck All"
                Else
                    Me.LblAllEmail.Text = "Email: Check All"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblAllPrint_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAllPrint.LinkClicked
        Dim RowDGV As DataGridViewRow
        Try
            If Me.LblAllPrint.Text = "Print: Check All" Then
                For Each RowDGV In Me.DGVSalaryCalculation.Rows
                    RowDGV.Cells(Me.ColChkPrint.Index).Value = True
                Next
                Me.LblAllPrint.Text = "Print: Uncheck All"
            Else
                For Each RowDGV In Me.DGVSalaryCalculation.Rows
                    RowDGV.Cells(Me.ColChkPrint.Index).Value = False
                Next
                Me.LblAllPrint.Text = "Print: Check All"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblAllPaid_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblAllPaid.Click
        Dim RowDGV As DataGridViewRow
        Try
            If Me.LblAllPaid.Text = "Paid: Check All" Then
                For Each RowDGV In Me.DGVSalaryCalculation.Rows
                    RowDGV.Cells(Me.ColPay.Index).Value = True
                Next
                Me.LblAllPaid.Text = "Paid: Uncheck All"
            Else
                For Each RowDGV In Me.DGVSalaryCalculation.Rows
                    RowDGV.Cells(Me.ColPay.Index).Value = False
                Next
                Me.LblAllPaid.Text = "Paid: Check All"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblAllDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAllDelete.LinkClicked
        Dim RowDGV As DataGridViewRow
        Try
            If Me.LblAllDelete.Text = "Delete: Check All" Then
                For Each RowDGV In Me.DGVSalaryCalculation.Rows
                    RowDGV.Cells(Me.ColDelete.Index).Value = True
                Next
                Me.LblAllDelete.Text = "Delete: Uncheck All"
            Else
                For Each RowDGV In Me.DGVSalaryCalculation.Rows
                    RowDGV.Cells(Me.ColDelete.Index).Value = False
                Next
                Me.LblAllDelete.Text = "Delete: Check All"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnNew.Click
        Dim FrmCari As New FrmTrans_Gaji_Quick2
        Try
            FrmCari.UserCode = Me.UserCode
            FrmCari.BackColorAll = Me.BackColorAll
            FrmCari.PeriodFrom = Me.DtpPeriod.Value.Date
            FrmCari.PeriodTo = Me.DtpTo.Value.Date
            FrmCari.RboAll.Checked = True
            FrmCari.dtpPeriod.Value = Now
            FrmCari.TxtPrimaryTaxExpAcc.Text = "-1"
            FrmCari.TxtExpAcc.Clear()
            FrmCari.TxtExpAccDesc.Clear()
            FrmCari.TxtPrimaryGainAcc.Text = "-1"
            FrmCari.TxtGainAccount.Clear()
            FrmCari.TxtPrimaryLoss.Text = "-1"
            FrmCari.TxtLossAccount.Clear()
            FrmCari.ChkDepartment.Checked = False
            FrmCari.ChkEmployee.Checked = False
            FrmCari.chkJobPosition.Checked = False
            FrmCari.chkSalaryFormula.Checked = False
            FrmCari.Lbl_Dept_Checked.Enabled = False
            FrmCari.Lbl_Dept_Unchecked.Enabled = False
            FrmCari.Lbl_JP_checked.Enabled = False
            FrmCari.Lbl_JP_Unchecked.Enabled = False
            FrmCari.Lbl_SF_Checked.Enabled = False
            FrmCari.Lbl_SF_Unchecked.Enabled = False
            FrmCari.ShowDialog()

            'add by ahdar
            If Me.CboTypeDateFilter.SelectedIndex = 0 Then
                Me.DtpPeriod.Value = FrmCari.DtpDate.Value
                Me.DtpTo.Value = FrmCari.DtpDate.Value
            Else
                Me.DtpPeriod.Value = FrmCari.dtpPeriod.Value
                Me.DtpTo.Value = FrmCari.DtpTo.Value
            End If

            Me.BtnShow.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCari.Dispose()
        End Try
    End Sub

    Public Function ApakahSelamaBulan1sampai11TransaksiBerjalanUdahAdaTransaksiBulan12(ByVal vTanggalTransSalary As Date, _
    ByVal vPrimaryEmployee As Double) As Boolean
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim TBL As New DseTransaksiSalary.tbltrans_salary2DataTable
        Dim Row As DseTransaksiSalary.tbltrans_salary2Row
        Dim StartDate As Date = #1/1/1950#
        Dim EndDate As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim Retval As String = ""
        Dim TipeTax As String = ""
        Dim ij As Integer = 1
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartDate = ObjCD.GetFirstDateofTheMonth(ObjCD.GetLastDateOfTheYear(vTanggalTransSalary))
            EndDate = ObjCD.GetLastDateOfTheYear(vTanggalTransSalary)
            DTA.FillBy_TglDariSampai_TanggalTransaksi(TBL, EndDate, StartDate, vPrimaryEmployee)
            If TBL.Rows.Count > 0 Then
                Row = TBL.Rows(0)
                'MsgBox("This data cannot edit or delete because this data has already used at period december " + vTanggalTransSalary.Year.ToString + ", " + vbCrLf + _
                '"document no :" + Row.No_Transaksi + ", date : " + Format(Row.Tanggal_Transaksi, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Private Sub ApakahBolehHapus(ByVal TBLMsgBox As DseTransaksiSalary.TBLMessageBoxDataTable, _
    ByVal ListPrimarySalary As List(Of Double), ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean)
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim RowMsgBox As DseTransaksiSalary.TBLMessageBoxRow
        Dim pCreatedBy As String = ""
        Dim ObjItem As New ClsTransaksiLoanPayment_Item
        Dim InfoData As String = ""
        Dim ObjSalary As New ClsTransaksiSalary
        Dim TanggalTransaksiSalary As Date = #1/1/1950#
        Dim Ketemu As Boolean = False
        Dim RowDGV As DataGridViewRow
        Dim ListString As New List(Of String)
        Dim PaymentStatus As ClsTransaksiSalary.EnumPaymentStatus = ClsTransaksiSalary.EnumPaymentStatus.paid
        Dim ObjSalaryPayment As New ClsTrans_Payment

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each RowDGV In Me.DGVSalaryCalculation.Rows
                Ketemu = False
                If RowDGV.Cells(Me.ColDelete.Index).Value = True Then
                    '1. tidak boleh hapus jika melanggar aturan strict edit
                    If vStrictDelete = True Then
                        If IsNothing(DTA.ScalarQuery_GetCreatedBy(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value)) = False Then
                            pCreatedBy = DTA.ScalarQuery_GetCreatedBy(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value)
                        Else
                            pCreatedBy = ""
                        End If
                        If pCreatedBy <> Me.UserCode Then
                            If vShowMsg = True Then
                                If ListString.Contains(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value.ToString + "-" + "User that can delete this data is: " + pCreatedBy.ToString) = False Then
                                    ListString.Add(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value.ToString + "-" + "User that can delete this data is: " + pCreatedBy.ToString)
                                    RowMsgBox = TBLMsgBox.NewRow
                                    RowMsgBox.Employee_Code = RowDGV.Cells(Me.ColEmployeeCode.Index).Value.ToString
                                    RowMsgBox.Employee_Name = RowDGV.Cells(Me.ColEmployeeName.Index).Value.ToString
                                    RowMsgBox.Reason = "User that can delete this data is: " + pCreatedBy.ToString
                                    RowMsgBox.Document_No = RowDGV.Cells(Me.ColDocumentNo.Index).Value.ToString
                                    TBLMsgBox.Rows.Add(RowMsgBox)

                                    RowDGV.Cells(Me.ColDelete.Index).Value = Boolean.FalseString
                                End If
                                'MsgBox("User that can delete this data is: " + pCreatedBy, _
                                'MsgBoxStyle.Information, "Information")
                            End If
                            Ketemu = True
                        End If
                    End If

                    '2. Berkaitan dengan transaksi salary Payment
                    PaymentStatus = DTA.ScalarQuery_PaymentStatus1(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value)
                    If PaymentStatus = ClsTransaksiSalary.EnumPaymentStatus.paid Then
                        If vShowMsg = True Then
                            If ListString.Contains(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value.ToString + "-" + "User that can delete this data is: " + pCreatedBy.ToString) = False Then
                                ListString.Add(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value.ToString + "-" + "User that can delete this data is: " + pCreatedBy.ToString)
                                RowMsgBox = TBLMsgBox.NewRow
                                RowMsgBox.Employee_Code = RowDGV.Cells(Me.ColEmployeeCode.Index).Value.ToString
                                RowMsgBox.Employee_Name = RowDGV.Cells(Me.ColEmployeeName.Index).Value.ToString
                                RowMsgBox.Reason = "Salary payment has been made on " + _
                                ObjSalaryPayment.GetDocumentNo_ByPrimarySalary(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value)
                                RowMsgBox.Document_No = RowDGV.Cells(Me.ColDocumentNo.Index).Value.ToString
                                TBLMsgBox.Rows.Add(RowMsgBox)

                                RowDGV.Cells(Me.ColDelete.Index).Value = Boolean.FalseString
                            End If
                        End If
                    End If

                    '3. Batasan bulan 1 sampai 11, jika sudah ada transaksi di bulan 12, maka tidak boleh edit, dan hapus salary
                    If Ketemu = False Then
                        TanggalTransaksiSalary = ObjSalary.GetTanggalTransaksi(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value)
                        If TanggalTransaksiSalary.Month <> 12 Then
                            If Me.ApakahSelamaBulan1sampai11TransaksiBerjalanUdahAdaTransaksiBulan12(TanggalTransaksiSalary, _
                            ObjSalary.GetPrimaryEmployee(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value)) = False Then
                                If vShowMsg = True Then
                                    If ListString.Contains(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value.ToString + "-" + "This data cannot delete because this data has already used at period december " + TanggalTransaksiSalary.Year.ToString) = False Then
                                        ListString.Add(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value.ToString + "-" + "This data cannot delete because this data has already used at period december " + TanggalTransaksiSalary.Year.ToString)
                                        RowMsgBox = TBLMsgBox.NewRow
                                        RowMsgBox.Employee_Code = RowDGV.Cells(Me.ColEmployeeCode.Index).Value.ToString
                                        RowMsgBox.Employee_Name = RowDGV.Cells(Me.ColEmployeeName.Index).Value.ToString
                                        RowMsgBox.Reason = "This data cannot delete because this data has already used at period december " + TanggalTransaksiSalary.Year.ToString
                                        RowMsgBox.Document_No = RowDGV.Cells(Me.ColDocumentNo.Index).Value.ToString

                                        TBLMsgBox.Rows.Add(RowMsgBox)

                                        RowDGV.Cells(Me.ColDelete.Index).Value = Boolean.FalseString
                                    End If
                                End If
                                Ketemu = True
                            End If
                        End If
                    End If

                    '4. Input kedalam list calculate salary
                    If Ketemu = False Then
                        If ListPrimarySalary.Contains(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value) = False Then
                            ListPrimarySalary.Add(RowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value)
                        End If
                    End If
                    '-----------------------
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjSalary = Nothing
            ListString = Nothing
            ObjSalaryPayment = Nothing
        End Try
    End Sub

    Private Function CekApakahBolehHapus(Optional ByVal vShowMsg As Boolean = True) As Boolean
        Dim RowDGV As DataGridViewRow
        Dim Ketemu As Boolean = False
        Try
            If Me.DGVSalaryCalculation.Rows.Count > 0 Then
                For Each RowDGV In Me.DGVSalaryCalculation.Rows
                    If RowDGV.Cells(Me.ColDelete.Index).Value = True Then
                        Ketemu = True
                        Exit For
                    End If
                Next
            End If
            If Ketemu = False Then
                If vShowMsg = True Then
                    MsgBox("Please choose data delete.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click
        Dim Hasil As MsgBoxResult
        Dim ListPrimarySalary As New List(Of Double)
        Dim PrimarySalary As Double = -1
        Dim TBLMsgBox As New DseTransaksiSalary.TBLMessageBoxDataTable
        Dim FrmError As New FrmTrans_Gaji_MsgBox
        Dim ObjSalary As New ClsTransaksiSalary
        Dim FrmGaji As New FrmTrans_Gaji
        Try
            If Me.CekApakahBolehHapus() = True Then
                Me.ApakahBolehHapus(TBLMsgBox, ListPrimarySalary, True, True)
                If TBLMsgBox.Rows.Count > 0 Then
                    TBLMsgBox.DefaultView.Sort = "[Employee Code], [Employee Name]"
                    FrmError.dgvItem.DataSource = TBLMsgBox.DefaultView
                    FrmError.dgvItem.Columns(0).Visible = False
                    FrmError.LblProgress.Text = "Total Rows = " + TBLMsgBox.Rows.Count.ToString
                    FrmError.Text = "The following employee list can't be deleted"
                    FrmError.ShowDialog()
                End If
            End If
            If Me.CekApakahBolehHapus(False) = True Then
                If ListPrimarySalary.Count > 0 Then
                    Hasil = MsgBox("Delete data?", MsgBoxStyle.YesNo, "Confirm")
                    If Hasil = MsgBoxResult.Yes Then
                        FrmGaji.Show()
                        FrmGaji.Top = (Me.FrmSalary.Height * -1) - 1
                        FrmGaji.ControlBox = False
                        FrmGaji.Hide()
                        For Each PrimarySalary In ListPrimarySalary
                            If ObjSalary.GetJmlByNoTransaksi(ObjSalary.GetNoTransaksi(PrimarySalary.ToString)) > 0 Then
                                FrmGaji.ShowQuick_Edit = True
                                FrmGaji.NewQuick_Data = False
                                FrmGaji.DeleteQuick_Data = True
                                FrmGaji.PrimaryQuick = PrimarySalary
                                FrmGaji.UserCode = Me.UserCode
                                FrmGaji.BackColorAll = Me.BackColorAll
                                FrmGaji.FungsiLoad_Event()
                                'Delete Transaksi
                                FrmGaji.FungsiBtnHapus_Quick()
                                FrmGaji.DeleteQuick_Data = False
                            End If
                        Next

                        MsgBox("Data has deleted", MsgBoxStyle.Information, "Information")
                        Me.BtnShow.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimarySalary = Nothing
            TBLMsgBox.Dispose()
            FrmError.Dispose()
            ObjSalary = Nothing
            FrmGaji.Dispose()
        End Try
    End Sub

    Private Sub DGVSalaryCalculation_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DGVSalaryCalculation.Sorted
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            ObjCD.SetNumber(Me.DGVSalaryCalculation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Private Sub DtpPeriod_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpPeriod.Enter
        Try
            Me.Buffer = Me.DtpPeriod.Value.Date.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpPeriod_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpPeriod.Leave
        'Try
        '    If CDate(Me.Buffer).Date <> Me.DtpPeriod.Value.Date Then
        '        Me.DGVSalaryCalculation.Rows.Clear()
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
        Try
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(5)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If CDate(Me.Buffer).Date <> Me.DtpPeriod.Value.Date Then
                Me.DGVSalaryCalculation.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpTo_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpTo.Enter
        Try
            Me.Buffer = Me.DtpTo.Value.Date.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpTo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpTo.Leave
        'Try
        '    If CDate(Me.Buffer).Date <> Me.DtpTo.Value.Date Then
        '        Me.DGVSalaryCalculation.Rows.Clear()
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
        Try
            Me.BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(5)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If CDate(Me.Buffer).Date <> Me.DtpTo.Value.Date Then
                Me.DGVSalaryCalculation.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblAllEmail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAllEmail.LinkClicked
        Dim RowDGV As DataGridViewRow
        Try
            If Me.LblAllEmail.Text = "Email: Check All" Then
                For Each RowDGV In Me.DGVSalaryCalculation.Rows
                    RowDGV.Cells(Me.ColChkEmail.Index).Value = True
                Next
                Me.LblAllEmail.Text = "Email: Uncheck All"
            Else
                For Each RowDGV In Me.DGVSalaryCalculation.Rows
                    RowDGV.Cells(Me.ColChkEmail.Index).Value = False
                Next
                Me.LblAllEmail.Text = "Email: Check All"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEmail.Click
        Dim ij As New FrmTrans_Gaji_SendEmail
        Dim ListPrimary As New List(Of Double)
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DGVSalaryCalculation.Rows
                If rowDGV.Cells(Me.ColChkEmail.Index).Value = Boolean.TrueString Then
                    ListPrimary.Add(rowDGV.Cells(Me.ColPrimaryTransaksi.Index).Value)
                End If
            Next

            If ListPrimary.Count > 0 Then
                ij.UserCode = Me.UserCode
                ij.BackColorAll = Me.BackColorAll
                ij.ListPrimarySalary = ListPrimary
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub lllSetDataSorting_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lllSetDataSorting.LinkClicked
        Dim i As New FrmTransaksiQuick_Sorting

        Try
            i.UserCode = Me.UserCode
            i.BackColorAll = Me.BackColorAll
            i.TipeTransaksi = ClsTransaksiQuick_Sorting.EnumTipeTransaksi.QuickSalary
            i.Text = "Quick Salary Data Sorting"
            i.ShowDialog()
            If i.ShowUlang = True Then
                Me.BtnShow.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class