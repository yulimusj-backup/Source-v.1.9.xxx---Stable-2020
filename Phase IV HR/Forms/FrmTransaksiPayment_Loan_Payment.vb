Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms
Public Class FrmTransaksiPayment_Loan_Payment
    Public mPrimarySalaryFromPayment As Double = -1
    Public mDef_Prec_Payment As Double = 0
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Public PrimaryCashIn As String = ""
    Public PrimaryGain As String = ""
    Public PrimaryLoss As String = ""
    Public TBLLoan As New DseTransaksiLoanPayment.DataTable2DataTable
    Public TBLCBOItem As New DseTransaksiLoanPayment.DataTable2DataTable

    Public ShowFromSalary As Boolean = False
    Public PrimaryLoanPaymentFromSalary As Double = -1
    Public ShowFromLinkLabel As Boolean = False
    Public PengaturanFocusDataBenar As Double = -1
    Public Baru As Boolean = False
    Public Obj As New ClsTransaksiLoanPayment
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Public Def_Prec As Double = Me.ObjCurr.GetDecimal_DefaultCurrency
    Public TotalLoanInstallment As Decimal = 0
    Public TBLLoanStatus As New DseTransaksiLoanPayment.TBLUpdateStatus_LoanDataTable

    Dim ObjItem As New ClsTransaksiLoanPayment_Item
    Dim ObjLoan As New ClsTransaksiLoan
    Dim ObjSalary As New ClsTransaksiSalary
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjAkun As New AccModule.ClsMasterAkun
    Dim PrimaryCurr_Default As Double = 1
    Dim ObjKrywn As New Core_Module.ClsMasterKaryawan

    Dim Buffer As String = ""
    Dim ApakahSimpanData As Boolean = False
    Dim TBLCash As New DseTransaksiLoanPayment.tblmaster_akunDataTable
    Dim ObjSetting As New ClsTransaksiSetting

    Public Enum EnumListFocus As Integer
        GaAda
        EmployeeCode
        GainAcc
        LossAcc
        Remarks
        DGVItem
    End Enum

    Private Sub FrmTransaksiTicketing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Frm As New FrmTransaksiLoan_Payment
        Try
            Me.NonAktifkanDetail()
            Me.KosongkanField()
            If Not Me.mPrimarySalaryFromPayment = -1 Then
                If Me.Obj.CariDetail(Me.Obj.GetPrimaryByPrimarySalary(Me.mPrimarySalaryFromPayment), Me.ObjCD, Me.dgvItem, Me.Def_Prec, Frm) = True Then
                    Me.SalinKeField()
                End If
            End If
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm.Dispose()
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Fungsi Standar
    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.GboItem.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.BackColorDefault()

            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToDeleteRows = False

            Me.KosongkanFieldItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.GboItem.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()

            'set dgv
            Me.dgvItem.ReadOnly = False
            Me.dgvItem.AllowUserToDeleteRows = True

            If Me.dgvItem.SelectedRows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            'Posisi Edit
            If Me.Baru = False Then
                Me.TxtEmployeeCode1.Enabled = False
                Me.BtnFindEmployee.Enabled = False
            End If

            If Me.ShowFromSalary = True Then
                Me.btnCancel.Enabled = False
                Me.dtpDate.Enabled = False
            End If

            If Me.ShowFromLinkLabel = True Then
                Me.dtpDate.Enabled = False
            End If

            'Untuk TBLLoanStatus, untuk settingan update status
            Me.TBLLoanStatus.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.GboItem.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.GboItem.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer4.Panel1.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer4.Panel2.Controls, True, True, False)
            Me.TxtPrimaryLoanPayment.Text = "-1"
            Me.TxtPrimarySalary.Text = "-1"
            Me.TxtPrimaryJurnal.Text = "-1"
            Me.TxtPrimaryEmployee.Text = "-1"
            Me.TxtPrimaryCashAcc.Text = "-1"
            Me.TxtPrimaryGainAcc.Text = "-1"
            Me.TxtPrimaryLoss.Text = "-1"
            Me.CboCashAccount.DataSource = Nothing
            Me.CboCashAccount.Items.Clear()
            If String.IsNullOrEmpty(Me.PrimaryCashIn) = True Then
                Me.PrimaryCashIn = -1
            End If
            If String.IsNullOrEmpty(Me.PrimaryGain) = True Then
                Me.PrimaryGain = -1
            End If
            If String.IsNullOrEmpty(Me.PrimaryLoss) = True Then
                Me.PrimaryLoss = -1
            End If
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.KosongkanFieldItem()
            Me.dgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanFieldItem()
        Try
            Me.ObjCD.KosongkanControl(Me.GboItem.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.TxtPrimaryLoan.Text = "-1"
            Me.TxtIdentifier.Clear()
            Me.CboLoanNumber.DataSource = Nothing
            Me.CboLoanNumber.Items.Clear()
            Me.TxtInstallmentAt.Text = "0"
            Me.TxtBase.Text = "0"
            Me.TxtInterest.Text = "0"
            Me.TxtTotalInstallment.Text = "0"
            Me.TxtPayment.Text = "0"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function HapusSpasi(ByVal Data As String) As String
        Try
            Do While Data.EndsWith(" ") And Data.Length > 0
                Data = Data.Remove(Data.Length - 1, 1)
            Loop
            Return Data
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .DocumentNo = Me.TxtDocumentNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .PrimarySalary = Me.TxtPrimarySalary.Text
                .PrimaryJurnal = Me.TxtPrimaryJurnal.Text
                .PrimaryKaryawan = Me.TxtPrimaryEmployee.Text
                .PrimaryCash = Me.TxtPrimaryCashAcc.Text
                .PrimaryGain = Me.TxtPrimaryGainAcc.Text
                .PrimaryLoss = Me.TxtPrimaryLoss.Text
                .Remarks = Me.TxtRemark.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Dim RowCash As DseTransaksiLoanPayment.tblmaster_akunRow
        Dim RV As Data.DataRowView
        Dim IndexCash As Double = -1
        Dim Frm As New FrmTransaksiLoan_Payment
        Try
            With Me.Obj
                Me.TxtPrimaryLoanPayment.Text = .Primary
                Me.TxtDocumentNo.Text = .DocumentNo
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtPrimarySalary.Text = .PrimarySalary
                Me.TxtPrimaryJurnal.Text = .PrimaryJurnal
                Me.TxtPrimaryEmployee.Text = .PrimaryKaryawan
                Me.TxtEmployeeCode1.Text = Me.ObjKrywn.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                Me.TxtEmployeeName.Text = Me.ObjKrywn.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                Me.Obj.IsiTBLCashCode(Me.CboCashAccount, Me.PrimaryCurr_Default, Me.TBLCash, Frm)
                Me.TxtPrimaryCashAcc.Text = .PrimaryCash
                For Each RV In Me.TBLCash.DefaultView
                    RowCash = RV.Row
                    IndexCash += 1
                    If RowCash.Primary = Me.TxtPrimaryCashAcc.Text Then
                        Me.CboCashAccount.SelectedIndex = IndexCash
                        Me.TxtDescCash.Text = RowCash.Description
                        Exit For
                    End If
                Next
                Me.TxtPrimaryGainAcc.Text = .PrimaryGain
                Me.TxtGainAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryGainAcc.Text)
                Me.TxtDescGain.Text = Me.ObjAkun.GetDescription(Me.TxtGainAccount.Text)
                Me.TxtPrimaryLoss.Text = .PrimaryLoss
                Me.TxtLossAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryLoss.Text)
                Me.TxtDescLoss.Text = Me.ObjAkun.GetDescription(Me.TxtLossAccount.Text)
                Me.TxtRemark.Text = .Remarks
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm.Dispose()
        End Try
    End Sub

    'Cash Account
    Private Sub CboCashOutAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboCashAccount.SelectedIndexChanged
        Dim row As DseTransaksiLoanPayment.tblmaster_akunRow
        Try
            If Me.CboCashAccount.SelectedIndex = -1 Then
                Me.TxtPrimaryCashAcc.Text = "-1"
            Else
                row = Me.TBLCash.DefaultView(Me.CboCashAccount.SelectedIndex).Row
                Me.TxtPrimaryCashAcc.Text = row.Primary
                Me.TxtDescCash.Text = row.Description
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Up Down
    Private Sub DGVItemSorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Try
            Me.dgvItem.TabStop = True
            If Me.TxtRemarksItem.Enabled = True Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItem.TabStop = False
                If Me.TxtRemarksItem.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e _
    As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
            'Posisi Employee Code 
            If Me.dgvItem.Rows.Count <= 0 Then
                If Me.Baru = True Then
                    Me.TxtPrimaryEmployee.Text = "-1"
                    Me.TxtEmployeeCode1.Clear()
                    Me.TxtEmployeeName.Clear()
                    Me.TBLLoan.Rows.Clear()
                    Me.TBLCBOItem.Rows.Clear()
                    Me.TxtPrimaryLoan.Text = "-1"
                    Me.TxtIdentifier.Text = ""
                    Me.CboLoanNumber.DataSource = Nothing
                    Me.CboLoanNumber.Items.Clear()
                    Me.TxtEmployeeCode1.Enabled = True
                    Me.BtnFindEmployee.Enabled = True
                    Me.KosongkanFieldItem()
                    Me.TxtEmployeeCode1.Focus()
                Else
                    Me.CboLoanNumber.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles dgvItem.SelectionChanged
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.dgvItem.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)

                If row.Index <> 0 Then
                    Dim indexbaru = row.Index - 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True)
                    End If
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False

                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> Me.dgvItem.Rows.Count - 1 Then
                    Dim indexbaru = row.Index + 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRowProcess(row, False)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvItem.Rows(row.Index).Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False) Then
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Payment Amount
    Private Sub TxtPayment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPayment.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPayment, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayment_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPayment.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPayment, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Back Color
    Private Sub TxtEmployeeCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode1.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
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

    Private Sub TxtEmployeeCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode1.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtEmployeeCode1.Enabled = True Then
                    If Me.TxtEmployeeCode1.ReadOnly = False Then
                        Me.TxtEmployeeCode1.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        Me.TxtEmployeeCode1.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub GotFocusTxtGainAccount()
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtGainAccount.Enabled = True Then
                    If Me.TxtGainAccount.ReadOnly = False Then
                        TxtGainAccount.BackColor = Me.BackColorAll
                    Else
                        TxtGainAccount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGainAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGainAccount.GotFocus
        Try
            Me.GotFocusTxtGainAccount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub GotLostFocusTxtGainAccount()
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtGainAccount.Enabled = True Then
                    If Me.TxtGainAccount.ReadOnly = False Then
                        TxtGainAccount.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtGainAccount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGainAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGainAccount.LostFocus
        Try
            Me.GotLostFocusTxtGainAccount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLossAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLossAccount.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtLossAccount.Enabled = True Then
                    If Me.TxtLossAccount.ReadOnly = False Then
                        TxtLossAccount.BackColor = Me.BackColorAll
                    Else
                        TxtLossAccount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLossAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLossAccount.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtLossAccount.Enabled = True Then
                    If Me.TxtLossAccount.ReadOnly = False Then
                        TxtLossAccount.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtLossAccount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayment_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayment.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtPayment.Enabled = True Then
                    If Me.TxtPayment.ReadOnly = False Then
                        TxtPayment.BackColor = Me.BackColorAll
                    Else
                        TxtPayment.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayment_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayment.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtPayment.Enabled = True Then
                    If Me.TxtPayment.ReadOnly = False Then
                        TxtPayment.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtPayment.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarksItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarksItem.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtRemarksItem.Enabled = True Then
                    If Me.TxtRemarksItem.ReadOnly = False Then
                        TxtRemarksItem.BackColor = Me.BackColorAll
                    Else
                        TxtRemarksItem.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarksItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarksItem.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtRemarksItem.Enabled = True Then
                    If Me.TxtRemarksItem.ReadOnly = False Then
                        TxtRemarksItem.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtRemarksItem.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtRemark.Enabled = True Then
                    If Me.TxtRemark.ReadOnly = False Then
                        TxtRemark.BackColor = Me.BackColorAll
                    Else
                        TxtRemark.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtRemark.Enabled = True Then
                    If Me.TxtRemark.ReadOnly = False Then
                        TxtRemark.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtRemark.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub TxtPayment_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayment.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtPayment, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPayment_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPayment.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtPayment, sender, e, Me.Def_Prec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Data Item Payment 
    Private Function ApakahDataItemBenar() As Boolean
        Try
            If Me.TxtPrimaryLoan.Text = "-1" Then
                MsgBox("Please choose loan number.", MsgBoxStyle.Information, "Information")
                Me.CboLoanNumber.Focus()
                Return False
            End If
            If Me.TxtPrimaryEmployee.Text = "-1" Then
                MsgBox("Please input employee code.", MsgBoxStyle.Information, "Information")
                If Me.TxtEmployeeCode1.Enabled = True Then
                    Me.TxtEmployeeCode1.Focus()
                End If
                Return False
            End If
            If Me.ObjCD.CekBilanganPositif(Me.TxtPayment.Text) = False Then
                MsgBox("Please input positive number.", MsgBoxStyle.Information, "Infomation")
                Me.TxtPayment.Focus()
                Return False
            End If
            If Me.ObjCD.CekDataKosong(Me.TxtRemarksItem.Text, False) = False Then
                MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndexItem(ByVal vIdentifierLoan As String) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColIdentifier.Index).Value = vIdentifierLoan Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub CboLoanNumber_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboLoanNumber.SelectedIndexChanged
        Dim RowCBOItem As DseTransaksiLoanPayment.DataTable2Row
        Try

            If Me.CboLoanNumber.Text <> "" Then
                If Me.TxtPrimaryLoan.Text = "-1" Then
                    Me.TxtIdentifier.Text = ""
                    Me.TxtPrimaryLoan.Text = "-1"
                    Me.TxtInstallmentAt.Text = "0"
                    Me.TxtBase.Text = "0"
                    Me.TxtInterest.Text = "0"
                    Me.TxtTotalInstallment.Text = "0"
                    Me.TxtPayment.Text = "0"
                    Me.CboLoanNumber.DataSource = Nothing
                Else
                    RowCBOItem = Me.TBLCBOItem(Me.CboLoanNumber.SelectedIndex)
                    Me.TxtIdentifier.Text = RowCBOItem.Identifier
                    Me.TxtPrimaryLoan.Text = RowCBOItem.Primary
                    Me.TxtInstallmentAt.Text = RowCBOItem.Installment_Period
                    Me.TxtBase.Text = Me.ObjCD.DigitGrouping(RowCBOItem.Base, Me.Def_Prec)
                    Me.TxtInterest.Text = Me.ObjCD.DigitGrouping(RowCBOItem.Interest, Me.Def_Prec)
                    Me.TxtTotalInstallment.Text = Me.ObjCD.DigitGrouping(CDec(RowCBOItem.Base + RowCBOItem.Interest), Me.Def_Prec)
                    Me.TxtPayment.Text = "0"
                    Me.TxtPrimaryRev.Text = RowCBOItem.Primary_Receivable_Account
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class