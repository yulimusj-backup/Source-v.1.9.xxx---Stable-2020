Option Explicit On
Option Strict Off

Public Class FrmMasterCommission
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ObjCurr As New Core_Module.clsMasterCurrency
    Public vDefCurr As Decimal = ObjCurr.GetDecimal_DefaultCurrency
    Dim ObjType As New ClsMasterCommission_Type
    Dim ObjTypeItem As New ClsMasterCommission_Type_Item
    Dim ObjAllocate As New ClsMasterCommission_Allocate
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim Baru As Boolean
    Dim Buffer As String = ""
    Dim EraseRowIndex As Integer = 0
    Dim DecimalMin As Decimal = 0
    Dim DecimalMax As Decimal = 0
    Dim DecimalCommValue As Decimal = 0

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim KoleksiNamaData1 As New Collection
        Dim KoleksiTextBox1 As New Collection
        Try

            Me.txtCode.Text = Trim(Me.txtCode.Text)
            Me.txtRemarks.Text = Trim(Me.txtRemarks.Text)

            'cek data yang harus diisi
            KoleksiNamaData.Add("Code no")
            KoleksiTextBox.Add(Me.txtCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.txtDescription)

            If Core_Module.ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            Me.txtRemarks.Text = Me.txtRemarks.Text.ToString.Replace("'", "").Replace("%", "").Replace("*", "")

            If Me.Baru = True Then
                If Me.ObjType.ApakahSudahAda(txtCode.Text) = True Then
                    MsgBox("Code No duplication, please change Code No.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.ObjType.GetCode(Me.txtPrimary.Text)
                If ProductGroupLama <> txtCode.Text Then
                    If Me.ObjType.ApakahSudahAda(txtCode.Text) = True Then
                        MsgBox("Code No duplication, please change Code No.", _
                        MsgBoxStyle.Information, "Information")
                        Me.txtCode.Focus()
                        Return False
                    End If
                End If
                'data salah jika nilai timestamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjType.GetTimestamp(Me.txtPrimary.Text)
                If Me.ObjType.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjType.CariDetail(Me.ObjType.Primary, Me.dgvItem, Me, Me.ObjCD)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            Dim Salah_Str As String = ""
            If ApakahDataItemBenar(Salah_Str) = False Then
                If Salah_Str = "MIN" Then
                    MsgBox("Min value must be > 0, please enter your data.", _
                                            MsgBoxStyle.Information, "Information")
                ElseIf Salah_Str = "MAX" Then
                    MsgBox("Max value must be > 0, please enter your data.", _
                                            MsgBoxStyle.Information, "Information")
                ElseIf Salah_Str = "COM" Then
                    MsgBox("Commission value must be > 0, please enter your data.", _
                                            MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub AktifkanDetail()
        Try
            ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimary.BackColor, False)
            ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ObjCD.KendalikanControl(Me.TabPage1.Controls, False, False, True, False, False, False, False, False)
            ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.dgvItem.AllowUserToAddRows = True
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            Me.dgvItem.TabStop = True
            Me.MatikanFungsiCbo()
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ObjCD.KendalikanControl(Me.TabPage1.Controls, False, False, True, True, False, False, False, False)
            ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, txtReadOnly.BackColor, False)
            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.TabStop = False
            Me.CancelButton = Me.btnExit
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ObjCD.KosongkanControl(Me.gboDetail.Controls, True, False, True)
            Me.txtPrimary.Text = "-1"
            Me.dgvItem.Rows.Clear()
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRequestBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.GotFocus, txtDescription.GotFocus, txtRemarks.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRequestBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.LostFocus, txtDescription.LostFocus, txtRemarks.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimary.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCommission_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.btnFind.Select()
            Me.Baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.cboLimitType.SelectedIndex = -1
            Me.cboBaseCalculate.SelectedIndex = -1
            Me.cboProgresiveCalculation.SelectedIndex = -1
            Me.KosongkanField()
            Me.AturForm_CommissionAllocate()
            Me.Text = "Sales Person Commission"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.Baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.txtCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColMin.Index Or e.ColumnIndex = Me.ColMax.Index Or e.ColumnIndex = Me.ColCommision.Index Or e.ColumnIndex = Me.Col_ConditionMin.Index Or e.ColumnIndex = Me.Col_ConditionMax.Index Then
                If e.ColumnIndex = Me.Col_ConditionMin.Index Or e.ColumnIndex = Me.Col_ConditionMax.Index Then
                    '1. cek bila baris di bawahnya sudah ada isinya jika sudah tidak boleh di edit
                    If e.RowIndex < dgvItem.RowCount - 1 Then
                        If CDec(Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex + 1).Value) > 0 Or _
                        CDec(Me.dgvItem.Item(Me.ColMax.Index, e.RowIndex + 1).Value) > 0 Then
                            MsgBox("Please delete value in the next row", MsgBoxStyle.Information, "Information")
                            e.Cancel = True
                        End If
                    ElseIf e.RowIndex = dgvItem.RowCount - 1 Then
                        '2. cek bila baris yang atasnya yaitu min dan max value <= 0 maka cancel
                        If e.RowIndex > 1 Then
                            Dim vRow As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex - 1)
                            If CDec(Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex - 1).Value) <= 0 Or CDec(Me.dgvItem.Item(Me.ColMax.Index, e.RowIndex - 1).Value) = 0 Then
                                MsgBox("Please enter value on the previous row", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                vRow.Selected = True
                                e.Cancel = True
                            End If
                        ElseIf e.RowIndex = 1 Then
                            Dim vRow As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex - 1)
                            If CDec(Me.dgvItem.Item(Me.ColMax.Index, e.RowIndex - 1).Value) <= 0 Then
                                MsgBox("Please enter value on the previous row", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                vRow.Selected = True
                                e.Cancel = True
                            End If
                        End If
                    End If
                End If

                If e.ColumnIndex = Me.ColMin.Index Or e.ColumnIndex = Me.ColMax.Index Then
                    If e.RowIndex = 1 And e.ColumnIndex = Me.ColMax.Index Then
                        If CDec(Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value) = 0 Then
                            MsgBox("Please enter value for data Min.", MsgBoxStyle.Information, "Information")
                            e.Cancel = True
                        End If
                    ElseIf e.RowIndex = 1 And e.ColumnIndex = Me.ColMin.Index Then
                        If CDec(Me.dgvItem.Item(Me.ColMax.Index, 0).Value) = 0 Then
                            MsgBox("Please enter value on the previous row.", MsgBoxStyle.Information, "Information")
                            e.Cancel = True
                        End If
                    End If

                    If e.RowIndex > 1 Then
                        If CDec(Me.dgvItem.Item(Me.ColMax.Index, e.RowIndex - 1).Value) <= 0 Or CDec(Me.dgvItem.Item(Me.ColMax.Index, e.RowIndex - 1).Value) = 0 Then
                            Dim vRow As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex - 1)
                            MsgBox("Please enter value on the previous row", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.ClearSelection()
                            vRow.Selected = True
                            e.Cancel = True
                        End If
                    End If
                End If

                If e.ColumnIndex = Me.ColCommision.Index Then
                    If CDec(Me.dgvItem.Item(Me.ColMax.Index, e.RowIndex).Value) > 0 Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim vRow As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Dim vRowIdx As Integer = -1
        Try
            If Me.cboLimitType.SelectedIndex = 0 Then
                DecimalMin = 2
                DecimalMax = 2
            ElseIf Me.cboLimitType.SelectedIndex = 1 Then
                DecimalMin = vDefCurr
                DecimalMax = vDefCurr
            End If

            If Me.cboBaseCalculate.SelectedIndex = 0 Then
                DecimalCommValue = 0
            ElseIf Me.cboBaseCalculate.SelectedIndex = 1 Then
                DecimalCommValue = 2
            End If

            With vRow
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If e.ColumnIndex = Me.Col_ConditionMin.Index Then
                        .Cells(Me.ColMin.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DecimalMin)
                        Me.MatikanFungsiCbo()
                    ElseIf e.ColumnIndex = Me.Col_ConditionMax.Index Then
                        .Cells(Me.ColMax.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DecimalMax)
                        Me.MatikanFungsiCbo()
                    ElseIf e.ColumnIndex = Me.ColMin.Index Then
                        If Me.ObjCD.CekBilanganPositif(.Cells(e.ColumnIndex).Value) = True Then
                            If Me.ApakahAngkaMinBenar(e.RowIndex) = True Then
                                .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(.Cells(e.ColumnIndex).Value, Me.DecimalMin)
                            Else
                                If e.RowIndex > 0 Then
                                    If CDec(.Cells(Me.ColMax.Index).Value) > 0 Then
                                        MsgBox("Please input min value must be > " & Me.dgvItem.Item(Me.ColMax.Index, e.RowIndex - 1).Value & _
                                                                                  " and < " & .Cells(Me.ColMax.Index).Value, MsgBoxStyle.Information, "Information")
                                    Else
                                        If Me.ObjTypeItem.GetInt_ByConditionMin(Me.dgvItem.Item(Me.Col_ConditionMax.Index, e.RowIndex - 1).Value) = ClsMasterCommission_Type_Item.Enum_ConditionMin.LebihBesar_SamaDengan Then
                                            MsgBox("Please input min value must be > " & Me.dgvItem.Item(Me.ColMax.Index, e.RowIndex - 1).Value, MsgBoxStyle.Information, "Information")
                                        Else
                                            MsgBox("Please input min value must be >= " & Me.dgvItem.Item(Me.ColMax.Index, e.RowIndex - 1).Value, MsgBoxStyle.Information, "Information")
                                        End If
                                    End If
                                Else
                                    If CDec(.Cells(Me.ColMax.Index).Value) > 0 Then
                                        MsgBox("Please input min value must be < " & Me.dgvItem.Item(Me.ColMax.Index, e.RowIndex).Value, MsgBoxStyle.Information, "Information")
                                    End If
                                End If
                            If String.IsNullOrEmpty(Me.Buffer) = True Or (Me.Buffer = 0 And .Cells(Me.ColMax.Index).Value = 0) Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                                Me.MatikanFungsiCbo()
                            Else
                                .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(Buffer, Me.DecimalMin)
                            End If
                            End If
                            Me.MatikanFungsiCbo()
                            Me.ObjCD.SetNumber(Me.dgvItem, False)
                            Exit Sub
                        Else
                            If e.RowIndex <> 0 Then
                                MsgBox("Please positive number for data", MsgBoxStyle.Information, "Information")
                                If String.IsNullOrEmpty(Me.Buffer) = True Or (Me.Buffer = 0 And .Cells(Me.ColMax.Index).Value = 0) Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                    Me.MatikanFungsiCbo()
                                Else
                                    .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(Buffer, Me.DecimalMin)
                                End If
                            Else
                                If IsNumeric(.Cells(e.ColumnIndex).Value) = True Then
                                    If CDec(.Cells(e.ColumnIndex).Value) >= 0 Then
                                        .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(.Cells(e.ColumnIndex).Value, Me.DecimalMin)
                                    Else
                                        MsgBox("Please positive number for data", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                            Me.MatikanFungsiCbo()
                                        Else
                                            .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(Buffer, Me.DecimalMin)
                                        End If
                                    End If

                                Else
                                    MsgBox("Please positive number for data", MsgBoxStyle.Information, "Information")
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                        Me.MatikanFungsiCbo()
                                    Else
                                        .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(Buffer, Me.DecimalMin)
                                    End If
                                End If
                            End If

                            Me.MatikanFungsiCbo()
                            Me.ObjCD.SetNumber(Me.dgvItem, False)
                            Exit Sub
                        End If

                    ElseIf e.ColumnIndex = Me.ColMax.Index Then
                        If Me.ObjCD.CekBilanganPositif(.Cells(e.ColumnIndex).Value) = True Then
                            If Me.ApakahAngkaMaxBenar(e.RowIndex) = True Then
                                .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(.Cells(e.ColumnIndex).Value, Me.DecimalMax)
                            Else
                                If e.RowIndex > 0 Then
                                    If e.RowIndex < Me.dgvItem.Rows.Count - 1 Then
                                        If CDec(Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex + 1).Value) > 0 Then
                                            If Me.ObjTypeItem.GetInt_ByConditionMin(Me.dgvItem.Item(Me.Col_ConditionMin.Index, e.RowIndex + 1).Value) = ClsMasterCommission_Type_Item.Enum_ConditionMin.LebihBesar_SamaDengan Then
                                                MsgBox("Please input max value must be > " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value & _
                                                                                          " and < " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex + 1).Value, MsgBoxStyle.Information, "Information")
                                            Else
                                                MsgBox("Please input max value must be > " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value & _
                                                                                          " and <= " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex + 1).Value, MsgBoxStyle.Information, "Information")

                                            End If
                                        Else
                                            If CDec(Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value) > 0 Then
                                                MsgBox("Please input max value must be > " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value, MsgBoxStyle.Information, "Information")
                                            Else
                                                MsgBox("Please input max value must be > " & Me.dgvItem.Item(Me.ColMax.Index, e.RowIndex - 1).Value, MsgBoxStyle.Information, "Information")
                                            End If
                                        End If
                                    Else
                                        MsgBox("Please input max value must be > " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value, MsgBoxStyle.Information, "Information")
                                    End If
                                Else
                                    If e.RowIndex = 0 Then
                                        If CDec(Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex + 1).Value) > 0 Then
                                            If CDec(Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value) > 0 Then
                                                If Me.ObjTypeItem.GetInt_ByConditionMin(Me.dgvItem.Item(Me.Col_ConditionMin.Index, e.RowIndex + 1).Value) = ClsMasterCommission_Type_Item.Enum_ConditionMin.LebihBesar_SamaDengan Then
                                                    MsgBox("Please input max value must be > " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value & _
                                                                                              " and < " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex + 1).Value, MsgBoxStyle.Information, "Information")
                                                Else
                                                    MsgBox("Please input max value must be > " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value & _
                                                                                              " and <= " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex + 1).Value, MsgBoxStyle.Information, "Information")

                                                End If

                                            Else

                                                MsgBox("Please input max value must be < " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex + 1).Value, MsgBoxStyle.Information, "Information")
                                            End If
                                        Else
                                            If CDec(Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value) > 0 Then
                                                MsgBox("Please input max value must be > " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value, MsgBoxStyle.Information, "Information")
                                            End If
                                        End If
                                    ElseIf e.RowIndex < Me.dgvItem.Rows.Count - 1 Then
                                        If CDec(Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex + 1).Value) > 0 Then
                                            MsgBox("Please input max value must be > " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value & _
                                                                                      " and < " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex + 1).Value, MsgBoxStyle.Information, "Information")
                                        Else
                                            MsgBox("Please input max value must be > " & Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value, MsgBoxStyle.Information, "Information")
                                        End If
                                    End If
                                End If

                                If String.IsNullOrEmpty(Me.Buffer) Or (Me.Buffer = 0 And .Cells(Me.ColMin.Index).Value = 0) Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                    Me.MatikanFungsiCbo()
                                Else
                                    .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(Buffer, Me.DecimalMax)
                                End If
                                Me.MatikanFungsiCbo()
                                Me.ObjCD.SetNumber(Me.dgvItem, False)
                                Exit Sub
                            End If
                        Else
                            MsgBox("Please positive number for data", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.Buffer) = True Or (Me.Buffer = 0 And .Cells(Me.ColMin.Index).Value = 0) Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                                Me.MatikanFungsiCbo()
                            Else
                                .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(Buffer, Me.DecimalMin)
                            End If
                            Me.MatikanFungsiCbo()
                            Me.ObjCD.SetNumber(Me.dgvItem, False)
                            Exit Sub
                        End If

                    ElseIf e.ColumnIndex = Me.ColCommision.Index Then
                        If Me.ObjCD.CekBilanganPositif(.Cells(e.ColumnIndex).Value) = True Then
                            .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(.Cells(e.ColumnIndex).Value, Me.DecimalCommValue)
                        Else
                            MsgBox("Please positive number for data", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                                Me.MatikanFungsiCbo()
                            Else
                                .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(Buffer, Me.DecimalCommValue)
                            End If
                            Me.ObjCD.SetNumber(Me.dgvItem, False)
                            Me.MatikanFungsiCbo()
                            Exit Sub
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColMin.Index Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(0, Me.DecimalMin)
                        Else
                            .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, Me.DecimalMin)
                        End If
                    ElseIf e.ColumnIndex = Me.ColMax.Index Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(0, Me.DecimalMax)
                        Else
                            .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, Me.DecimalMax)
                        End If
                    ElseIf e.ColumnIndex = Me.ColCommision.Index Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(0, Me.DecimalCommValue)
                        Else
                            .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, Me.DecimalCommValue)
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError
        'jangan di hapus buat nghandle error
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Try
            Me.dgvItem.Item(Me.ColMin.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping("0", Me.DecimalMin)
            Me.dgvItem.Item(Me.ColMax.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping("0", Me.DecimalMax)
            Me.dgvItem.Item(Me.ColCommision.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping("0", Me.DecimalCommValue)
            Me.dgvItem.Item(Me.Col_ConditionMax.Index, e.RowIndex).Value = Core_Module.ClsMasterCommission_Type_Item.Enum_ConditionMax.LebihKecil
            Me.dgvItem.Item(Me.Col_ConditionMin.Index, e.RowIndex).Value = Core_Module.ClsMasterCommission_Type_Item.Enum_ConditionMin.LebihBesar
            Me.ObjCD.SetNumber(Me.dgvItem, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvItem, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Public Function ApakahDataItemBenar(ByRef SalahBy As String) As Boolean
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.dgvItem.Rows.Count = 0 Then
                MsgBox("Please enter item at least one item", MsgBoxStyle.Information, "Information")
                Return False
            End If
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                With row
                    If row.Index > 0 Then
                        If CDec(.Cells(Me.ColMin.Index).Value) <= 0 Then
                            SalahBy = "MIN"
                            Return False
                        End If
                    End If

                    If CDec(.Cells(Me.ColMax.Index).Value) <= 0 Then
                        SalahBy = "MAX"
                        Return False
                    End If
                    If CDec(.Cells(Me.ColCommision.Index).Value) <= 0 Then
                        SalahBy = "COM"
                        Return False
                    End If
                End With
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvItem.AllowUserToAddRows = True
        End Try
    End Function

    Public Function ApakahAngkaMinBenar(ByVal vRowIdx As Integer) As Boolean
        Try
            '1. bandingkan dengan max pada baris yang disorot
            If Me.dgvItem.Item(Me.ColMax.Index, vRowIdx).Value > 0 Then
                If CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx).Value) <= CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx).Value) Then
                    Return False
                End If
            End If


            '2. bandingkan dengan max dan min pada baris diatasnya
            If vRowIdx > 0 Then
                'a. jika kondisinya ">=" maka 
                If Me.ObjTypeItem.GetInt_ByConditionMin(Me.dgvItem.Item(Me.Col_ConditionMin.Index, vRowIdx).Value) = ClsMasterCommission_Type_Item.Enum_ConditionMin.LebihBesar_SamaDengan Then

                    'a.1 jika baris diatas kondisinya "<" maka 
                    If Me.ObjTypeItem.GetInt_ByConditionMax(Me.dgvItem.Item(Me.Col_ConditionMax.Index, vRowIdx - 1).Value) = ClsMasterCommission_Type_Item.Enum_ConditionMax.LebihKecil Then
                        If CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx).Value) < CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx - 1).Value) Then
                            Return False
                        End If
                    End If

                    'a.2 jika baris diatas kondisinya "<=" maka 
                    If Me.ObjTypeItem.GetInt_ByConditionMax(Me.dgvItem.Item(Me.Col_ConditionMax.Index, vRowIdx - 1).Value) = ClsMasterCommission_Type_Item.Enum_ConditionMax.LebihKecil_SamaDengan Then
                        If CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx).Value) <= CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx - 1).Value) Then
                            Return False
                        End If
                    End If

                    'b. jika kondisinya ">" maka
                ElseIf Me.ObjTypeItem.GetInt_ByConditionMin(Me.dgvItem.Item(Me.Col_ConditionMin.Index, vRowIdx).Value) = ClsMasterCommission_Type_Item.Enum_ConditionMin.LebihBesar Then
                    If CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx).Value) < CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx - 1).Value) Then
                        Return False
                    End If
                    If CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx).Value) < CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx - 1).Value) Then
                        Return False
                    End If
                End If
            End If


            '3. bandingkan dengan max dan min pada baris dibawahnya
            If Me.dgvItem.Rows.Count - 1 > vRowIdx Then
                If Me.dgvItem.Item(Me.ColMin.Index, vRowIdx + 1).Value > 0 Then
                    If CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx + 1).Value) <= CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx).Value) Then
                        Return False
                    End If
                End If
                If Me.dgvItem.Item(Me.ColMax.Index, vRowIdx + 1).Value > 0 Then
                    If CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx + 1).Value) <= CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx).Value) Then
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvItem.AllowUserToAddRows = True
        End Try
    End Function

    Public Function ApakahAngkaMaxBenar(ByVal vRowIdx As Integer) As Boolean
        Try
            '1. bandingkan dengan minimum baris yang di sorot
            If CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx).Value) >= CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx).Value) Then
                Return False
            End If

            '2. bandingkan dengan baris di atasnya jika ada
            If vRowIdx > 0 Then
                If CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx - 1).Value) >= CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx).Value) Then
                    Return False
                End If
                If CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx - 1).Value) >= CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx).Value) Then
                    Return False
                End If
            End If

            '3. bandingkan dengan baris di bawahnya jika ada
            If Me.dgvItem.Rows.Count - 1 > vRowIdx Then
                '1. jika kondisinya "<" maka
                If Me.ObjTypeItem.GetInt_ByConditionMax(Me.dgvItem.Item(Me.Col_ConditionMax.Index, vRowIdx).Value) = ClsMasterCommission_Type_Item.Enum_ConditionMax.LebihKecil Then
                    If Me.dgvItem.Item(Me.ColMin.Index, vRowIdx + 1).Value > 0 Then
                        If CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx).Value) > CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx + 1).Value) Then
                            Return False
                        End If
                    End If

                    If Me.dgvItem.Item(Me.ColMax.Index, vRowIdx + 1).Value > 0 Then
                        If CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx).Value) > CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx + 1).Value) Then
                            Return False
                        End If
                    End If

                    '1. jika kondisinya "<=" maka
                ElseIf Me.ObjTypeItem.GetInt_ByConditionMax(Me.dgvItem.Item(Me.Col_ConditionMax.Index, vRowIdx).Value) = ClsMasterCommission_Type_Item.Enum_ConditionMax.LebihKecil_SamaDengan Then

                    '1.a jika baris dibawah kondisinya ">" maka

                    If Me.ObjTypeItem.GetInt_ByConditionMin(Me.dgvItem.Item(Me.Col_ConditionMin.Index, vRowIdx + 1).Value) = ClsMasterCommission_Type_Item.Enum_ConditionMin.LebihBesar Then
                        If CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx + 1).Value) > 0 Then
                            If CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx).Value) > CDec(Me.dgvItem.Item(Me.ColMin.Index, vRowIdx + 1).Value) Then
                                Return False
                            End If
                        End If
                    End If

                    '1.b jika baris dibawah kondisinya ">=" maka
                    If Me.ObjTypeItem.GetInt_ByConditionMin(Me.dgvItem.Item(Me.Col_ConditionMin.Index, vRowIdx + 1).Value) = ClsMasterCommission_Type_Item.Enum_ConditionMin.LebihBesar_SamaDengan Then
                        If CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx + 1).Value) > 0 Then
                            If CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx).Value) >= CDec(Me.dgvItem.Item(Me.ColMax.Index, vRowIdx + 1).Value) Then
                                Return False
                            End If
                        End If
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvItem.Rows.RemoveAt(EraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub AturForm_CommissionAllocate()
        Try
            Me.dgvFilterComm.Rows.Clear()
            Me.dgvItemProductGroup.Rows.Clear()
            Me.dgvAllocate.Rows.Clear()
            Me.dgvFilterComm.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvAllocate.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvAllocate.AllowUserToDeleteRows = False
            Me.dgvAllocate.AllowUserToAddRows = False
            Me.dgvFilterComm.AllowUserToAddRows = False
            Me.dgvItemProductGroup.AllowUserToAddRows = False
            Me.chkSalesPeriod.Checked = True
            Me.dtpFrom.Value = Now
            Me.dtpTo.Value = Now
            Me.chkCommission.Checked = False
            Me.dgvFilterComm.ReadOnly = False
            Me.chkProduct.Checked = False
            Me.dgvItemProductGroup.ReadOnly = False
            Me.dgvAllocate.ReadOnly = False
            Me.ColProductGroup_PrimaryAcuan.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function HapusSpasi(ByVal Data As String) As String
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

    Private Function CekDataRemarks(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim i As Int16
        Dim data As String
        Dim Txt As Windows.Forms.TextBox

        Try
            Do While i < KoleksiTextBox.Count
                Txt = TryCast(KoleksiTextBox.Item(i + 1), Windows.Forms.TextBox)
                If IsNothing(Txt) = False Then
                    Txt.Text = Txt.Text.Trim
                End If
                data = KoleksiTextBox.Item(i + 1).text
                data = HapusSpasi(data)


                If data.Contains("'") = True Or data.Contains("*") = True Or data.Contains("%") = True Then
                    MsgBox("" + KoleksiNamaData.Item(i + 1) _
                        + " data can not contain the following characters: ' , % , or * .", MsgBoxStyle.Information, "Information")
                    KoleksiTextBox.Item(i + 1).focus()
                    Return False
                End If
                i += 1
            Loop

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.ObjType
                .Primary = Me.txtPrimary.Text
                .Code = Me.txtCode.Text
                .Desc = Me.txtDescription.Text
                .Base_Calculation = Me.cboBaseCalculate.SelectedIndex
                .LimitType = Me.cboLimitType.SelectedIndex
                .ProgresiveCalculation = Me.cboProgresiveCalculation.SelectedIndex
                .Remarks = Me.txtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjType
                Me.txtPrimary.Text = .Primary
                Me.txtCode.Text = .Code
                Me.txtDescription.Text = .Desc
                Me.cboBaseCalculate.SelectedIndex = .Base_Calculation
                Me.cboLimitType.SelectedIndex = .LimitType
                Me.cboProgresiveCalculation.SelectedIndex = .ProgresiveCalculation
                Me.txtRemarks.Text = .Remarks

                If Me.cboLimitType.SelectedIndex = 0 Then
                    DecimalMin = 2
                    DecimalMax = 2
                ElseIf Me.cboLimitType.SelectedIndex = 1 Then
                    DecimalMin = vDefCurr
                    DecimalMax = vDefCurr
                End If

                If Me.cboBaseCalculate.SelectedIndex = 0 Then
                    DecimalCommValue = 0
                ElseIf Me.cboBaseCalculate.SelectedIndex = 1 Then
                    DecimalCommValue = 2
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjType.SimpanBaru(Me.dgvItem, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    If Me.ObjType.ApakahBolehEdit(Me.txtPrimary.Text, UserCode) = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ObjType.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtCode.Text <> "" Then
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjType.CariDetail(Me.txtPrimary.Text, Me.dgvItem, Me, Me.ObjCD)
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

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterCommission_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()

            If Not hj.Primary = -1 Then
                If Me.ObjType.CariDetail(hj.Primary, Me.dgvItem, Me, Me.ObjCD) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.txtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.txtPrimary.Text <> "-1" Then
                If Me.ObjType.ApakahBolehEdit(Me.txtPrimary.Text, UserCode) = True Then
                    Me.Baru = False
                    Me.ObjType.CariDetail(Me.txtPrimary.Text, Me.dgvItem, Me, Me.ObjCD)
                    Me.AktifkanDetail()
                    Me.SalinKeField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtCode.Text <> "" Then
                If Me.ObjType.ApakahBolehHapus(Me.txtPrimary.Text, UserCode, Me.txtCode.Text) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjType.Hapus(Me.txtPrimary.Text, Me.UserCode, Me, Me.txtCode.Text) = True Then
                            Me.KosongkanField()
                            Me.NonAktifkanDetail()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
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

    Private Sub FrmMasterLoad_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Baru = True Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjType.SimpanBaru(Me.dgvItem, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjType.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) = True Then
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

    Private Sub cboLimitType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLimitType.SelectedIndexChanged
        Try
            If Me.cboLimitType.SelectedIndex = 0 Then
                DecimalMin = 2
                DecimalMax = 2
            ElseIf Me.cboLimitType.SelectedIndex = 1 Then
                DecimalMin = vDefCurr
                DecimalMax = vDefCurr
            End If
            If Me.dgvItem.Rows.Count > 0 Then
                Me.dgvItem.Item(Me.ColMin.Index, 0).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(Me.ColMin.Index, 0).Value, Me.DecimalMin)
                Me.dgvItem.Item(Me.ColMax.Index, 0).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(Me.ColMax.Index, 0).Value, Me.DecimalMax)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboBaseCalculate_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBaseCalculate.SelectedIndexChanged
        Try
            If Me.cboBaseCalculate.SelectedIndex = 0 Then
                DecimalCommValue = 0
            ElseIf Me.cboBaseCalculate.SelectedIndex = 1 Then
                DecimalCommValue = 2
            End If
            If Me.dgvItem.Rows.Count > 0 Then
                Me.dgvItem.Item(Me.ColCommision.Index, 0).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(Me.ColCommision.Index, 0).Value, Me.DecimalCommValue)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub MatikanFungsiCbo()
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.dgvItem.Rows.Count >= 1 Then
                Me.cboBaseCalculate.Enabled = False
                Me.cboLimitType.Enabled = False
            Else
                Me.cboBaseCalculate.Enabled = True
                Me.cboLimitType.Enabled = True
            End If
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.MatikanFungsiCbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class