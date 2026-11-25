Option Explicit On
Option Strict Off
Imports Core_Module, System.Drawing, System.Windows.Forms

Public Class FrmLaporanLogUser
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjCd As New Core_Module.ClsCekData
    Dim ObjUser As New Core_Module.ClsUser
    Dim objSF As New Core_Module.ClsMasterSalaryFormula

    Dim DS As New DseLaporanLogUser
    Dim report As New CrpLaporanLogUser

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Starting period must be smaller than ending period", MsgBoxStyle.Information, "Information")
                Me.dtpFrom.Focus()
                Exit Sub
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.CetakDetail()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakDetail()
        Dim Decimal_Multi As Integer = 0

        Dim DTACetak As New DseLaporanLogUserTableAdapters.tbl_log_userTableAdapter
        Dim TBLCetak As DseLaporanLogUser.tbl_log_userDataTable = DS.tbl_log_user
        Dim RowCetak As DseLaporanLogUser.tbl_log_userRow
        Dim TBL As New DseLaporanLogUser.tbl_log_userDataTable
        Dim Row As DseLaporanLogUser.tbl_log_userRow

        Dim DV As DataRowView
       

        Dim paramStartDate As Date = Now
        Dim paramEndDate As Date = Now
        Dim ParamUser As String = ""
        Dim ParamSalaryFormulaCode As String = ""
        Dim filter As String = ""
        Try
            DTACetak.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'kosongkan tabel
            TBLCetak.Rows.Clear()

            paramStartDate = Format(Me.dtpFrom.Value, "yyyy-MM-dd 00:00")
            paramEndDate = Format(Me.dtpTo.Value, "yyyy-MM-dd 23:59")
            DTACetak.FillBy(TBL, paramStartDate, paramEndDate)

            filter = " [Primary] = [Primary] "

            'If IsDate(paramStartDate) Then
            '    filter &= " And [Activity Time] >= #" & Format(paramStartDate, "yyyy-MM-dd") & "#"
            'End If

            'If IsDate(paramEndDate) Then
            '    filter &= " And [Activity Time] <= #" & Format(paramEndDate, "yyyy-MM-dd") & "#"
            'End If

            If Me.txtUser.Text <> "" Then
                filter &= " And [User] = '" & Me.txtUser.Text & "'"
                ParamUser = Me.txtUser.Text
            End If


            If Me.txtPrimarySalaryFormula.Text <> -1 Then
                filter &= " And [Primary Salary Formula] = " & Me.txtPrimarySalaryFormula.Text
                ParamSalaryFormulaCode = Me.txtSalaryFormulaCode.Text
            End If


            If Me.CboStatusPO.SelectedIndex <> 0 Then
                filter &= " And [Activity Type] = " & Me.CboStatusPO.SelectedIndex
            End If

            TBL.DefaultView.RowFilter = Nothing
            TBL.DefaultView.RowFilter = filter
            TBL.DefaultView.Sort = "[Activity Time] Desc "
            Dim ListRemarks As New List(Of String)
            If TBL.DefaultView.Count > 0 Then
                For Each DV In TBL.DefaultView
                    Row = DV.Row
                    RowCetak = TBLCetak.Newtbl_log_userRow
                    RowCetak.Primary = Row.Primary
                    RowCetak.User = Row.User
                    RowCetak.Activity_Time = Row.Activity_Time
                    RowCetak.Activity_Type = Row.Activity_Type
                    RowCetak.Remarks = Row.Remarks
                    RowCetak.Primary_Salary_Formula = Row.Primary_Salary_Formula
                    RowCetak.code = Row.code
                    If Row.Activity_Type = Core_Module.ClsLogUser.EnumActivityType.Create Then
                        RowCetak.Type = "Create"
                    ElseIf Row.Activity_Type = Core_Module.ClsLogUser.EnumActivityType.Modify Then
                        RowCetak.Type = "Modify"
                    ElseIf Row.Activity_Type = Core_Module.ClsLogUser.EnumActivityType.Delete Then
                        RowCetak.Type = "Delete"
                    End If
                    TBLCetak.Rows.Add(RowCetak)
                    'TBLCetak.DefaultView.RowFilter = " [User] = '" & Row.User & _
                    '"' And [Primary Salary Formula] = " & Row.Primary_Salary_Formula 

                    'If TBLCetak.DefaultView.Count > 0 Then
                    '    For Each dv2 As Data.DataRowView In TBLCetak.DefaultView
                    '        RowCetak = dv2.Row
                    '        RowCetak.Remarks &= Row.Remarks
                    '    Next
                    'Else
                    '    RowCetak = TBLCetak.Newtbl_log_userRow
                    '    RowCetak.Primary = Row.Primary
                    '    RowCetak.User = Row.User
                    '    RowCetak.Activity_Time = Row.Activity_Time
                    '    RowCetak.Activity_Type = Row.Activity_Type
                    '    RowCetak.Remarks = Row.Remarks
                    '    RowCetak.Primary_Salary_Formula = Row.Primary_Salary_Formula
                    '    RowCetak.code = Row.code
                    '    If Row.Activity_Type = Core_Module.ClsLogUser.EnumActivityType.Create Then
                    '        RowCetak.Type = "Create"
                    '    ElseIf Row.Activity_Type = Core_Module.ClsLogUser.EnumActivityType.Modify Then
                    '        RowCetak.Type = "Modify"
                    '    ElseIf Row.Activity_Type = Core_Module.ClsLogUser.EnumActivityType.Delete Then
                    '        RowCetak.Type = "Delete"
                    '    End If
                    '    TBLCetak.Rows.Add(RowCetak)
                    'End If
                Next
            End If



            'generate report
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_ParameterStartDate.ParameterFieldName, paramStartDate)
            report.SetParameterValue(report.Parameter_ParameterEndDate.ParameterFieldName, paramEndDate)
            report.SetParameterValue(report.Parameter_ParameterUser.ParameterFieldName, ParamUser)
            report.SetParameterValue(report.Parameter_SalaryFormulaCode.ParameterFieldName, ParamSalaryFormulaCode)
            crv.ReportSource = report
            crv.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTACetak.Dispose()
            TBLCetak.Dispose()
            TBL.Dispose()
            DS.Dispose()
        End Try
    End Sub

    Private Sub txtUser_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUser.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.txtUser.Text <> "" Then
                    If Me.ObjUser.ApakahSudahAda(Me.txtUser.Text) = True Then
                        Me.txtUser.Text = Me.txtUser.Text
                        Me.btnFilter.Focus()
                    Else
                        Me.btnFindUser.PerformClick()
                    End If
                Else
                    Me.btnFindUser.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtUser_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUser.Leave
        Dim hj As New Core_Module.FrmCariMasterUser
        'hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.txtUser.Text <> "" Then
                If Me.ObjUser.ApakahSudahAda(Me.txtUser.Text) = True Then
                    Me.txtPrimaryUser.Text = CStr(Me.ObjUser.GetPrimary(Me.txtUser.Text))
                    Me.txtUser.Text = Me.ObjUser.GetKodeUser(CDbl(Me.txtPrimaryUser.Text))
                Else
                    If Me.txtUser.Text <> "" Then
                        hj.TxtCari.Text = Me.txtUser.Text
                    End If
                    hj.ShowDialog()

                    If hj.Kodeuser <> "" Then
                        Me.txtPrimaryUser.Text = Me.ObjUser.GetPrimary(hj.Kodeuser)
                        Me.txtUser.Text = hj.Kodeuser
                        Me.btnFilter.Focus()
                    Else
                        Me.txtPrimaryUser.Text = "-1"
                        Me.txtUser.Clear()
                        Me.txtUser.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryUser.Text = "-1"
                Me.txtUser.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFindUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindUser.Click
        Dim hj As New Core_Module.FrmCariMasterUser
        Try
            'hj.BackColorAll = Me.BAckColorAll
            If Me.txtUser.Text <> "" Then
                hj.TxtCari.Text = Me.txtUser.Text
            End If

            hj.ShowDialog()
            If hj.Kodeuser <> "" Then
                Me.txtUser.Text = hj.Kodeuser
                Me.txtPrimaryUser.Text = Me.ObjUser.GetPrimary(hj.Kodeuser)
                Me.btnFilter.Focus()
            Else
                Me.txtPrimaryUser.Text = -1
                Me.txtUser.Clear()
                Me.txtUser.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub


    Private Sub txtProductCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUser.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnFilter.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadonly.BackColor
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

    Private Sub txtUser_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUser.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnFilter.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadonly.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimaryUser.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseOrderReportByDateDetail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
           
            Me.CboStatusPO.SelectedIndex = 0

            Me.txtSalaryFormulaCode.Visible = False
            Me.txtPrimarySalaryFormula.Visible = False
            Me.btnFindSalaryFormula.Visible = False
            Me.Label4.Visible = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindSalaryFormula_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindSalaryFormula.Click
        Dim hj As New Core_Module.FrmMasterSalaryFormula_Cari
        Try
            'hj.BackColorAll = Me.BAckColorAll
            If Me.txtSalaryFormulaCode.Text <> "" Then
                hj.txtCari.Text = Me.txtSalaryFormulaCode.Text
            End If

            hj.ShowDialog()
            If hj.Primary > -1 Then
                Me.txtPrimarySalaryFormula.Text = hj.Primary
                Me.txtSalaryFormulaCode.Text = objSF.GetCodeByPrimary(hj.Primary)
                Me.btnFilter.Focus()
            Else
                Me.txtPrimarySalaryFormula.Text = -1
                Me.txtSalaryFormulaCode.Clear()
                Me.txtSalaryFormulaCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub txtSalaryFormulaCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalaryFormulaCode.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnFilter.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadonly.BackColor
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

    Private Sub txtSalaryFormulaCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSalaryFormulaCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.txtSalaryFormulaCode.Text <> "" Then
                    If Me.objSF.ApakahSudahAda(Me.txtSalaryFormulaCode.Text) = True Then
                        Me.txtSalaryFormulaCode.Text = Me.txtSalaryFormulaCode.Text
                        Me.btnFilter.Focus()
                    Else
                        Me.btnFindSalaryFormula.PerformClick()
                    End If
                Else
                    Me.btnFindSalaryFormula.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSalaryFormulaCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalaryFormulaCode.Leave
        Dim hj As New Core_Module.FrmMasterSalaryFormula_Cari
        'hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.txtSalaryFormulaCode.Text <> "" Then
                If Me.objSF.ApakahSudahAda(Me.txtSalaryFormulaCode.Text) = True Then
                    Me.txtPrimarySalaryFormula.Text = CStr(Me.objSF.GetPrimaryByCode(Me.txtSalaryFormulaCode.Text))
                    Me.txtSalaryFormulaCode.Text = Me.objSF.GetCodeByPrimary(CDbl(Me.txtPrimarySalaryFormula.Text))
                Else
                    If Me.txtSalaryFormulaCode.Text <> "" Then
                        hj.TxtCari.Text = Me.txtSalaryFormulaCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.Primary > -1 Then
                        Me.txtPrimarySalaryFormula.Text = hj.Primary
                        Me.txtSalaryFormulaCode.Text = Me.objSF.GetCodeByPrimary(hj.Primary)
                        Me.btnFilter.Focus()
                    Else
                        Me.txtPrimarySalaryFormula.Text = "-1"
                        Me.txtSalaryFormulaCode.Clear()
                        Me.txtSalaryFormulaCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryUser.Text = "-1"
                Me.txtUser.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub txtSalaryFormulaCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalaryFormulaCode.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnFilter.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadonly.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimaryUser.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class