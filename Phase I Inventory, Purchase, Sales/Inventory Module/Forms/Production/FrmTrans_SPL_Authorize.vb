Option Explicit On
Option Strict Off

Public Class FrmTrans_SPL_Authorize
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim buffer As String = ""
    Dim EraseRowIndex As Integer = 0

    Dim objCD As New Core_Module.ClsCekData
    Dim objSPL As New ClsTrans_SPL
    Dim objSPLAuth As New ClsTrans_SPL_Authorize
    Dim objUser As New Core_Module.ClsUser
    Dim objAA As New ClsTrans_SPL_Authorize_Access

    Private Sub FrmTrans_SPL_Authorize_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        
    End Sub

    Private Sub FrmTrans_SPL_Authorize_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim frm As New FrmTrans_SPL
        Try

            Me.ReadOnlyDetail()
            objSPLAuth.IsiDGVAuthorize(TxtPrimarySPL.Text, Me.UserCode, Me.DgvOtorisasi, Me.DgvCancelOtorisasi, Me.DgvAllautorisasi, Me)
            If DgvOtorisasi.RowCount > 0 Then
                If Me.objSPL.CariDetail(Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, 0).Value, Me.DgvWorkshift, Me.DgvWorkunit, frm) = True Then
                    Me.SalinKeField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ReadOnlyDetail()
        Try

            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, True, False, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, True, False, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, False, True, True)
            Core_Module.ModCekData.KendalikanControl(Me.GroupBox2.Controls, True, True, True, True, True, False, True, True)

            Me.DgvWorkshift.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvWorkunit.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvOtorisasi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvCancelOtorisasi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvAllautorisasi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvUser.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.CancelButton = Me.btnExit
            Me.CboStatus.Enabled = False
            Me.CboStatus.SelectedIndex = 2
          
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub TboAuthorize_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TboAuthorize.Click
        Dim frm As New FrmTrans_SPL
        Try

            Me.CboStatus.Enabled = False
            Me.ChkSPLDate.Checked = False
            Me.DtpFrom.Value = Now
            Me.DtpTo.Value = Now
            Me.TsbAuthorize.Visible = True
            Me.tsbPrintAuthorize.Visible = False
            Me.TxtNoSPL.Clear()
            Me.DgvUser.Rows.Clear()
            Me.KosongkanField()
            If Me.TboAuthorize.SelectedIndex = 0 Then
                Me.CboStatus.SelectedIndex = 2
                Me.TsbAuthorize.Text = "Authorize"
                objSPLAuth.IsiDGVAuthorize(TxtPrimarySPL.Text, Me.UserCode, Me.DgvOtorisasi, Me.DgvCancelOtorisasi, Me.DgvAllautorisasi, Me)
                If DgvOtorisasi.RowCount > 0 Then
                    If Me.objSPL.CariDetail(Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, 0).Value, Me.DgvWorkshift, Me.DgvWorkunit, frm) = True Then
                        Me.SalinKeField()
                    End If
                End If
            ElseIf Me.TboAuthorize.SelectedIndex = 1 Then
                Me.CboStatus.SelectedIndex = 1
                Me.TsbAuthorize.Text = "Un-Authorize"
                objSPLAuth.IsiDGVAuthorize(TxtPrimarySPL.Text, Me.UserCode, Me.DgvOtorisasi, Me.DgvCancelOtorisasi, Me.DgvAllautorisasi, Me)
                If DgvCancelOtorisasi.RowCount > 0 Then
                    If Me.objSPL.CariDetail(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, 0).Value, Me.DgvWorkshift, Me.DgvWorkunit, frm) = True Then
                        Me.SalinKeField()
                    End If
                End If
            Else
                Me.CboStatus.SelectedIndex = 0
                Me.CboStatus.Enabled = True
                Me.ChkSPLDate.Checked = True
                Me.DtpFrom.Value = Now.AddMonths(-3)
                Me.TsbAuthorize.Visible = False
                Me.tsbPrintAuthorize.Visible = True
                objSPLAuth.IsiDGVAuthorize(TxtPrimarySPL.Text, Me.UserCode, Me.DgvOtorisasi, Me.DgvCancelOtorisasi, Me.DgvAllautorisasi, Me)
                If DgvAllautorisasi.RowCount > 0 Then
                    If Me.objSPL.CariDetail(Me.DgvAllautorisasi.Item(Me.col3_Primary.Index, 0).Value, Me.DgvWorkshift, Me.DgvWorkunit, frm) = True Then
                        Me.SalinKeField()
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvOtorisasi_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvOtorisasi.CellClick
        Dim frm As New FrmTrans_SPL
     
        Try
            If DgvOtorisasi.RowCount > 0 Then
                If Me.objSPL.CariDetail(Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, Me.DgvOtorisasi.CurrentRow.Index).Value, Me.DgvWorkshift, Me.DgvWorkunit, frm) Then
                    Me.SalinKeField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub DgvOtorisasi_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvOtorisasi.DataError
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objSPL

                Me.txtPrimary.Text = .Primary
                Me.txtDocumentNo.Text = .NoDocument
                Me.dtpDate.Value = .Tanggal
                Me.txtJobDesc.Text = .JobDesc
                Me.txtRemarks.Text = .Remarks
                Me.DtpFromDetail.Value = .FromDate
                Me.DtpToDetail.Value = .ToDate

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.txtDocumentNo.Clear()
            Me.txtJobDesc.Clear()
            Me.dtpDate.Value = Today
            Me.DtpFromDetail.Value = Today
            Me.DtpToDetail.Value = Today
            Me.txtRemarks.Text = ""

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvWorkshift.Rows.Clear()
            Me.dgvWorkUnit.Rows.Clear()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCancelOtorisasi_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCancelOtorisasi.CellClick
        Dim frm As New FrmTrans_SPL
        Try
            If DgvCancelOtorisasi.RowCount > 0 Then
                If Me.objSPL.CariDetail(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DgvCancelOtorisasi.CurrentRow.Index).Value, Me.DgvWorkshift, Me.DgvWorkunit, frm) = True Then
                    Me.SalinKeField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub DgvAllautorisasi_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvAllautorisasi.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.col3_Print.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvAllautorisasi_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvAllautorisasi.CellClick
        Dim frm As New FrmTrans_SPL
        Try
            If DgvAllautorisasi.RowCount > 0 Then
                If Me.objSPL.CariDetail(Me.DgvAllautorisasi.Item(Me.col3_Primary.Index, Me.DgvAllautorisasi.CurrentRow.Index).Value, Me.DgvWorkshift, Me.DgvWorkunit, frm) = True Then
                    Me.SalinKeField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub LlbCheckAll_Cancel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll_Cancel.LinkClicked
        Try
            Dim i As Integer
            If Me.TboAuthorize.SelectedIndex = 1 Then
                For i = 0 To Me.DgvCancelOtorisasi.RowCount - 1
                    Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Cancel.Index, i).Value = True
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub LlbCheckAll_Authorize_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll_Authorize.LinkClicked
        Try
            Dim i As Integer
            If Me.TboAuthorize.SelectedIndex = 0 Then
                For i = 0 To Me.DgvOtorisasi.RowCount - 1
                    Me.DgvOtorisasi.Item(Me.ColDgvAuth_Authorize.Index, i).Value = True
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub LlbUncheckAll_Athorize_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll_Athorize.LinkClicked
        Try
            Dim i As Integer
            If Me.TboAuthorize.SelectedIndex = 0 Then
                For i = 0 To Me.DgvOtorisasi.RowCount - 1
                    Me.DgvOtorisasi.Item(Me.ColDgvAuth_Authorize.Index, i).Value = False
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub


    Private Sub LlbUncheckAll_Cancel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll_Cancel.LinkClicked
        Try
            Dim i As Integer
            If Me.TboAuthorize.SelectedIndex = 1 Then
                For i = 0 To Me.DgvCancelOtorisasi.RowCount - 1
                    Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Cancel.Index, i).Value = False
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub TsbAuthorize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TsbAuthorize.Click
        Dim i As Integer
        Dim dp As New Core_Module.DialogPassword
        Dim hasil As MsgBoxResult
        Dim frm As New FrmTrans_SPL
        Dim kosong As Boolean = True

        Try
            If Me.TboAuthorize.SelectedIndex = 0 Then
                Me.DgvOtorisasi.EndEdit()
                For x As Integer = 0 To Me.DgvOtorisasi.RowCount - 1
                    If Me.DgvOtorisasi.Item(Me.ColDgvAuth_Authorize.Index, x).Value = True Then
                        kosong = False
                        Exit For
                    End If
                Next
            ElseIf Me.TboAuthorize.SelectedIndex = 1 Then
                Me.DgvCancelOtorisasi.EndEdit()
                For x As Integer = 0 To Me.DgvCancelOtorisasi.RowCount - 1
                    If Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Cancel.Index, x).Value = True Then
                        kosong = False
                        Exit For
                    End If
                Next
            End If

            If kosong = False Then
                hasil = MsgBox("" & Me.TsbAuthorize.Text & " ?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.DgvUser.Focus()
                    dp.txtUserName.Text = Me.UserCode
                    dp.ShowDialog()

                    If dp.DialogResult = Windows.Forms.DialogResult.OK Then
                        Me.KosongkanField()
                        If Me.TboAuthorize.SelectedIndex = 0 Then

                            For i = 0 To Me.DgvOtorisasi.RowCount - 1
                                If Me.DgvOtorisasi.Item(Me.ColDgvAuth_Authorize.Index, i).Value = True Then
                                    Me.objSPLAuth.updateSPL(1, Me.objUser.GetPrimary(Me.UserCode), Me.DgvOtorisasi.Item(Me.ColDgvAuth_SPLNo.Index, i).Value)
                                End If
                            Next
                            objSPLAuth.IsiDGVAuthorize(TxtPrimarySPL.Text, Me.UserCode, Me.DgvOtorisasi, Me.DgvCancelOtorisasi, Me.DgvAllautorisasi, Me)
                            If DgvOtorisasi.RowCount > 0 Then
                                If Me.objSPL.CariDetail(Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, 0).Value, Me.DgvWorkshift, Me.DgvWorkunit, frm) = True Then
                                    Me.SalinKeField()
                                End If
                            End If
                        ElseIf Me.TboAuthorize.SelectedIndex = 1 Then

                            For i = 0 To Me.DgvCancelOtorisasi.RowCount - 1
                                If Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Cancel.Index, i).Value = True Then
                                    Me.objSPLAuth.updateSPL(0, -1, Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_SPLNo.Index, i).Value)
                                End If
                            Next
                            objSPLAuth.IsiDGVAuthorize(TxtPrimarySPL.Text, Me.UserCode, Me.DgvCancelOtorisasi, Me.DgvCancelOtorisasi, Me.DgvAllautorisasi, Me)
                            If DgvCancelOtorisasi.RowCount > 0 Then
                                If Me.objSPL.CariDetail(Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, 0).Value, Me.DgvWorkshift, Me.DgvWorkunit, frm) = True Then
                                    Me.SalinKeField()
                                End If
                            End If
                        End If

                        MsgBox("" & Me.TsbAuthorize.Text & " Succeed", MsgBoxStyle.Information, "Information")
                    End If
                End If
            Else
                MsgBox("There must be at least one spl that will be " & Me.TsbAuthorize.Text, MsgBoxStyle.Information, "Information")
            End If

            

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFilter.Click
        Try
            objSPLAuth.IsiDGVAuthorize(TxtPrimarySPL.Text, Me.UserCode, Me.DgvOtorisasi, Me.DgvCancelOtorisasi, Me.DgvAllautorisasi, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNoSPL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNoSPL.KeyDown
        Dim objCode As New FrmTrans_SPL_Cari

        Try
            If e.KeyValue = 13 Then
                If Me.TxtNoSPL.Text <> "" Then
                    If objSPL.ApakahSudahAda(Me.TxtNoSPL.Text) = True Then
                        Me.TxtPrimarySPL.Text = objSPL.GetPrimary(Me.TxtNoSPL.Text)
                        Me.TxtNoSPL.Text = objSPL.GetNoDokumen(Me.TxtPrimarySPL.Text)
                        Me.BtnFindSPL.Focus()
                    Else
                        Me.BtnFindSPL.PerformClick()
                    End If
                Else
                    Me.BtnFindSPL.PerformClick()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
        End Try
    End Sub

    Private Sub TxtNoMO_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoSPL.Leave
        Dim i As New FrmTrans_SPL_Cari

        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtNoSPL.Text <> "" Then
                If objSPL.ApakahSudahAda(Me.TxtNoSPL.Text) = True Then
                    Me.TxtPrimarySPL.Text = objSPL.GetPrimary(Me.TxtNoSPL.Text)
                    Me.TxtNoSPL.Text = objSPL.GetNoDokumen(Me.TxtPrimarySPL.Text)
                Else
                    i.txtCari.Text = Me.TxtNoSPL.Text
                    i.ShowDialog()
                    If i.PrimaryPP > 0 Then
                        Me.TxtPrimarySPL.Text = i.PrimaryPP
                        Me.TxtNoSPL.Text = objSPL.GetNoDokumen(Me.TxtPrimarySPL.Text)
                    Else
                        Me.TxtNoSPL.Clear()
                        Me.TxtPrimarySPL.Text = "-1"
                        Me.BtnFindSPL.Focus()
                    End If
                End If
            Else
                Me.TxtNoSPL.Clear()
                Me.TxtPrimarySPL.Text = "-1"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()

        End Try
    End Sub

    Private Sub ChkMODate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkSPLDate.CheckedChanged

        If Me.ChkSPLDate.Checked = True Then
            Me.DtpFrom.Enabled = True
            Me.DtpTo.Enabled = True
        Else
            Me.DtpFrom.Enabled = False
            Me.DtpTo.Enabled = False
          
        End If
    End Sub

    Private Sub BtnFindMO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindSPL.Click
        Dim i As New FrmTrans_SPL_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtNoSPL.Text <> "" Then
                i.txtCari.Text = Me.TxtNoSPL.Text
            End If
            i.ShowDialog()
            If i.PrimaryPP > 0 Then
                Me.TxtPrimarySPL.Text = i.PrimaryPP
                Me.TxtNoSPL.Text = objSPL.GetNoDokumen(Me.TxtPrimarySPL.Text)
                Me.ChkSPLDate.Focus()
            Else
                Me.TxtNoSPL.Clear()
                Me.TxtPrimarySPL.Text = "-1"
                Me.BtnFindSPL.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    
    Private Sub DgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvUser.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCreatedBy_Name.Index Then
                e.Cancel = True
            Else
                If e.ColumnIndex <> Me.ColCreatedBy.Index Then
                    If Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = "-1" Then
                        e.Cancel = True
                    End If
                Else
                    buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvUser.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvUser.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryUser As Double = -1
        Dim RowIdx As Integer = e.RowIndex

        Try
            With row

                If IsNothing(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColCreatedBy.Index Then
                        Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value.ToString.Replace("'", Me.buffer).Replace("*", Me.buffer).Replace("#", Me.buffer).Replace("@", Me.buffer)
                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.objUser.ApakahSudahAda(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryUser = Me.objAA.GetPrimary(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemUser(DgvUser.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)
                            If index < 0 Then
                                Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = PrimaryUser
                                Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.objUser.GetUserCode(PrimaryUser)
                                Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(PrimaryUser)

                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else

                                Dim ij As New Core_Module.FrmCariMasterUser
                                ij.TxtCari.Text = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.backColorAll = Me.BackColorAll
                                ij.ShowDialog()
                                PrimaryUser = Me.objAA.GetPrimary(ij.Kodeuser)

                                If PrimaryUser <> -1 Then
                                Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = PrimaryUser
                                Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.objUser.GetUserCode(PrimaryUser)
                                Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(PrimaryUser)

                                    index = Me.GetIndexItemUser(DgvUser.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)

                                    If index > -1 Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = False Then
                                            Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.objUser.GetPrimary(Me.buffer)
                                        Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value)
                                        'Me.dgvUser.Item(Me.ColIdetifierHeader.Index, e.RowIndex).Value = Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value
                                            Exit Sub
                                        Else
                                            Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                        End If
                                    Else
                                    Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.objUser.GetUserCode(PrimaryUser)
                                    Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(PrimaryUser)
                                        
                                    End If
                                ElseIf PrimaryUser = -1 Then
                                    If Me.DgvUser.Rows.Count - 1 <> e.RowIndex Then
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.objCD.SetNumber(Me.DgvUser)
                                        Exit Sub
                                    End If
                                If Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.objCD.SetNumber(Me.DgvUser)
                                    Exit Sub
                                End If
                                End If
                                ij = Nothing
                        End If
                        Me.objCD.SetNumber(Me.DgvUser)
                    End If
                Else
                    Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    Else
                        Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.buffer
                    End If
                    Me.objCD.SetNumber(Me.DgvUser)
                    Exit Sub
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Function GetIndexItemUser(ByVal vUsercode As String, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Try
            'pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.dgvUser.Rows.Count - 1
                RowDGV = Me.dgvUser.Rows(ij)
                If RowDGV.Cells(Me.ColCreatedBy.Index).Value <> "" Then
                    If vRowIdxFocus <> ij Then
                        If RowDGV.Cells(Me.ColCreatedBy.Index).Value = vUsercode Then
                            Return Me.DgvUser.Rows.IndexOf(RowDGV)
                        End If
                    End If
                End If
            Next
            Return -1
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
            If DgvUser.RowCount > 1 Then
                Me.DgvUser.Rows.RemoveAt(EraseRowIndex)
                Me.objCD.SetNumber(Me.DgvUser)
            End If
           

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tsbPrintAuthorize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbPrintAuthorize.Click
        Dim frmCetakSPL As New FrmTrans_SPL_Cetak
        Dim kosong As Boolean = True
        Try
            Me.DgvAllautorisasi.EndEdit()
            For x As Integer = 0 To Me.DgvAllautorisasi.RowCount - 1
                If Me.DgvAllautorisasi.Item(Me.col3_Print.Index, x).Value = True Then
                    kosong = False
                    Exit For
                End If
            Next
            
            If kosong = False Then
                If Me.objSPL.CetakanSPLAuthorize(Me.GetListPrimary, frmCetakSPL.crvSPL) = True Then
                    If objSPL.print = True Then
                        frmCetakSPL.ShowDialog()
                    End If
                End If
            Else
                MsgBox("Please select the spl which will be in print", MsgBoxStyle.Information)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetListPrimary() As List(Of Double)
        Dim val As New List(Of Double)
        Try
            For Each dr As DataGridViewRow In Me.DgvAllautorisasi.Rows
                If dr.Cells(Me.col3_Primary.Index).Value <> -1 Then
                    If dr.Cells(Me.col3_Print.Index).Value = True Then
                        val.Add(dr.Cells(Me.col3_Primary.Index).Value)
                    End If
                End If
            Next
            Return val
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return val
        End Try
    End Function

    Private Sub LlbCheckAll_ALL_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll_ALL.LinkClicked
        Try
            Dim i As Integer
            If Me.TboAuthorize.SelectedIndex = 2 Then
                For i = 0 To Me.DgvAllautorisasi.RowCount - 1
                    Me.DgvAllautorisasi.Item(Me.col3_Print.Index, i).Value = True
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbUncheckAll_All_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll_All.LinkClicked
        Try
            Dim i As Integer
            If Me.TboAuthorize.SelectedIndex = 2 Then
                For i = 0 To Me.DgvAllautorisasi.RowCount - 1
                    Me.DgvAllautorisasi.Item(Me.col3_Print.Index, i).Value = False
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub tsbExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbExit.Click
        Try
            Me.Close()
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


End Class