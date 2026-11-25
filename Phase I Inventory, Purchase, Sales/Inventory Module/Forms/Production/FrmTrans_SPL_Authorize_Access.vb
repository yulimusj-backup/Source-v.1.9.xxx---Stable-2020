Option Explicit On
Option Strict Off

Public Class FrmTrans_SPL_Authorize_Access
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim objCD As New Core_Module.ClsCekData
    Dim objUser As New Core_Module.ClsUser
    Dim objAA As New ClsTrans_SPL_Authorize_Access
    Dim objAAItem As New ClsTrans_SPL_Authorize_AccessItem
    Dim objSetting As New Core_Module.ClsTransaksiSetting


    Dim buffer As String = ""
    Dim EraseRowIndex As Integer = 0
    Dim dgvToRemove As Integer = 0
    Public change As Integer = 0

    Public tblItemUp As New DseTrans_SPL_Autorize_Access.tbltrans_spl_authaccess1DataTable
    Public rowItemUp As DseTrans_SPL_Autorize_Access.tbltrans_spl_authaccess1Row
    Public tblItemDown As New DseTrans_SPL_Autorize_Access.tbltrans_spl_authaccess_itemDataTable
    Public rowItemDown As DseTrans_SPL_Autorize_Access.tbltrans_spl_authaccess_itemRow

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUser.LinkClicked
        Me.ContextMenuStrip1.Show(Me.LlbUser, 0, 20)
    End Sub

    Private Sub FrmTrans_SPL_Authorize_Access_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Dim nilai As Integer = 0
            If Me.buffer <> "" Or change <> (Me.dgvManager.Rows.Count - 1) + (Me.dgvUser.Rows.Count - 1) Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    change = (Me.dgvManager.Rows.Count - 1) + (Me.dgvUser.Rows.Count - 1)
                    buffer = ""
                    If Me.ApakahDataBenar() = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objAA.Simpan(Me) Then
                                If Me.ChkActive.Checked = True Then
                                    nilai = 1
                                End If
                                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Overtime_Order, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Activate_Access_Retriction, nilai)
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.Close()
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

    Private Sub FrmTrans_SPL_Authorize_Access_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Leave

    End Sub

    Private Sub FrmTrans_SPL_Authorize_Access_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.dgvManager.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvUser.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.CancelButton = Me.btnExit
            If Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Overtime_Order, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Activate_Access_Retriction) = "1" Then
                Me.ChkActive.Checked = True
            Else
                Me.ChkActive.Checked = False
            End If
            objAA.IsiDGVHeader(Me.dgvManager, Me)
            change = (Me.dgvManager.Rows.Count - 1) + (Me.dgvUser.Rows.Count - 1)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim nilai As Integer = 0
        Try
            Dim hasil As MsgBoxResult
            If Me.ApakahDataBenar() = True Then
                hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If Me.objAA.Simpan(Me) Then
                        If Me.ChkActive.Checked = True Then
                            nilai = 1
                        End If
                        Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Overtime_Order, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Activate_Access_Retriction, nilai)
                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Me.buffer = ""
                        change = (Me.dgvManager.Rows.Count - 1) + (Me.dgvUser.Rows.Count - 1)
                        Me.Close()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub dgvManager_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvManager.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colNo.Index Or e.ColumnIndex = Me.colUsername.Index Then
                e.Cancel = True
            Else
                If e.ColumnIndex <> Me.colUserCode.Index Then
                    If Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value = "" Then
                        e.Cancel = True
                    End If
                    If e.ColumnIndex = Me.colAccessType.Index Then
                        buffer = Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    buffer = Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub TambahBarisAtas(ByVal vPrimary As Double, _
    ByVal vUserCode As String, ByVal vAuthType As String, ByVal vIdentifier As String)
        Try
            Me.rowItemUp = Me.tblItemUp.Newtbltrans_spl_authaccess1Row
            With Me.rowItemUp
                .primary = vPrimary
                .usercode = vUserCode
                If vAuthType = "Full Access" Then
                    .auth_type = 0
                Else
                    .auth_type = 1
                End If
                .Identifier = vIdentifier
            End With
            Me.tblItemUp.Rows.Add(Me.rowItemUp)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Sub TambahBarisBawah(ByVal vPrimary As Double, _
    ByVal vPrimaryUser As Double, _
    ByVal vIdentifierHead As String)
        Try
            Me.rowItemDown = Me.tblItemDown.Newtbltrans_spl_authaccess_itemRow
            With Me.rowItemDown
                .primary = vPrimary
                .primary_user = vPrimaryUser
                .Identifier_Header = vIdentifierHead

            End With
            Me.tblItemDown.Rows.Add(Me.rowItemDown)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvManager_CellBorderStyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvManager.CellBorderStyleChanged

    End Sub

    Private Sub dgvManager_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvManager.CellClick
        Try
            If Me.dgvManager.Item(Me.colUserCode.Index, Me.dgvManager.CurrentRow.Index).Value <> "" And _
                    Me.dgvManager.Item(Me.colAccessType.Index, Me.dgvManager.CurrentRow.Index).Value = Me.colAccessType.Items.Item(1) Then
                Me.LlbUser.Enabled = True
                Me.dgvUser.Enabled = True
            Else
                Me.LlbUser.Enabled = False
                Me.dgvUser.Rows.Clear()
                Me.dgvUser.Enabled = False
            End If
            If e.RowIndex > -1 Then
                reloadUser(Me.dgvManager.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub reloadUser(ByVal vIdentifier As String)
        Try
            Dim dr As DataGridViewRow
            tblItemDown.DefaultView.RowFilter = "[Identifier Header] = '" & vIdentifier & "'"
            Me.dgvUser.Rows.Clear()
            Me.dgvUser.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvUser.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvUser.AllowUserToAddRows = False
            For i As Integer = 0 To tblItemDown.DefaultView.Count - 1
                rowItemDown = tblItemDown.DefaultView(i).Row
                dgvUser.Rows.Add()
                dr = dgvUser.Rows(dgvUser.Rows.Count - 1)
                dr.Cells(Me.col2No.Index).Value = Me.dgvUser.Rows.Count
                dr.Cells(Me.col2UserCode.Index).Value = Me.objUser.GetUserCode(rowItemDown.primary_user)
                dr.Cells(Me.col2Username.Index).Value = Me.objUser.GetNamaByPrimary(rowItemDown.primary_user)
                dr.Cells(Me.col2PrimaryUser.Index).Value = rowItemDown.primary_user
            Next
            Me.dgvUser.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub dgvManager_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvManager.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvManager.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryManager As Double = -1
        Dim RowIdx As Integer = e.RowIndex

        Try
            With row

                If IsNothing(Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.buffer = Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colUserCode.Index Then

                        Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value = Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value.ToString.Replace("'", "").Replace("*", "").Replace("#", "").Replace("%", "").ToUpper
                        'kemungkinan
                        If Me.objUser.ApakahSudahAda(Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryManager = Me.objAA.GetPrimary(Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemManager(dgvManager.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)
                            If index < 0 Then
                                'Me.dgvManager.Item(Me.colPrimary.Index, e.RowIndex).Value = PrimaryManager
                                Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value = Me.objUser.GetUserCode(PrimaryManager)
                                Me.dgvManager.Item(Me.colUsername.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(PrimaryManager)
                                tblItemUp.DefaultView.RowFilter = "[Identifier]='" & Me.dgvManager.Item(Me.ColIdentifier.Index, e.RowIndex).Value & "'"
                                If tblItemUp.DefaultView.Count > 0 Then
                                    rowItemUp = tblItemUp.DefaultView(0).Row
                                    rowItemUp.usercode = Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value
                                Else
                                    Me.dgvManager.Item(Me.ColIdentifier.Index, e.RowIndex).Value = Me.GetIdentifierBaru
                                    TambahBarisAtas(Me.dgvManager.Item(Me.colPrimary.Index, e.RowIndex).Value, _
                                    Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value, _
                                    Me.dgvManager.Item(Me.colAccessType.Index, e.RowIndex).Value, _
                                    Me.dgvManager.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                                End If
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.dgvToRemove = 0
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.objCD.SetNumber(Me.dgvManager)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmCariMasterUser
                            ij.TxtCari.Text = Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.backColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            PrimaryManager = Me.objAA.GetPrimary(ij.Kodeuser)

                            If PrimaryManager <> -1 Then
                                'Me.dgvManager.Item(Me.colPrimary.Index, e.RowIndex).Value = PrimaryManager
                                Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value = Me.objUser.GetUserCode(PrimaryManager)
                                Me.dgvManager.Item(Me.colUsername.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(PrimaryManager)
                                index = Me.GetIndexItemManager(dgvManager.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)
                                If index > -1 Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    If String.IsNullOrEmpty(Me.buffer) = False Then
                                        Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Me.dgvManager.Item(Me.colPrimary.Index, e.RowIndex).Value = Me.objUser.GetPrimary(Me.buffer)
                                        Me.dgvManager.Item(Me.colUsername.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(Me.dgvManager.Item(Me.colPrimary.Index, e.RowIndex).Value)
                                        Exit Sub
                                    Else
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.dgvToRemove = 0
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                Else
                                    'Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value = Me.objUser.GetUserCode(PrimaryManager)
                                    'Me.dgvManager.Item(Me.colUsername.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(PrimaryManager)
                                    tblItemUp.DefaultView.RowFilter = "[Identifier]='" & Me.dgvManager.Item(Me.ColIdentifier.Index, e.RowIndex).Value & "'"
                                    If tblItemUp.DefaultView.Count > 0 Then
                                        rowItemUp = tblItemUp.DefaultView(0).Row
                                        rowItemUp.usercode = Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value
                                    Else
                                        Me.dgvManager.Item(Me.ColIdentifier.Index, e.RowIndex).Value = Me.GetIdentifierBaru
                                        TambahBarisAtas(Me.dgvManager.Item(Me.colPrimary.Index, e.RowIndex).Value, _
                                        Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value, _
                                        Me.dgvManager.Item(Me.colAccessType.Index, e.RowIndex).Value, _
                                        Me.dgvManager.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                                    End If
                                End If
                            ElseIf PrimaryManager = -1 Then
                                If Me.dgvManager.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.dgvToRemove = 0
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.objCD.SetNumber(Me.dgvManager)
                                    Exit Sub
                                End If
                                If Me.dgvManager.Item(Me.col2UserCode.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.dgvToRemove = 0
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.objCD.SetNumber(Me.dgvManager)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.objCD.SetNumber(Me.dgvManager)

                    ElseIf e.ColumnIndex = Me.colAccessType.Index Then
                        If Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value = Me.colAccessType.Items.Item(1) Then
                            Me.dgvUser.Enabled = True
                            Me.LlbUser.Enabled = True
                        Else
                            Me.LlbUser.Enabled = False
                            Me.dgvUser.Rows.Clear()
                            Me.dgvUser.Enabled = False
                        End If

                        tblItemUp.DefaultView.RowFilter = "[Identifier]='" & Me.dgvManager.Item(Me.ColIdentifier.Index, e.RowIndex).Value & "'"
                        If tblItemUp.DefaultView.Count > 0 Then
                            rowItemUp = tblItemUp.DefaultView(0).Row
                            If Me.dgvManager.Item(Me.colAccessType.Index, e.RowIndex).Value = "Full Access" Then
                                rowItemUp.auth_type = 0
                                Dim identifier As String

                                identifier = Me.dgvManager.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                                Me.tblItemDown.DefaultView.RowFilter = "[Identifier Header] = '" & identifier & "'"
                                For x As Integer = 0 To Me.tblItemDown.DefaultView.Count - 1
                                    If Me.tblItemDown.DefaultView.Count > 0 Then
                                        Me.rowItemDown = Me.tblItemDown.DefaultView(i).Row
                                        Me.tblItemDown.Rows.Remove(Me.rowItemDown)
                                    End If
                                Next

                            Else
                                rowItemUp.auth_type = 1
                            End If
                        End If
                    End If
                Else
                    Me.dgvManager.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.dgvToRemove = 0
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    Else
                        Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value = Me.buffer
                    End If
                    Me.objCD.SetNumber(Me.dgvManager)
                    Exit Sub
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Function GetIndexItemManager(ByVal vUserCode As String, _
       ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Try
            'pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.dgvManager.Rows.Count - 1
                RowDGV = Me.dgvManager.Rows(ij)
                If RowDGV.Cells(Me.colUserCode.Index).Value <> "" Then
                    If vRowIdxFocus <> ij Then
                        If RowDGV.Cells(Me.colUserCode.Index).Value = vUserCode Then
                            Return Me.dgvManager.Rows.IndexOf(RowDGV)
                        End If
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Function GetIndexItemUser(ByVal vUsercode As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Try
            'pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.dgvUser.Rows.Count - 1
                RowDGV = Me.dgvUser.Rows(ij)
                If RowDGV.Cells(Me.col2UserCode.Index).Value <> "" Then
                    If vRowIdxFocus <> ij Then
                        If RowDGV.Cells(Me.col2UserCode.Index).Value = vUsercode Then
                            Return Me.dgvUser.Rows.IndexOf(RowDGV)
                        End If
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.dgvToRemove = 0 Then
                Me.dgvManager.Rows.RemoveAt(EraseRowIndex)
                Me.objCD.SetNumber(Me.dgvManager)
            Else
                Me.dgvUser.Rows.RemoveAt(EraseRowIndex)
                Me.objCD.SetNumber(Me.dgvUser)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvUser.CellBeginEdit
        Try
            If e.ColumnIndex = Me.col2No.Index Or e.ColumnIndex = Me.col2Username.Index Then
                e.Cancel = True
            Else
                If e.ColumnIndex <> Me.col2UserCode.Index Then
                    If Me.dgvUser.Item(Me.col2PrimaryUser.Index, e.RowIndex).Value = "-1" Then
                        e.Cancel = True
                    End If
                Else
                    buffer = Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvUser.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvUser.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryUser As Double = -1
        Dim RowIdx As Integer = e.RowIndex


        Try
            With row

                If IsNothing(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.buffer = Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If

                    If e.ColumnIndex = Me.colUserCode.Index Then
                        Me.dgvUser.Item(Me.col2UserCode.Index, e.RowIndex).Value = Me.dgvUser.Item(Me.col2UserCode.Index, e.RowIndex).Value.ToString.Replace("'", Me.buffer).Replace("*", Me.buffer).Replace("#", Me.buffer).Replace("@", Me.buffer)
                        Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.objUser.ApakahSudahAda(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryUser = Me.objAA.GetPrimary(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemUser(dgvUser.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)
                            If index < 0 Then
                                If Me.dgvUser.Item(Me.col2UserCode.Index, e.RowIndex).Value = _
                                Me.dgvManager.Item(Me.colUserCode.Index, Me.dgvManager.CurrentRow.Index).Value Then

                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvUser.Item(Me.colUserCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.dgvToRemove = 1
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.objCD.SetNumber(Me.dgvUser)
                                    Exit Sub
                                Else
                                    Me.dgvUser.Item(Me.col2PrimaryUser.Index, e.RowIndex).Value = PrimaryUser
                                    Me.dgvUser.Item(Me.col2UserCode.Index, e.RowIndex).Value = Me.objUser.GetUserCode(PrimaryUser)
                                    Me.dgvUser.Item(Me.col2Username.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(PrimaryUser)
                                    'Me.dgvUser.Item(Me.ColIdetifierHeader.Index, e.RowIndex).Value = Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value
                                    tblItemDown.DefaultView.RowFilter = "[primary user]='" & Me.objUser.GetPrimary(Me.buffer) & "'"
                                    If tblItemDown.DefaultView.Count > 0 Then
                                        rowItemDown = tblItemDown.DefaultView(0).Row
                                        rowItemDown.primary_user = PrimaryUser
                                    Else
                                        TambahBarisBawah(Me.dgvUser.Item(Me.col2PrimaryUser.Index, e.RowIndex).Value, _
                                        PrimaryUser, _
                                        Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value)
                                    End If

                                End If
                            Else
                                'If e.RowIndex <> index Then
                                MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                Me.dgvUser.Item(Me.colUserCode.Index, e.RowIndex).Value = ""
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.dgvToRemove = 1
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                                Me.objCD.SetNumber(Me.dgvUser)
                                Exit Sub
                                'End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmCariMasterUser
                            ij.TxtCari.Text = Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.backColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            PrimaryUser = Me.objAA.GetPrimary(ij.Kodeuser)

                            If PrimaryUser <> -1 Then
                                Me.dgvUser.Item(Me.col2PrimaryUser.Index, e.RowIndex).Value = PrimaryUser
                                Me.dgvUser.Item(Me.col2UserCode.Index, e.RowIndex).Value = Me.objUser.GetUserCode(PrimaryUser)
                                Me.dgvUser.Item(Me.col2Username.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(PrimaryUser)
                                'Me.dgvUser.Item(Me.ColIdetifierHeader.Index, e.RowIndex).Value = Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value
                                index = Me.GetIndexItemUser(dgvUser.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)
                                
                                If index > -1 Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    If String.IsNullOrEmpty(Me.buffer) = False Then
                                        Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Me.dgvUser.Item(Me.col2PrimaryUser.Index, e.RowIndex).Value = Me.objUser.GetPrimary(Me.buffer)
                                        Me.dgvUser.Item(Me.col2Username.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(Me.dgvUser.Item(Me.col2PrimaryUser.Index, e.RowIndex).Value)
                                        'Me.dgvUser.Item(Me.ColIdetifierHeader.Index, e.RowIndex).Value = Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value
                                        Exit Sub
                                    Else
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.dgvToRemove = 1
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                Else
                                    'Me.dgvManager.Item(Me.colUserCode.Index, e.RowIndex).Value = Me.objUser.GetUserCode(PrimaryManager)
                                    'Me.dgvManager.Item(Me.colUsername.Index, e.RowIndex).Value = Me.objUser.GetNamaByPrimary(PrimaryManager)
                                    tblItemDown.DefaultView.RowFilter = "[primary user]='" & Me.objUser.GetPrimary(Me.buffer) & "'"
                                    If tblItemDown.DefaultView.Count > 0 Then
                                        rowItemDown = tblItemDown.DefaultView(0).Row
                                        rowItemDown.primary_user = PrimaryUser
                                    Else
                                        TambahBarisBawah(Me.dgvUser.Item(Me.col2PrimaryUser.Index, e.RowIndex).Value, _
                                        PrimaryUser, _
                                        Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value)
                                    End If
                                End If
                            ElseIf PrimaryUser = -1 Then
                                If Me.dgvUser.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.dgvToRemove = 1
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.objCD.SetNumber(Me.dgvUser)
                                    Exit Sub
                                End If
                                If Me.dgvUser.Item(Me.col2UserCode.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.objCD.SetNumber(Me.dgvUser)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.objCD.SetNumber(Me.dgvUser)
                    End If
                Else
                    Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.dgvToRemove = 1
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    Else
                        Me.dgvUser.Item(Me.col2UserCode.Index, e.RowIndex).Value = Me.buffer
                    End If
                    Me.objCD.SetNumber(Me.dgvUser)
                    Exit Sub
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyToAllWokUnitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApplyToAllUserToolStripMenuItem.Click
        Dim DTA As New DseTrans_SPL_Autorize_AccessTableAdapters.tbluserTableAdapter
        Dim TBL As New DseTrans_SPL_Autorize_Access.tbluserDataTable
        Dim rowItem As DseTrans_SPL_Autorize_Access.tbluserRow
        Dim rowDGV As DataGridViewRow

        Try

            Me.dgvUser.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvUser.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Me.dgvUser.Rows.Clear()
            Me.tblItemDown.Rows.Clear()
            Me.dgvUser.AllowUserToAddRows = False
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(TBL) > 0 Then
                For Each rowItem In TBL
                    
                    If Me.GetIndexItemUser(rowItem.Kode_User, -1) = -1 Then
                        Me.dgvUser.Rows.Add()
                        rowDGV = Me.dgvUser.Rows(Me.dgvUser.Rows.Count - 1)
                        rowDGV = Me.dgvUser.Rows(Me.dgvUser.Rows.Count - 1)
                        rowDGV.Cells(Me.col2No.Index).Value = Me.dgvUser.Rows.Count
                        rowDGV.Cells(Me.col2UserCode.Index).Value = rowItem.Kode_User
                        rowDGV.Cells(Me.col2Username.Index).Value = rowItem.Nama_User
                        rowDGV.Cells(Me.col2PrimaryUser.Index).Value = rowItem.Primary
                        TambahBarisBawah(Me.dgvUser.Item(Me.col2PrimaryUser.Index, Me.dgvUser.Rows.Count - 1).Value, _
                        rowItem.Primary, _
                        Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value)
                    End If


                    'If Me.dgvUser.Item(Me.colUserCode.Index, Me.dgvUser.Rows.Count - 1).Value = Me.dgvManager.Item(Me.col2UserCode.Index, Me.dgvManager.CurrentRow.Index).Value Then
                    '    Me.EraseRowIndex = Me.dgvUser.Rows.Count - 1
                    '    Me.dgvToRemove = 1
                    '    Me.BackgroundWorker1.RunWorkerAsync()
                    'Else
                    '    TambahBarisBawah(Me.dgvUser.Item(Me.col2PrimaryUser.Index, Me.dgvUser.Rows.Count - 1).Value, _
                    '    rowItem.Primary, _
                    '    Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value)
                    'End If
                    'Else
                    '    Me.dgvUser.Rows.Remove(rowDGV)
                    'End If
                Next
                Me.objCD.SetNumber(Me.dgvUser)
            End If

            Me.dgvUser.AllowUserToAddRows = True
            dgvUser.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvUser.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub ApplyToUserGroupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApplyToUserGroupToolStripMenuItem.Click
        Dim DTA As New DseTrans_SPL_Autorize_AccessTableAdapters.tbluserTableAdapter
        Dim TBL As New DseTrans_SPL_Autorize_Access.tbluserDataTable
        Dim rowItem As DseTrans_SPL_Autorize_Access.tbluserRow
        Dim rowDGV As DataGridViewRow
        Dim ij As New Core_Module.FrmCariMasterHakAkses2
        'Dim stat As Boolean
        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryLPB = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.dgvUser.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvUser.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                'Me.dgvWorkUnit.Rows.Clear()
                Me.dgvUser.AllowUserToAddRows = False
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                If DTA.FillBy_PrimaryByPrimaryGroup(TBL, ij.PrimaryLPB) > 0 Then

                    For Each rowItem In TBL
                        Me.dgvUser.Rows.Add()
                        rowDGV = Me.dgvUser.Rows(Me.dgvUser.Rows.Count - 1)
                        If Me.GetIndexItemUser(rowItem.Kode_User, -1) = -1 And Me.GetIndexItemManager(rowItem.Kode_User, -1) = -1 Then
                            rowDGV = Me.dgvUser.Rows(Me.dgvUser.Rows.Count - 1)
                            rowDGV.Cells(Me.col2No.Index).Value = Me.dgvUser.Rows.Count
                            rowDGV.Cells(Me.col2UserCode.Index).Value = rowItem.Kode_User
                            rowDGV.Cells(Me.col2Username.Index).Value = rowItem.Nama_User
                            rowDGV.Cells(Me.col2PrimaryUser.Index).Value = rowItem.Primary

                            If Me.dgvUser.Item(Me.colUserCode.Index, Me.dgvUser.Rows.Count - 1).Value = Me.dgvManager.Item(Me.col2UserCode.Index, Me.dgvManager.CurrentRow.Index).Value Then
                                Me.EraseRowIndex = Me.dgvUser.Rows.Count - 1
                                Me.dgvToRemove = 1
                                Me.BackgroundWorker1.RunWorkerAsync()
                            Else
                                TambahBarisBawah(Me.dgvUser.Item(Me.col2PrimaryUser.Index, Me.dgvUser.Rows.Count - 1).Value, _
                            rowItem.Primary, _
                            Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value)
                            End If
                        Else
                            Me.dgvUser.Rows.Remove(rowDGV)
                        End If
                    Next
                End If
            End If
            Me.dgvUser.AllowUserToAddRows = True
            dgvUser.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvUser.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Public Function GetIdentifierBaru() As String
        Dim IdentifierD As Integer = 0
        Dim NilaiMax As Integer = 0
        Dim Retval As String = ""

        Try
            For Each rowDGV As DataGridViewRow In Me.dgvManager.Rows
                IdentifierD = rowDGV.Cells(Me.ColIdentifier.Index).Value.ToString.IndexOf("-", 1)
                IdentifierD = rowDGV.Cells(Me.ColIdentifier.Index).Value.ToString.Substring(IdentifierD + 1)
                If NilaiMax < IdentifierD Then
                    NilaiMax = IdentifierD
                End If
            Next
            NilaiMax += 1

            Retval = "-1" & "-" & NilaiMax.ToString
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
    Private Sub dgvManager_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvManager.RowsAdded
        Dim r As DataGridViewRow
        Try
            r = Me.dgvManager.Rows(Me.dgvManager.Rows.Count - 1)
            r.Cells(Me.colNo.Index).Value = Me.dgvManager.Rows.Count
            r.Cells(Me.colUserCode.Index).Value = ""
            r.Cells(Me.colUsername.Index).Value = ""
            r.Cells(Me.colAccessType.Index).Value = "Full Access"
            r.Cells(Me.colPrimary.Index).Value = -1
            r.Cells(Me.ColIdentifier.Index).Value = "-1-0"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvManager_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvManager.SelectionChanged
        Try

            If Me.dgvManager.SelectedCells.Count > 0 Then
                If Me.dgvManager.Item(Me.colUserCode.Index, Me.dgvManager.SelectedCells(0).RowIndex).Value <> Nothing And _
                                    Me.dgvManager.Item(Me.colAccessType.Index, Me.dgvManager.SelectedCells(0).RowIndex).Value = Me.colAccessType.Items.Item(1) Then
                    Me.LlbUser.Enabled = True
                    Me.dgvUser.Enabled = True
                Else
                    Me.LlbUser.Enabled = False
                    Me.dgvUser.Rows.Clear()
                    Me.dgvUser.Enabled = False
                End If
                If Me.dgvManager.CurrentRow.Index > -1 Then
                    reloadUser(Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvManager_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvManager.Sorted
        Me.objCD.SetNumber(Me.dgvManager)
    End Sub

    Private Sub dgvManager_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvManager.UserDeletedRow
        Me.objCD.SetNumber(Me.dgvManager)
    End Sub

    Private Sub dgvManager_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvManager.UserDeletingRow
        Dim identifier As String
        Try
            identifier = Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value
            Me.tblItemDown.DefaultView.RowFilter = "[Identifier Header] = '" & identifier & "'"
            For x As Integer = 0 To Me.tblItemDown.DefaultView.Count - 1
                If Me.tblItemDown.DefaultView.Count > 0 Then
                    Me.rowItemDown = Me.tblItemDown.DefaultView(0).Row
                    Me.tblItemDown.Rows.Remove(Me.rowItemDown)
                End If
            Next
            identifier = Me.dgvManager.Item(Me.ColIdentifier.Index, e.Row.Index).Value
            Me.tblItemUp.DefaultView.RowFilter = "[identifier] = '" & identifier & "'"
            If Me.tblItemUp.DefaultView.Count > 0 Then
                Me.rowItemUp = Me.tblItemUp.DefaultView(0).Row
                Me.tblItemUp.Rows.Remove(Me.rowItemUp)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvUser_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvUser.RowsAdded
        Dim r As DataGridViewRow
        Try
            r = Me.dgvUser.Rows(Me.dgvUser.Rows.Count - 1)
            r.Cells(Me.col2No.Index).Value = Me.dgvUser.Rows.Count
            r.Cells(Me.col2UserCode.Index).Value = ""
            r.Cells(Me.col2Username.Index).Value = ""
            r.Cells(Me.col2PrimaryUser.Index).Value = "-1"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvUser_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUser.Sorted
        Me.objCD.SetNumber(Me.dgvUser)
    End Sub

    Private Sub dgvUser_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvUser.UserDeletedRow
        Me.objCD.SetNumber(Me.dgvUser)
    End Sub


    Private Sub dgvUser_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvUser.UserDeletingRow
        Dim identifier As String
        Try
            identifier = Me.dgvManager.Item(Me.ColIdentifier.Index, Me.dgvManager.CurrentRow.Index).Value
            Me.tblItemDown.DefaultView.RowFilter = "[primary user] = " & Me.dgvUser.Item(Me.col2PrimaryUser.Index, e.Row.Index).Value & " AND [Identifier Header] = '" & identifier & "'"
            If Me.tblItemDown.DefaultView.Count > 0 Then
                Me.rowItemDown = Me.tblItemDown.DefaultView(0).Row
                Me.tblItemDown.Rows.Remove(Me.rowItemDown)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim rowDGV As DataGridViewRow
        Dim identifier As String
        Try
            'cek minimal ada satu user Authorize
            rowDGV = Me.dgvManager.Rows(0)
            If rowDGV.Cells(Me.colUsername.Index).Value = "" And ChkActive.Checked = True Then
                MsgBox("There must be one User Authorize", MsgBoxStyle.Information, "Information")
                Return False
            Else
                For im As Integer = 0 To Me.dgvManager.Rows.Count - 2
                    rowDGV = Me.dgvManager.Rows(im)
                    If rowDGV.Cells(Me.colAccessType.Index).Value = Me.colAccessType.Items.Item(1) Then
                        identifier = Me.dgvManager.Item(Me.ColIdentifier.Index, im).Value
                        Me.tblItemDown.DefaultView.RowFilter = "[Identifier Header] = '" & identifier & "'"
                        If Me.tblItemDown.DefaultView.Count = 0 Then
                            MsgBox("There must be one User Item Authorize in User '" & rowDGV.Cells.Item(Me.col2UserCode.Index).Value _
                            & "'", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                    'Exit For
                Next
            End If
            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function


    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
            If hasil = MsgBoxResult.Yes Then
                Me.KosongkanDGV()
                objAA.IsiDGVHeader(Me.dgvManager, Me)
                Me.buffer = ""
                Me.Close()
            Else
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanDGV()
        Try
            Me.tblItemUp.Rows.Clear()
            Me.tblItemDown.Rows.Clear()

            Me.dgvManager.Rows.Clear()
            Me.dgvUser.Rows.Clear()
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