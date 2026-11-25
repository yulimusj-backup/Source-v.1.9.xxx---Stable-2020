''@rosadinnn
''201508111022

Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Drawing

Public Class FrmMasterAttendanceServer

    Public UserCode As String = "USER"
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Dim Baru As Boolean = False
    Dim Buffer As String = ""
    Dim DgvRowIdx As Integer = -1
    Dim Obj As New ClsMasterAttendanceServer
    Dim ObjCd As New Core_Module.ClsCekData
    Dim ControlNitgenMagic As Boolean = True

#Region "common"
    Private Sub AktifkanDetail()
        Try
            Me.ObjCd.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCd.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCd.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCd.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.DgvItem.ReadOnly = False
            Me.DgvItem.AllowUserToDeleteRows = True
            Me.DgvItem.AllowUserToAddRows = True
            Me.CancelButton = Me.btnCancel
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCd.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCd.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCd.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, True)
            Me.ObjCd.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.DgvItem.ReadOnly = True
            Me.DgvItem.AllowUserToDeleteRows = False
            Me.DgvItem.AllowUserToAddRows = False
            Me.CancelButton = Me.btnExit
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCd.KosongkanControl(SplitContainer2.Panel1.Controls, True, True, True)
            Me.ObjCd.KosongkanControl(SplitContainer2.Panel2.Controls, True, True, True)
            Me.CboServer.SelectedIndex = 0
            Me.TxtPrimary.Text = "-1"
            Me.DgvItem.Rows.Clear()
            Me.ChkIsActive.Checked = True
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtServerName.Text = .ServerName
                Me.TxtLocation.Text = .Location
                If .IsActive = 0 Then
                    Me.ChkIsActive.Checked = False
                Else
                    Me.ChkIsActive.Checked = True
                End If
                Me.TxtRemarks.Text = .Remarks
                Me.TxtDataSource.Text = .DataSource
                Me.TxtInitialcatalog.Text = .InitialCatalog
                Me.TxtUserId.Text = .UserId
                Me.TxtPassword.Text = .Password
                Me.ControlNitgenMagic = False
                Me.CboServer.SelectedIndex = .MachineType
                Me.ControlNitgenMagic = True
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .Primary = Me.TxtPrimary.Text
                .ServerName = Me.TxtServerName.Text
                .Location = Me.TxtLocation.Text
                If Me.ChkIsActive.Checked = False Then
                    .IsActive = 0
                Else
                    .IsActive = 1
                End If
                .Remarks = Me.TxtRemarks.Text
                .DataSource = Me.TxtDataSource.Text
                .InitialCatalog = Me.TxtInitialcatalog.Text
                .UserId = Me.TxtUserId.Text
                .Password = Me.TxtPassword.Text
                If Me.Baru = True Then
                    .CreatedBy = Me.UserCode
                    .CreatedDate = Now
                End If
                .ModifiedBy = Me.UserCode
                .ModifiedDate = Now
                .TimeStamp = Now
                .MachineType = Me.CboServer.SelectedIndex
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            KoleksiNamaData.Add("Server name")
            KoleksiTextBox.Add(Me.TxtServerName)
            KoleksiNamaData.Add("Location")
            KoleksiTextBox.Add(Me.TxtLocation)
            KoleksiNamaData.Add(Me.Label3.Text)
            KoleksiTextBox.Add(Me.TxtDataSource)
            KoleksiNamaData.Add("Database")
            KoleksiTextBox.Add(Me.TxtInitialcatalog)
            KoleksiNamaData.Add("User ID")
            KoleksiTextBox.Add(Me.TxtUserId)
            If Me.CboServer.SelectedIndex = 0 Then
                KoleksiNamaData.Add("Password")
                KoleksiTextBox.Add(Me.TxtPassword)
            End If
            If ObjCd.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            If Me.Baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtServerName.Text) = True Then
                    MsgBox(Me.TxtServerName.Text & " is already exist.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                If Me.Obj.ApakahSudahAda(Me.TxtServerName.Text) = True Then
                    If Me.Obj.GetServerName(Me.TxtPrimary.Text) <> Me.TxtServerName.Text Then
                        MsgBox(Me.TxtServerName.Text & " is already exist.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            If Me.DgvItem.RowCount = 1 Then
                MsgBox("Please input at least 1 machine.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Private Function ApakahDataItemBenar() As Boolean
        Try
            For Each DgvRow As DataGridViewRow In Me.DgvItem.Rows
                If IsNothing(DgvRow.Cells(Me.ColMachineCode.Index)) = False Then
                    If DgvRow.Cells(Me.ColMachineCode.Index).Value.ToString <> "" Then
                        If String.IsNullOrEmpty(DgvRow.Cells(Me.ColTerminalID.Index).Value) = True Then
                            MsgBox("Terminal ID cannot be an empty string.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                        If String.IsNullOrEmpty(DgvRow.Cells(Me.ColRemarks.Index).Value) = True Then
                            MsgBox("Terminal location cannot be an empty string.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function
#End Region

#Region "textbox focus"
    Private Sub Txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtServerName.GotFocus, TxtLocation.GotFocus, TxtRemarks.GotFocus, TxtDataSource.GotFocus, TxtInitialcatalog.GotFocus, TxtUserId.GotFocus, TxtPassword.GotFocus
        Dim tb As TextBox
        Try
            tb = TryCast(sender, TextBox)
            If tb.Enabled = True Then
                If tb.ReadOnly = False Then
                    tb.BackColor = Me.BackColorAll
                Else
                    tb.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtServerName.LostFocus, TxtLocation.LostFocus, TxtRemarks.LostFocus, TxtDataSource.LostFocus, TxtInitialcatalog.LostFocus, TxtUserId.LostFocus, TxtPassword.LostFocus
        Dim tb As TextBox
        Try
            tb = TryCast(sender, TextBox)
            If tb.Enabled = True Then
                tb.BackColor = Color.White
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    Private Sub FrmMasterAttendanceServer_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If (Me.TxtRemarks.Enabled = True) And (Me.TxtServerName.Text.Trim.Length > 0 Or Me.TxtLocation.Text.Trim.Length > 0 Or _
                Me.TxtRemarks.Text.Trim.Length > 0 Or Me.TxtDataSource.Text.Trim.Length > 0 Or Me.TxtInitialcatalog.Text.Trim.Length > 0 Or _
                Me.TxtUserId.Text.Trim.Length > 0 Or Me.TxtPassword.Text.Trim.Length > 0) Then
                Dim hasil As MsgBoxResult
                hasil = MsgBox("Do you want to save data before exit?", MsgBoxStyle.YesNoCancel, "Confirmation")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Me.Baru = True Then
                            Me.Cursor = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me, Me.Baru) = True Then
                                'MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            Me.Cursor = Cursors.WaitCursor
                            If Me.Obj.SimpanPerubahan(Me, Me.Baru) = True Then
                                'MsgBox("Data change saved.", MsgBoxStyle.Information, "Information")
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
            MsgBox(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub FrmMasterAttendanceServer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.NonAktifkanDetail()
            Me.KosongkanField()

            'ahdar
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbTest_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbTest.LinkClicked
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim Connection_ As Object
        If Me.CboServer.SelectedIndex = 0 Then
            Connection_ = New SqlClient.SqlConnection
        ElseIf Me.CboServer.SelectedIndex = 1 Then
            Connection_ = New OleDb.OleDbConnection
        End If
        Try
            If Me.TxtRemarks.Enabled = False And Me.TxtPrimary.Text = "-1" Then
                MsgBox("Open the server first.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            KoleksiNamaData.Add(Me.Label3.Text)
            KoleksiTextBox.Add(Me.TxtDataSource)
            KoleksiNamaData.Add("Database")
            KoleksiTextBox.Add(Me.TxtInitialcatalog)
            KoleksiNamaData.Add("User ID")
            KoleksiTextBox.Add(Me.TxtUserId)
            If Me.CboServer.SelectedIndex = ClsMasterAttendanceServer.EnumMachineType.NITGEN Then
                KoleksiNamaData.Add("Password")
                KoleksiTextBox.Add(Me.TxtPassword)
            End If
            If ObjCd.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Exit Sub
            End If
            If Me.CboServer.SelectedIndex = ClsMasterAttendanceServer.EnumMachineType.NITGEN Then
                Connection_.ConnectionString = "Data Source=" & Me.TxtDataSource.Text & ";Initial Catalog=" & Me.TxtInitialcatalog.Text & ";Persist Security Info=True;User ID=" & Me.TxtUserId.Text & ";Password=" & Me.TxtPassword.Text & ";Connect Timeout=10"
            ElseIf Me.CboServer.SelectedIndex = ClsMasterAttendanceServer.EnumMachineType.MAGIC Then
                Connection_.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Me.TxtDataSource.Text & ";Persist Security Info=True;User ID=" & Me.TxtUserId.Text & ";Jet OLEDB:Database Password=" & Me.TxtPassword.Text
            End If

            Try
                Connection_.Open()
                If Connection_.State = ConnectionState.Open Then
                    MsgBox("Test connection succeeded.", MsgBoxStyle.Information, "Information")
                    Connection_.Close()
                End If
            Catch exAcc As OleDb.OleDbException
                If exAcc.ErrorCode = -2147217843 Then
                    MsgBox("Invalid User ID or Password.", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Test connection failed.", MsgBoxStyle.Information, "Information")
                End If
            Catch ex As SqlClient.SqlException
                If ex.Number = -1 Then
                    MsgBox("The server was not found.", MsgBoxStyle.Information, "Information")
                ElseIf ex.Number = 4060 Then
                    MsgBox("Database does not exist.", MsgBoxStyle.Information, "Information")
                ElseIf ex.Number = 18456 Then
                    MsgBox("Login failed for user '" & Me.TxtUserId.Text & "'.", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Test connection failed.", MsgBoxStyle.Information, "Information")
                End If
            End Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Connection_.Dispose()
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColNo.Index Then
                e.Cancel = True
            Else
                If e.ColumnIndex <> Me.ColMachineCode.Index And Me.DgvItem.Item(Me.ColMachineCode.Index, e.RowIndex).Value = "" Then
                    e.Cancel = True
                Else
                    Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Function GetIndex(ByVal vCurrentRowIndex As Integer, ByVal vMachineName As String) As Integer
        Try
            For Each DgvRow As DataGridViewRow In Me.DgvItem.Rows
                If Me.DgvItem.Rows.IndexOf(DgvRow) <> vCurrentRowIndex Then
                    If DgvRow.Cells(Me.ColMachineCode.Index).Value = vMachineName Then
                        Return Me.DgvItem.Rows.IndexOf(DgvRow)
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Private Function GetIndex2(ByVal vCurrentRowIndex As Integer, ByVal vTerminalId As String) As Integer
        Try
            For Each DgvRow As DataGridViewRow In Me.DgvItem.Rows
                If Me.DgvItem.Rows.IndexOf(DgvRow) <> vCurrentRowIndex Then
                    If DgvRow.Cells(Me.ColTerminalID.Index).Value = vTerminalId Then
                        Return Me.DgvItem.Rows.IndexOf(DgvRow)
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Try
            If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString <> "" Then
                    If e.ColumnIndex = Me.ColMachineCode.Index Then
                        If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Or _
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Then
                            MsgBox("Machine code cannot contains these characters; ', #, %, or *.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Else
                                If e.RowIndex <> Me.DgvItem.Rows.Count - 1 Then
                                    Me.DgvRowIdx = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                End If
                            End If
                            Exit Sub
                        End If
                        If Me.GetIndex(e.RowIndex, Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) <> -1 Then
                            MsgBox(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString & " is already exist.", MsgBoxStyle.Information, "Information")
                            If Me.Buffer <> "" Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Else
                                If e.RowIndex <> Me.DgvItem.Rows.Count - 1 Then
                                    Me.DgvRowIdx = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                End If
                            End If
                            Exit Sub
                        End If
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper.Trim
                    End If
                    If e.ColumnIndex = Me.ColMachineName.Index Then
                        If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Or _
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Then
                            MsgBox("Machine name cannot contains these characters; ', #, %, or *.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    If e.ColumnIndex = Me.ColTerminalID.Index Then
                        If IsNumeric(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            MsgBox("Allowed terminal ID is: 1-250.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains(".") Or _
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains(",") Then
                            MsgBox("Allowed terminal ID is: 1-250.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If CDec(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) < 1 Or _
                            CDec(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString > 250) Then
                            MsgBox("Allowed terminal ID is: 1-250.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If Me.GetIndex2(e.RowIndex, Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) <> -1 Then
                            MsgBox(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString & " is already exist.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Or _
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Then
                            MsgBox("Remarks cannot contains these characters; ', #, %, or *.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                Else
                    If e.ColumnIndex <> Me.ColRemarks.Index And e.ColumnIndex <> Me.ColMachineName.Index Then
                        If Me.Buffer <> "" Then
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            If e.RowIndex <> Me.DgvItem.Rows.Count - 1 Then
                                Me.DgvRowIdx = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                        End If
                    End If
                End If
            Else
                If e.ColumnIndex <> Me.ColRemarks.Index And e.ColumnIndex <> Me.ColMachineName.Index Then
                    If Me.Buffer <> "" Then
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Else
                        If e.RowIndex <> Me.DgvItem.Rows.Count - 1 Then
                            Me.DgvRowIdx = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
                End If
                End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItem.RowsAdded
        Try
            Me.DgvItem.Item(Me.ColNo.Index, Me.DgvItem.Rows.Count - 1).Value = ""
            Me.DgvItem.Item(Me.ColMachineCode.Index, Me.DgvItem.Rows.Count - 1).Value = ""
            Me.DgvItem.Item(Me.ColMachineName.Index, Me.DgvItem.Rows.Count - 1).Value = ""
            Me.DgvItem.Item(Me.ColTerminalID.Index, Me.DgvItem.Rows.Count - 1).Value = ""
            Me.DgvItem.Item(Me.ColIsActive.Index, Me.DgvItem.Rows.Count - 1).Value = Boolean.TrueString
            Me.DgvItem.Item(Me.ColRemarks.Index, Me.DgvItem.Rows.Count - 1).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvItem.RowsRemoved
        Try
            Me.ObjCd.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Try
            Me.ObjCd.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.ApakahDataBenar() = True Then
                If Me.ApakahDataItemBenar = True Then
                    Me.SalinKeObjek()
                    If Me.Baru = True Then
                        If MsgBox("Are you sure to save?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                            Me.Cursor = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me, Me.Baru) = True Then
                                MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            End If
                        End If
                    Else
                        If MsgBox("Are you sure to save?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                            Me.Cursor = Cursors.WaitCursor
                            If Me.Obj.SimpanPerubahan(Me, Me.Baru) = True Then
                                MsgBox("Data change saved.", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If Me.TxtServerName.Text <> "" Or Me.TxtLocation.Text <> "" Or Me.TxtRemarks.Text <> "" Or Me.TxtDataSource.Text <> "" Or _
                    Me.TxtUserId.Text <> "" Or Me.TxtPassword.Text <> "" Or Me.TxtInitialcatalog.Text <> "" Then
                If MsgBox("Are you sure to cancel and discard data changes?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Confirmation") = MsgBoxResult.Yes Then
                    Me.NonAktifkanDetail()
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.Obj.CariDetail(Me.TxtPrimary.Text, Me)
                        Me.SalinKeField()
                    End If
                End If
            Else
                Me.KosongkanField()
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim frm As New FrmMasterAttendanceServer_Cari
        Try
            frm.BackColorAll = Me.BackColorAll
            frm.ShowDialog()
            If Not frm.mPrimary = -1 Then
                If Me.Obj.CariDetail(frm.mPrimary, Me) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            frm.Dispose()
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.Baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            If Me.CboServer.SelectedIndex = 0 Then
                Me.CboServer.SelectedIndex = 1
                Me.CboServer.SelectedIndex = 0
            ElseIf Me.CboServer.SelectedIndex = 1 Then
                Me.CboServer.SelectedIndex = 0
                Me.CboServer.SelectedIndex = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimary.Text <> "-1" Then
                Me.Baru = False
                Me.AktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If Me.TxtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimary.Text <> "-1" Then
                If MsgBox("Are you sure to delete?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    Me.Cursor = Cursors.WaitCursor
                    If Me.Obj.Hapus(Me.TxtPrimary.Text) = True Then
                        MsgBox("Data deleted.", MsgBoxStyle.Information, "Information")
                        KosongkanField()
                        NonAktifkanDetail()
                        btnNew.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvItem.Rows.RemoveAt(Me.dgvRowIdx)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TxtServerName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtServerName.Leave
        Try
            Me.TxtServerName.Text = Me.TxtServerName.Text.Trim
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DgvItem_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvItem.UserAddedRow
        Try
            Me.ObjCd.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub CboServer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboServer.SelectedIndexChanged
        Try
            If Me.CboServer.SelectedIndex = 0 Then
                Me.ControlNitgen(Me.ControlNitgenMagic)
            ElseIf Me.CboServer.SelectedIndex = 1 Then
                Me.ControlMagic(Me.ControlNitgenMagic)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ControlNitgen(ByVal vControlNitgenMagic As Boolean)
        Try
            Me.TxtDataSource.ReadOnly = False
            Me.TxtDataSource.Size = Me.TxtInitialcatalog.Size
            Me.BtnFindAccessFile.Visible = False
            Me.TxtInitialcatalog.ReadOnly = True
            Me.TxtUserId.ReadOnly = True
            Me.Label3.Text = "Data Source"
            If vControlNitgenMagic = True Then
                Me.TxtDataSource.Clear()
                Me.TxtInitialcatalog.Text = "NitgenAccessManager"
                Me.TxtUserId.Text = "sa"
                Me.TxtPassword.Clear()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ControlMagic(ByVal vControlNitgenMagic As Boolean)
        Try
            Me.TxtDataSource.ReadOnly = True
            Me.TxtDataSource.Size = New System.Drawing.Size(Me.TxtInitialcatalog.Size.Width - 5 - Me.BtnFindAccessFile.Size.Width, Me.TxtDataSource.Height)
            Me.BtnFindAccessFile.Visible = True
            Me.TxtInitialcatalog.ReadOnly = True
            Me.TxtUserId.ReadOnly = False
            Me.Label3.Text = "Ms.Access File"
            If vControlNitgenMagic = True Then
                Me.TxtDataSource.Clear()
                Me.TxtInitialcatalog.Clear()
                Me.TxtUserId.Clear()
                Me.TxtPassword.Clear()
            End If
        Catch ex As Exception

        End Try
    End Sub


    Private Sub BtnFindAccessFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindAccessFile.Click
        Dim Buffer As String = ""
        Try
            Buffer = Me.TxtDataSource.Text
            If Me.OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TxtDataSource.Text = OpenFileDialog1.FileName.ToString
                Me.TxtInitialcatalog.Text = My.Computer.FileSystem.GetName(OpenFileDialog1.FileName.ToString)
            Else
                Me.TxtDataSource.Text = Buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class