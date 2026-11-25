''@rosadinnn
''20150422

Option Explicit On
Option Strict Off

Public Class FrmEmailSettings_LogReport

    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim tblEmailAccounts As New DseEmailSettings.tbltrans_emailsettingsDataTable
    Dim rowEmaiAccounts As DseEmailSettings.tbltrans_emailsettingsRow
    Dim objCd As New Core_Module.ClsCekData
    Dim obj As New ClsEmailSettings_LogReport
    Dim objEmail As New ClsEmailSettings
    Dim objCust As New Core_Module.clsMasterCustomer
    Dim ibuffer As String = ""
    Dim iSelectedAccountCount As Integer = 0
    Dim iListEmailAccountFilter As List(Of Double) = New List(Of Double)
    Dim iFilterCustomer As Boolean = False
    Dim iCustomerCount As Integer = 0
    Dim iListCustomerFilter As List(Of Double) = New List(Of Double)
    Dim vTblEmailLog As New DseEmailSettings.DataTable3DataTable
    Dim vRowEmailLog As DseEmailSettings.DataTable3Row
    Dim iLog As Boolean = False

    Dim Order1, Order2, Order3, Order4, Order5 As String
    Dim Tipe1, Tipe2, Tipe3, Tipe4, Tipe5 As String
    Dim OrderTemp As String = ""
    Dim OrderDirectionTemp As String = "ASC"


    Private Sub FrmEmailSettings_LogReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.LoadEmailAccounts()
            Me.LoadSortingSetting()
            Me.chkShowAllCustomer.Checked = True
            Me.cboDeliveryStatus.SelectedIndex = 0
            Me.btnPrint.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LoadSortingSetting()
        Dim rowDgv As DataGridViewRow
        Try
            For i As Integer = 0 To 5
                Me.dgvSortBy.Rows.Add("")
                rowDgv = Me.dgvSortBy.Rows(i)
                rowDgv.Cells(Me.ColNo.Index).Value = (i + 1).ToString
                Select Case i
                    Case 0
                        rowDgv.Cells(Me.ColField.Index).Value = "Sender"
                    Case 1
                        rowDgv.Cells(Me.ColField.Index).Value = "Receiver"
                    Case 2
                        rowDgv.Cells(Me.ColField.Index).Value = "Transaction"
                    Case 3
                        rowDgv.Cells(Me.ColField.Index).Value = "Date"
                    Case 4
                        rowDgv.Cells(Me.ColField.Index).Value = "Doc No"
                    Case 5
                        rowDgv.Cells(Me.ColField.Index).Value = "Delivery Status"
                End Select
                rowDgv.Cells(Me.ColOrder.Index).Value = (i + 1).ToString
                rowDgv.Cells(Me.ColDirection.Index).Value = "Ascending"
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub LoadEmailAccounts()
        Try
            Me.chkListEmailAccounts.Items.Clear()
            If Me.obj.LoadEmailAccounts(Me.tblEmailAccounts) = True Then
                If Me.tblEmailAccounts.Count > 0 Then
                    For Each Me.rowEmaiAccounts In Me.tblEmailAccounts
                        Me.chkListEmailAccounts.Items.Add(rowEmaiAccounts.Account_Code & " - " & rowEmaiAccounts.Email_From)
                    Next
                    CheckUncheckAllEmailAccounts(True)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub CheckUncheckAllEmailAccounts(ByVal vBol As Boolean)
        Try
            For i As Integer = 0 To chkListEmailAccounts.Items.Count - 1
                If vBol = True Then
                    Me.chkListEmailAccounts.SetItemChecked(i, True)
                Else
                    Me.chkListEmailAccounts.SetItemChecked(i, False)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            Me.CheckUncheckAllEmailAccounts(True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Me.CheckUncheckAllEmailAccounts(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub cboShowAllCustomer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowAllCustomer.CheckedChanged
        Try
            If Me.chkShowAllCustomer.Checked = True Then
                Me.dgvCustomer.Enabled = False
            Else
                Me.dgvCustomer.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvCustomer_CellBeginEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCustomer.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColCustCode2.Index Then
                e.Cancel = True
            Else
                Me.ibuffer = Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvCustomer_CellEndEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCustomer.CellEndEdit
        Dim frm As New Core_Module.FrmCariMasterCustomer
        Try
            If IsNothing(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                    ''cek to other rows
                    If Me.IsExistRow(e.RowIndex, Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = True Then
                        MsgBox("The Customer Code already exist.", MsgBoxStyle.Information, "Information")
                        ''restore/remove
                        If Me.ibuffer <> "" Then
                            Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.ibuffer
                        Else
                            Me.RemoveDgvCustAt(e.RowIndex)
                        End If
                        Exit Sub
                    End If
                    If Me.objCust.ApakahSudahAda(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvCustomer.Item(Me.ColPrimaryCust2.Index, e.RowIndex).Value = Me.objCust.GetPrimary(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvCustomer.Item(Me.ColCustCode2.Index, e.RowIndex).Value = Me.objCust.GetCustomerCode(Me.dgvCustomer.Item(Me.ColPrimaryCust2.Index, e.RowIndex).Value)
                        Me.dgvCustomer.Item(Me.ColCustName2.Index, e.RowIndex).Value = Me.objCust.GetCustName(Me.dgvCustomer.Item(Me.ColPrimaryCust2.Index, e.RowIndex).Value)
                    Else
                        frm.txtCari.Text = Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
                        frm.ShowDialog()
                        If frm.PrimaryCustomer > 0 Then
                            Me.dgvCustomer.Item(Me.ColPrimaryCust2.Index, e.RowIndex).Value = frm.PrimaryCustomer
                            Me.dgvCustomer.Item(Me.ColCustCode2.Index, e.RowIndex).Value = Me.objCust.GetCustomerCode(frm.PrimaryCustomer)
                            Me.dgvCustomer.Item(Me.ColCustName2.Index, e.RowIndex).Value = Me.objCust.GetCustName(frm.PrimaryCustomer)
                        Else
                            ''restore/remove
                            If Me.ibuffer <> "" Then
                                Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.ibuffer
                            Else
                                Me.RemoveDgvCustAt(e.RowIndex)
                            End If
                        End If
                    End If
                Else
                    ''restore/remove
                    If Me.ibuffer <> "" Then
                        Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.ibuffer
                    Else
                        Me.RemoveDgvCustAt(e.RowIndex)
                    End If
                End If
            Else
                ''restore/remove
                If Me.ibuffer <> "" Then
                    Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.ibuffer
                Else
                    Me.RemoveDgvCustAt(e.RowIndex)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm.Dispose()
        End Try
    End Sub

    Private Function IsExistRow(ByVal vIdx As Integer, ByVal vCode As String) As Boolean
        Try
            For Each rowDgv As DataGridViewRow In Me.dgvCustomer.Rows
                If Me.dgvCustomer.Rows.IndexOf(rowDgv) <> vIdx Then
                    If Me.dgvCustomer.Rows.IndexOf(rowDgv) <> Me.dgvCustomer.Rows.Count - 1 Then
                        If rowDgv.Cells(Me.ColCustCode2.Index).Value.ToString.ToUpper = vCode.ToUpper Then
                            Return True
                        End If
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally

        End Try
    End Function

    Private Sub RemoveDgvCustAt(ByVal vIdx As Integer)
        Try
            If vIdx <> -1 Then
                Me.dgvCustomer.Rows.RemoveAt(vIdx)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvCustomer_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvCustomer.RowsAdded
        Try
            Me.objCd.SetNumber(Me.dgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvCustomer_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvCustomer.RowsRemoved
        Try
            Me.objCd.SetNumber(Me.dgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvCustomer_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCustomer.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvSortBy_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSortBy.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColOrder.Index Or e.ColumnIndex = Me.ColDirection.Index Then
                If e.ColumnIndex = Me.ColOrder.Index Then
                    Me.ibuffer = Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value.ToString
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvSortBy_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSortBy.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColOrder.Index Then '' jika kolom order
                If IsNothing(Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value) = False Then ''jika tidak kosong
                    If Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value.ToString <> "" Then ''jika tidak kosong
                        If IsNumeric(Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value) = True Then ''jika angka
                            Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCd.DigitGrouping(Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value, 0)
                            If (CInt(Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value) >= 1) And (CInt(Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value <= 6)) Then ''jika nilai angka diantara 1-5
                                Me.SwitchOrder(e.RowIndex, Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
                            Else
                                Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.ibuffer
                            End If
                        Else
                            Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.ibuffer
                        End If
                    Else
                        Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.ibuffer
                    End If
                Else
                    Me.dgvSortBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.ibuffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Function SwitchOrder(ByVal vDgvRowIdx As Integer, ByVal vNewOrder As String) As Boolean
        Try
            If Me.dgvSortBy.RowCount > 0 Then
                For Each rowDgv As DataGridViewRow In Me.dgvSortBy.Rows
                    If Me.dgvSortBy.Rows.IndexOf(rowDgv) <> vDgvRowIdx Then
                        If rowDgv.Cells(Me.ColOrder.Index).Value.ToString = vNewOrder Then
                            rowDgv.Cells(Me.ColOrder.Index).Value = Me.ibuffer
                            Return True
                        End If
                    End If
                Next
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally

        End Try
    End Function

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Me.dgvData.Rows.Clear()
        Try
            ''validasi email account
            If Me.chkListEmailAccounts.Items.Count = 0 Then
                MsgBox("There is no email account", MsgBoxStyle.Information, "Information")
                Me.btnShow.Focus()
                Exit Sub
            End If

            Me.iSelectedAccountCount = 0
            Me.iListEmailAccountFilter.Clear()
            For i2 As Integer = 0 To Me.chkListEmailAccounts.Items.Count - 1
                If Me.chkListEmailAccounts.GetItemCheckState(i2) = CheckState.Checked Then
                    Me.iSelectedAccountCount += 1
                    Me.rowEmaiAccounts = Me.tblEmailAccounts.Rows(i2)
                    Me.iListEmailAccountFilter.Add(Me.rowEmaiAccounts.Primary)
                End If
            Next
            If Me.iSelectedAccountCount <= 0 Then
                MsgBox("Please select at least 1 email account", MsgBoxStyle.Information, "Information")
                Me.chkListEmailAccounts.Focus()
                Exit Sub
            End If

            ''validasi filter customer
            Me.iFilterCustomer = False
            Me.iCustomerCount = 0
            Me.iListCustomerFilter.Clear()
            If Me.chkShowAllCustomer.Checked = False Then
                Me.iFilterCustomer = True
                For Each rowDgv As DataGridViewRow In Me.dgvCustomer.Rows
                    If CInt(Me.dgvCustomer.Rows.IndexOf(rowDgv)) <> CInt(Me.dgvCustomer.Rows.Count - 1) Then
                        If rowDgv.Cells(Me.ColPrimaryCust2.Index).Value.ToString <> "" And rowDgv.Cells(Me.ColPrimaryCust2.Index).Value.ToString <> "-1" Then
                            iCustomerCount += 1
                            Me.iListCustomerFilter.Add(CDbl(rowDgv.Cells(Me.ColPrimaryCust2.Index).Value.ToString))
                        End If
                    End If
                Next
                If Me.iCustomerCount <= 0 Then
                    MsgBox("Please input at least 1 customer", MsgBoxStyle.Information, "Information")
                    Me.dgvCustomer.Focus()
                    Exit Sub
                End If
            End If

            ''validasi dtp
            Me.dtpFrom.Value = Format(CDate(Me.dtpFrom.Value), "yyyy-MM-dd 00:00:00")
            Me.dtpTo.Value = Format(CDate(Me.dtpTo.Value), "yyyy-MM-dd 23:59:00")
            If Me.dtpTo.Value < Me.dtpFrom.Value Then
                MsgBox("Start period value cannot more than end period", MsgBoxStyle.Information, "Information")
                Me.dtpFrom.Focus()
                Exit Sub
            End If

            Me.iLog = False
            If Me.cboDeliveryStatus.SelectedIndex = 0 Then
                If Me.obj.ShowLog(Me.vTblEmailLog, Me.dtpFrom.Value, Me.dtpTo.Value) = True Then
                    Me.iLog = True
                End If
            Else
                If Me.cboDeliveryStatus.SelectedIndex = 1 Then
                    If Me.obj.ShowLog(Me.vTblEmailLog, Me.dtpFrom.Value, Me.dtpTo.Value, True) Then
                        Me.iLog = True
                    End If
                ElseIf Me.cboDeliveryStatus.SelectedIndex = 2 Then
                    If Me.obj.ShowLog(Me.vTblEmailLog, Me.dtpFrom.Value, Me.dtpTo.Value, False) Then
                        Me.iLog = True
                    End If
                End If
            End If
            If Me.iLog = True Then
                Me.Cursor = Cursors.WaitCursor
                For a As Integer = 0 To Me.vTblEmailLog.Rows.Count - 1
                    If Me.vTblEmailLog.Count < 0 Or a > Me.vTblEmailLog.Count - 1 Then
                        Exit For
                    End If
                    Me.vRowEmailLog = Me.vTblEmailLog.Rows(a)
                    ''filter email account::
                    If Me.iListEmailAccountFilter.Contains(Me.vRowEmailLog.Primary_Email_Settings) = False Then
                        Me.vTblEmailLog.Rows.RemoveAt(a)
                        a = a - 1
                    Else
                        ''filter customer::
                        If Me.iFilterCustomer = True Then
                            If Me.iListCustomerFilter.Contains(Me.vRowEmailLog.Primary_Customer) = False Then
                                Me.vTblEmailLog.Rows.RemoveAt(a)
                                a = a - 1
                            End If
                        End If
                    End If
                Next
                For Each Me.vRowEmailLog In Me.vTblEmailLog
                    ''filter email account
                    If Me.iListEmailAccountFilter.Contains(Me.vRowEmailLog.Primary_Email_Settings) = False Then
                        Me.vTblEmailLog.Rows.Remove(Me.vRowEmailLog)
                    End If
                    ''filter customer
                    If Me.iFilterCustomer = True Then
                        If Me.iListCustomerFilter.Contains(Me.vRowEmailLog.Primary_Customer) = False Then
                            Me.vTblEmailLog.Rows.Remove(Me.vRowEmailLog)
                        End If
                    End If
                Next
                If Me.vTblEmailLog.Count > 0 Then
                    ''sort order
                    Me.OrderTemp = "[Account Code]"
                    OrderDirectionTemp = "ASC"
                    For Each rowDgv As DataGridViewRow In Me.dgvSortBy.Rows
                        With rowDgv
                            Select Case .Cells(Me.ColField.Index).Value.ToString
                                Case "Sender"
                                    Me.OrderTemp = "[Account Code]"
                                Case "Receiver"
                                    Me.OrderTemp = "[Customer Code]"
                                Case "Transaction"
                                    Me.OrderTemp = "[Tipe Transaksi]"
                                Case "Date"
                                    Me.OrderTemp = "[Date]"
                                Case "Doc No"
                                    Me.OrderTemp = "[Doc No]"
                                Case "Delivery Status"
                                    Me.OrderTemp = "[Status]"
                            End Select
                            If .Cells(Me.ColDirection.Index).Value.ToString = "Ascending" Then
                                Me.OrderDirectionTemp = "ASC"
                            Else
                                Me.OrderDirectionTemp = "DESC"
                            End If
                            If .Cells(Me.ColOrder.Index).Value.ToString = "1" Then
                                Me.Order1 = Me.OrderTemp
                                Me.Tipe1 = Me.OrderDirectionTemp
                            End If
                            If .Cells(Me.ColOrder.Index).Value.ToString = "2" Then
                                Me.Order2 = Me.OrderTemp
                                Me.Tipe2 = Me.OrderDirectionTemp
                            End If
                            If .Cells(Me.ColOrder.Index).Value.ToString = "3" Then
                                Me.Order3 = Me.OrderTemp
                                Me.Tipe3 = Me.OrderDirectionTemp
                            End If
                            If .Cells(Me.ColOrder.Index).Value.ToString = "4" Then
                                Me.Order4 = Me.OrderTemp
                                Me.Tipe4 = Me.OrderDirectionTemp
                            End If
                            If .Cells(Me.ColOrder.Index).Value.ToString = "5" Then
                                Me.Order5 = Me.OrderTemp
                                Me.Tipe5 = Me.OrderDirectionTemp
                            End If
                        End With
                    Next
                    Me.vTblEmailLog.DefaultView.Sort = Me.Order1 & " " & Me.Tipe1 & "," & Me.Order2 & " " & Me.Tipe2 & "," & Me.Order3 & " " & Me.Tipe3 & "," & Me.Order4 & " " & Me.Tipe4 & "," & Me.Order5 & " " & Me.Tipe5
                    Me.vTblEmailLog.DefaultView.RowFilter = ""

                    ''show on DGV
                    Dim dgvRow As DataGridViewRow

                    For i As Integer = 0 To Me.vTblEmailLog.DefaultView.Count - 1
                        Me.vRowEmailLog = Me.vTblEmailLog.DefaultView(i).Row
                        Me.dgvData.Rows.Add("")
                        dgvRow = Me.dgvData.Rows(Me.dgvData.Rows.Count - 1)
                        With dgvRow
                            .Cells(Me.ColNo3.Index).Value = ""
                            .Cells(Me.ColAccount3.Index).Value = vRowEmailLog.Account_Code & "-" & vRowEmailLog.Email_From
                            .Cells(Me.ColCustCode3.Index).Value = vRowEmailLog.Customer_Code
                            .Cells(Me.ColCustName3.Index).Value = vRowEmailLog.Customer_Name
                            If vRowEmailLog.Tipe_Transaksi = ClsEmailSettings.EnumTipeTransaksiEmail.CollectionLetter Then
                                .Cells(Me.ColTransaction3.Index).Value = "Collection Letter"
                            ElseIf vRowEmailLog.Tipe_Transaksi = ClsEmailSettings.EnumTipeTransaksiEmail.Invoice Then
                                .Cells(Me.ColTransaction3.Index).Value = "Invoice"
                            End If
                            .Cells(Me.ColDate3.Index).Value = Format(vRowEmailLog._Date, "dd MMM yyyy HH:mm")
                            .Cells(Me.ColDocNo3.Index).Value = vRowEmailLog.Doc_No
                            If vRowEmailLog.Status = 0 Then
                                .Cells(Me.ColStatus3.Index).Value = False
                            ElseIf vRowEmailLog.Status = 1 Then
                                .Cells(Me.ColStatus3.Index).Value = True
                            End If
                            .Cells(Me.ColPrimaryAccount3.Index).Value = vRowEmailLog.Primary_Email_Settings
                            .Cells(Me.ColPrimaryCust3.Index).Value = vRowEmailLog.Primary_Customer
                            .Cells(Me.ColPrimaryTransAcuan3.Index).Value = vRowEmailLog.Primary_Transaksi_Acuan
                            .Cells(Me.ColAction3.Index).Value = "View / Resend"
                            .Cells(Me.ColAccountCode3.Index).Value = vRowEmailLog.Account_Code

                            .Cells(Me.ColAccountName3.Index).Value = vRowEmailLog.Description
                            .Cells(Me.ColEmailFrom3.Index).Value = vRowEmailLog.Email_From
                            .Cells(Me.ColEmailTo3.Index).Value = vRowEmailLog.Email_Receiver.Replace(";", "")
                            .Cells(Me.ColCc3.Index).Value = vRowEmailLog.CC
                            .Cells(Me.ColSubject3.Index).Value = vRowEmailLog.Subject
                            .Cells(Me.ColMessage3.Index).Value = vRowEmailLog.Message
                            .Cells(Me.ColAttachmentFilename3.Index).Value = vRowEmailLog.Attachment_Filename

                        End With
                    Next
                    Me.objCd.SetNumber(Me.dgvData)
                    Me.btnPrint.Enabled = True
                Else
                    MsgBox("There is no log", MsgBoxStyle.Information, "Information")
                    Me.btnShow.Focus()
                    Me.btnPrint.Enabled = False
                    Exit Sub
                End If
            Else
                MsgBox("There is no log", MsgBoxStyle.Information, "Information")
                Me.btnShow.Focus()
                Me.btnPrint.Enabled = False
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvData_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvData.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvData_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellClick
        Dim frm As New FrmEmailSettings_LogReport_View
        Try
            If e.ColumnIndex = Me.ColAction3.Index And e.RowIndex <> -1 Then
                Me.Cursor = Cursors.WaitCursor
                If Me.dgvData.Item(Me.ColTransaction3.Index, e.RowIndex).Value.ToString = "Collection Letter" Then
                    frm.TipeTrans = ClsEmailSettings.EnumTipeTransaksiEmail.CollectionLetter
                ElseIf Me.dgvData.Item(Me.ColTransaction3.Index, e.RowIndex).Value.ToString = "Invoice" Then
                    frm.TipeTrans = ClsEmailSettings.EnumTipeTransaksiEmail.Invoice
                End If
                frm.PrimaryTransaksiAcuan = Me.dgvData.Item(Me.ColPrimaryTransAcuan3.Index, e.RowIndex).Value.ToString
                frm.DocNo = Me.dgvData.Item(Me.ColDocNo3.Index, e.RowIndex).Value.ToString
                frm.PrimaryCustomer = Me.dgvData.Item(Me.ColPrimaryCust3.Index, e.RowIndex).Value.ToString
                frm.UserCode = Me.UserCode
                frm.txtPrimaryAccount.Text = Me.dgvData.Item(Me.ColPrimaryAccount3.Index, e.RowIndex).Value.ToString
                frm.txtAccountCode.Text = Me.dgvData.Item(Me.ColAccountCode3.Index, e.RowIndex).Value.ToString
                frm.txtAccountName.Text = Me.dgvData.Item(Me.ColAccountName3.Index, e.RowIndex).Value.ToString
                frm.txtEmailFrom.Text = Me.dgvData.Item(Me.ColEmailFrom3.Index, e.RowIndex).Value.ToString
                frm.txtTo.Text = Me.dgvData.Item(Me.ColEmailTo3.Index, e.RowIndex).Value.ToString
                frm.txtCc.Text = Me.dgvData.Item(Me.ColCc3.Index, e.RowIndex).Value.ToString
                frm.txtSubject.Text = Me.dgvData.Item(Me.ColSubject3.Index, e.RowIndex).Value.ToString
                frm.txtMessage.Text = Me.dgvData.Item(Me.ColMessage3.Index, e.RowIndex).Value.ToString
                frm.vAttachmentFilename = Me.dgvData.Item(Me.ColAttachmentFilename3.Index, e.RowIndex).Value.ToString
                If My.Computer.FileSystem.FileExists(Me.dgvData.Item(Me.ColAttachmentFilename3.Index, e.RowIndex).Value.ToString) = False Then
                    MsgBox("The attachment file (" & My.Computer.FileSystem.GetName(Me.dgvData.Item(Me.ColAttachmentFilename3.Index, e.RowIndex).Value.ToString) & ") doesn't exist, you cannot resend this email.", MsgBoxStyle.Information, "Information")
                    Me.btnShow.Focus()
                    Exit Sub
                End If
                frm.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub dgvData_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvData.Sorted
        Try
            Me.objCd.SetNumber(Me.dgvData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim DS As New DseEmailSettings
        Dim vTblCetak As DseEmailSettings.DataTable4DataTable = DS.DataTable4
        Dim vRowCetak As DseEmailSettings.DataTable4Row
        Dim frmReportView As New FrmEmailSettings_LogReport_Print
        Dim report As New CrpEmailSettings_LogReport
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.vTblEmailLog.Rows.Count > 0 Then
                For i As Integer = 0 To Me.vTblEmailLog.DefaultView.Count - 1
                    Me.vRowEmailLog = Me.vTblEmailLog.DefaultView(i).Row
                    vRowCetak = vTblCetak.NewDataTable4Row
                    vRowCetak.No = i + 1
                    vRowCetak.Sender = Me.vRowEmailLog.Account_Code & " - " & Me.vRowEmailLog.Email_From
                    vRowCetak.Customer_Code = Me.vRowEmailLog.Customer_Code
                    vRowCetak.Customer_Name = Me.vRowEmailLog.Customer_Name
                    If vRowEmailLog.Tipe_Transaksi = ClsEmailSettings.EnumTipeTransaksiEmail.CollectionLetter Then
                        vRowCetak.Transaction = "Collection Letter"
                    ElseIf vRowEmailLog.Tipe_Transaksi = ClsEmailSettings.EnumTipeTransaksiEmail.Invoice Then
                        vRowCetak.Transaction = "Invoice"
                    End If
                    vRowCetak._Date = Me.vRowEmailLog._Date
                    vRowCetak.Doc_No = Me.vRowEmailLog.Doc_No
                    If Me.vRowEmailLog.Status = 0 Then
                        vRowCetak.Status = "No"
                    ElseIf Me.vRowEmailLog.Status = 1 Then
                        vRowCetak.Status = "Yes"
                    End If
                    vTblCetak.Rows.Add(vRowCetak)
                Next
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_P_PeridFrom.ParameterFieldName, Me.dtpFrom.Value)
                report.SetParameterValue(report.Parameter_P_PeridTo.ParameterFieldName, Me.dtpTo.Value)
                frmReportView.CrystalReportViewer1.ReportSource = report
                frmReportView.CrystalReportViewer1.Zoom(90)
                frmReportView.CrystalReportViewer1.Refresh()
                frmReportView.CrystalReportViewer1.Show()
                frmReportView.ShowDialog()
            Else
                MsgBox("There is no log, please change the filter", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frmReportView.Dispose()
            Me.Cursor = Cursors.Default
        End Try
    End Sub
End Class