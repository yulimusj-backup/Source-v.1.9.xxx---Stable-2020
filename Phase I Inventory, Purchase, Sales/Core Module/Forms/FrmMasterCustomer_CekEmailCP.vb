Option Explicit On
Option Strict Off

'form ini dipanggil ketika akan mengirimkan email secara sekaligus pada form surat tagihan2
Public Class FrmMasterCustomer_CekEmailCP
    Public ListPrimaryCust As New List(Of Double)
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCust As New Core_Module.clsMasterCustomer
    Dim buffer As String = ""
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ReferenceIndex As Integer = -1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterCustomer_CekEmailCP_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.CekApakahBenar = False Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCustomer_CekEmailCP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim PrimaryCustomer As Double = -1
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_itemcp1TableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_itemcp1DataTable
        Dim row, rowDummy As dseCustomer.tblmaster_customer_itemcp1Row
        Dim TBLDummy As New dseCustomer.tblmaster_customer_itemcp1DataTable
        Dim DV As DataRowView
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBLDummy)

            'load data customer + contact personnya
            For Each PrimaryCustomer In Me.ListPrimaryCust
                TBLDummy.DefaultView.RowFilter = "[Primary Customer] = " + PrimaryCustomer.ToString
                If TBLDummy.DefaultView.Count = 0 Then
                    'tambahkan baris kosong untuk customer ybs
                    row = TBL.Newtblmaster_customer_itemcp1Row
                    With row
                        .Customer_Code = Me.ObjCust.GetCustomerCode(PrimaryCustomer)
                        .Customer_Name = Me.ObjCust.GetCustName(PrimaryCustomer)
                        .Fax_No = ""
                        .Identifier = PrimaryCustomer.ToString + "-1"
                        .Name = ""
                        .Phone_No = ""
                        .Position = ""
                        .Primary = -1
                        .Primary_Customer = PrimaryCustomer
                        .Remark = ""
                        .Send_Collection_Letter = 1
                        ._E_mail_Address = ""
                    End With
                    TBL.Rows.Add(row)
                Else
                    For ij As Integer = 0 To TBLDummy.DefaultView.Count - 1
                        rowDummy = TBLDummy.DefaultView(ij).Row
                        row = TBL.Newtblmaster_customer_itemcp1Row
                        For jk As Integer = 0 To TBL.Columns.Count - 1
                            row.Item(jk) = rowDummy.Item(jk)
                        Next
                        TBL.Rows.Add(row)
                    Next
                End If
            Next

            'sort berdasarkan customer dan nama contact person
            TBL.DefaultView.Sort = "[Customer Code], Name"

            'tampilkan ke dalam dgv
            Me.dgvCP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvCP.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvCP.Rows.Clear()
            For Each DV In TBL.DefaultView
                row = DV.Row
                Me.dgvCP.Rows.Add()
                rowDGV = Me.dgvCP.Rows(Me.dgvCP.Rows.Count - 1)
                With row
                    rowDGV.Cells(Me.ColAdd.Index).Value = ""
                    rowDGV.Cells(Me.ColPerson_Customer.Index).Value = .Customer_Code + " - " + .Customer_Name
                    rowDGV.Cells(Me.ColPerson_PrimaryCust.Index).Value = row.Primary_Customer
                    rowDGV.Cells(Me.ColPerson_Remarks.Index).Value = .Remark
                    rowDGV.Cells(Me.ColPersonCode.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(Me.ColPersonEmail.Index).Value = ._E_mail_Address
                    rowDGV.Cells(Me.ColPersonFax.Index).Value = .Fax_No
                    rowDGV.Cells(Me.ColPersonIdentifier.Index).Value = .Identifier
                    rowDGV.Cells(Me.ColPersonName.Index).Value = .Name
                    rowDGV.Cells(Me.ColPersonPhoneNumber.Index).Value = .Phone_No
                    rowDGV.Cells(Me.ColPersonPosition.Index).Value = .Position
                    If .Send_Collection_Letter = 1 Then
                        rowDGV.Cells(Me.ColPersonSendCL.Index).Value = Boolean.TrueString
                    Else
                        rowDGV.Cells(Me.ColPersonSendCL.Index).Value = Boolean.FalseString
                    End If
                    rowDGV.Cells(Me.ColPersonPrimaryItem.Index).Value = .Primary
                End With
            Next
            Me.dgvCP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.dgvCP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            TBLDummy.Dispose()
        End Try
    End Sub

    Private Function CekApakahBenar() As Boolean
        Dim row As DataGridViewRow
        Dim cell As DataGridViewCell
        Dim ListNamaCustomer As New List(Of String)
        Dim DummyCust As String = ""
        Dim AdaSendCollectionLetter As Boolean = True

        Try
            'Untuk item contact person:
            '1. tidak boleh ada nama contact person yang kosong
            '2. jika send collection letter, maka pastikan emailnya tepat
            '3. untuk setiap customer, harus ada minimal 1 yang send collection letter
            Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
            For Each row In Me.dgvCP.Rows
                With row
                    For Each cell In row.Cells
                        If IsNothing(cell.Value) = True Then
                            cell.Value = ""
                        End If
                    Next
                    If .Cells(1).Value = "" Then
                        'MsgBox("Wrong contact person's name, please fill in contact person's name data.", _
                        'MsgBoxStyle.Information, "Information")
                        'Me.dgvCP.ClearSelection()
                        '.Selected = True
                        'Return False
                        If ListNamaCustomer.Contains(row.Cells(Me.ColPerson_Customer.Index).Value) = False Then
                            ListNamaCustomer.Add(row.Cells(Me.ColPerson_Customer.Index).Value)
                        End If
                    End If

                    'utnuk setiap item contact person, pastikan email yang diberikan tepat
                    If .Cells(Me.ColPersonSendCL.Index).Value = Boolean.TrueString Then
                        .Cells(Me.ColPersonEmail.Index).Value = .Cells(Me.ColPersonEmail.Index).Value.ToString.Trim
                        If Expression.IsMatch(.Cells(Me.ColPersonEmail.Index).Value.ToString) = False Then
                            'MsgBox("The email address is not valid for the following contact person: " + .Cells(Me.ColPersonName.Index).Value.ToString + _
                            '". Please check again.", MsgBoxStyle.Information, "Information")
                            'Me.TabControl1.SelectedIndex = 1
                            'Me.dgvCP.ClearSelection()
                            'row.Selected = True

                            'Return False
                            If ListNamaCustomer.Contains(row.Cells(Me.ColPerson_Customer.Index).Value) = False Then
                                ListNamaCustomer.Add(row.Cells(Me.ColPerson_Customer.Index).Value)
                            End If
                        End If
                    End If

                    If row.Cells(Me.ColPerson_Customer.Index).Value <> DummyCust Then
                        If DummyCust = "" Then
                            AdaSendCollectionLetter = False
                            DummyCust = row.Cells(Me.ColPerson_Customer.Index).Value
                        Else
                            If AdaSendCollectionLetter = False Then
                                If ListNamaCustomer.Contains(DummyCust) = False Then
                                    ListNamaCustomer.Add(DummyCust)
                                End If
                            End If
                            DummyCust = row.Cells(Me.ColPerson_Customer.Index).Value
                            AdaSendCollectionLetter = False
                        End If
                        If row.Cells(Me.ColPersonSendCL.Index).Value = Boolean.TrueString Then
                            AdaSendCollectionLetter = True
                        End If
                    Else
                        If row.Cells(Me.ColPersonSendCL.Index).Value = Boolean.TrueString Then
                            AdaSendCollectionLetter = True
                        End If
                    End If
                    If row.Index = Me.dgvCP.RowCount - 1 Then
                        If row.Cells(Me.ColPersonSendCL.Index).Value = Boolean.TrueString Then
                            AdaSendCollectionLetter = True
                        End If
                        If AdaSendCollectionLetter = False Then
                            If ListNamaCustomer.Contains(DummyCust) = False Then
                                ListNamaCustomer.Add(DummyCust)
                            End If
                        End If
                    End If
                End With
            Next
            Expression = Nothing

            If ListNamaCustomer.Count > 0 Then
                DummyCust = ""
                For Each Customer As String In ListNamaCustomer
                    If DummyCust = "" Then
                        DummyCust = Customer
                    Else
                        DummyCust = DummyCust + ", " + vbCrLf + Customer
                    End If
                Next
                If MsgBox("The following customer have invalid data for sending collection letter:" + vbCrLf + DummyCust + "." + vbCrLf + vbCrLf + _
                "Please check that each customer comply the following criteria:" + vbCrLf + _
                "1. Make sure each customer has at least 1 valid contact person to send collection letter" + vbCrLf + _
                "2. Make sure the email address is valid." + vbCrLf + vbCrLf + _
                "Ignore this exception and close?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListNamaCustomer = Nothing
        End Try
    End Function

    Private Sub dgvCP_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCP.CellBeginEdit
        Try
            'yang bisa diedit: semua kecuali nama customer dan nomor
            If e.ColumnIndex = Me.ColPerson_Customer.Index Or e.ColumnIndex = Me.ColPersonCode.Index Then
                e.Cancel = True
            Else
                'jika nama cp-nya kosong, yang pertama kali boleh diedit adalah nama cp-nya
                If IsNothing(Me.dgvCP.Item(Me.ColPersonName.Index, e.RowIndex).Value) = True Then
                    If e.ColumnIndex <> Me.ColPersonName.Index Then
                        e.Cancel = True
                    End If
                ElseIf Me.dgvCP.Item(Me.ColPersonName.Index, e.RowIndex).Value.ToString.Trim = "" Then
                    If e.ColumnIndex <> Me.ColPersonName.Index Then
                        e.Cancel = True
                    End If
                Else
                    Me.buffer = Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCP_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCP.CellEndEdit
        Try
            If IsNothing(Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColPersonName.Index Then
                    Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                ElseIf e.ColumnIndex = Me.ColPersonSendCL.Index Then
                    Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = CBool(Boolean.FalseString)
                ElseIf e.ColumnIndex = Me.ColPersonEmail.Index Then
                    If Me.dgvCP.Item(Me.ColPersonSendCL.Index, e.RowIndex).Value = Boolean.TrueString Then
                        Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                Else
                    Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
                If Me.dgvCP.Item(Me.ColPersonName.Index, e.RowIndex).Value <> "" Then
                    Me.ReferenceIndex = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            Else
                Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                If e.ColumnIndex = Me.ColPersonName.Index Then
                    'untuk ItemCP nama tidak boleh kosong
                    If Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        MsgBox("Contact person name cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColPersonEmail.Index Then
                    Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                        Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                        If Expression.IsMatch(Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                            Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Expression = Nothing
                            Exit Sub
                        End If
                        Expression = Nothing
                    End If
                End If
                If Me.dgvCP.Item(Me.ColPersonName.Index, e.RowIndex).Value <> "" Then
                    Me.ReferenceIndex = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateKeDB1(ByVal vRowIndex As Integer)
        Dim rowDGV As DataGridViewRow
        Dim SimpanBaru As Boolean = True
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim SendCL As Integer = 0

        Try
            rowDGV = Me.dgvCP.Rows(vRowIndex)
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If rowDGV.Cells(Me.ColPersonName.Index).Value.ToString.Trim <> "" Then
                If rowDGV.Cells(Me.ColPersonPrimaryItem.Index).Value = -1 Then
                    'simpan baru
                    SimpanBaru = True
                Else
                    'cek apakah sudah ada didalam database primarynya/ berdasarkan kombinasi nama contact dan primary customer
                    If Me.ObjCust.ApakahSudahAda_CP(rowDGV.Cells(Me.ColPersonPrimaryItem.Index).Value, rowDGV.Cells(Me.ColPersonName.Index).Value, _
                    rowDGV.Cells(Me.ColPerson_PrimaryCust.Index).Value) = True Then
                        'simpan perubahan, bila ditemukan pada database
                        SimpanBaru = False
                    Else
                        'simpan baru jika tidak ditemukan pada database
                        SimpanBaru = True
                    End If
                End If

                If rowDGV.Cells(Me.ColPersonSendCL.Index).Value = Boolean.TrueString Then
                    SendCL = 1
                Else
                    SendCL = 0
                End If
                If SimpanBaru = True Then
                    'simpan baru
                    'insert
                    DTA.InsertQuery(rowDGV.Cells(Me.ColPerson_PrimaryCust.Index).Value, rowDGV.Cells(Me.ColPersonName.Index).Value, _
                    rowDGV.Cells(Me.ColPersonPosition.Index).Value, rowDGV.Cells(Me.ColPersonPhoneNumber.Index).Value, _
                    rowDGV.Cells(Me.ColPersonFax.Index).Value, rowDGV.Cells(Me.ColPersonEmail.Index).Value, SendCL, _
                    rowDGV.Cells(Me.ColPerson_Remarks.Index).Value, rowDGV.Cells(Me.ColPersonIdentifier.Index).Value)
                    'update primary
                    rowDGV.Cells(Me.ColPersonPrimaryItem.Index).Value = DTA.ScalarQuery_GetPrimary(rowDGV.Cells(Me.ColPersonIdentifier.Index).Value)
                Else
                    'simpan perubahan
                    DTA.UpdateQuery(rowDGV.Cells(Me.ColPersonName.Index).Value, _
                    rowDGV.Cells(Me.ColPersonPosition.Index).Value, rowDGV.Cells(Me.ColPersonPhoneNumber.Index).Value, _
                    rowDGV.Cells(Me.ColPersonFax.Index).Value, rowDGV.Cells(Me.ColPersonEmail.Index).Value, _
                    rowDGV.Cells(Me.ColPerson_Remarks.Index).Value, SendCL, rowDGV.Cells(Me.ColPersonPrimaryItem.Index).Value, _
                    rowDGV.Cells(Me.ColPersonIdentifier.Index).Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            DTA.Dispose()
        End Try
    End Sub

    Private Sub dgvCP_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCP.CellContentClick
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_itemcp1TableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_itemcp1DataTable
        Dim row As dseCustomer.tblmaster_customer_itemcp1Row
        Dim PrimaryCust As Double = -1
        Dim IndexAcuan As Double = -1
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If e.ColumnIndex = Me.ColAdd.Index And e.RowIndex <> -1 Then
                PrimaryCust = Me.dgvCP.Item(Me.ColPerson_PrimaryCust.Index, e.RowIndex).Value

                Dim ij As New Core_Module.FrmMasterCustomer
                ij.FromCekEmailCP = True
                ij.vPrimaryCustCekEmailCP = PrimaryCust
                ij.ShowDialog()
                ij = Nothing

                'isikan ulang data contact person untuk customer ybs
                DTA.FillBy_PrimaryCust(TBL, PrimaryCust)
                For jk As Double = Me.dgvCP.Rows.Count - 1 To 0 Step -1
                    rowDGV = Me.dgvCP.Rows(jk)
                    If rowDGV.Cells(Me.ColPerson_PrimaryCust.Index).Value = PrimaryCust Then
                        IndexAcuan = rowDGV.Index
                        Me.dgvCP.Rows.Remove(rowDGV)
                    End If
                Next
                If TBL.Rows.Count = 0 Then
                    'tambahkan baris kosong
                    row = TBL.Newtblmaster_customer_itemcp1Row
                    With row
                        .Customer_Code = Me.ObjCust.GetCustomerCode(PrimaryCust)
                        .Customer_Name = Me.ObjCust.GetCustName(PrimaryCust)
                        .Fax_No = ""
                        .Identifier = PrimaryCust.ToString + "-1"
                        .Name = ""
                        .Phone_No = ""
                        .Position = ""
                        .Primary = -1
                        .Primary_Customer = PrimaryCust
                        .Remark = ""
                        .Send_Collection_Letter = 1
                        ._E_mail_Address = ""
                    End With
                    TBL.Rows.Add(row)
                End If

                For Each row In TBL.Rows
                    Me.dgvCP.Rows.Insert(CInt(IndexAcuan), 1)
                    rowDGV = Me.dgvCP.Rows(IndexAcuan)
                    With row
                        rowDGV.Cells(Me.ColAdd.Index).Value = ""
                        rowDGV.Cells(Me.ColPerson_Customer.Index).Value = .Customer_Code + " - " + .Customer_Name
                        rowDGV.Cells(Me.ColPerson_PrimaryCust.Index).Value = row.Primary_Customer
                        rowDGV.Cells(Me.ColPerson_Remarks.Index).Value = .Remark
                        rowDGV.Cells(Me.ColPersonCode.Index).Value = rowDGV.Index + 1
                        rowDGV.Cells(Me.ColPersonEmail.Index).Value = ._E_mail_Address
                        rowDGV.Cells(Me.ColPersonFax.Index).Value = .Fax_No
                        rowDGV.Cells(Me.ColPersonIdentifier.Index).Value = .Identifier
                        rowDGV.Cells(Me.ColPersonName.Index).Value = .Name
                        rowDGV.Cells(Me.ColPersonPhoneNumber.Index).Value = .Phone_No
                        rowDGV.Cells(Me.ColPersonPosition.Index).Value = .Position
                        If .Send_Collection_Letter = 1 Then
                            rowDGV.Cells(Me.ColPersonSendCL.Index).Value = Boolean.TrueString
                        Else
                            rowDGV.Cells(Me.ColPersonSendCL.Index).Value = Boolean.FalseString
                        End If
                        rowDGV.Cells(Me.ColPersonPrimaryItem.Index).Value = .Primary
                    End With
                    IndexAcuan += 1
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub dgvCP_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvCP.UserDeletingRow
        'tidak bisa delete item cp, kalo mau delete, langsung edit detailnya via form master customer saja
        e.Cancel = True
    End Sub

    Private Sub dgvCP_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCP.Sorted
        Try
            ObjCD.SetNumber(Me.dgvCP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.ReferenceIndex <> -1 Then
                Me.UpdateKeDB1(Me.ReferenceIndex)
                Me.ReferenceIndex = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class