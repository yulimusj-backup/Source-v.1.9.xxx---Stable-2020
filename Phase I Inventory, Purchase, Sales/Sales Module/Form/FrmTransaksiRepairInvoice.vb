Option Explicit On
Option Strict Off

Public Class FrmTransaksiRepairInvoice
    Dim DTA As New DseTransaksiRepairSalesPersonTableAdapters.DataTable1TableAdapter
    Dim TBL As New DseTransaksiRepairSalesPerson.DataTable1DataTable
    Dim row As DseTransaksiRepairSalesPerson.DataTable1Row
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Buffer As String = ""
    Dim ObjSalesP As New Core_Module.ClsMasterSalesPerson
    Dim ObjCust As New Core_Module.clsMasterCustomer
    Dim DTA1 As New DseTransaksiRepairSalesPersonTableAdapters.DataTable2TableAdapter
    Dim TBL1 As New DseTransaksiRepairSalesPerson.DataTable2DataTable
    Dim row1 As DseTransaksiRepairSalesPerson.DataTable2Row

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""

    Private Sub btnInvRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvRefresh.Click
        Try
            If MsgBox("Reload sales order data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.ReloadDataSO()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ReloadDataSO()
        Dim rowDGV As DataGridViewRow

        Try
            TBL.Rows.Clear()
            TBL1.Rows.Clear()

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.Fill(TBL, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpInvPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpInvPeriod.Value))
            DTA1.Fill(TBL1, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpInvPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpInvPeriod.Value))

            'remove row yang sudah ada sales personnya jika ditentukan demikian
            If Me.chkInvShowAll.Checked = False Then
                For ij As Integer = TBL.Rows.Count - 1 To 0 Step -1
                    row = TBL.Rows(ij)
                    If row.Sales_Person_Code <> "" Then
                        'remove row
                        TBL.Rows.RemoveAt(ij)
                    End If
                Next
                For ij As Integer = TBL1.Rows.Count - 1 To 0 Step -1
                    row1 = TBL1.Rows(ij)
                    If row1.Sales_Person_Code <> "" Then
                        'remove row
                        TBL1.Rows.RemoveAt(ij)
                    End If
                Next
            End If

            'pindahkan baris sr ke dalam baris invoice
            For Each row1 In TBL1.Rows
                TBL.DefaultView.RowFilter = "[Primary SO] = " + row1.Primary_SO.ToString
                If TBL.DefaultView.Count = 0 Then
                    row = TBL.NewDataTable1Row
                    With row
                        .Customer_Code = row1.Customer_Code
                        .Customer_Name = row1.Customer_Name
                        .Primary_Customer = row1.Primary_Customer
                        .primary_sales_person = row1.primary_sales_person
                        .Primary_SO = row1.Primary_SO
                        .Sales_Person_Code = row1.Sales_Person_Code
                        .Sales_Person_Name = row1.Sales_Person_Name
                        .SO_Date = row1.SO_Date
                        .So_Number = row1.SO_Number
                    End With
                    TBL.Rows.Add(row)
                End If
            Next

            'tampilkan ke dalam dgv
            Me.dgvSO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvSO.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvSO.Rows.Clear()
            For Each row In Me.TBL.Rows
                Me.dgvSO.Rows.Add()
                rowDGV = Me.dgvSO.Rows(Me.dgvSO.Rows.Count - 1)
                rowDGV.Cells(Me.ColSO_CustCode.Index).Value = row.Customer_Code
                rowDGV.Cells(Me.ColSO_CustName.Index).Value = row.Customer_Name
                rowDGV.Cells(Me.ColSO_Date.Index).Value = row.SO_Date
                rowDGV.Cells(Me.ColSO_No.Index).Value = Me.dgvSO.Rows.Count
                rowDGV.Cells(Me.ColSO_PrimaryCust.Index).Value = row.Primary_Customer
                rowDGV.Cells(Me.ColSO_PrimarySalesPerson.Index).Value = row.primary_sales_person
                rowDGV.Cells(Me.ColSO_SalesPersonCode.Index).Value = row.Sales_Person_Code
                rowDGV.Cells(Me.ColSO_SalesPersonName.Index).Value = row.Sales_Person_Name
                rowDGV.Cells(Me.ColSO_SONo.Index).Value = row.So_Number
                rowDGV.Cells(Me.ColSO_PrimarySO.Index).Value = row.Primary_SO
            Next
            Me.ColSO_Date.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.dgvSO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub dgvSO_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSO.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColSO_SalesPersonCode.Index Then
                Me.Buffer = Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSO_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSO.CellEndEdit
        Try
            If IsNothing(Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColSO_SalesPersonCode.Index Then
                    'kalo kode sales person dihapus, maka kembalikan ke kondisi buffer
                    Me.dgvSO.Item(Me.ColSO_SalesPersonCode.Index, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            Else
                If e.ColumnIndex = Me.ColSO_SalesPersonCode.Index Then
                    'code
                    'pastikan ditrim dan huruf besar, max 250 karakter
                    Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'cek apakah ada ada dalam master sales person
                    'kemungkinan
                    If Me.ObjSalesP.ApakahSudahAda(Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode sales person ada  --> load data nama + primary
                        Me.dgvSO.Item(Me.ColSO_PrimarySalesPerson.Index, e.RowIndex).Value = Me.ObjSalesP.GetPrimary(Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvSO.Item(Me.ColSO_SalesPersonName.Index, e.RowIndex).Value = Me.ObjSalesP.GetName(Me.dgvSO.Item(Me.ColSO_PrimarySalesPerson.Index, e.RowIndex).Value)
                    Else
                        '2. kode sales person yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmMasterSalesPerson_Cari
                        ij.txtCari.Text = Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimarySalesPerson <> -1 Then
                            '2.a. sales person yang dicari ada --> load data deskripsi + primary + kode
                            Me.dgvSO.Item(Me.ColSO_PrimarySalesPerson.Index, e.RowIndex).Value = ij.PrimarySalesPerson
                            Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjSalesP.GetCode(ij.PrimarySalesPerson)
                            Me.dgvSO.Item(Me.ColSO_SalesPersonName.Index, e.RowIndex).Value = Me.ObjSalesP.GetName(ij.PrimarySalesPerson)
                        ElseIf ij.PrimarySalesPerson = -1 Then
                            '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvSO.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.dgvSO.Item(Me.ColSO_SalesPersonCode.Index, e.RowIndex).Value = Me.Buffer
                                ij = Nothing
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                    'kalo sampe sini update
                    Me.DTA.UpdateQuery(Me.dgvSO.Item(Me.ColSO_PrimarySalesPerson.Index, e.RowIndex).Value, _
                    Me.dgvSO.Item(Me.ColSO_PrimarySO.Index, e.RowIndex).Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiRepairInvoice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dtpInvPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpInvPeriod.Value)
            Me.dtpCust_Period.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpCust_Period.Value)
            Me.ReloadDataSO()
            Me.ReloadDataCust()
            Me.dgvCustomer.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ReloadDataCust()
        Dim rowDGV As DataGridViewRow
        Dim ListPrimaryCust As New List(Of Double)
        Dim PrimarySalesP As Double = -1
        Dim PrimaryCust As Double = -1

        Try
            TBL.Rows.Clear()
            TBL1.Rows.Clear()

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.Fill(TBL, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpInvPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpInvPeriod.Value))
            DTA1.Fill(TBL1, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpInvPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpInvPeriod.Value))

            'remove row yang sudah ada sales personnya jika ditentukan demikian
            If Me.chkInvShowAll.Checked = False Then
                For ij As Integer = TBL.Rows.Count - 1 To 0 Step -1
                    row = TBL.Rows(ij)
                    If row.Sales_Person_Code <> "" Then
                        'remove row
                        TBL.Rows.RemoveAt(ij)
                    End If
                Next
                For ij As Integer = TBL1.Rows.Count - 1 To 0 Step -1
                    row1 = TBL1.Rows(ij)
                    If row1.Sales_Person_Code <> "" Then
                        'remove row
                        TBL1.Rows.RemoveAt(ij)
                    End If
                Next
            End If

            'pindahkan baris sr ke dalam baris invoice
            For Each row1 In TBL1.Rows
                TBL.DefaultView.RowFilter = "[Primary SO] = " + row1.Primary_SO.ToString
                If TBL.DefaultView.Count = 0 Then
                    row = TBL.NewDataTable1Row
                    With row
                        .Customer_Code = row1.Customer_Code
                        .Customer_Name = row1.Customer_Name
                        .Primary_Customer = row1.Primary_Customer
                        .primary_sales_person = row1.primary_sales_person
                        .Primary_SO = row1.Primary_SO
                        .Sales_Person_Code = row1.Sales_Person_Code
                        .Sales_Person_Name = row1.Sales_Person_Name
                        .SO_Date = row1.SO_Date
                        .So_Number = row1.SO_Number
                    End With
                    TBL.Rows.Add(row)
                End If
            Next

            'tampilkan ke dalam dgv
            Me.dgvCustomer.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvCustomer.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvCustomer.Rows.Clear()
            For Each row In Me.TBL.Rows
                'cek apakah sudah ada primary sales person untuk cust ybs
                If row.primary_sales_person = -1 Then
                    If ListPrimaryCust.Contains(row.Primary_Customer) = False Then
                        ListPrimaryCust.Add(row.Primary_Customer)
                    End If
                End If
            Next

            For Each PrimaryCust In ListPrimaryCust
                Me.dgvCustomer.Rows.Add()
                rowDGV = Me.dgvCustomer.Rows(Me.dgvCustomer.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColCust_CustCode.Index).Value = Me.ObjCust.GetCustomerCode(PrimaryCust)
                    .Cells(Me.ColCust_Customer.Index).Value = Me.ObjCust.GetCustName(PrimaryCust)
                    .Cells(Me.ColCust_No.Index).Value = ""
                    .Cells(Me.ColCust_PrimaryCust.Index).Value = PrimaryCust
                    .Cells(Me.ColCust_PrimarySalesPerson.Index).Value = -1
                    .Cells(Me.ColCust_SalesPersonCode.Index).Value = ""
                    .Cells(Me.ColCust_SalesPersonName.Index).Value = ""
                End With
            Next
            Me.dgvCustomer.Sort(Me.ColCust_CustCode, System.ComponentModel.ListSortDirection.Ascending)
            Me.ObjCD.SetNumber(Me.dgvCustomer)
            Me.dgvCustomer.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            ListPrimaryCust = Nothing
        End Try
    End Sub

    Private Sub brnReloadCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles brnReloadCust.Click
        Try
            If MsgBox("Reload customer data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.ReloadDataCust()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedTab.Name = Me.TabPage_Customer.Name Then
                Me.ReloadDataCust()
            ElseIf Me.TabControl1.SelectedTab.Name = Me.TabPage_SO.Name Then
                Me.ReloadDataSO()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCustomer_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCustomer.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCust_SalesPersonCode.Index Then
                Me.Buffer = Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCustomer_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCustomer.CellEndEdit
        Try
            If IsNothing(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColCust_SalesPersonCode.Index Then
                    'kalo kode sales person dihapus, maka kembalikan ke kondisi buffer
                    Me.dgvCustomer.Item(Me.ColCust_SalesPersonCode.Index, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            Else
                If e.ColumnIndex = Me.ColCust_SalesPersonCode.Index Then
                    'code
                    'pastikan ditrim dan huruf besar, max 250 karakter
                    Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'cek apakah ada ada dalam master sales person
                    'kemungkinan
                    If Me.ObjSalesP.ApakahSudahAda(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode sales person ada  --> load data nama + primary
                        Me.dgvCustomer.Item(Me.ColCust_PrimarySalesPerson.Index, e.RowIndex).Value = Me.ObjSalesP.GetPrimary(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvCustomer.Item(Me.ColCust_SalesPersonName.Index, e.RowIndex).Value = Me.ObjSalesP.GetName(Me.dgvCustomer.Item(Me.ColCust_PrimarySalesPerson.Index, e.RowIndex).Value)
                    Else
                        '2. kode sales person yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmMasterSalesPerson_Cari
                        ij.txtCari.Text = Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimarySalesPerson <> -1 Then
                            '2.a. sales person yang dicari ada --> load data deskripsi + primary + kode
                            Me.dgvCustomer.Item(Me.ColCust_PrimarySalesPerson.Index, e.RowIndex).Value = ij.PrimarySalesPerson
                            Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjSalesP.GetCode(ij.PrimarySalesPerson)
                            Me.dgvCustomer.Item(Me.ColCust_SalesPersonName.Index, e.RowIndex).Value = Me.ObjSalesP.GetName(ij.PrimarySalesPerson)
                        ElseIf ij.PrimarySalesPerson = -1 Then
                            '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.dgvCustomer.Item(Me.ColCust_SalesPersonCode.Index, e.RowIndex).Value = Me.Buffer
                                ij = Nothing
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                    'kalo sampe sini update
                    Me.DTA.UpdateQuery_ByPrimaryCust(Me.dgvCustomer.Item(Me.ColCust_PrimarySalesPerson.Index, e.RowIndex).Value, _
                    Me.dgvCustomer.Item(Me.ColCust_PrimaryCust.Index, e.RowIndex).Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class