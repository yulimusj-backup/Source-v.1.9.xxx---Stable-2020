Option Explicit On
Option Strict Off

Public Class FrmLaporanCustomerList
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCD As New Core_Module.ClsCekData
    Dim objCust As New clsMasterCustomer
    Dim ObjCustGroup As New clsMasterCustomerGroup
    Dim ObjCustType As New clsMasterCustomerType
    Dim ObjCustSalesArea As New clsMasterCustomerSalesArea
    Dim ObjCustSegmentation As New clsMasterCustomerMarketSegmentation
    Dim Buffer As String = ""
    Dim index As Integer = -1

    Enum EnumGroup As Integer
        Customer_Group
        Sales_Area
        Customer_Segmentation
        Customer_Type
        Customer_Reference
    End Enum

    Private Sub FrmLaporanSupplierList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.DgvCustGroup.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvSalesArea.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvMarketSegmentation.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvCustRef.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvCustomerType.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.CboGroup.SelectedIndex = EnumGroup.Customer_Group
            Me.ChkJoinDate1.Checked = True
            Me.ChkJoinDate1.Checked = False
            Me.ColMS_Primary.Visible = False
            Me.Text = "Customer List Report"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShow.Click
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Try
            Me.Cetak()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub Cetak()
        Dim Ds As New DseLaporanCustomerList
        Dim Dta As New DseLaporanCustomerListTableAdapters.tblmaster_customerTableAdapter
        Dim DtaAdd As New DseLaporanCustomerListTableAdapters.tblmaster_item_addressTableAdapter
        Dim DtaCP As New DseLaporanCustomerListTableAdapters.tblmaster_item_cpTableAdapter
        Dim DtaBank As New DseLaporanCustomerListTableAdapters.tblmaster_item_bankTableAdapter
        Dim Tbl As New DseLaporanCustomerList.tblmaster_customerDataTable
        Dim TblAdd As New DseLaporanCustomerList.tblmaster_item_addressDataTable
        Dim TblCP As New DseLaporanCustomerList.tblmaster_item_cpDataTable
        Dim TblBank As New DseLaporanCustomerList.tblmaster_item_bankDataTable
        Dim TblC As DseLaporanCustomerList.tbl_cetakDataTable = Ds.tbl_cetak
        Dim Row As DseLaporanCustomerList.tblmaster_customerRow
        Dim RowAdd As DseLaporanCustomerList.tblmaster_item_addressRow
        Dim RowCP As DseLaporanCustomerList.tblmaster_item_cpRow
        Dim RowBank As DseLaporanCustomerList.tblmaster_item_bankRow
        Dim RowC As DseLaporanCustomerList.tbl_cetakRow
        Dim count As Integer

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            DtaAdd.Connection.ConnectionString = Core_Module.GetConstring
            DtaCP.Connection.ConnectionString = Core_Module.GetConstring
            DtaBank.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ChkJoinDate1.Checked = True Then
                Dta.FillBy(Tbl, Me.DtpFrom.Value, Me.DtpTo.Value)
            Else
                Dta.Fill(Tbl)
            End If

            If Me.ChkMultipleAddress.Checked = True Then
                DtaAdd.Fill(TblAdd) 'fill semua item address
            Else
                DtaAdd.Fill_2(TblAdd) 'fill item address hanya baris pertama pada tiap customer
            End If
            If Me.ChkMultipleCP.Checked = True Then
                DtaCP.Fill(TblCP) 'fill semua item contact person
            Else
                DtaCP.Fill_2(TblCP) 'fill item contact person hanya baris pertama pada tiap customer
            End If
            If Me.ChkMultipleBank.Checked = True Then
                DtaBank.Fill(TblBank) 'fill semua item bank
            Else
                DtaBank.Fill_2(TblBank) 'fill item bank hanya baris pertama pada tiap customer
            End If

            'filter customer berdasarkan filter di dgvfilter
            FilterFromDgv(Tbl)

            For Each Row In Tbl.Rows
                With Row
                    TblAdd.DefaultView.RowFilter = "[Primary Customer] = " & .Primary
                    TblCP.DefaultView.RowFilter = "[Primary Customer] = " & .Primary
                    TblBank.DefaultView.RowFilter = "[Primary Customer] = " & .Primary
                    If TblAdd.DefaultView.Count > TblCP.DefaultView.Count Then
                        count = TblAdd.DefaultView.Count
                    Else
                        count = TblCP.DefaultView.Count
                    End If
                    If count < TblBank.DefaultView.Count Then
                        count = TblBank.DefaultView.Count
                    End If

                    For i As Integer = 0 To count - 1
                        RowC = TblC.Newtbl_cetakRow
                        'header
                        RowC.Primary_Customer = .Primary
                        RowC.Customer_Code = .Customer_Code
                        RowC.Customer_Name = .Customer_Name
                        If i = 0 Then
                            RowC.NPWP_No = .NPWP_No
                            RowC.Tax_Address = "-"
                            If Me.ChkTaxData.Checked = True Then
                                Dim Array_() As String = {"", "Blok ", "No ", "RT ", "RW ", "Kel ", "Kec ", "Kab ", "Propinsi "}
                                Dim Tax_AddressArray() As String = {.Nama_Jalan, .Blok, .Nomor, .RT, .RW, .Kelurahan, .Kecamatan, .Kabupaten, .Propinsi}
                                For i2 As Integer = 0 To Tax_AddressArray.Length - 1
                                    If Tax_AddressArray(i2) <> "" And Tax_AddressArray(i2) <> "0" And Tax_AddressArray(i2) <> "-" Then
                                        If RowC.Tax_Address = "-" Then
                                            RowC.Tax_Address = Array_(i2).ToString & Tax_AddressArray(i2).ToString
                                        Else
                                            RowC.Tax_Address &= ", " & Array_(i2).ToString & Tax_AddressArray(i2).ToString
                                        End If
                                    End If
                                Next
                            End If
                        Else
                            RowC.NPWP_No = "-"
                            RowC.Tax_Address = "-"
                        End If
                        'item address 
                        RowC.Address = "-"
                        'item cp
                        RowC.Contact_Person = "-"
                        'item bank
                        RowC._Bank_s_Account = "-"
                        RowC._Bank_s_Name = "-"
                        RowC.Beneficiary_Name = "-"
                        RowC.Currency_Code = ""
                        If .Code_Ref = "" Then
                            RowC.Code_Reference = "-"
                            RowC.Join_Date = "-"
                        Else
                            RowC.Code_Reference = .Code_Ref
                            RowC.Join_Date = Format(.Join_Date, "dd MMM yy")
                        End If
                        'berkaitan group
                        If Me.CboGroup.SelectedIndex = EnumGroup.Customer_Group Then
                            RowC.Group = .Customer_Group_Code & " - " & .Customer_Group_Desc
                        ElseIf Me.CboGroup.SelectedIndex = EnumGroup.Sales_Area Then
                            RowC.Group = .Sales_Area_Code & " - " & .Sales_Area_Desc
                        ElseIf Me.CboGroup.SelectedIndex = EnumGroup.Customer_Segmentation Then
                            RowC.Group = .Segmentation_Code & " - " & .Segmentation_Desc
                        ElseIf Me.CboGroup.SelectedIndex = EnumGroup.Customer_Type Then
                            RowC.Group = .Type_Code & " - " & .Type_Desc
                        ElseIf Me.CboGroup.SelectedIndex = EnumGroup.Customer_Reference Then
                            RowC.Group = .Code_Ref & " - " & .Desc_Ref
                        Else : RowC.Group = "-"
                        End If
                        RowC.Count = count
                        TblC.Rows.Add(RowC)

                        'masukin nilai item address, untuk baris ke i
                        If TblAdd.DefaultView.Count > i Then
                            RowAdd = TblAdd.DefaultView(i).Row
                            Dim AddressArray() As String = {RowAdd.Line_1, RowAdd.Line_2, RowAdd.Line_3, RowAdd.Line_4, RowAdd.Line_5, RowAdd.Phone_No, RowAdd.Fax_No}
                            For i3 As Integer = 0 To AddressArray.Length - 1
                                If AddressArray(i3) <> "" Then
                                    If RowC.Address = "-" Then
                                        RowC.Address = AddressArray(i3).ToString
                                    Else
                                        RowC.Address &= ", " & AddressArray(i3).ToString
                                    End If
                                End If
                            Next
                        End If

                        'masukin nilai item contact person, untuk baris ke i
                        If TblCP.DefaultView.Count > i Then
                            RowCP = TblCP.DefaultView(i).Row
                            RowC.Contact_Person = RowCP.Name
                        End If

                        'masukin nilai item contact person, untuk baris ke i
                        If TblBank.DefaultView.Count > i Then
                            RowBank = TblBank.DefaultView(i).Row
                            RowC._Bank_s_Account = RowBank.Account_No
                            RowC._Bank_s_Name = RowBank._Bank_s_Name
                            If RowBank.Beneficiary_Name <> "" Then
                                RowC.Beneficiary_Name = RowBank.Beneficiary_Name
                            End If
                            RowC.Currency_Code = " (" & RowBank.Currency_Code & ")"
                        End If
                    Next
                End With
            Next

            'Sorting(TblC)
            '
            Dim tblDummy As New DseLaporanCustomerList.tbl_cetakDataTable
            Dim rowDummy As DseLaporanCustomerList.tbl_cetakRow
            Dim Sort As String = ""
            Dim List As New List(Of Double)
            TblC.DefaultView.Sort = "[Group] ASC, [Customer Code] ASC"
            For i As Integer = 0 To TblC.DefaultView.Count - 1
                RowC = TblC.DefaultView(i).Row
                rowDummy = tblDummy.NewRow
                With rowDummy
                    .Primary_Customer = RowC.Primary_Customer
                    .Customer_Code = RowC.Customer_Code
                    .Customer_Name = RowC.Customer_Name
                    .Address = RowC.Address
                    .NPWP_No = RowC.NPWP_No
                    .Tax_Address = RowC.Tax_Address
                    .Contact_Person = RowC.Contact_Person
                    ._Bank_s_Account = RowC._Bank_s_Account
                    ._Bank_s_Name = RowC._Bank_s_Name
                    .Beneficiary_Name = RowC.Beneficiary_Name
                    .Currency_Code = RowC.Currency_Code
                    .Group = RowC.Group
                    .Count = RowC.Count
                    .Code_Reference = RowC.Code_Reference
                    .Desc_Reference = RowC.Desc_Reference
                    .Join_Date = RowC.Join_Date
                    If List.Contains(.Primary_Customer) = False Then
                        .Tipe = 1
                        List.Add(.Primary_Customer)
                    Else
                        .Tipe = 0
                    End If
                End With
                tblDummy.Rows.Add(rowDummy)
            Next

            TblC.Rows.Clear()
            For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                rowDummy = tblDummy.DefaultView(i).Row
                RowC = TblC.NewRow
                For ii As Integer = 0 To tblDummy.Columns.Count - 1
                    RowC.Item(ii) = rowDummy.Item(ii)
                Next
                TblC.Rows.Add(RowC)
            Next
            '
            If Me.ChkTaxData.Checked = True Then
                Dim Report As New CrpLaporanCustomerList

                Report.SetDataSource(Ds)
                If Me.CboGroup.SelectedIndex = EnumGroup.Customer_Group Then
                    Report.SetParameterValue(Report.Parameter_ParamGroup.ParameterFieldName, "Customer Group :")
                ElseIf Me.CboGroup.SelectedIndex = EnumGroup.Sales_Area Then
                    Report.SetParameterValue(Report.Parameter_ParamGroup.ParameterFieldName, "Customer Sales Area :")
                ElseIf Me.CboGroup.SelectedIndex = EnumGroup.Customer_Segmentation Then
                    Report.SetParameterValue(Report.Parameter_ParamGroup.ParameterFieldName, "Customer Segmentation :")
                ElseIf Me.CboGroup.SelectedIndex = EnumGroup.Customer_Type Then
                    Report.SetParameterValue(Report.Parameter_ParamGroup.ParameterFieldName, "Customer Type :")
                ElseIf Me.CboGroup.SelectedIndex = EnumGroup.Customer_Reference Then
                    Report.SetParameterValue(Report.Parameter_ParamGroup.ParameterFieldName, "Customer Reference :")
                End If
                Crv.ReportSource = Report
            Else
                Dim Report As New CrpLaporanCustomerList_WithoutTaxAddress

                Report.SetDataSource(Ds)
                If Me.CboGroup.SelectedIndex = EnumGroup.Customer_Group Then
                    Report.SetParameterValue(Report.Parameter_ParamGroup.ParameterFieldName, "Customer Group :")
                ElseIf Me.CboGroup.SelectedIndex = EnumGroup.Sales_Area Then
                    Report.SetParameterValue(Report.Parameter_ParamGroup.ParameterFieldName, "Customer Sales Area :")
                ElseIf Me.CboGroup.SelectedIndex = EnumGroup.Customer_Segmentation Then
                    Report.SetParameterValue(Report.Parameter_ParamGroup.ParameterFieldName, "Customer Segmentation :")
                ElseIf Me.CboGroup.SelectedIndex = EnumGroup.Customer_Type Then
                    Report.SetParameterValue(Report.Parameter_ParamGroup.ParameterFieldName, "Customer Type :")
                ElseIf Me.CboGroup.SelectedIndex = EnumGroup.Customer_Reference Then
                    Report.SetParameterValue(Report.Parameter_ParamGroup.ParameterFieldName, "Customer Reference :")
                End If
                Crv.ReportSource = Report
            End If
            
            Crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Ds.Dispose()
            Dta.Dispose()
            DtaAdd.Dispose()
            DtaCP.Dispose()
            DtaBank.Dispose()
            Tbl.Dispose()
            TblAdd.Dispose()
            TblCP.Dispose()
            TblBank.Dispose()
            TblC.Dispose()
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub Sorting(ByRef vTbl As DseLaporanCustomerList.tbl_cetakDataTable)
        Dim tblDummy As New DseLaporanCustomerList.tbl_cetakDataTable
        Dim rowDummy As DseLaporanCustomerList.tbl_cetakRow
        Dim Sort As String = ""
        Dim Row As DseLaporanCustomerList.tbl_cetakRow
        Dim List As New List(Of Double)
        Try
            vTbl.DefaultView.Sort = "[Customer Code] ASC"
            For i As Integer = 0 To vTbl.DefaultView.Count - 1
                Row = vTbl.DefaultView(i).Row
                rowDummy = tblDummy.NewRow
                With rowDummy
                    .Primary_Customer = Row.Primary_Customer
                    .Customer_Code = Row.Customer_Code
                    .Customer_Name = Row.Customer_Name
                    .Address = Row.Address
                    .NPWP_No = Row.NPWP_No
                    .Tax_Address = Row.Tax_Address
                    .Contact_Person = Row.Contact_Person
                    ._Bank_s_Account = Row._Bank_s_Account
                    ._Bank_s_Name = Row._Bank_s_Name
                    .Beneficiary_Name = Row.Beneficiary_Name
                    .Currency_Code = Row.Currency_Code
                    .Group = Row.Group
                    .Count = Row.Count
                    .Code_Reference = Row.Code_Reference
                    .Desc_Reference = Row.Desc_Reference
                    .Join_Date = Row.Join_Date
                    If List.Contains(.Primary_Customer) = False Then
                        .Tipe = 1
                        List.Add(.Primary_Customer)
                    Else
                        .Tipe = 0
                    End If
                End With
                tblDummy.Rows.Add(rowDummy)
            Next
            'vTbl.Rows.Clear()
            vTbl = tblDummy.Copy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterFromDgv(ByRef vTbl As Data.DataTable)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim filter As String = ""
        Try
            tblDummy = vTbl.Copy
            'For Each dr As DataGridViewRow In Me.DgvFilter.Rows
            '    With dr
            '        If dr.Cells(Me.colPrimary.Index).Value <> -1 Then
            '            If dr.Cells(Me.colFilter.Index).Value = "Customer_Group" Then
            '                If filter = "" Then
            '                    filter = "[Primary Customer Group] = " & .Cells(Me.colPrimary.Index).Value.ToString
            '                Else
            '                    filter += " OR [Primary Customer Group] = " & .Cells(Me.colPrimary.Index).Value.ToString
            '                End If
            '            ElseIf dr.Cells(Me.colFilter.Index).Value = "Sales Area" Then
            '                If filter = "" Then
            '                    filter = "[Primary Sales Area] = " & .Cells(Me.colPrimary.Index).Value.ToString
            '                Else
            '                    filter += " OR [Primary Sales Area] = " & .Cells(Me.colPrimary.Index).Value.ToString
            '                End If
            '            ElseIf dr.Cells(Me.colFilter.Index).Value = "Customer Segmentation" Then
            '                If filter = "" Then
            '                    filter = "[Primary Customer Segmentasi] = " & .Cells(Me.colPrimary.Index).Value.ToString
            '                Else
            '                    filter += " OR [Primary Customer Segmentasi] = " & .Cells(Me.colPrimary.Index).Value.ToString
            '                End If
            '            ElseIf dr.Cells(Me.colFilter.Index).Value = "Customer Type" Then
            '                If filter = "" Then
            '                    filter = "[Primary Customer Type] = " & .Cells(Me.colPrimary.Index).Value.ToString
            '                Else
            '                    filter += " OR [Primary Customer Type] = " & .Cells(Me.colPrimary.Index).Value.ToString
            '                End If
            '            End If
            '        End If
            '    End With
            'Next

            tblDummy.DefaultView.RowFilter = Me.FilterByDgvItem
            vTbl.Rows.Clear()
            For i As Integer = 0 To tblDummy.DefaultView.Count - 1
                rowDummy = tblDummy.DefaultView(i).Row
                rowTbl = vTbl.NewRow
                For ii As Integer = 0 To tblDummy.Columns.Count - 1
                    rowTbl.Item(ii) = rowDummy.Item(ii)
                Next
                vTbl.Rows.Add(rowTbl)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblDummy.Dispose()
        End Try
    End Sub

    'Private Sub DgvFilter_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs)
    '    Try
    '        If e.ColumnIndex = Me.colFilter.Index Or e.ColumnIndex = Me.colCode.Index Then
    '            Me.Buffer = Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value
    '        Else
    '            e.Cancel = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Function getIndex(ByVal vRowException As Integer, ByVal vFilterLevel As String, ByVal vPrimary As Double) As Integer
    '    Try
    '        For Each dr As DataGridViewRow In Me.DgvFilter.Rows
    '            With dr
    '                If Me.DgvFilter.Rows.IndexOf(dr) <> vRowException Then
    '                    If dr.Cells(Me.colPrimary.Index).Value <> -1 Then
    '                        If dr.Cells(Me.colFilter.Index).Value = vFilterLevel And dr.Cells(Me.colPrimary.Index).Value = vPrimary Then
    '                            Return Me.DgvFilter.Rows.IndexOf(dr)
    '                        End If
    '                    End If
    '                End If
    '            End With
    '        Next

    '        Return -1
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Private Sub DgvFilter_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    '    Try
    '        If IsNothing(Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
    '            If String.IsNullOrEmpty(Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
    '                If IsNothing(Me.Buffer) = True Then
    '                    Me.index = e.RowIndex
    '                    Me.Bgw.RunWorkerAsync()
    '                    Exit Sub
    '                End If
    '                If Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
    '                    Exit Sub
    '                End If
    '                If e.ColumnIndex = Me.colFilter.Index Then
    '                    cellEndEditFilter(sender, e)
    '                End If
    '                If e.ColumnIndex = Me.colCode.Index Then
    '                    cellEndEditCode(sender, e)
    '                End If
    '            Else
    '                If IsNothing(Me.Buffer) = False Then
    '                    If String.IsNullOrEmpty(Buffer) = False Then
    '                        Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
    '                        Exit Sub
    '                    End If
    '                Else
    '                    Me.index = e.RowIndex
    '                    Me.Bgw.RunWorkerAsync()
    '                End If
    '            End If
    '        Else
    '            If IsNothing(Me.Buffer) = False Then
    '                If String.IsNullOrEmpty(Buffer) = False Then
    '                    Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
    '                    Exit Sub
    '                End If
    '            End If
    '            Me.index = e.RowIndex
    '            Me.Bgw.RunWorkerAsync()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub cellEndEditFilter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
    '    Dim frmFindCustGroup As New Core_Module.FrmCariMasterCustomerGroup
    '    Dim frmFindSalesArea As New Core_Module.FrmCariMasterSalesArea
    '    Dim frmFindSegmentasi As New Core_Module.FrmCariMasterCustomerSegmentation
    '    Dim frmFindType As New Core_Module.frmCariMasterCustomerType

    '    Try
    '        Select Case Me.DgvFilter.Item(Me.colFilter.Index, e.RowIndex).Value
    '            Case "Customer Group"
    '                frmFindCustGroup.txtCari.Text = Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value
    '                frmFindCustGroup.ShowDialog()
    '                If frmFindCustGroup.PrimaryCustomerGroup > 0 Then
    '                    Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value = frmFindCustGroup.PrimaryCustomerGroup
    '                    Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value = ObjCustGroup.GetGroupCode(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                    Me.DgvFilter.Item(Me.colDesc.Index, e.RowIndex).Value = ObjCustGroup.GetGroupDesc(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                Else
    '                    Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
    '                    Me.index = e.RowIndex
    '                    Me.Bgw.RunWorkerAsync()
    '                End If
    '            Case "Sales Area"
    '                frmFindSalesArea.txtCari.Text = Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value
    '                frmFindSalesArea.ShowDialog()
    '                If frmFindSalesArea.PrimaryCustomerSalesArea > 0 Then
    '                    Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value = frmFindSalesArea.PrimaryCustomerSalesArea
    '                    Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value = ObjCustSalesArea.GetSalesAreaCode(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                    Me.DgvFilter.Item(Me.colDesc.Index, e.RowIndex).Value = ObjCustSalesArea.GetSalesAreaDesc(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                Else
    '                    Me.index = e.RowIndex
    '                    Me.Bgw.RunWorkerAsync()
    '                End If
    '            Case "Customer Segmentation"
    '                frmFindSegmentasi.txtCari.Text = Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value
    '                frmFindSegmentasi.ShowDialog()
    '                If frmFindSegmentasi.PrimaryCustomerMarketSegmentation > 0 Then
    '                    Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value = frmFindSegmentasi.PrimaryCustomerMarketSegmentation
    '                    Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value = ObjCustSegmentation.GetSegmentationCode(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                    Me.DgvFilter.Item(Me.colDesc.Index, e.RowIndex).Value = ObjCustSegmentation.GetDesc(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                Else
    '                    Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
    '                    Me.index = e.RowIndex
    '                    Me.Bgw.RunWorkerAsync()
    '                End If
    '            Case "Customer Type"
    '                frmFindType.txtCari.Text = Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value
    '                frmFindType.ShowDialog()
    '                If frmFindType.PrimaryCustomerType > 0 Then
    '                    Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value = frmFindType.PrimaryCustomerType
    '                    Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value = ObjCustType.GetTypeCode(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                    Me.DgvFilter.Item(Me.colDesc.Index, e.RowIndex).Value = ObjCustType.GetTypeDesc(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                Else
    '                    Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
    '                End If
    '        End Select
    '        If Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value <> -1 Then
    '            If Me.getIndex(e.RowIndex, Me.DgvFilter.Item(Me.colFilter.Index, e.RowIndex).Value, Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value) <> -1 Then
    '                MsgBox(Me.DgvFilter.Item(Me.colFilter.Index, e.RowIndex).Value & " " & Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
    '                Me.index = e.RowIndex
    '                Me.Bgw.RunWorkerAsync()
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub cellEndEditCode(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
    '    Dim frmFindCustGroup As New Core_Module.FrmCariMasterCustomerGroup
    '    Dim frmFindSalesArea As New Core_Module.FrmCariMasterSalesArea
    '    Dim frmFindSegmentasi As New Core_Module.FrmCariMasterCustomerSegmentation
    '    Dim frmFindType As New Core_Module.frmCariMasterCustomerType

    '    Try
    '        Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
    '        Select Case Me.DgvFilter.Item(Me.colFilter.Index, e.RowIndex).Value
    '            Case "Group"
    '                If Me.ObjCustGroup.ApakahSudahAda(Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
    '                    Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value = ObjCustGroup.GetPrimary(Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value)
    '                    Me.DgvFilter.Item(Me.colDesc.Index, e.RowIndex).Value = ObjCustGroup.GetGroupDesc(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                Else
    '                    frmFindCustGroup.txtCari.Text = Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value
    '                    frmFindCustGroup.ShowDialog()
    '                    If frmFindCustGroup.PrimaryCustomerGroup > 0 Then
    '                        Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value = frmFindCustGroup.PrimaryCustomerGroup
    '                        Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value = ObjCustGroup.GetGroupCode(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                        Me.DgvFilter.Item(Me.colDesc.Index, e.RowIndex).Value = ObjCustGroup.GetGroupDesc(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                    Else
    '                        If IsNothing(Me.Buffer) = False Then
    '                            If String.IsNullOrEmpty(Buffer) = False Then
    '                                Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
    '                                Exit Sub
    '                            End If
    '                        End If
    '                        Me.index = e.RowIndex
    '                        Me.Bgw.RunWorkerAsync()
    '                    End If
    '                End If
    '            Case "Sales Area"
    '                If Me.ObjCustSalesArea.ApakahSudahAda(Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
    '                    Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value = ObjCustSalesArea.GetPrimary(Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value)
    '                    Me.DgvFilter.Item(Me.colDesc.Index, e.RowIndex).Value = ObjCustSalesArea.GetSalesAreaDesc(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                Else
    '                    frmFindSalesArea.txtCari.Text = Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value
    '                    frmFindSalesArea.ShowDialog()
    '                    If frmFindSalesArea.PrimaryCustomerSalesArea > 0 Then
    '                        Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value = frmFindSalesArea.PrimaryCustomerSalesArea
    '                        Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value = ObjCustSalesArea.GetSalesAreaCode(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                        Me.DgvFilter.Item(Me.colDesc.Index, e.RowIndex).Value = ObjCustSalesArea.GetSalesAreaDesc(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                    Else
    '                        If IsNothing(Me.Buffer) = False Then
    '                            If String.IsNullOrEmpty(Buffer) = False Then
    '                                Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
    '                                Exit Sub
    '                            End If
    '                        End If
    '                        Me.index = e.RowIndex
    '                        Me.Bgw.RunWorkerAsync()
    '                    End If
    '                End If
    '            Case "Customer Segmentation"
    '                If Me.ObjCustSegmentation.ApakahSudahAda(Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
    '                    Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value = ObjCustSegmentation.GetPrimary(Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value)
    '                    Me.DgvFilter.Item(Me.colDesc.Index, e.RowIndex).Value = ObjCustSegmentation.GetDesc(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                Else
    '                    frmFindSegmentasi.txtCari.Text = Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value
    '                    frmFindSegmentasi.ShowDialog()
    '                    If frmFindSegmentasi.PrimaryCustomerMarketSegmentation > 0 Then
    '                        Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value = frmFindSegmentasi.PrimaryCustomerMarketSegmentation
    '                        Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value = ObjCustSegmentation.GetSegmentationCode(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                        Me.DgvFilter.Item(Me.colDesc.Index, e.RowIndex).Value = ObjCustSegmentation.GetDesc(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                    Else
    '                        If IsNothing(Me.Buffer) = False Then
    '                            If String.IsNullOrEmpty(Buffer) = False Then
    '                                Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
    '                                Exit Sub
    '                            End If
    '                        End If
    '                        Me.index = e.RowIndex
    '                        Me.Bgw.RunWorkerAsync()
    '                    End If
    '                End If
    '            Case "Customer Type"
    '                If Me.ObjCustType.ApakahSudahAda(Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
    '                    Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value = ObjCustType.GetPrimary(Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value)
    '                    Me.DgvFilter.Item(Me.colDesc.Index, e.RowIndex).Value = ObjCustType.GetTypeDesc(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                Else
    '                    frmFindType.txtCari.Text = Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value
    '                    frmFindType.ShowDialog()
    '                    If frmFindType.PrimaryCustomerType > 0 Then
    '                        Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value = frmFindType.PrimaryCustomerType
    '                        Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value = ObjCustType.GetTypeCode(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                        Me.DgvFilter.Item(Me.colDesc.Index, e.RowIndex).Value = ObjCustType.GetTypeDesc(Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value)
    '                    Else
    '                        If IsNothing(Me.Buffer) = False Then
    '                            If String.IsNullOrEmpty(Buffer) = False Then
    '                                Me.DgvFilter.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
    '                                Exit Sub
    '                            End If
    '                        End If
    '                        Me.index = e.RowIndex
    '                        Me.Bgw.RunWorkerAsync()
    '                    End If
    '                End If
    '        End Select
    '        If Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value <> -1 Then
    '            If Me.getIndex(e.RowIndex, Me.DgvFilter.Item(Me.colFilter.Index, e.RowIndex).Value, Me.DgvFilter.Item(Me.colPrimary.Index, e.RowIndex).Value) <> -1 Then
    '                MsgBox(Me.DgvFilter.Item(Me.colFilter.Index, e.RowIndex).Value & " " & Me.DgvFilter.Item(Me.colCode.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
    '                Me.index = e.RowIndex
    '                Me.Bgw.RunWorkerAsync()
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub DgvFilter_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs)
    '    Dim dr As DataGridViewRow
    '    Try
    '        dr = Me.DgvFilter.Rows(Me.DgvFilter.Rows.Count - 1)
    '        dr.Cells(Me.colNo.Index).Value = Me.DgvFilter.Rows.Count
    '        dr.Cells(Me.colFilter.Index).Value = "Group"
    '        dr.Cells(Me.colCode.Index).Value = ""
    '        dr.Cells(Me.colDesc.Index).Value = ""
    '        dr.Cells(Me.colPrimary.Index).Value = -1
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub DgvFilter_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs)
    '    Try
    '        Me.ObjCD.SetNumber(Me.DgvFilter)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub DgvFilter_Sorted(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Try
    '        Me.ObjCD.SetNumber(Me.DgvFilter)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub Bgw_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles Bgw.DoWork
    '    Try
    '        Threading.Thread.Sleep(1)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub Bgw_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles Bgw.RunWorkerCompleted
    '    Try
    '        Me.DgvFilter.Rows.RemoveAt(Me.index)
    '        Me.ObjCD.SetNumber(Me.DgvFilter)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub DgvFilter_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)

    'End Sub

#Region "Filteran Customer Group"
    Dim EraseIndex_CG As Integer = -1
    Private Sub DgvCustGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvCustGroup.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCG_CustGroup.Index Then
                Me.Buffer = Me.DgvCustGroup(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCustGroup.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvCustGroup.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjSA As New Core_Module.clsMasterCustomerGroup
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvCustGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCG_CustGroup.Index Then
                        .Cells(Me.ColCG_CustGroup.Index).Value = .Cells(Me.ColCG_CustGroup.Index).Value.ToString.ToUpper
                        If ObjSA.ApakahSudahAda(.Cells(Me.ColCG_CustGroup.Index).Value) = True Then
                            primary = ObjSA.GetPrimary(.Cells(Me.ColCG_CustGroup.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_CG(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColCG_Primary.Index).Value = primary
                                .Cells(Me.ColCG_CustGroup.Index).Value = ObjSA.GetGroupCode(primary)
                                .Cells(Me.ColCG_Desc.Index).Value = ObjSA.GetGroupDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvCustGroup.Item(Me.ColCG_CustGroup.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_CG = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.DgvCustGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustGroup)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterCustomerGroup
                            ij.txtCari.Text = Me.DgvCustGroup.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerGroup <> -1 Then
                                index = Me.GetIndexItem_CG(ij.PrimaryCustomerGroup, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColCG_Primary.Index).Value = ij.PrimaryCustomerGroup
                                    .Cells(Me.ColCG_CustGroup.Index).Value = ObjSA.GetGroupCode(ij.PrimaryCustomerGroup)
                                    .Cells(Me.ColCG_Desc.Index).Value = ObjSA.GetGroupDesc(ij.PrimaryCustomerGroup)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_CG = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvCustGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvCustGroup)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColCG_Primary.Index).Value = ij.PrimaryCustomerGroup
                                        .Cells(Me.ColCG_CustGroup.Index).Value = ObjSA.GetGroupCode(ij.PrimaryCustomerGroup)
                                        .Cells(Me.ColCG_Desc.Index).Value = ObjSA.GetGroupDesc(ij.PrimaryCustomerGroup)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerGroup = -1 Then
                                If Me.DgvCustGroup.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_CG = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.DgvCustGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustGroup)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvCustGroup)
                        Me.DgvCustGroup.Item(Me.ColCG_CustGroup.Index, e.RowIndex).Value = Me.DgvCustGroup.Item(Me.ColCG_CustGroup.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvCustGroup.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColCG_CustGroup.Index Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_CG = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            Else
                                Me.DgvCustGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvCustGroup)
                            Exit Sub
                        Else
                            Me.DgvCustGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColCG_CustGroup.Index Then
                            Me.DgvCustGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvCustGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub DgvCustGroup_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvCustGroup.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvCustGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustGroup_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCustGroup.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvCustGroup, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork1(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted1(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvCustGroup.Rows.RemoveAt(EraseIndex_CG)
            Me.ObjCD.SetNumber(Me.DgvCustGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_CG(ByVal vPrimary As Double, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvCustGroup.Rows.Count - 1
                RowDGV = Me.DgvCustGroup.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCG_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
#End Region

#Region "Filteran Sales Area"
    Dim EraseIndex_SA As Integer = -1
    Private Sub DgvSalesArea_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvSalesArea.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColSA_SalesArea.Index Then
                Me.Buffer = Me.DgvSalesArea(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvSalesArea.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvSalesArea.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjSA As New Core_Module.clsMasterCustomerSalesArea
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColSA_SalesArea.Index Then
                        .Cells(Me.ColSA_SalesArea.Index).Value = .Cells(Me.ColSA_SalesArea.Index).Value.ToString.ToUpper
                        If ObjSA.ApakahSudahAda(.Cells(Me.ColSA_SalesArea.Index).Value) = True Then
                            primary = ObjSA.GetPrimary(.Cells(Me.ColSA_SalesArea.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_SA(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColSA_Primary.Index).Value = primary
                                .Cells(Me.ColSA_SalesArea.Index).Value = ObjSA.GetSalesAreaCode(primary)
                                .Cells(Me.ColSA_Desc.Index).Value = ObjSA.GetSalesAreaDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvSalesArea.Item(Me.ColSA_SalesArea.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_SA = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalesArea)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterSalesArea
                            ij.txtCari.Text = Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerSalesArea <> -1 Then
                                index = Me.GetIndexItem_SA(ij.PrimaryCustomerSalesArea, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColSA_Primary.Index).Value = ij.PrimaryCustomerSalesArea
                                    .Cells(Me.ColSA_SalesArea.Index).Value = ObjSA.GetSalesAreaCode(ij.PrimaryCustomerSalesArea)
                                    .Cells(Me.ColSA_Desc.Index).Value = ObjSA.GetSalesAreaDesc(ij.PrimaryCustomerSalesArea)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_SA = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvSalesArea)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColSA_Primary.Index).Value = ij.PrimaryCustomerSalesArea
                                        .Cells(Me.ColSA_SalesArea.Index).Value = ObjSA.GetSalesAreaCode(ij.PrimaryCustomerSalesArea)
                                        .Cells(Me.ColSA_Desc.Index).Value = ObjSA.GetSalesAreaDesc(ij.PrimaryCustomerSalesArea)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerSalesArea = -1 Then
                                If Me.DgvSalesArea.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_SA = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalesArea)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvSalesArea)
                        Me.DgvSalesArea.Item(Me.ColSA_SalesArea.Index, e.RowIndex).Value = Me.DgvSalesArea.Item(Me.ColSA_SalesArea.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvSalesArea.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColSA_SalesArea.Index Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_SA = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            Else
                                Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvSalesArea)
                            Exit Sub
                        Else
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColSA_SalesArea.Index Then
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub DgvSalesArea_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvSalesArea.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvSalesArea)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvSalesArea.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvSalesArea, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.DgvSalesArea.Rows.RemoveAt(EraseIndex_SA)
            Me.ObjCD.SetNumber(Me.DgvSalesArea)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_SA(ByVal vPrimary As Double, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvSalesArea.Rows.Count - 1
                RowDGV = Me.DgvSalesArea.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColSA_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
#End Region

#Region "Filteran Customer Segmentation"
    Dim EraseIndex_MS As Integer = -1
    Dim ObjMS As New Core_Module.clsMasterCustomerMarketSegmentation
    Private Sub DgvMarketSegmentation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvMarketSegmentation.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                Me.Buffer = Me.DgvMarketSegmentation(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMarketSegmentation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvMarketSegmentation.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvMarketSegmentation.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjMS As New Core_Module.clsMasterCustomerMarketSegmentation
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                        .Cells(Me.ColMS_MarketSeg.Index).Value = .Cells(Me.ColMS_MarketSeg.Index).Value.ToString.ToUpper
                        If ObjMS.ApakahSudahAda(.Cells(Me.ColMS_MarketSeg.Index).Value) = True Then
                            primary = ObjMS.GetPrimary(.Cells(Me.ColMS_MarketSeg.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_MS(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColMS_Primary.Index).Value = primary
                                .Cells(Me.ColMS_MarketSeg.Index).Value = ObjMS.GetSegmentationCode(primary)
                                .Cells(Me.ColMS_Desc.Index).Value = ObjMS.GetDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvMarketSegmentation.Item(Me.ColMS_MarketSeg.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_MS = e.RowIndex
                                        Me.BackgroundWorker3.RunWorkerAsync()
                                    Else
                                        Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterCustomerSegmentation
                            ij.txtCari.Text = Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerMarketSegmentation <> -1 Then
                                index = Me.GetIndexItem_MS(ij.PrimaryCustomerMarketSegmentation, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColMS_Primary.Index).Value = ij.PrimaryCustomerMarketSegmentation
                                    .Cells(Me.ColMS_MarketSeg.Index).Value = ObjMS.GetSegmentationCode(ij.PrimaryCustomerMarketSegmentation)
                                    .Cells(Me.ColMS_Desc.Index).Value = ObjMS.GetDesc(ij.PrimaryCustomerMarketSegmentation)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_MS = e.RowIndex
                                            Me.BackgroundWorker3.RunWorkerAsync()
                                        Else
                                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColMS_Primary.Index).Value = ij.PrimaryCustomerMarketSegmentation
                                        .Cells(Me.ColMS_MarketSeg.Index).Value = ObjMS.GetSegmentationCode(ij.PrimaryCustomerMarketSegmentation)
                                        .Cells(Me.ColMS_Desc.Index).Value = ObjMS.GetDesc(ij.PrimaryCustomerMarketSegmentation)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerMarketSegmentation = -1 Then
                                If Me.DgvMarketSegmentation.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_MS = e.RowIndex
                                        Me.BackgroundWorker3.RunWorkerAsync()
                                    Else
                                        Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                        Me.DgvMarketSegmentation.Item(Me.ColMS_MarketSeg.Index, e.RowIndex).Value = Me.DgvMarketSegmentation.Item(Me.ColMS_MarketSeg.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvMarketSegmentation.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_MS = e.RowIndex
                                Me.BackgroundWorker3.RunWorkerAsync()
                            Else
                                Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                            Exit Sub
                        Else
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColMS_MarketSeg.Index Then
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetIndexItem_MS(ByVal vPrimary As Double, _
     ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvMarketSegmentation.Rows.Count - 1
                RowDGV = Me.DgvMarketSegmentation.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColMS_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvMarketSegmentation_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvMarketSegmentation.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvMarketSegmentation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.DgvMarketSegmentation.Rows.RemoveAt(EraseIndex_MS)
            Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Filteran Customer Type"
    Dim EraseIndex_CT As Integer = -1
    Dim ObjCT As New Core_Module.clsMasterCustomerType
    Private Sub DgvCustomerType_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvCustomerType.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCT_CustType.Index Then
                Me.Buffer = Me.DgvCustomerType(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomerType_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCustomerType.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvCustomerType.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjCT As New Core_Module.clsMasterCustomerType
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCT_CustType.Index Then
                        .Cells(Me.ColCT_CustType.Index).Value = .Cells(Me.ColCT_CustType.Index).Value.ToString.ToUpper
                        If ObjCT.ApakahSudahAda(.Cells(Me.ColCT_CustType.Index).Value) = True Then
                            primary = ObjCT.GetPrimary(.Cells(Me.ColCT_CustType.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_CT(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColCT_Primary.Index).Value = primary
                                .Cells(Me.ColCT_CustType.Index).Value = ObjCT.GetTypeCode(primary)
                                .Cells(Me.ColCT_Desc.Index).Value = ObjCT.GetTypeDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvCustomerType.Item(Me.ColCT_CustType.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_CT = e.RowIndex
                                        Me.BackgroundWorker4.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomerType)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterCustomerType
                            ij.txtCari.Text = Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerType <> -1 Then
                                index = Me.GetIndexItem_CT(ij.PrimaryCustomerType, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColCT_Primary.Index).Value = ij.PrimaryCustomerType
                                    .Cells(Me.ColCT_CustType.Index).Value = ObjCT.GetTypeCode(ij.PrimaryCustomerType)
                                    .Cells(Me.ColCT_Desc.Index).Value = ObjCT.GetTypeDesc(ij.PrimaryCustomerType)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_CT = e.RowIndex
                                            Me.BackgroundWorker4.RunWorkerAsync()
                                        Else
                                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvCustomerType)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColCT_Primary.Index).Value = ij.PrimaryCustomerType
                                        .Cells(Me.ColCT_CustType.Index).Value = ObjCT.GetTypeCode(ij.PrimaryCustomerType)
                                        .Cells(Me.ColCT_Desc.Index).Value = ObjCT.GetTypeDesc(ij.PrimaryCustomerType)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerType = -1 Then
                                If Me.DgvCustomerType.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_CT = e.RowIndex
                                        Me.BackgroundWorker4.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomerType)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvCustomerType)
                        Me.DgvCustomerType.Item(Me.ColCT_CustType.Index, e.RowIndex).Value = Me.DgvCustomerType.Item(Me.ColCT_CustType.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvCustomerType.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColCT_CustType.Index Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_CT = e.RowIndex
                                Me.BackgroundWorker4.RunWorkerAsync()
                            Else
                                Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvCustomerType)
                            Exit Sub
                        Else
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColCT_CustType.Index Then
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub DgvCustomerType_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvCustomerType.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvCustomerType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomerType_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCustomerType.Sorted
        Try
            ObjCD.SetNumber(Me.DgvCustomerType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            Me.DgvCustomerType.Rows.RemoveAt(EraseIndex_CT)
            Me.ObjCD.SetNumber(Me.DgvCustomerType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_CT(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvCustomerType.Rows.Count - 1
                RowDGV = Me.DgvCustomerType.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCT_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
#End Region

#Region "Filteran Customer Reference"
    Dim EraseIndexCR As Integer = -1
    Private Sub DgvFilterCustRef_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvCustRef.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCustomerReference.Index = True Then
                Me.Buffer = Me.DgvCustRef(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvFilterCustRef_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCustRef.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvCustRef.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjRef As New Core_Module.ClsMasterCustomerReference
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvCustRef.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCustomerReference.Index Then
                        .Cells(Me.ColCustomerReference.Index).Value = .Cells(Me.ColCustomerReference.Index).Value.ToString.ToUpper
                        If ObjRef.ApakahSudahAda(.Cells(Me.ColCustomerReference.Index).Value) = True Then
                            primary = ObjRef.GetPrimary(.Cells(Me.ColCustomerReference.Index).Value.ToString.ToUpper)
                            index = Me.getindexitemRef(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColCustomerReference_Primary.Index).Value = primary
                                .Cells(Me.ColCustomerReference.Index).Value = ObjRef.GetCode(primary)
                                .Cells(Me.ColCustomerReference_Desc.Index).Value = ObjRef.GetDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvCustRef.Item(Me.ColCustomerReference.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndexCR = e.RowIndex
                                        Me.BackgroundWorker5.RunWorkerAsync()
                                    Else
                                        Me.DgvCustRef.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustRef)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmMasterCustomer_Reference_Cari
                            ij.txtCari.Text = Me.DgvCustRef.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustRef <> -1 Then
                                index = Me.getindexitemRef(ij.PrimaryCustRef, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColCustomerReference_Primary.Index).Value = ij.PrimaryCustRef
                                    .Cells(Me.ColCustomerReference.Index).Value = ObjRef.GetCode(ij.PrimaryCustRef)
                                    .Cells(Me.ColCustomerReference_Desc.Index).Value = ObjRef.GetDesc(ij.PrimaryCustRef)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndexCR = e.RowIndex
                                            Me.BackgroundWorker5.RunWorkerAsync()
                                        Else
                                            Me.DgvCustRef.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvCustRef)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColCustomerReference_Primary.Index).Value = ij.PrimaryCustRef
                                        .Cells(Me.ColCustomerReference.Index).Value = ObjRef.GetCode(ij.PrimaryCustRef)
                                        .Cells(Me.ColCustomerReference_Desc.Index).Value = ObjRef.GetDesc(ij.PrimaryCustRef)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustRef = -1 Then
                                If Me.DgvCustRef.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndexCR = e.RowIndex
                                        Me.BackgroundWorker5.RunWorkerAsync()
                                    Else
                                        Me.DgvCustRef.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustRef)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvCustRef)
                        Me.DgvCustRef.Item(Me.ColCustomerReference.Index, e.RowIndex).Value = Me.DgvCustRef.Item(Me.ColCustomerReference.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvCustRef.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColCustomerReference.Index Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndexCR = e.RowIndex
                                Me.BackgroundWorker5.RunWorkerAsync()
                            Else
                                Me.DgvCustRef.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvCustRef)
                            Exit Sub
                        Else
                            Me.DgvCustRef.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColCustomerReference.Index Then
                            Me.DgvCustRef.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvCustRef.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function getindexitemRef(ByVal vPrimary As Double, _
     ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvCustRef.Rows.Count - 1
                RowDGV = Me.DgvCustRef.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCustomerReference_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvFilterCustRef_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvCustRef.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvCustRef)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvFilterCustRef_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCustRef.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvCustRef, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvFilterCustRef_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvCustRef.UserDeletedRow
        ObjCD.SetNumber(Me.DgvCustRef)
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
        Try
            Me.DgvCustRef.Rows.RemoveAt(EraseIndexCR)
            Me.ObjCD.SetNumber(Me.DgvCustRef)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    Private Function FilterByDgvItem() As String
        Dim FilterBy_CustomerType As String = ""
        Dim FilterBy_CustomerGroup As String = ""
        Dim FilterBy_Market As String = ""
        Dim FilterBy_SalesArea As String = ""
        Dim FilterBy_Ref As String = ""
        Dim Filter As String = ""
        Try
            If Me.DgvCustGroup.RowCount > 0 Then
                For Each row As DataGridViewRow In Me.DgvCustGroup.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColCG_CustGroup.Index).Value) = False Then
                        If FilterBy_CustomerGroup = "" Then
                            FilterBy_CustomerGroup = "[Customer Group Code]='" & row.Cells(Me.ColCG_CustGroup.Index).Value & "'"
                        Else
                            FilterBy_CustomerGroup &= " OR [Customer Group Code]='" & row.Cells(Me.ColCG_CustGroup.Index).Value & "'"
                        End If
                    End If
                Next
                If FilterBy_CustomerGroup <> "" Then
                    If Filter <> "" Then
                        Filter &= "AND (" & FilterBy_CustomerGroup & ")"
                    Else
                        Filter &= "(" & FilterBy_CustomerGroup & ")"
                    End If
                End If
            End If

            If Me.DgvCustomerType.RowCount > 0 Then
                For Each row As DataGridViewRow In Me.DgvCustomerType.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColCT_CustType.Index).Value) = False Then
                        If FilterBy_CustomerType = "" Then
                            FilterBy_CustomerType = "[Type Code]='" & row.Cells(Me.ColCT_CustType.Index).Value & "'"
                        Else
                            FilterBy_CustomerType &= " OR [Type Code]='" & row.Cells(Me.ColCT_CustType.Index).Value & "'"
                        End If
                    End If
                Next
                If FilterBy_CustomerType <> "" Then
                    If Filter <> "" Then
                        Filter &= "AND (" & FilterBy_CustomerType & ")"
                    Else
                        Filter &= "(" & FilterBy_CustomerType & ")"
                    End If
                End If
            End If


            If Me.DgvMarketSegmentation.RowCount > 0 Then
                For Each row As DataGridViewRow In Me.DgvMarketSegmentation.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColMS_MarketSeg.Index).Value) = False Then
                        If FilterBy_Market = "" Then
                            FilterBy_Market = "[Segmentation Code]='" & row.Cells(Me.ColMS_MarketSeg.Index).Value & "'"
                        Else
                            FilterBy_Market &= " OR [Segmentation Code]='" & row.Cells(Me.ColMS_MarketSeg.Index).Value & "'"
                        End If
                    End If
                Next

                If FilterBy_Market <> "" Then
                    If Filter <> "" Then
                        Filter &= "AND (" & FilterBy_Market & ")"
                    Else
                        Filter = "(" & FilterBy_Market & ")"
                    End If
                End If
            End If

            If Me.DgvSalesArea.RowCount > 0 Then
                For Each row As DataGridViewRow In Me.DgvSalesArea.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColSA_SalesArea.Index).Value) = False Then
                        If FilterBy_SalesArea = "" Then
                            FilterBy_SalesArea = "[Sales Area Code]='" & row.Cells(Me.ColSA_SalesArea.Index).Value & "'"
                        Else
                            FilterBy_SalesArea &= " OR [Sales Area Code]='" & row.Cells(Me.ColSA_SalesArea.Index).Value & "'"
                        End If
                    End If
                Next
                If FilterBy_SalesArea <> "" Then
                    If Filter <> "" Then
                        Filter &= "AND (" & FilterBy_SalesArea & ")"
                    Else
                        Filter = "(" & FilterBy_SalesArea & ")"
                    End If
                End If
            End If

            If Me.DgvCustRef.RowCount > 0 Then
                For Each row As DataGridViewRow In Me.DgvCustRef.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColCustomerReference.Index).Value) = False Then
                        If FilterBy_Ref = "" Then
                            FilterBy_Ref = "[Code Ref]='" & row.Cells(Me.ColCustomerReference.Index).Value & "'"
                        Else
                            FilterBy_Ref &= " OR [Code Ref]='" & row.Cells(Me.ColCustomerReference.Index).Value & "'"
                        End If
                    End If
                Next
                If FilterBy_Ref <> "" Then
                    If Filter <> "" Then
                        Filter &= "AND (" & FilterBy_Ref & ")"
                    Else
                        Filter &= "(" & FilterBy_Ref & ")"
                    End If
                End If
            End If

            Return Filter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub ChkJoinDate1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkJoinDate1.CheckedChanged
        Try
            If Me.ChkJoinDate1.Checked = True Then
                Me.DtpFrom.Enabled = True
                Me.DtpTo.Enabled = True
            Else
                Me.DtpFrom.Enabled = False
                Me.DtpTo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class