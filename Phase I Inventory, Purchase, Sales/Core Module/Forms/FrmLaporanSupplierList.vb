Option Explicit On
Option Strict Off
Public Class FrmLaporanSupplierList

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click

        Dim DS As New DseLaporanSupplierList
        Dim DtaSupplier As New DseLaporanSupplierListTableAdapters.tblmaster_supplierTableAdapter
        Dim TblSupplier As New DseLaporanSupplierList.tblmaster_supplierDataTable
        Dim RowSupplier As DseLaporanSupplierList.tblmaster_supplierRow

        Dim DtaItemAddress As New DseLaporanSupplierListTableAdapters.tblItem_AddressTableAdapter
        Dim TblItemAddress As New DseLaporanSupplierList.tblItem_AddressDataTable
        Dim RowItemAddress As DseLaporanSupplierList.tblItem_AddressRow


        Dim DtaItemCp As New DseLaporanSupplierListTableAdapters.tblItem_CPTableAdapter
        Dim TblItemCp As New DseLaporanSupplierList.tblItem_CPDataTable
        Dim RowItemCp As DseLaporanSupplierList.tblItem_CPRow

        Dim DtaItemBa As New DseLaporanSupplierListTableAdapters.tblItem_BankTableAdapter
        Dim TblItemBa As New DseLaporanSupplierList.tblItem_BankDataTable
        Dim RowItemBa As DseLaporanSupplierList.tblItem_BankRow

        Dim TBLCetak As DseLaporanSupplierList.tbl_cetakDataTable = DS.tbl_cetak
        Dim RowCetak As DseLaporanSupplierList.tbl_cetakRow

        Dim Report As New CrpLaporanSupplierList

        Dim a(2) As Integer
        Dim counter As Integer
        Try
            DtaSupplier.Connection.ConnectionString = Core_Module.GetConstring
            DtaItemAddress.Connection.ConnectionString = Core_Module.GetConstring
            DtaItemCp.Connection.ConnectionString = Core_Module.GetConstring
            DtaItemBa.Connection.ConnectionString = Core_Module.GetConstring
            If cboGrupSupplier.Text = "Local Only" Then
                DtaSupplier.FillByForeign(TblSupplier, 0, 0)
            ElseIf cboGrupSupplier.Text = "Foreign Only" Then
                DtaSupplier.FillByForeign(TblSupplier, 1, 1)
            Else
                DtaSupplier.FillByForeign(TblSupplier, 0, 1)
            End If

            If chkMul_address.Checked = True Then
                DtaItemAddress.Fill(TblItemAddress)
            Else
                DtaItemAddress.FillBy_PrimarySupplier(TblItemAddress)
            End If
            If chkConPerson.Checked = True Then
                DtaItemCp.Fill(TblItemCp)
            Else
                DtaItemCp.FillBy_PrimarySupplier(TblItemCp)
            End If
            If chkBankaccount.Checked = True Then
                DtaItemBa.Fill(TblItemBa)
            Else
                DtaItemBa.FillBy_PrimarySupplier(TblItemBa)
            End If

            If TblSupplier.Count > 0 Then
                For x As Integer = 0 To TblSupplier.Count - 1
                    RowSupplier = TblSupplier.DefaultView(x).Row
                    TblItemAddress.DefaultView.RowFilter = "[Primary Supplier]  = " & RowSupplier.Primary
                    TblItemCp.DefaultView.RowFilter = "[Primary Supplier]  = " & RowSupplier.Primary
                    TblItemBa.DefaultView.RowFilter = "[Primary Supplier]  = " & RowSupplier.Primary

                    a(0) = TblItemAddress.DefaultView.Count
                    a(1) = TblItemCp.DefaultView.Count
                    a(2) = TblItemBa.DefaultView.Count
                    counter = 0
                    For ij As Integer = 0 To 2
                        If a(ij) >= counter Then
                            counter = a(ij)
                        End If
                    Next

                    For z As Integer = 0 To counter - 1
                        RowCetak = TBLCetak.Newtbl_cetakRow
                        RowCetak.PrimarySupplier = RowSupplier.Primary
                        RowCetak.Supplier_Code = RowSupplier.Supplier_Code
                        RowCetak.Supplier_Name = RowSupplier.Supplier_Name

                        RowCetak.Address = "-"
                        RowCetak.NPWP = "-"
                        RowCetak.Contact = "-"
                        RowCetak.Bank_Account = "-"
                        RowCetak.Bank_Name = "-"
                        RowCetak.Beneficiary_Name = "-"
                        RowCetak.Curr = "-"
                        RowCetak.Supplier_Group = "-"
                        RowCetak.Count = counter
                        If TblItemAddress.DefaultView.Count - 1 >= z Then
                            RowItemAddress = TblItemAddress.DefaultView(z).Row
                            If RowItemAddress.Line_1 <> "" Then
                                RowCetak.Address = RowItemAddress.Line_1
                            End If
                            If RowItemAddress.Line_2 <> "" And RowItemAddress.Line_1 <> "" Then
                                RowCetak.Address = RowCetak.Address & "," & RowItemAddress.Line_2
                            ElseIf RowItemAddress.Line_2 <> "" And RowItemAddress.Line_1 = "" Then
                                RowCetak.Address = RowItemAddress.Line_2
                            End If
                            If RowItemAddress.Line_3 <> "" And RowItemAddress.Line_2 <> "" Then
                                RowCetak.Address = RowCetak.Address & "," & RowItemAddress.Line_3
                            ElseIf RowItemAddress.Line_3 <> "" And RowItemAddress.Line_2 = "" Then
                                RowCetak.Address = RowItemAddress.Line_3
                            End If
                            If RowItemAddress.Line_4 <> "" And RowItemAddress.Line_3 <> "" Then
                                RowCetak.Address = RowCetak.Address & "," & RowItemAddress.Line_4
                            ElseIf RowItemAddress.Line_4 <> "" And RowItemAddress.Line_3 = "" Then
                                RowCetak.Address = RowItemAddress.Line_4
                            End If
                            If RowItemAddress.Line_5 <> "" And RowItemAddress.Line_4 <> "" Then
                                RowCetak.Address = RowCetak.Address & "," & RowItemAddress.Line_5
                            ElseIf RowItemAddress.Line_5 <> "" And RowItemAddress.Line_4 = "" Then
                                RowCetak.Address = RowItemAddress.Line_4
                            End If
                            If RowCetak.Address <> "" And RowItemAddress.Phone_No <> "" Then
                                RowCetak.Address = RowCetak.Address & "," & RowItemAddress.Phone_No
                            ElseIf RowCetak.Address = "" And RowItemAddress.Phone_No <> "" Then
                                RowCetak.Address = RowItemAddress.Phone_No
                            End If
                            If RowCetak.Address <> "" And RowItemAddress.Fax_No <> "" Then
                                RowCetak.Address = RowCetak.Address & "," & RowItemAddress.Fax_No
                            ElseIf RowCetak.Address = "" And RowItemAddress.Fax_No <> "" Then
                                RowCetak.Address = RowItemAddress.Fax_No
                            End If

                            'RowCetak.Address = RowCetak.Address & "," & RowItemAddress.Phone_No & "," & RowItemAddress.Fax_No
                        End If

                        If z = 0 Then
                            RowCetak.NPWP = RowSupplier.NPWP_No
                        End If
                        If TblItemCp.DefaultView.Count - 1 >= z Then
                            RowItemCp = TblItemCp.DefaultView(z).Row
                            RowCetak.Contact = RowItemCp.Phone_No
                        End If
                        If TblItemBa.DefaultView.Count - 1 >= z Then
                            RowItemBa = TblItemBa.DefaultView(z).Row
                            RowCetak.Bank_Account = RowItemBa.Account_No
                            RowCetak.Bank_Name = RowItemBa._Bank_s_Name
                            RowCetak.Beneficiary_Name = RowItemBa.Beneficiary_Name
                            RowCetak.Curr = RowItemBa.Currency_Code
                        End If
                        If RowSupplier.Foreign_Supplier = 0 Then
                            RowCetak.Supplier_Group = "LOCAL SUPPLIER"
                        ElseIf RowSupplier.Foreign_Supplier = 1 Then
                            RowCetak.Supplier_Group = "FOREIGN SUPPLIER"
                        End If
                        TBLCetak.Rows.Add(RowCetak)

                    Next
                Next

                Report.SetDataSource(DS)
                Me.crvSupplierList.ReportSource = Report
                Me.crvSupplierList.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSupplierList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.cboGrupSupplier.SelectedIndex = 2
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class