Option Strict Off
Imports Core_Module


Public Class FrmPenCetakanAmplop
    Dim BackColorAll As Color = Color.LightGoldenrodYellow
    Dim TBLOffice As New DseCetakAmplopCreative.tblmaster_customer_item_addressDataTable
    Dim TBLCP As New DseCetakAmplopCreative.tblmaster_customer_itemcpDataTable
    Dim TBLOfficeSupp As New DseCetakAmplopCreative.tblmaster_supplier_item_addressDataTable
    Dim TBLCPSupp As New DseCetakAmplopCreative.tblmaster_supplier_itemcpDataTable

    Dim report As New crpCetakAmplop4
    Dim report2_Biasa As New crpCetakAmplop5
    Dim report2 As New crpCetakAmplop5_Dot_Matix
    Dim Report3 As New crpCetakAmplop4_Dot_Matrix

    Private Sub TxtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSupplierCode.GotFocus
        Try
            If Me.TxtSupplierCode.Enabled = True Then
                If Me.TxtSupplierCode.ReadOnly = False Then
                    Me.TxtSupplierCode.BackColor = Me.BackColorAll
                Else
                    TxtSupplierCode.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSupplierCode.LostFocus
        Try
            If Me.TxtSupplierCode.Enabled = True Then
                If Me.TxtSupplierCode.ReadOnly = False Then
                    Me.TxtSupplierCode.BackColor = Me.TxtPrimarySupplier.BackColor
                Else
                    TxtSupplierCode.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCode.GotFocus
        Try
            If Me.TxtCustomerCode.Enabled = True Then
                If Me.TxtCustomerCode.ReadOnly = False Then
                    Me.TxtCustomerCode.BackColor = Me.BackColorAll
                Else
                    TxtCustomerCode.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomerCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCode.LostFocus
        Try
            If Me.TxtCustomerCode.Enabled = True Then
                If Me.TxtCustomerCode.ReadOnly = False Then
                    Me.TxtCustomerCode.BackColor = Me.TxtPrimaryCustomer.BackColor
                Else
                    TxtCustomerCode.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BersihkanObjek()
        Try
            Me.TxtIdentifierCP.Text = "-1"
            Me.TxtIdentifierOffice.Text = "-1"
            Me.CBOcp.DataSource = Nothing
            Me.CBOOffice.DataSource = Nothing
            If Me.rboCustomer.Checked = True Then
                Me.TxtPrimaryCustomer.Text = "-1"
                Me.TxtCustomerCode.Text = ""
            End If
            If Me.rboSupplier.Checked = True Then
                Me.TxtPrimarySupplier.Text = "-1"
                Me.TxtSupplierCode.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InvisibleObjek()
        Try
            If Me.rboCustomer.Checked = True Then
                Me.LblSupp.Visible = False
                Me.TxtSupplierCode.Visible = False
                Me.btnFindSupplier.Visible = False
                Me.LblCust.Visible = True
                Me.TxtCustomerCode.Visible = True
                Me.BtnFindCustomer.Visible = True
                Me.TxtSupplierCode.TabStop = False
                Me.btnFindSupplier.TabStop = False
                Me.TxtCustomerCode.TabStop = True
            End If
            If Me.rboSupplier.Checked = True Then
                Me.LblCust.Visible = False
                Me.TxtCustomerCode.Visible = False
                Me.BtnFindCustomer.Visible = False
                Me.LblSupp.Visible = True
                Me.TxtSupplierCode.Visible = True
                Me.btnFindSupplier.Visible = True
                Me.TxtSupplierCode.TabStop = True
                Me.btnFindSupplier.TabStop = False
                Me.TxtCustomerCode.TabStop = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Setting()
        Try
            If Me.CboPrint.SelectedIndex = 0 Then
                Me.Label3.Visible = False
                Me.cboFormatLaporan.Visible = False
            Else
                Me.Label3.Visible = True
                Me.cboFormatLaporan.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmPenCetakanAmplop_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.cboFormatLaporan.SelectedIndex = 0
            Me.CboPrint.SelectedIndex = 1
            Me.Setting()
            Me.InvisibleObjek()
            Me.BersihkanObjek()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboSupplier_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboSupplier.CheckedChanged
        Try
            Me.InvisibleObjek()
            Me.BersihkanObjek()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboCustomer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboCustomer.CheckedChanged
        Try
            Me.InvisibleObjek()
            Me.BersihkanObjek()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ''Find Code No
    Private Sub TxtSupplierCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtSupplierCode.KeyDown
        Dim objCode As New clsMasterSupplier
        Try
            If e.KeyValue = 13 Then
                If Me.TxtSupplierCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtSupplierCode.Text) = True Then
                        Me.TxtPrimarySupplier.Text = CStr(objCode.GetPrimary(Me.TxtSupplierCode.Text))
                        Me.TxtSupplierCode.Text = CStr(objCode.GetSupplierCode(CDbl(Me.TxtPrimarySupplier.Text)))
                        Me.CBOOffice.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindSupplier.PerformClick()
                    End If
                Else
                    Me.btnFindSupplier.PerformClick()
                End If
            End If
            Me.IsiCBOOffice()
            Me.IsiCBOCP()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
        End Try
    End Sub

    Private Sub TxtSupplierCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtSupplierCode.Leave
        Dim i As New FrmCariMasterSupplier
        Dim objCode As New clsMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtSupplierCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtSupplierCode.Text) = True Then
                    Me.TxtPrimarySupplier.Text = CStr(objCode.GetPrimary(Me.TxtSupplierCode.Text))
                    Me.TxtSupplierCode.Text = objCode.GetSupplierCode(CDbl(Me.TxtPrimarySupplier.Text))
                Else
                    If Me.TxtSupplierCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSupplierCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.TxtPrimarySupplier.Text = CStr(i.PrimarySupplier)
                        Me.TxtSupplierCode.Text = objCode.GetSupplierCode(CDbl(Me.TxtPrimarySupplier.Text))
                    Else
                        Me.TxtSupplierCode.Clear()
                        Me.TxtPrimarySupplier.Text = "-1"
                        Me.TxtSupplierCode.Focus()
                        Me.BersihkanObjek()
                    End If
                End If
            Else
                Me.TxtSupplierCode.Clear()
                Me.TxtPrimarySupplier.Text = "-1"
                Me.BersihkanObjek()
            End If
            Me.IsiCBOOffice()
            Me.IsiCBOCP()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFindSupplier.Click
        Dim i As New FrmCariMasterSupplier
        Dim objCode As New clsMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtSupplierCode.Text <> "" Then
                i.txtCari.Text = Me.TxtSupplierCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.TxtPrimarySupplier.Text = CStr(i.PrimarySupplier)
                Me.TxtSupplierCode.Text = objCode.GetSupplierCode(CDbl(Me.TxtPrimarySupplier.Text))
                Me.CBOOffice.Focus()
            Else
                Me.TxtSupplierCode.Clear()
                Me.TxtPrimarySupplier.Text = "-1"
                Me.TxtSupplierCode.Focus()
                Me.BersihkanObjek()
            End If
            Me.IsiCBOOffice()
            Me.IsiCBOCP()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    ''Find Code No
    Private Sub TxtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtCustomerCode.KeyDown
        Dim objCode As New clsMasterCustomer
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCustomerCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtCustomerCode.Text) = True Then
                        Me.TxtPrimaryCustomer.Text = CStr(objCode.GetPrimary(Me.TxtCustomerCode.Text))
                        Me.TxtCustomerCode.Text = CStr(objCode.GetCustomerCode(CDbl(Me.TxtPrimaryCustomer.Text)))
                        Me.CBOOffice.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnFindCustomer.PerformClick()
                    End If
                Else
                    Me.BtnFindCustomer.PerformClick()
                End If
            End If
            Me.IsiCBOOffice()
            Me.IsiCBOCP()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
        End Try
    End Sub

    Private Sub TxtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCustomerCode.Leave
        Dim i As New FrmCariMasterCustomer
        Dim objCode As New clsMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCustomerCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtCustomerCode.Text) = True Then
                    Me.TxtPrimaryCustomer.Text = CStr(objCode.GetPrimary(Me.TxtCustomerCode.Text))
                    Me.TxtCustomerCode.Text = objCode.GetCustomerCode(CDbl(Me.TxtPrimaryCustomer.Text))
                Else
                    If Me.TxtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtPrimaryCustomer.Text = CStr(i.PrimaryCustomer)
                        Me.TxtCustomerCode.Text = objCode.GetCustomerCode(CDbl(Me.TxtPrimaryCustomer.Text))
                    Else
                        Me.TxtCustomerCode.Clear()
                        Me.TxtPrimaryCustomer.Text = "-1"
                        Me.TxtCustomerCode.Focus()
                        Me.BersihkanObjek()
                    End If
                End If
            Else
                Me.TxtCustomerCode.Clear()
                Me.TxtPrimaryCustomer.Text = "-1"
                Me.BersihkanObjek()
            End If
            Me.IsiCBOOffice()
            Me.IsiCBOCP()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles BtnFindCustomer.Click
        Dim i As New FrmCariMasterCustomer
        Dim objCode As New clsMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCustomerCode.Text <> "" Then
                i.txtCari.Text = Me.TxtCustomerCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtPrimaryCustomer.Text = CStr(i.PrimaryCustomer)
                Me.TxtCustomerCode.Text = objCode.GetCustomerCode(CDbl(Me.TxtPrimaryCustomer.Text))
                Me.CBOOffice.Focus()
            Else
                Me.TxtCustomerCode.Clear()
                Me.TxtPrimaryCustomer.Text = "-1"
                Me.TxtCustomerCode.Focus()
                Me.BersihkanObjek()
            End If
            Me.IsiCBOOffice()
            Me.IsiCBOCP()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub IsiTBLCP(ByVal vTBLCP As DseCetakAmplopCreative.tblmaster_customer_itemcpDataTable, _
    ByVal vTBLCPSupp As DseCetakAmplopCreative.tblmaster_supplier_itemcpDataTable, ByVal vPrimaryCust As Double)
        Dim DTA As New DseCetakAmplopCreativeTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim DTASupp As New DseCetakAmplopCreativeTableAdapters.tblmaster_supplier_itemcpTableAdapter
        Try
            If Me.rboCustomer.Checked = True Then
                DTA.Fill(vTBLCP, vPrimaryCust)
            End If
            If Me.rboSupplier.Checked = True Then
                DTASupp.Fill(vTBLCPSupp, vPrimaryCust)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTASupp.Dispose()
        End Try
    End Sub

    Private Sub IsiCBOCP()
        Try
            Me.CBOcp.DataSource = Nothing
            Me.CBOcp.Items.Clear()

            If Me.rboCustomer.Checked = True Then
                If Me.TxtPrimaryCustomer.Text <> "-1" Then
                    If Me.GetJmlCustCP(Me.TxtPrimaryCustomer.Text) > 0 Then
                        Me.IsiTBLCP(Me.TBLCP, Me.TBLCPSupp, Me.TxtPrimaryCustomer.Text)
                        Me.CBOcp.DataSource = Me.TBLCP.DefaultView
                        Me.TBLCP.DefaultView.Sort = "Identifier"
                        Me.CBOcp.DisplayMember = "Name"
                        Me.TxtIdentifierCP.Text = Me.TBLCP(Me.CBOcp.SelectedIndex).Identifier
                    End If
                End If
            End If

            If Me.rboSupplier.Checked = True Then
                If Me.TxtPrimarySupplier.Text <> "-1" Then
                    If Me.GetJmlSuppCP(Me.TxtPrimarySupplier.Text) > 0 Then
                        Me.IsiTBLCP(Me.TBLCP, Me.TBLCPSupp, Me.TxtPrimarySupplier.Text)
                        Me.CBOcp.DataSource = Me.TBLCPSupp.DefaultView
                        Me.TBLCPSupp.DefaultView.Sort = "Identifier"
                        Me.CBOcp.DisplayMember = "Name"
                        Me.TxtIdentifierCP.Text = Me.TBLCPSupp(Me.CBOcp.SelectedIndex).Identifier
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCP_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBOcp.SelectedIndexChanged
        Dim row As DseCetakAmplopCreative.tblmaster_customer_itemcpRow
        Dim rowSupp As DseCetakAmplopCreative.tblmaster_supplier_itemcpRow
        Try
            If Me.CBOcp.SelectedIndex = -1 Then
                Me.TxtIdentifierCP.Text = "-1"
            Else
                If Me.rboCustomer.Checked = True Then
                    row = Me.TBLCP.DefaultView(Me.CBOcp.SelectedIndex).Row
                    Me.TxtIdentifierCP.Text = row.Identifier
                End If
                If Me.rboSupplier.Checked = True Then
                    rowSupp = Me.TBLCPSupp.DefaultView(Me.CBOcp.SelectedIndex).Row
                    Me.TxtIdentifierCP.Text = rowSupp.Identifier
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiTBLOffice(ByVal vTBLOffice As DseCetakAmplopCreative.tblmaster_customer_item_addressDataTable, _
    ByVal vTBLOfficeSupp As DseCetakAmplopCreative.tblmaster_supplier_item_addressDataTable, ByVal vPrimaryCust As Double)
        Dim DTA As New DseCetakAmplopCreativeTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim DTASupp As New DseCetakAmplopCreativeTableAdapters.tblmaster_supplier_item_addressTableAdapter
        Try
            If Me.rboCustomer.Checked = True Then
                DTA.Fill(vTBLOffice, vPrimaryCust)
            End If

            If Me.rboSupplier.Checked = True Then
                DTASupp.Fill(vTBLOfficeSupp, vPrimaryCust)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTASupp.Dispose()
        End Try
    End Sub

    Private Sub IsiCBOOffice()
        Try
            Me.CBOOffice.DataSource = Nothing
            Me.CBOOffice.Items.Clear()

            If Me.rboCustomer.Checked = True Then
                If Me.TxtPrimaryCustomer.Text <> "-1" Then
                    If Me.GetJmlCustAdd(Me.TxtPrimaryCustomer.Text) > 0 Then
                        Me.IsiTBLOffice(Me.TBLOffice, Me.TBLOfficeSupp, Me.TxtPrimaryCustomer.Text)
                        Me.CBOOffice.DataSource = Me.TBLOffice.DefaultView
                        Me.TBLOffice.DefaultView.Sort = "Identifier"
                        Me.CBOOffice.DisplayMember = "Address Code"
                        Me.TxtIdentifierOffice.Text = Me.TBLOffice(Me.CBOOffice.SelectedIndex).Identifier
                    End If
                End If
            End If

            If Me.rboSupplier.Checked = True Then
                If Me.TxtPrimarySupplier.Text <> "-1" Then
                    If Me.GetJmlSuppAdd(Me.TxtPrimarySupplier.Text) > 0 Then
                        Me.IsiTBLOffice(Me.TBLOffice, Me.TBLOfficeSupp, Me.TxtPrimarySupplier.Text)
                        Me.CBOOffice.DataSource = Me.TBLOfficeSupp.DefaultView
                        Me.TBLOfficeSupp.DefaultView.Sort = "Identifier"
                        Me.CBOOffice.DisplayMember = "Address Code"
                        Me.TxtIdentifierOffice.Text = Me.TBLOfficeSupp(Me.CBOOffice.SelectedIndex).Identifier
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboOffice_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBOOffice.SelectedIndexChanged
        Dim row As DseCetakAmplopCreative.tblmaster_customer_item_addressRow
        Dim RowSupp As DseCetakAmplopCreative.tblmaster_supplier_item_addressRow
        Try
            If Me.CBOOffice.SelectedIndex = -1 Then
                Me.TxtIdentifierOffice.Text = "-1"
            Else
                If Me.rboCustomer.Checked = True Then
                    row = Me.TBLOffice.DefaultView(Me.CBOOffice.SelectedIndex).Row
                    Me.TxtIdentifierOffice.Text = row.Identifier
                End If
                If Me.rboSupplier.Checked = True Then
                    RowSupp = Me.TBLOfficeSupp.DefaultView(Me.CBOOffice.SelectedIndex).Row
                    Me.TxtIdentifierOffice.Text = RowSupp.Identifier
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cetak()
        report.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperEnvelopeDL
        Dim ds As New DseCetakAmplopCreative
        Dim TBL As DseCetakAmplopCreative.TBL_CetakDataTable = ds.TBL_Cetak
        Dim row As DseCetakAmplopCreative.TBL_CetakRow
        Dim Nama As String = ""

        Dim DTASupp As New DseCetakAmplopCreativeTableAdapters.tblmaster_supplierTableAdapter
        Dim TBLSupp As New DseCetakAmplopCreative.tblmaster_supplierDataTable
        Dim RowSupp As DseCetakAmplopCreative.tblmaster_supplierRow
        Dim DTASuppAdd As New DseCetakAmplopCreativeTableAdapters.tblmaster_supplier_item_addressTableAdapter
        Dim TBLSuppAdd As New DseCetakAmplopCreative.tblmaster_supplier_item_addressDataTable
        Dim RowSuppAdd As DseCetakAmplopCreative.tblmaster_supplier_item_addressRow
        Dim DTASuppCP As New DseCetakAmplopCreativeTableAdapters.tblmaster_supplier_itemcpTableAdapter
        Dim TBLSuppCP As New DseCetakAmplopCreative.tblmaster_supplier_itemcpDataTable
        Dim RowSuppCP As DseCetakAmplopCreative.tblmaster_supplier_itemcpRow
        Dim DTACust As New DseCetakAmplopCreativeTableAdapters.tblmaster_customerTableAdapter
        Dim TBLCust As New DseCetakAmplopCreative.tblmaster_customerDataTable
        Dim RowCust As DseCetakAmplopCreative.tblmaster_customerRow
        Dim DTACustAdd As New DseCetakAmplopCreativeTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim TBLCustAdd As New DseCetakAmplopCreative.tblmaster_customer_item_addressDataTable
        Dim RowCustAdd As DseCetakAmplopCreative.tblmaster_customer_item_addressRow
        Dim DTACustCP As New DseCetakAmplopCreativeTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim TBLCustCP As New DseCetakAmplopCreative.tblmaster_customer_itemcpDataTable
        Dim RowCustCP As DseCetakAmplopCreative.tblmaster_customer_itemcpRow
        ''
        Dim Line As String = ""
        Dim CP As String = ""

        Try
            'kosongkan crv
            Me.DisposeCRV()

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            TBL.Rows.Clear()
            If Me.rboSupplier.Checked = True Then
                DTASupp.Fill(TBLSupp, Me.TxtPrimarySupplier.Text)
                DTASuppAdd.FillBy_Identifier(TBLSuppAdd, Me.TxtIdentifierOffice.Text)
                DTASuppCP.FillBy_Identifier(TBLSuppCP, Me.TxtIdentifierCP.Text)
                row = TBL.NewRow
                With row
                    If TBLSupp.Rows.Count > 0 Then
                        RowSupp = TBLSupp.Rows(0)
                        .Primary = RowSupp.Primary
                        .Name = RowSupp.Supplier_Name
                    End If
                    If TBLSuppAdd.Rows.Count > 0 Then
                        RowSuppAdd = TBLSuppAdd.Rows(0)
                        .Line1 = RowSuppAdd.Line_1
                        .Line2 = RowSuppAdd.Line_2
                        .Line3 = RowSuppAdd.Line_3
                        .Line4 = RowSuppAdd.Line_4
                        .Line5 = RowSuppAdd.Line_5
                        .PhoneNo = RowSuppAdd.Phone_No
                    End If
                    If TBLSuppCP.Rows.Count > 0 Then
                        RowSuppCP = TBLSuppCP.Rows(0)
                        .UP_To = RowSuppCP.Name
                    End If
                End With
                TBL.AddTBL_CetakRow(row)

                For Each row In TBL.Rows
                    Nama = row.Name
                    If row.Line1 <> "" Then
                        Line += row.Line1
                    End If
                    If row.Line2 <> "" Then
                        Line += vbCrLf + row.Line2
                    End If
                    If row.Line3 <> "" Then
                        Line += vbCrLf + row.Line3
                    End If
                    If row.Line4 <> "" Then
                        Line += vbCrLf + row.Line4
                    End If
                    If row.Line5 <> "" Then
                        Line += vbCrLf + row.Line5
                    End If
                    If row.PhoneNo <> "" Then
                        Line += vbCrLf + "Phone: " + row.PhoneNo
                    End If
                    If row.UP_To <> "" Then
                        CP += row.UP_To
                    End If
                    For i As Integer = 0 To 40
                        If i = 0 Then
                            Line += vbCrLf + "="
                        Else
                            Line += "="
                        End If
                        If i = 40 Then
                            Line += vbCrLf + "UP. " + CP
                        End If
                    Next
                Next
            End If
            If Me.rboCustomer.Checked = True Then
                DTACust.Fill(TBLCust, Me.TxtPrimaryCustomer.Text)
                DTACustAdd.FillBy_Identifier(TBLCustAdd, Me.TxtIdentifierOffice.Text)
                DTACustCP.FillBy_Identifier(TBLCustCP, Me.TxtIdentifierCP.Text)
                row = TBL.NewRow
                With row
                    If TBLCust.Rows.Count > 0 Then
                        RowCust = TBLCust.Rows(0)
                        .Primary = RowCust.Primary
                        .Name = RowCust.Customer_Name
                    End If
                    If TBLCustAdd.Rows.Count > 0 Then
                        RowCustAdd = TBLCustAdd.Rows(0)
                        .Line1 = RowCustAdd.Line_1
                        .Line2 = RowCustAdd.Line_2
                        .Line3 = RowCustAdd.Line_3
                        .Line4 = RowCustAdd.Line_4
                        .Line5 = RowCustAdd.Line_5
                        .PhoneNo = RowCustAdd.Phone_No
                    End If
                    If TBLCustCP.Rows.Count > 0 Then
                        RowCustCP = TBLCustCP.Rows(0)
                        .UP_To = RowCustCP.Name
                    End If
                End With
                TBL.AddTBL_CetakRow(row)

                For Each row In TBL.Rows
                    Nama = row.Name
                    If row.Line1 <> "" Then
                        Line += row.Line1
                    End If
                    If row.Line2 <> "" Then
                        Line += vbCrLf + row.Line2
                    End If
                    If row.Line3 <> "" Then
                        Line += vbCrLf + row.Line3
                    End If
                    If row.Line4 <> "" Then
                        Line += vbCrLf + row.Line4
                    End If
                    If row.Line5 <> "" Then
                        Line += vbCrLf + row.Line5
                    End If
                    If row.PhoneNo <> "" Then
                        Line += vbCrLf + "Phone: " + row.PhoneNo
                    End If
                    If row.UP_To <> "" Then
                        CP += row.UP_To
                    End If
                    For i As Integer = 0 To 40
                        If i = 0 Then
                            Line += vbCrLf + "="
                        Else
                            Line += "="
                        End If
                        If i = 40 Then
                            Line += vbCrLf + "UP. " + CP
                        End If
                    Next
                Next
            End If
            'End If

            report.Refresh()
            report2.Refresh()
            Report3.Refresh()
            report2_Biasa.Refresh()
            ''Report3_Biasa.Refresh()

            report.SetDataSource(ds)
            report2.SetDataSource(ds)
            Report3.SetDataSource(ds)
            report2_Biasa.SetDataSource(ds)
            ''Report3_Biasa.SetDataSource(ds)

            report.SetParameterValue(report.Parameter_Line.ParameterFieldName, Line)
            report2.SetParameterValue(report2.Parameter_Line.ParameterFieldName, Line)
            Report3.SetParameterValue(Report3.Parameter_Line.ParameterFieldName, Line)
            report2_Biasa.SetParameterValue(report2_Biasa.Parameter_Line.ParameterFieldName, Line)
            report.SetParameterValue(report.Parameter_Name.ParameterFieldName, Nama)
            report2.SetParameterValue(report2.Parameter_Name.ParameterFieldName, Nama)
            Report3.SetParameterValue(Report3.Parameter_Name.ParameterFieldName, Nama)
            report2_Biasa.SetParameterValue(report2_Biasa.Parameter_Name.ParameterFieldName, Nama)
            ''Report3_Biasa.SetParameterValue(Report3_Biasa.Parameter_Line.ParameterFieldName, Line)

            Me.crv1.ReportSource = report
            Me.CRV_Matrix_Potrait.ReportSource = report2
            Me.CRV_Matrix_Landscape.ReportSource = Report3
            Me.CRV_Biasa_Potrait.ReportSource = report2_Biasa
            ''Me.CRV_Biasa_Landscape.ReportSource = Report3_Biasa

            Me.crv1.Show()
            Me.CRV_Biasa_Potrait.Show()
            Me.CRV_Biasa_Landscape.Show()
            Me.CRV_Matrix_Potrait.Show()
            Me.CRV_Matrix_Landscape.Show()
            'Windows.Forms.Cursor.Current = Cursors.Arrow
            'Me.ShowDialog()
            'Me.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASupp.Dispose()
            TBLSupp.Dispose()
            DTASuppAdd.Dispose()
            TBLSuppAdd.Dispose()
            DTASuppCP.Dispose()
            TBLSuppCP.Dispose()
            DTACust.Dispose()
            TBLCust.Dispose()
            DTACustAdd.Dispose()
            TBLCustAdd.Dispose()
            DTACustCP.Dispose()
            TBLCustCP.Dispose()
        End Try
    End Sub

    Private Sub DisposeCRV()
        Try
            Me.crv1.ReportSource = Nothing
            Me.CRV_Matrix_Potrait.ReportSource = Nothing
            Me.CRV_Matrix_Landscape.ReportSource = Nothing
            Me.CRV_Biasa_Potrait.ReportSource = Nothing
            'Me.CRV_Biasa_Landscape.ReportSource = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Try
            If Me.CboPrint.SelectedIndex = 0 Then
                If Me.TxtIdentifierOffice.Text <> "-1" Then
                    Me.CRV_Biasa_Potrait.PrintReport()
                Else
                    MsgBox("Please insert customer / supplier code", MsgBoxStyle.Information, "Information")
                    Me.DisposeCRV()
                End If
            End If
            If Me.CboPrint.SelectedIndex = 1 Then
                If Me.cboFormatLaporan.SelectedIndex = 0 Then
                    If Me.TxtIdentifierOffice.Text <> "-1" Then
                        Me.CRV_Matrix_Potrait.PrintReport()
                    Else
                        MsgBox("Please insert customer / supplier code", MsgBoxStyle.Information, "Information")
                        Me.DisposeCRV()
                    End If
                End If
                If Me.cboFormatLaporan.SelectedIndex = 1 Then
                    If Me.TxtIdentifierOffice.Text <> "-1" Then
                        Me.CRV_Matrix_Landscape.PrintReport()
                    Else
                        MsgBox("Please insert customer / supplier code", MsgBoxStyle.Information, "Information")
                        Me.DisposeCRV()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            If Me.TxtIdentifierOffice.Text <> "-1" Then
                Me.Cetak()
            Else
                MsgBox("Please insert customer / supplier code", MsgBoxStyle.Information, "Information")
                If Me.rboCustomer.Checked = True Then
                    Me.TxtCustomerCode.Focus()
                Else
                    Me.TxtSupplierCode.Focus()
                End If
                Me.DisposeCRV()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboPrint_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboPrint.SelectedIndexChanged
        Try
            Me.Setting()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlCustAdd(ByVal vPrimaryCust As Double) As Double
        Dim DTA As New DseCetakAmplopCreativeTableAdapters.tblmaster_customer_item_addressTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByPrimaryCust(vPrimaryCust).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimaryCust(vPrimaryCust).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetJmlCustCP(ByVal vPrimaryCust As Double) As Double
        Dim DTA As New DseCetakAmplopCreativeTableAdapters.tblmaster_customer_itemcpTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByPrimaryCust(vPrimaryCust).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimaryCust(vPrimaryCust).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetJmlSuppAdd(ByVal vPrimarySupp As Double) As Double
        Dim DTA As New DseCetakAmplopCreativeTableAdapters.tblmaster_supplier_item_addressTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByPrimarySupp(vPrimarySupp).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimarySupp(vPrimarySupp).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetJmlSuppCP(ByVal vPrimarySupp As Double) As Double
        Dim DTA As New DseCetakAmplopCreativeTableAdapters.tblmaster_supplier_itemcpTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByPrimarySupp(vPrimarySupp).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimarySupp(vPrimarySupp).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class