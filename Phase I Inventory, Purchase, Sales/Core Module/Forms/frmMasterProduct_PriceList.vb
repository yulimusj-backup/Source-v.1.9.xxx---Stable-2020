Option Explicit On
Option Strict Off

Public Class frmMasterProduct_PriceList
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ShowLangsung As Boolean = False
    Public PrimaryProduct As Double = -1

    Dim i As String = "-"
    Dim Def_Prec_Disc As Double = 2
    Dim Buffer As String = ""

    Dim TblMasterCurrency As New DseMasterProduct_EditPriceList.tblmaster_currencyDataTable
    Dim RowMasterCurrency As DseMasterProduct_EditPriceList.tblmaster_currencyRow
    Dim ObjMP As New clsMasterProduct_PriceList
    Dim ObjCur As New clsMasterCurrency
    Dim Defprec As Double = ObjCur.GetDecimal_DefaultCurrency
    Dim ObjCD As New ClsCekData
    Dim ObjProd As New clsMasterProduct
    Dim InsertRowIndex As Integer = -1
    Dim TBLCustomerGroup As New DseMasterProduct_EditPriceList.tblmaster_product_pricelist_groupDataTable

    Dim TBLCboCustomerGroup As New DseMasterProduct_EditPriceList.tblmaster_customer_groupDataTable
    Dim TBLCBOCustomerSalesArea As New DseMasterProduct_EditPriceList.tblmaster_customer_sales_areaDataTable
    Dim TBLCBOCustomerSegmentation As New DseMasterProduct_EditPriceList.tblmaster_customer_segmentationDataTable
    Dim TBLTCBOypeCode As New DseMasterProduct_EditPriceList.tblmaster_customer_typeDataTable



    Public Enum EnumCustomerGroupTipe As Integer
        Customer_Group
        Customer_Sales_Area
        Customer_Segmentation
        Customer_Type
    End Enum

    Private Sub InsertRow()
        Dim rowDGV As DataGridViewRow
        Try
            Me.DGVHarga.Rows.InsertCopy(10, 11)
            Me.DGVHarga.Rows.InsertRange(5, rowDGV)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboCurrency()
        Try
            Me.CboCurrency.DataSource = Nothing
            Me.CboCurrency.Items.Clear()

            Me.ObjMP.IsiTBLCurrency(Me.TblMasterCurrency)

            Me.CboCurrency.DataSource = Me.TblMasterCurrency.DefaultView
            Me.txtPrimaryCurrency.Text = Me.TblMasterCurrency.Rows(0).Item(0).ToString
            Me.TblMasterCurrency.DefaultView.Sort = "Primary"
            Me.CboCurrency.DisplayMember = "Currency Code"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboCurrency_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboCurrency.SelectedValueChanged
        Try
            If Me.CboCurrency.SelectedIndex = -1 Then
                Me.txtPrimaryCurrency.Text = "-1"
            Else
                Me.RowMasterCurrency = Me.TblMasterCurrency.DefaultView(Me.CboCurrency.SelectedIndex).Row
                Me.txtPrimaryCurrency.Text = Me.RowMasterCurrency.Primary
                Me.DGVHarga.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

   
    Private Sub frmMasterProduct_PriceList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.KosongkanField()
            Me.DGVHarga.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.IsiCboCurrency()
            Me.ObjMP.IsiTBLCustomerGroup(Me.TBLCboCustomerGroup)
            Me.ObjMP.IsiTBLCustomerSalesArea(Me.TBLCBOCustomerSalesArea)
            Me.ObjMP.IsiTBLSegmentationCode(Me.TBLCBOCustomerSegmentation)
            Me.ObjMP.IsiTBLTypeCode(Me.TBLTCBOypeCode)
            If Me.cboCustomerCode.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Group Then
                Me.IsiCboCustomerGroup()
            ElseIf Me.cboCustomerCode.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Sales_Area Then
                Me.IsiCboCustomerSalesArea()
            ElseIf Me.cboCustomerCode.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Segmentation Then
                Me.IsiCboCustomerSegmentation()
            ElseIf Me.cboCustomerCode.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Type Then
                Me.IsiCboCustomerTypeCode()
            End If

            If Me.ShowLangsung = True Then
                Me.TxtKodeProduk.Text = Me.ObjProd.GetProductCode(Me.PrimaryProduct)
                Me.MinimizeBox = False
                Me.MaximizeBox = False
                Me.Size = Me.MinimumSize
                If Me.TxtKodeProduk.Text <> "" Then
                    Me.BtnFilter.PerformClick()
                End If
            End If

            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cmdKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKeluar.Click
        Me.Close()
    End Sub

    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.DGVHarga.Controls, True, True, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.GroupBox1.Controls, True, True, False)
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtprimaryCust.Text = "-1"
            Me.TxtPrimaryCustGroup.Text = "-1"
            Me.TxtIndexCustType.Text = "-1"
            Me.rboGeneral.Checked = True
            Me.ChkUpdateCustGroup.Checked = False
            Me.ChkUpdateCust.Checked = False
            Me.TxtCustomerCode.Text = ""
            Me.TxtCustomerName.Text = ""
            Me.DGVHarga.Rows.Clear()
            Me.TxtKodeProduk.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiBtnFilter()
        Try
            Me.DGVHarga.Rows.Clear()
            If Me.rboGeneral.Checked = True Then
                Me.ObjMP.IsiSemua(Me.DGVHarga, Me.Def_Prec_Disc, Me, Me.ObjCur.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
            ElseIf Me.rboCustomerGroup.Checked = True Then
                If CDbl(Me.TxtPrimaryCustGroup.Text) <> -1 Then
                    Me.ObjMP.isiCustomerGroup(Me.TxtKodeProduk.Text & "%", Me.txtPrimaryCurrency.Text, Me.TxtPrimaryCustGroup.Text, _
                    Me.DGVHarga, Me.Def_Prec_Disc, Me, Me.ObjCur.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                Else
                    MsgBox("Please choose customer group.", MsgBoxStyle.Information, "Information")
                    Me.DGVHarga.Rows.Clear()
                    Me.cboCustomerCode.Focus()
                    Exit Try
                End If
            ElseIf Me.rboCustomer.Checked = True Then
                If CDbl(Me.txtprimaryCust.Text) <> -1 Then
                    Me.ObjMP.IsiCustomer(Me.TxtKodeProduk.Text & "%", Me.txtPrimaryCurrency.Text, Me.txtprimaryCust.Text, _
                    Me.DGVHarga, Me.Def_Prec_Disc, Me, Me.cboCustomerGroup.SelectedIndex, Me.ObjCur.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                Else
                    MsgBox("Please input customer code.", MsgBoxStyle.Information, "Information")
                    Me.DGVHarga.Rows.Clear()
                    Me.TxtCustomerCode.Focus()
                    Exit Try
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFilter.Click
        Try
            Me.FungsiBtnFilter()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub rboGeneral_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboGeneral.CheckedChanged
        Try
            If Me.rboGeneral.Checked = True Then
                Me.ChkGeneralAll.Enabled = True
                Me.cboCustomerGroup.Enabled = False
                Me.cboCustomerCode.Enabled = False
                Me.btnFindCustomer.Enabled = False
                Me.btnDelete_ByGroup.Enabled = False
                Me.btnDelete_ByCustomer.Enabled = False
                Me.BtnPrintCusGrouop.Enabled = False
                Me.BtnPrintCust.Enabled = False
                Me.Button1.Enabled = False
                Me.Button2.Enabled = False
                Me.TxtCustomerCode.Enabled = False
                Me.TxtCustomerCode.Text = ""
                Me.TxtCustomerName.Text = ""
                Me.TxtCustomerName.Enabled = False
                Me.cboCustomerCode.DataSource = Nothing
                Me.ChkUpdateCust.Enabled = True
                Me.ChkUpdateCustGroup.Enabled = True
                Me.DGVHarga.Rows.Clear()
            ElseIf Me.rboCustomerGroup.Checked = True Then
                Me.ChkGeneralAll.Enabled = False
                Me.ChkUpdateCust.Enabled = True
                Me.ChkUpdateCust.Enabled = False
                Me.cboCustomerGroup.Enabled = True
                Me.cboCustomerCode.Enabled = True
                Me.btnDelete_ByGroup.Enabled = True
                Me.btnDelete_ByCustomer.Enabled = False
                Me.BtnPrintCusGrouop.Enabled = True
                Me.BtnPrintCust.Enabled = False
                Me.Button1.Enabled = True
                Me.Button2.Enabled = False
                Me.ChkUpdateCust.Enabled = True
                Me.ChkUpdateCustGroup.Enabled = False
                Me.DGVHarga.Rows.Clear()

            ElseIf Me.rboCustomer.Checked = True Then
                Me.ChkGeneralAll.Enabled = False
                Me.TxtCustomerCode.Enabled = True
                Me.TxtCustomerName.Enabled = False
                Me.btnDelete_ByCustomer.Enabled = True
                Me.btnFindCustomer.Enabled = True
                Me.cboCustomerGroup.Enabled = False
                Me.cboCustomerCode.Enabled = False
                Me.btnDelete_ByGroup.Enabled = False
                Me.BtnPrintCusGrouop.Enabled = False
                Me.BtnPrintCust.Enabled = True
                Me.Button1.Enabled = False
                Me.Button2.Enabled = True
                Me.cboCustomerCode.DataSource = Nothing
                Me.ChkUpdateCust.Enabled = False
                Me.ChkUpdateCustGroup.Enabled = False
                Me.DGVHarga.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboCustomerGroup_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboCustomerGroup.CheckedChanged
        Dim RowCustomerGroup As DseMasterProduct_EditPriceList.tblmaster_product_pricelist_groupRow
        Try
            If Me.rboCustomerGroup.Checked = True Then
                Me.ChkGeneralAll.Enabled = False
                Me.Button1.Enabled = True
                Me.Button2.Enabled = False
                Me.cboCustomerCode.Enabled = True
                Me.btnDelete_ByGroup.Enabled = True
                Me.TxtCustomerCode.Enabled = False
                Me.ChkUpdateCustGroup.Enabled = False
                Me.TxtCustomerCode.Text = ""
                Me.TxtCustomerName.Text = ""
                Me.txtprimaryCust.Text = "-1"
                Me.TxtCustomerName.Enabled = False
                Me.ChkUpdateCust.Enabled = False
                Me.ChkUpdateCust.Enabled = False
                Me.btnDelete_ByCustomer.Enabled = False
                Me.btnFindCustomer.Enabled = False
                Me.BtnPrintCusGrouop.Enabled = True
                Me.BtnPrintCust.Enabled = False
                Me.DGVHarga.Rows.Clear()
            ElseIf Me.rboCustomer.Checked = True Then
                Me.ChkGeneralAll.Enabled = False
                Me.ChkUpdateCust.Enabled = False
                Me.ChkUpdateCust.Enabled = False
                Me.TxtCustomerCode.Enabled = True
                Me.TxtCustomerName.Enabled = False
                Me.btnDelete_ByCustomer.Enabled = True
                Me.btnFindCustomer.Enabled = True
                Me.cboCustomerGroup.Enabled = False
                Me.cboCustomerCode.Enabled = False
                Me.btnDelete_ByGroup.Enabled = False
                Me.cboCustomerCode.DataSource = Nothing
                Me.DGVHarga.Rows.Clear()
                Me.Button1.Enabled = False
                Me.Button2.Enabled = True
                Me.BtnPrintCusGrouop.Enabled = False
                Me.BtnPrintCust.Enabled = True
            End If
            'Koding Customer Group
            If Me.rboCustomerGroup.Checked = True Then
                'Fill data customer group
                Me.ObjMP.FillMasterProductPriceList_Cust_Group(Me.TBLCustomerGroup)
                'TblMasterProeduct Group, ada datanya/ rownya ga
                'kalo = 0, maka cbo group enable=true, set index dipaling pertama
                If Me.TBLCustomerGroup.Rows.Count <= 0 Then
                    Me.cboCustomerGroup.Enabled = True
                    Me.cboCustomerGroup.SelectedIndex = 0
                Else
                    'kalo <>0, maka RowCustomergroup ambil ke index 0, Row.type = cboType.index, enable = false
                    RowCustomerGroup = Me.TBLCustomerGroup.Rows(0)
                    Me.cboCustomerGroup.Enabled = False
                    Me.cboCustomerGroup.SelectedIndex = RowCustomerGroup.tipe_group
                End If
                If Me.cboCustomerGroup.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Group Then
                    Me.IsiCboCustomerGroup()
                ElseIf Me.cboCustomerGroup.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Sales_Area Then
                    Me.IsiCboCustomerSalesArea()
                ElseIf Me.cboCustomerGroup.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Segmentation Then
                    Me.IsiCboCustomerSegmentation()
                ElseIf Me.cboCustomerGroup.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Type Then
                    Me.IsiCboCustomerTypeCode()
                End If
                Me.cboCustomerCode.SelectedIndex = -1
                If Me.cboCustomerCode.Items.Count > 0 Then
                    Me.cboCustomerCode.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function MasukanFormatStrip(ByVal vMaxQty As Decimal, ByVal vDefPrec As Double) As String
        Dim Retval As String = "-"
        Try
            If vMaxQty = 0 Then
                Retval = "-"
            Else
                Retval = Me.ObjCD.DigitGrouping(vMaxQty, vDefPrec)
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Error")
            Return "-"
        End Try
    End Function

    Public Function CekStrings(ByVal Data As String) As Boolean
        Try
            If Data.Contains("'") = True Or Data.Contains("*") = True Or Data.Contains("%") Or Data.Contains("#") = True Then
                MsgBox("Data cannot contain character : ' , % , atau * .", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "error")
        End Try
    End Function

    Private Sub DGVHarga_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVHarga.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColMaxQty.Index Then
                If Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value = "-" Then
                    Me.Buffer = 0
                Else
                    Me.Buffer = Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value
                    Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value)
                End If
            ElseIf e.ColumnIndex = Me.ColDiscount.Index Or e.ColumnIndex = Me.ColUnitPrice.Index _
                    Or e.ColumnIndex = Me.ColRemarks.Index Then
                Me.Buffer = Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value
                If e.ColumnIndex = Me.ColDiscount.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Then
                    Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value)
                End If
            Else
                e.Cancel = True
            End If

            If Me.TBLCboCustomerGroup.Rows.Count > 0 Then
                Me.cboCustomerGroup.Enabled = False
            Else
                Me.cboCustomerGroup.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVHarga_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVHarga.CellEndEdit
        Dim Row As DataGridViewRow = Me.DGVHarga.Rows(e.RowIndex)
        Dim nilainol As Double = 0
        Dim DTACusGroup As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Dim TblCusGroup As New DseMasterProduct_EditPriceList.tblmaster_product_pricelist_groupDataTable

        Dim DTACust As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_customerTableAdapter
        Dim TBlCust As New DseMasterProduct_EditPriceList.tblmaster_product_pricelist_customerDataTable
        'Jika tidak sama dengan buffer 
        'If CDec(Row.Cells(Me.ColMaxQty.Index).Value) <> CDec(Me.Buffer) Then
        '    Me.InsertRowIndex = e.RowIndex
        '    Me.BackgroundWorker1.RunWorkerAsync()
        'End If
        Dim TargetIndex As Integer = 0
        Dim rowTarget As DataGridViewRow
        Dim DummyTarget As String = ""

        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        If Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value = "0" Then
                            Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value = "-"
                        End If
                        Exit Sub
                    End If
                    'Set colomn yang boleh diedit Cek numeric dan non negative
                    If e.ColumnIndex = Me.ColMaxQty.Index Then
                        If IsNumeric(Row.Cells(Me.ColMaxQty.Index).Value) = False Then
                            MsgBox("Please input number in max qty", MsgBoxStyle.Information, "Information")
                            Row.Cells(Me.ColMaxQty.Index).Value = Me.MasukanFormatStrip(Me.Buffer, Row.Cells(Me.ColDecimal.Index).Value)
                            Exit Sub
                        End If
                        If Me.ObjCD.CekBilanganPositif(Row.Cells(Me.ColMaxQty.Index).Value) = False Then
                            MsgBox("Please input postive number", MsgBoxStyle.Information, "Information")
                            Row.Cells(Me.ColMaxQty.Index).Value = Me.MasukanFormatStrip(Me.Buffer, Row.Cells(Me.ColDecimal.Index).Value)
                            Exit Sub
                        End If
                        Me.InsertRowIndex = e.RowIndex
                        BackgroundWorker1.RunWorkerAsync()

                        If Me.ObjProd.GetPrimaryUnit(Row.Cells(Me.ColPrimaryCodeNo.Index).Value) <> Row.Cells(Me.ColPrimaryUnit.Index).Value Then
                            Row.Cells(Me.ColMaxQty.Index).Value = Me.ObjCD.DigitGrouping(CDbl(Row.Cells(Me.ColMaxQty.Index).Value), _
                            Me.ObjProd.GetAngkaDecimal(Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Row.Cells(Me.ColPrimaryUnit.Index).Value))
                        Else
                            Row.Cells(Me.ColMaxQty.Index).Value = Me.ObjCD.DigitGrouping(CDbl(Row.Cells(Me.ColMaxQty.Index).Value), _
                            Me.ObjProd.GetAngkaDecimal(Row.Cells(Me.ColPrimaryCodeNo.Index).Value))
                        End If
                    End If

                    If e.ColumnIndex = Me.ColUnitPrice.Index Then
                        If IsNumeric(Row.Cells(Me.ColUnitPrice.Index).Value) = False Then
                            MsgBox("Please input number in unit price", MsgBoxStyle.Information, "Information")
                            Row.Cells(Me.ColUnitPrice.Index).Value = 0 'Me.MasukanFormatStrip(Me.Buffer, Row.Cells(Me.ColDecimal.Index).Value)
                            Exit Sub
                        End If
                        If Me.ObjCD.CekNonNegatif(Row.Cells(Me.ColUnitPrice.Index).Value) = False Then
                            MsgBox("Please input positive number", MsgBoxStyle.Information, "Information")
                            Row.Cells(Me.ColUnitPrice.Index).Value = 0 'Me.MasukanFormatStrip(Me.Buffer, Row.Cells(Me.ColDecimal.Index).Value)
                            Exit Sub
                        End If

                        If Me.ChkUpdateCustGroup.Checked = True Then
                            DTACusGroup.Fill(TblCusGroup)
                            Me.ObjMP.HapusSatuanGroupCustomer(Me.txtPrimaryCurrency.Text, Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Row.Cells(Me.ColPrimaryUnit.Index).Value)
                        End If

                        If Me.ChkUpdateCust.Checked = True Then
                            DTACust.Fill(TBlCust)
                            Me.ObjMP.HapusSatuanCustomer(Me.txtPrimaryCurrency.Text, Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Row.Cells(Me.ColPrimaryUnit.Index).Value)
                        End If


                        Row.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(Row.Cells(Me.ColUnitPrice.Index).Value), Me.ObjMP.GetDefPrecCurrency(Me.txtPrimaryCurrency.Text))
                    End If

                    If e.ColumnIndex = Me.ColDiscount.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Or e.ColumnIndex = Me.ColRemarks.Index Then
                        If e.ColumnIndex = Me.ColRemarks.Index Then
                            If Me.CekStrings(Row.Cells(Me.ColRemarks.Index).Value) = False Then
                                Row.Cells(Me.ColRemarks.Index).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If
                        If IsNumeric(Row.Cells(Me.ColDiscount.Index).Value) = False Then
                            MsgBox("Please input number in discount", MsgBoxStyle.Information, "Information")
                            Row.Cells(Me.ColDiscount.Index).Value = 0 'Me.MasukanFormatStrip(Me.Buffer, Row.Cells(Me.ColDecimal.Index).Value)
                            Exit Sub
                        End If
                        If e.ColumnIndex = Me.ColDiscount.Index Then
                            Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        End If
                        If Me.ObjCD.CekNonNegatif(Row.Cells(Me.ColDiscount.Index).Value) = False Then
                            MsgBox("Please input positive number", MsgBoxStyle.Information, "Information")
                            Row.Cells(Me.ColDiscount.Index).Value = 0 'Me.MasukanFormatStrip(Me.Buffer, Row.Cells(Me.ColDecimal.Index).Value)
                            Exit Sub
                        End If
                        If CDec(Row.Cells(Me.ColDiscount.Index).Value) > 100 Then
                            MsgBox("Maximum discount is 100% ", MsgBoxStyle.Information, "Information")
                            Row.Cells(Me.ColDiscount.Index).Value = 0 'Me.MasukanFormatStrip(Me.Buffer, Row.Cells(Me.ColDecimal.Index).Value)
                            Exit Sub

                            If Me.ChkUpdateCustGroup.Checked = True Then
                                DTACusGroup.Fill(TblCusGroup)
                                Me.ObjMP.HapusSatuanGroupCustomer(Me.txtPrimaryCurrency.Text, Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Row.Cells(Me.ColPrimaryUnit.Index).Value)
                            End If

                            If Me.ChkUpdateCust.Checked = True Then
                                DTACust.Fill(TBlCust)
                                Me.ObjMP.HapusSatuanCustomer(Me.txtPrimaryCurrency.Text, Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Row.Cells(Me.ColPrimaryUnit.Index).Value)
                            End If

                            Row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(CDec(Row.Cells(Me.ColDiscount.Index).Value), Me.Def_Prec_Disc)
                        End If

                        If Row.Cells(Me.ColPrimary.Index).Value = -1 Then
                            'Radio Button General
                            If Me.rboGeneral.Checked = True Then
                                If Row.Cells(Me.ColMaxQty.Index).Value = "-" Then
                                    Me.ObjMP.SimpanGeneral(CDbl(Row.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(Row.Cells(Me.ColUnitPrice.Index).Value), _
                                    CDec(Row.Cells(Me.ColDiscount.Index).Value), nilainol, CDbl(Me.txtPrimaryCurrency.Text), _
                                    CDbl(Row.Cells(Me.ColPrimaryUnit.Index).Value), CStr(Row.Cells(Me.ColRemarks.Index).Value))
                                Else
                                    Me.ObjMP.SimpanGeneral(CDbl(Row.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(Row.Cells(Me.ColUnitPrice.Index).Value), _
                                    CDec(Row.Cells(Me.ColDiscount.Index).Value), CDbl(Row.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), _
                                    CDbl(Row.Cells(Me.ColPrimaryUnit.Index).Value), CStr(Row.Cells(Me.ColRemarks.Index).Value))
                                End If
                                Row.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryGeneral
                                'Radio Button Customer
                            ElseIf Me.rboCustomer.Checked = True Then
                                If Row.Cells(Me.ColMaxQty.Index).Value = "-" Then
                                    Me.ObjMP.SimpanCustomer(CDbl(Row.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(Row.Cells(Me.ColUnitPrice.Index).Value), _
                                    CDec(Row.Cells(Me.ColDiscount.Index).Value), nilainol, CDbl(Me.txtPrimaryCurrency.Text), CDbl(Me.txtprimaryCust.Text), _
                                     CDbl(Row.Cells(Me.ColPrimaryUnit.Index).Value), CStr(Row.Cells(Me.ColRemarks.Index).Value))
                                Else
                                    Me.ObjMP.SimpanCustomer(CDbl(Row.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(Row.Cells(Me.ColUnitPrice.Index).Value), _
                                    CDec(Row.Cells(Me.ColDiscount.Index).Value), CDbl(Row.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), CDbl(Me.txtprimaryCust.Text), _
                                    CDbl(Row.Cells(Me.ColPrimaryUnit.Index).Value), CStr(Row.Cells(Me.ColRemarks.Index).Value))
                                End If
                                Row.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryCustomer
                                'Radio Button CustomerGroup
                            ElseIf Me.rboCustomerGroup.Checked = True Then
                                If Row.Cells(Me.ColMaxQty.Index).Value = "-" Then
                                    Me.ObjMP.SimpanGroupCustomer(CDbl(Row.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(Row.Cells(Me.ColUnitPrice.Index).Value), _
                                    CDec(Row.Cells(Me.ColDiscount.Index).Value), nilainol, CDbl(Me.txtPrimaryCurrency.Text), CDbl(Me.TxtPrimaryCustGroup.Text), Me.cboCustomerGroup.SelectedIndex, _
                                     CDbl(Row.Cells(Me.ColPrimaryUnit.Index).Value), CStr(Row.Cells(Me.ColRemarks.Index).Value))
                                    Me.cboCustomerGroup.Enabled = False
                                Else
                                    Me.ObjMP.SimpanGroupCustomer(CDbl(Row.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(Row.Cells(Me.ColUnitPrice.Index).Value), _
                                    CDec(Row.Cells(Me.ColDiscount.Index).Value), CDbl(Row.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), CDbl(Me.TxtPrimaryCustGroup.Text), _
                                    Me.cboCustomerGroup.SelectedIndex, CDbl(Row.Cells(Me.ColPrimaryUnit.Index).Value), CStr(Row.Cells(Me.ColRemarks.Index).Value))
                                    Me.cboCustomerGroup.Enabled = False
                                End If
                                Me.cboCustomerGroup.Enabled = False
                                Row.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetmaxPrimaryGroup
                            End If
                        Else
                            'update discount
                            'Update General
                            If Me.rboGeneral.Checked = True Then
                                If Row.Cells(Me.ColMaxQty.Index).Value = "-" Then
                                    Me.ObjMP.UpdateGeneral(CDbl(Row.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(Row.Cells(Me.ColUnitPrice.Index).Value), _
                                    CDec(Row.Cells(Me.ColDiscount.Index).Value), nilainol, CDbl(Me.txtPrimaryCurrency.Text), _
                                    CDbl(Row.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(Row.Cells(Me.ColPrimary.Index).Value), CStr(Row.Cells(Me.ColRemarks.Index).Value))
                                Else
                                    Me.ObjMP.UpdateGeneral(CDbl(Row.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(Row.Cells(Me.ColUnitPrice.Index).Value), _
                                    CDec(Row.Cells(Me.ColDiscount.Index).Value), CDbl(Row.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), _
                                    CDbl(Row.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(Row.Cells(Me.ColPrimary.Index).Value), CStr(Row.Cells(Me.ColRemarks.Index).Value))
                                End If
                                'Row.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryGeneral

                                'Update Customer
                            ElseIf Me.rboCustomer.Checked = True Then
                                If Me.ObjMP.ApakahAdaCust(Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.txtPrimaryCurrency.Text, _
                               Me.txtprimaryCust.Text, Row.Cells(Me.ColPrimaryUnit.Index).Value) = True Then
                                    If Row.Cells(Me.ColMaxQty.Index).Value = "-" Then
                                        Me.ObjMP.UpdateCustomer(CDbl(Row.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(Row.Cells(Me.ColUnitPrice.Index).Value), _
                                        CDec(Row.Cells(Me.ColDiscount.Index).Value), nilainol, CDbl(Me.txtPrimaryCurrency.Text), _
                                        CDbl(Me.txtprimaryCust.Text), CDbl(Row.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(Row.Cells(Me.ColPrimary.Index).Value), CStr(Row.Cells(Me.ColRemarks.Index).Value))
                                    Else
                                        Me.ObjMP.UpdateCustomer(CDbl(Row.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(Row.Cells(Me.ColUnitPrice.Index).Value), _
                                        CDec(Row.Cells(Me.ColDiscount.Index).Value), CDbl(Row.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), _
                                        CDbl(Me.txtprimaryCust.Text), CDbl(Row.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(Row.Cells(Me.ColPrimary.Index).Value), CStr(Row.Cells(Me.ColRemarks.Index).Value))
                                    End If
                                    'Row.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryCustomer
                                Else
                                    'simpan baru
                                    '1. cari start index
                                    TargetIndex = Row.Index
                                    '2. cari acuan kombinasi primary product dan primary unit
                                    DummyTarget = Row.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + Row.Cells(Me.ColPrimaryUnit.Index).Value.ToString
                                    While TargetIndex > 0
                                        rowTarget = Me.DGVHarga.Rows(TargetIndex)
                                        If DummyTarget = rowTarget.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + rowTarget.Cells(Me.ColPrimaryUnit.Index).Value.ToString Then
                                            'ok
                                        Else
                                            TargetIndex += 1
                                            Exit While
                                        End If
                                        TargetIndex -= 1
                                    End While
                                    For ij As Integer = TargetIndex To Me.DGVHarga.Rows.Count - 1
                                        rowTarget = Me.DGVHarga.Rows(ij)
                                        If DummyTarget = rowTarget.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + rowTarget.Cells(Me.ColPrimaryUnit.Index).Value.ToString Then
                                            'simpan datanya
                                            With rowTarget
                                                If .Cells(Me.ColMaxQty.Index).Value = "-" Then
                                                    Me.ObjMP.SimpanCustomer(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColUnitPrice.Index).Value, _
                                                    .Cells(Me.ColDiscount.Index).Value, 0, Me.txtPrimaryCurrency.Text, _
                                                    Me.txtprimaryCust.Text, .Cells(Me.ColPrimaryUnit.Index).Value, .Cells(Me.ColRemarks.Index).Value)
                                                Else
                                                    Me.ObjMP.SimpanCustomer(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColUnitPrice.Index).Value, _
                                                    .Cells(Me.ColDiscount.Index).Value, .Cells(Me.ColMaxQty.Index).Value, Me.txtPrimaryCurrency.Text, _
                                                    Me.txtprimaryCust.Text, .Cells(Me.ColPrimaryUnit.Index).Value, .Cells(Me.ColRemarks.Index).Value)
                                                End If
                                                rowTarget.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryCustomer
                                            End With
                                        Else
                                            Exit For
                                        End If
                                    Next
                                End If

                                'Update Customer Group
                            ElseIf Me.rboCustomerGroup.Checked = True Then
                                If Me.ObjMP.ApakahAdaCustGroup(Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.txtPrimaryCurrency.Text, _
                                Me.TxtPrimaryCustGroup.Text, Row.Cells(Me.ColPrimaryUnit.Index).Value) = True Then
                                    If Row.Cells(Me.ColMaxQty.Index).Value = "-" Then
                                        Me.ObjMP.UpdateGroupCustomer(CDbl(Row.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(Row.Cells(Me.ColUnitPrice.Index).Value), _
                                        CDec(Row.Cells(Me.ColDiscount.Index).Value), nilainol, CDbl(Me.txtPrimaryCurrency.Text), _
                                       CDbl(Me.TxtPrimaryCustGroup.Text), Me.cboCustomerGroup.SelectedIndex, CDbl(Row.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(Row.Cells(Me.ColPrimary.Index).Value), CStr(Row.Cells(Me.ColRemarks.Index).Value))
                                    Else
                                        Me.ObjMP.UpdateGroupCustomer(CDbl(Row.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(Row.Cells(Me.ColUnitPrice.Index).Value), _
                                        CDec(Row.Cells(Me.ColDiscount.Index).Value), CDbl(Row.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), _
                                        CDbl(Me.TxtPrimaryCustGroup.Text), Me.cboCustomerGroup.SelectedIndex, CDbl(Row.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(Row.Cells(Me.ColPrimary.Index).Value), CStr(Row.Cells(Me.ColRemarks.Index).Value))
                                    End If

                                    'Row.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetmaxPrimaryGroup
                                Else
                                    'simpan baru
                                    '1. cari start index
                                    TargetIndex = Row.Index
                                    '2. cari acuan kombinasi primary product dan primary unit
                                    DummyTarget = Row.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + Row.Cells(Me.ColPrimaryUnit.Index).Value.ToString
                                    While TargetIndex > 0
                                        rowTarget = Me.DGVHarga.Rows(TargetIndex)
                                        If DummyTarget = rowTarget.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + rowTarget.Cells(Me.ColPrimaryUnit.Index).Value.ToString Then
                                            'ok
                                        Else
                                            TargetIndex += 1
                                            Exit While
                                        End If
                                        TargetIndex -= 1
                                    End While
                                    For ij As Integer = TargetIndex To Me.DGVHarga.Rows.Count - 1
                                        rowTarget = Me.DGVHarga.Rows(ij)
                                        If DummyTarget = rowTarget.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + rowTarget.Cells(Me.ColPrimaryUnit.Index).Value.ToString Then
                                            'simpan datanya
                                            With rowTarget
                                                If .Cells(Me.ColMaxQty.Index).Value = "-" Then
                                                    Me.ObjMP.SimpanGroupCustomer(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColUnitPrice.Index).Value, _
                                                    .Cells(Me.ColDiscount.Index).Value, 0, Me.txtPrimaryCurrency.Text, _
                                                    Me.TxtPrimaryCustGroup.Text, Me.cboCustomerGroup.SelectedIndex, .Cells(Me.ColPrimaryUnit.Index).Value, .Cells(Me.ColRemarks.Index).Value)
                                                Else
                                                    Me.ObjMP.SimpanGroupCustomer(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColUnitPrice.Index).Value, _
                                                    .Cells(Me.ColDiscount.Index).Value, .Cells(Me.ColMaxQty.Index).Value, Me.txtPrimaryCurrency.Text, _
                                                    Me.TxtPrimaryCustGroup.Text, Me.cboCustomerGroup.SelectedIndex, .Cells(Me.ColPrimaryUnit.Index).Value, .Cells(Me.ColRemarks.Index).Value)
                                                End If
                                                rowTarget.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryCustomer
                                            End With
                                        Else
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If

                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.DGVHarga.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim RowDGV_Source As DataGridViewRow = Me.DGVHarga.Rows(Me.InsertRowIndex)
        Dim RowDGV_Destination As DataGridViewRow
        Dim RowDGV_Before As DataGridViewRow
        Dim RowDGV_After As DataGridViewRow
        Dim nilainol As Double = 0
        Dim TargetIndex As Integer = 0
        Dim rowTarget As DataGridViewRow
        Dim DummyTarget As String = ""

        Try
            ''hanya tambah baris baru jika max qty awalnya "-"
            If Me.Buffer = "0" Then
                ''periksa ke atas
                ''hanya periksa yang bukan paling atas
                If Me.DGVHarga.Rows.IndexOf(RowDGV_Source) > 0 Then
                    RowDGV_Before = Me.DGVHarga.Rows(Me.InsertRowIndex - 1)
                    If RowDGV_Before.Cells(Me.ColMaxQty.Index).Value <> "-" Then
                        If CDbl(RowDGV_Source.Cells(Me.ColMaxQty.Index).Value) <= CDbl(RowDGV_Before.Cells(Me.ColMaxQty.Index).Value) Then
                            MsgBox("Max quantity cannot smaller or equals with previous row", MsgBoxStyle.Exclamation, "Information")
                            RowDGV_Source.Cells(Me.ColMaxQty.Index).Value = "-"
                            Exit Sub
                        End If
                    End If
                End If

                If Me.DGVHarga.Item(Me.ColMaxQty.Index, RowDGV_Source.Index).Value = "0" Then
                    MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                    Me.DGVHarga.Item(Me.ColMaxQty.Index, RowDGV_Source.Index).Value = Me.Buffer
                    Exit Sub
                End If
                ''salin data ke baris bawah
                Me.DGVHarga.Rows.InsertCopy(Me.InsertRowIndex, Me.InsertRowIndex + 1)
                RowDGV_Destination = Me.DGVHarga.Rows(Me.InsertRowIndex + 1)
                RowDGV_Destination.Cells(Me.ColProductCode.Index).Value = RowDGV_Source.Cells(Me.ColProductCode.Index).Value
                RowDGV_Destination.Cells(Me.ColDescription.Index).Value = RowDGV_Source.Cells(Me.ColDescription.Index).Value
                RowDGV_Destination.Cells(Me.ColUnit.Index).Value = RowDGV_Source.Cells(Me.ColUnit.Index).Value
                RowDGV_Destination.Cells(Me.ColUnitPrice.Index).Value = RowDGV_Source.Cells(Me.ColUnitPrice.Index).Value
                RowDGV_Destination.Cells(Me.ColDiscount.Index).Value = RowDGV_Source.Cells(Me.ColDiscount.Index).Value
                RowDGV_Destination.Cells(Me.ColMaxQty.Index).Value = "-"
                RowDGV_Destination.Cells(Me.ColRemarks.Index).Value = ""
                RowDGV_Destination.Cells(Me.ColPrimaryUnit.Index).Value = RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value
                RowDGV_Destination.Cells(Me.ColPrimaryCodeNo.Index).Value = RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value
                RowDGV_Destination.Cells(Me.ColPrimary.Index).Value = "-1"
                RowDGV_Destination.Cells(Me.ColDefrec.Index).Value = RowDGV_Source.Cells(Me.ColDefrec.Index).Value
                RowDGV_Destination.Cells(Me.ColDecimal.Index).Value = RowDGV_Source.Cells(Me.ColDecimal.Index).Value

                ''simpan ke database:
                If RowDGV_Destination.Cells(Me.ColPrimary.Index).Value = -1 Then
                    If Me.rboGeneral.Checked = True Then
                        ''update yg lagi dikerjain:
                        If RowDGV_Source.Cells(Me.ColPrimary.Index).Value <> "-1" Then
                            Me.ObjMP.UpdateGeneral(CDbl(RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Source.Cells(Me.ColUnitPrice.Index).Value), _
                                                            CDec(RowDGV_Source.Cells(Me.ColDiscount.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), _
                                                            CDbl(RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColPrimary.Index).Value), CStr(RowDGV_Source.Cells(Me.ColRemarks.Index).Value))
                        Else
                            Me.ObjMP.SimpanGeneral(CDbl(RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Source.Cells(Me.ColUnitPrice.Index).Value), _
                                                                CDec(RowDGV_Source.Cells(Me.ColDiscount.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), _
                                                                CDbl(RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value), CStr(RowDGV_Source.Cells(Me.ColRemarks.Index).Value))
                            ''get the last primary
                            RowDGV_Source.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryGeneral
                        End If
                        ''simpan baris yang baru:
                        If RowDGV_Destination.Cells(Me.ColMaxQty.Index).Value = "-" Then
                            Me.ObjMP.SimpanGeneral(CDbl(RowDGV_Destination.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Destination.Cells(Me.ColUnitPrice.Index).Value), _
                                    CDec(RowDGV_Destination.Cells(Me.ColDiscount.Index).Value), nilainol, CDbl(Me.txtPrimaryCurrency.Text), _
                                    CDbl(RowDGV_Destination.Cells(Me.ColPrimaryUnit.Index).Value), CStr(RowDGV_Destination.Cells(Me.ColRemarks.Index).Value))
                            ''get the last primary
                            RowDGV_Destination.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryGeneral
                        End If

                        ''didieu

                    ElseIf Me.rboCustomer.Checked = True Then
                        If Me.ObjMP.ApakahAdaCust(RowDGV_Destination.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.txtPrimaryCurrency.Text, _
                              Me.txtprimaryCust.Text, RowDGV_Destination.Cells(Me.ColPrimaryUnit.Index).Value) = True Then
                            ''simpan baris yang baru:
                            If RowDGV_Destination.Cells(Me.ColMaxQty.Index).Value = "-" Then
                                Me.ObjMP.SimpanCustomer(CDbl(RowDGV_Destination.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Destination.Cells(Me.ColUnitPrice.Index).Value), _
                                        CDec(RowDGV_Destination.Cells(Me.ColDiscount.Index).Value), nilainol, CDbl(Me.txtPrimaryCurrency.Text), _
                                         CDbl(Me.txtprimaryCust.Text), (RowDGV_Destination.Cells(Me.ColPrimaryUnit.Index).Value), CStr(RowDGV_Destination.Cells(Me.ColRemarks.Index).Value))
                            End If
                            ''update yg lagi dikerjain:
                            Me.ObjMP.UpdateCustomer(CDbl(RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Source.Cells(Me.ColUnitPrice.Index).Value), _
                                    CDec(RowDGV_Source.Cells(Me.ColDiscount.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), _
                                     CDbl(Me.txtprimaryCust.Text), CDbl(RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColPrimary.Index).Value), CStr(RowDGV_Source.Cells(Me.ColRemarks.Index).Value))
                            ''get the last primary
                            RowDGV_Destination.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryCustomer
                        Else
                            'simpan baru
                            '1. cari start index
                            TargetIndex = RowDGV_Destination.Index
                            '2. cari acuan kombinasi primary product dan primary unit
                            DummyTarget = RowDGV_Destination.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + RowDGV_Destination.Cells(Me.ColPrimaryUnit.Index).Value.ToString
                            While TargetIndex > 0
                                rowTarget = Me.DGVHarga.Rows(TargetIndex)
                                If DummyTarget = rowTarget.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + rowTarget.Cells(Me.ColPrimaryUnit.Index).Value.ToString Then
                                    'ok
                                Else
                                    TargetIndex += 1
                                    Exit While
                                End If
                                TargetIndex -= 1
                            End While
                            For ij As Integer = TargetIndex To Me.DGVHarga.Rows.Count - 1
                                rowTarget = Me.DGVHarga.Rows(ij)
                                If DummyTarget = rowTarget.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + rowTarget.Cells(Me.ColPrimaryUnit.Index).Value.ToString Then
                                    'simpan datanya
                                    With rowTarget
                                        If .Cells(Me.ColMaxQty.Index).Value = "-" Then
                                            Me.ObjMP.SimpanCustomer(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColUnitPrice.Index).Value, _
                                            .Cells(Me.ColDiscount.Index).Value, 0, Me.txtPrimaryCurrency.Text, _
                                            Me.txtprimaryCust.Text, .Cells(Me.ColPrimaryUnit.Index).Value, .Cells(Me.ColRemarks.Index).Value)
                                        Else
                                            Me.ObjMP.SimpanCustomer(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColUnitPrice.Index).Value, _
                                            .Cells(Me.ColDiscount.Index).Value, .Cells(Me.ColMaxQty.Index).Value, Me.txtPrimaryCurrency.Text, _
                                            Me.txtprimaryCust.Text, .Cells(Me.ColPrimaryUnit.Index).Value, .Cells(Me.ColRemarks.Index).Value)
                                        End If
                                        rowTarget.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryCustomer
                                    End With
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                    ElseIf Me.rboCustomerGroup.Checked = True Then
                        If Me.ObjMP.ApakahAdaCustGroup(RowDGV_Destination.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.txtPrimaryCurrency.Text, _
                             Me.TxtPrimaryCustGroup.Text, RowDGV_Destination.Cells(Me.ColPrimaryUnit.Index).Value) = True Then
                            ''simpan baris yang baru:
                            If RowDGV_Destination.Cells(Me.ColMaxQty.Index).Value = "-" Then
                                Me.ObjMP.SimpanGroupCustomer(CDbl(RowDGV_Destination.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Destination.Cells(Me.ColUnitPrice.Index).Value), _
                                        CDec(RowDGV_Destination.Cells(Me.ColDiscount.Index).Value), nilainol, CDbl(Me.txtPrimaryCurrency.Text), _
                                         CDbl(Me.TxtPrimaryCustGroup.Text), Me.cboCustomerGroup.SelectedIndex, (RowDGV_Destination.Cells(Me.ColPrimaryUnit.Index).Value), CStr(RowDGV_Destination.Cells(Me.ColRemarks.Index).Value))
                            End If
                            ''update yg lagi dikerjain:
                            'Console.WriteLine(CDec(RowDGV_Source.Cells(Me.ColUnitPrice.Index).Value))
                            'Console.WriteLine(CDec(RowDGV_Source.Cells(Me.ColDiscount.Index).Value))
                            'Console.WriteLine(CDbl(RowDGV_Source.Cells(Me.ColMaxQty.Index).Value))
                            Me.ObjMP.UpdateGroupCustomer(CDbl(RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Source.Cells(Me.ColUnitPrice.Index).Value), _
                                     CDec(RowDGV_Source.Cells(Me.ColDiscount.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), _
                                      CDbl(Me.TxtPrimaryCustGroup.Text), Me.cboCustomerGroup.SelectedIndex, CDbl(RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value), _
                                     CDbl(RowDGV_Source.Cells(Me.ColPrimary.Index).Value), CStr(RowDGV_Source.Cells(Me.ColRemarks.Index).Value))
                            ''get the last primary
                            RowDGV_Destination.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetmaxPrimaryGroup
                        Else
                            'simpan baru
                            '1. cari start index
                            TargetIndex = RowDGV_Destination.Index
                            '2. cari acuan kombinasi primary product dan primary unit
                            DummyTarget = RowDGV_Destination.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + RowDGV_Destination.Cells(Me.ColPrimaryUnit.Index).Value.ToString
                            While TargetIndex > 0
                                rowTarget = Me.DGVHarga.Rows(TargetIndex)
                                If DummyTarget = rowTarget.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + rowTarget.Cells(Me.ColPrimaryUnit.Index).Value.ToString Then
                                    'ok
                                Else
                                    TargetIndex += 1
                                    Exit While
                                End If
                                TargetIndex -= 1
                            End While
                            For ij As Integer = TargetIndex To Me.DGVHarga.Rows.Count - 1
                                rowTarget = Me.DGVHarga.Rows(ij)
                                If DummyTarget = rowTarget.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + rowTarget.Cells(Me.ColPrimaryUnit.Index).Value.ToString Then
                                    'simpan datanya
                                    With rowTarget
                                        If .Cells(Me.ColMaxQty.Index).Value = "-" Then
                                            Me.ObjMP.SimpanGroupCustomer(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColUnitPrice.Index).Value, _
                                            .Cells(Me.ColDiscount.Index).Value, 0, Me.txtPrimaryCurrency.Text, _
                                            Me.TxtPrimaryCustGroup.Text, Me.cboCustomerGroup.SelectedIndex, .Cells(Me.ColPrimaryUnit.Index).Value, .Cells(Me.ColRemarks.Index).Value)
                                        Else
                                            Me.ObjMP.SimpanGroupCustomer(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColUnitPrice.Index).Value, _
                                            .Cells(Me.ColDiscount.Index).Value, .Cells(Me.ColMaxQty.Index).Value, Me.txtPrimaryCurrency.Text, _
                                            Me.TxtPrimaryCustGroup.Text, Me.cboCustomerGroup.SelectedIndex, .Cells(Me.ColPrimaryUnit.Index).Value, .Cells(Me.ColRemarks.Index).Value)
                                        End If
                                        rowTarget.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryCustomer
                                    End With
                                Else
                                    Exit For
                                End If
                            Next
                        End If

                    End If
                End If
            Else
                ''jika merubah yang awal max qty <> "-"
                If DGVHarga.Rows.IndexOf(RowDGV_Source) > 0 Then
                    RowDGV_Before = Me.DGVHarga.Rows(Me.InsertRowIndex - 1)
                    If RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value = RowDGV_Before.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                            RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value = RowDGV_Before.Cells(Me.ColPrimaryUnit.Index).Value Then
                        If IsNumeric(RowDGV_Before.Cells(Me.ColMaxQty.Index).Value) = True Then
                            If CDbl(RowDGV_Source.Cells(Me.ColMaxQty.Index).Value) <= CDbl(RowDGV_Before.Cells(Me.ColMaxQty.Index).Value) Then
                                MsgBox("Max quantity cannot smaller or equals with previous row", MsgBoxStyle.Exclamation, "Information")
                                RowDGV_Source.Cells(Me.ColMaxQty.Index).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If
                    End If
                End If

                If DGVHarga.Rows.IndexOf(RowDGV_Source) < DGVHarga.Rows.Count - 1 Then
                    RowDGV_After = Me.DGVHarga.Rows(Me.InsertRowIndex + 1)
                    If RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value = RowDGV_After.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                            RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value = RowDGV_After.Cells(Me.ColPrimaryUnit.Index).Value Then
                        If RowDGV_After.Cells(Me.ColMaxQty.Index).Value <> "-" Then
                            If IsNumeric(RowDGV_After.Cells(Me.ColMaxQty.Index).Value) = True Then
                                If CDbl(RowDGV_Source.Cells(Me.ColMaxQty.Index).Value) >= CDbl(RowDGV_After.Cells(Me.ColMaxQty.Index).Value) Then
                                    MsgBox("Max Qty cannot bigger or equals with further row", MsgBoxStyle.Exclamation, "Information")
                                    RowDGV_Source.Cells(Me.ColMaxQty.Index).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                        End If
                        ''TIKAMARIIII!!!!!!!!!!!!!!!!
                    End If
                End If

                ''update ke db
                If Me.rboGeneral.Checked = True Then
                    If RowDGV_Source.Cells(Me.ColMaxQty.Index).Value = "-" Then
                        Me.ObjMP.UpdateGeneral(CDbl(RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Source.Cells(Me.ColUnitPrice.Index).Value), _
                        CDec(RowDGV_Source.Cells(Me.ColDiscount.Index).Value), nilainol, CDbl(Me.txtPrimaryCurrency.Text), _
                        CDbl(RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColPrimary.Index).Value), CStr(RowDGV_Source.Cells(Me.ColRemarks.Index).Value))
                    Else
                        Me.ObjMP.UpdateGeneral(CDbl(RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Source.Cells(Me.ColUnitPrice.Index).Value), _
                        CDec(RowDGV_Source.Cells(Me.ColDiscount.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), _
                        CDbl(RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColPrimary.Index).Value), CStr(RowDGV_Source.Cells(Me.ColRemarks.Index).Value))
                    End If

                    'RowDGV_Source.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryGeneral
                    '                    RowDGV_Source.Cells(Me.ColMaxQty.Index).Value = Me.ObjCD.DigitGrouping(RowDGV_Source.Cells(Me.ColMaxQty.Index).Value.ToString, CInt(RowDGV_Source.Cells(Me.ColDecimal.Index).Value))
                    ''didieu
                    If Me.ChkUpdateCustGroup.Checked = True Then
                        Me.ObjMP.HapusSatuanGroupCustomer(Me.txtPrimaryCurrency.Text, RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value, RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value)
                    End If

                    If Me.ChkUpdateCust.Checked = True Then
                        Me.ObjMP.HapusSatuanCustomer(Me.txtPrimaryCurrency.Text, RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value, RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value)
                    End If
                ElseIf Me.rboCustomerGroup.Checked = True Then
                    If Me.ObjMP.ApakahAdaCustGroup(RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.txtPrimaryCurrency.Text, _
                            Me.TxtPrimaryCustGroup.Text, RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value) = True Then
                        If RowDGV_Source.Cells(Me.ColMaxQty.Index).Value = "-" Then
                            Me.ObjMP.UpdateGroupCustomer(CDbl(RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Source.Cells(Me.ColUnitPrice.Index).Value), _
                            CDec(RowDGV_Source.Cells(Me.ColDiscount.Index).Value), nilainol, CDbl(Me.txtPrimaryCurrency.Text), _
                             CDbl(Me.TxtPrimaryCustGroup.Text), Me.cboCustomerGroup.SelectedIndex, (RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColPrimary.Index).Value), CStr(RowDGV_Source.Cells(Me.ColRemarks.Index).Value))
                        Else
                            Me.ObjMP.UpdateGroupCustomer(CDbl(RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Source.Cells(Me.ColUnitPrice.Index).Value), _
                            CDec(RowDGV_Source.Cells(Me.ColDiscount.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), _
                            CDbl(Me.TxtPrimaryCustGroup.Text), Me.cboCustomerGroup.SelectedIndex, CDbl(RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColPrimary.Index).Value), CStr(RowDGV_Source.Cells(Me.ColRemarks.Index).Value))
                        End If
                        'RowDGV_Source.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryGeneral
                    Else
                        'simpan baru
                        '1. cari start index
                        TargetIndex = RowDGV_Source.Index
                        '2. cari acuan kombinasi primary product dan primary unit
                        DummyTarget = RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value.ToString
                        While TargetIndex > 0
                            rowTarget = Me.DGVHarga.Rows(TargetIndex)
                            If DummyTarget = rowTarget.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + rowTarget.Cells(Me.ColPrimaryUnit.Index).Value.ToString Then
                                'ok
                            Else
                                TargetIndex += 1
                                Exit While
                            End If
                            TargetIndex -= 1
                        End While
                        For ij As Integer = TargetIndex To Me.DGVHarga.Rows.Count - 1
                            rowTarget = Me.DGVHarga.Rows(ij)
                            If DummyTarget = rowTarget.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + rowTarget.Cells(Me.ColPrimaryUnit.Index).Value.ToString Then
                                'simpan datanya
                                With rowTarget
                                    If .Cells(Me.ColMaxQty.Index).Value = "-" Then
                                        Me.ObjMP.SimpanGroupCustomer(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColUnitPrice.Index).Value, _
                                        .Cells(Me.ColDiscount.Index).Value, 0, Me.txtPrimaryCurrency.Text, _
                                        Me.TxtPrimaryCustGroup.Text, Me.cboCustomerGroup.SelectedIndex, .Cells(Me.ColPrimaryUnit.Index).Value, .Cells(Me.ColRemarks.Index).Value)
                                    Else
                                        Me.ObjMP.SimpanGroupCustomer(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColUnitPrice.Index).Value, _
                                        .Cells(Me.ColDiscount.Index).Value, .Cells(Me.ColMaxQty.Index).Value, Me.txtPrimaryCurrency.Text, _
                                        Me.TxtPrimaryCustGroup.Text, Me.cboCustomerGroup.SelectedIndex, .Cells(Me.ColPrimaryUnit.Index).Value, .Cells(Me.ColRemarks.Index).Value)
                                    End If
                                    rowTarget.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryCustomer
                                End With
                            Else
                                Exit For
                            End If
                        Next
                    End If
                ElseIf Me.rboCustomer.Checked = True Then
                    If Me.ObjMP.ApakahAdaCust(RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.txtPrimaryCurrency.Text, _
                           Me.txtprimaryCust.Text, RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value) = True Then
                        If RowDGV_Source.Cells(Me.ColMaxQty.Index).Value = "-" Then
                            Me.ObjMP.UpdateCustomer(CDbl(RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Source.Cells(Me.ColUnitPrice.Index).Value), _
                            CDec(RowDGV_Source.Cells(Me.ColDiscount.Index).Value), nilainol, CDbl(Me.txtPrimaryCurrency.Text), _
                            CDbl(Me.txtprimaryCust.Text), CDbl(RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColPrimary.Index).Value), CStr(RowDGV_Source.Cells(Me.ColRemarks.Index).Value))
                        Else
                            Me.ObjMP.UpdateCustomer(CDbl(RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value), CDec(RowDGV_Source.Cells(Me.ColUnitPrice.Index).Value), _
                            CDec(RowDGV_Source.Cells(Me.ColDiscount.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColMaxQty.Index).Value), CDbl(Me.txtPrimaryCurrency.Text), _
                            CDbl(Me.txtprimaryCust.Text), CDbl(RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value), CDbl(RowDGV_Source.Cells(Me.ColPrimary.Index).Value), CStr(RowDGV_Source.Cells(Me.ColRemarks.Index).Value))
                        End If
                        'RowDGV_Source.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryGeneral
                    Else
                        'simpan baru
                        '1. cari start index
                        TargetIndex = RowDGV_Source.Index
                        '2. cari acuan kombinasi primary product dan primary unit
                        DummyTarget = RowDGV_Source.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + RowDGV_Source.Cells(Me.ColPrimaryUnit.Index).Value.ToString
                        While TargetIndex > 0
                            rowTarget = Me.DGVHarga.Rows(TargetIndex)
                            If DummyTarget = rowTarget.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + rowTarget.Cells(Me.ColPrimaryUnit.Index).Value.ToString Then
                                'ok
                            Else
                                TargetIndex += 1
                                Exit While
                            End If
                            TargetIndex -= 1
                        End While
                        For ij As Integer = TargetIndex To Me.DGVHarga.Rows.Count - 1
                            rowTarget = Me.DGVHarga.Rows(ij)
                            If DummyTarget = rowTarget.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + " - " + rowTarget.Cells(Me.ColPrimaryUnit.Index).Value.ToString Then
                                'simpan datanya
                                With rowTarget
                                    If .Cells(Me.ColMaxQty.Index).Value = "-" Then
                                        Me.ObjMP.SimpanCustomer(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColUnitPrice.Index).Value, _
                                        .Cells(Me.ColDiscount.Index).Value, 0, Me.txtPrimaryCurrency.Text, _
                                        Me.txtprimaryCust.Text, .Cells(Me.ColPrimaryUnit.Index).Value, .Cells(Me.ColRemarks.Index).Value)
                                    Else
                                        Me.ObjMP.SimpanCustomer(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColUnitPrice.Index).Value, _
                                        .Cells(Me.ColDiscount.Index).Value, .Cells(Me.ColMaxQty.Index).Value, Me.txtPrimaryCurrency.Text, _
                                        Me.txtprimaryCust.Text, .Cells(Me.ColPrimaryUnit.Index).Value, .Cells(Me.ColRemarks.Index).Value)
                                    End If
                                    rowTarget.Cells(Me.ColPrimary.Index).Value = Me.ObjMP.GetMaxPrimaryCustomer
                                End With
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCustomerGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCustomerGroup.SelectedIndexChanged
        Try
            If Me.cboCustomerGroup.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Group Then
                Me.IsiCboCustomerGroup()
            ElseIf Me.cboCustomerGroup.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Sales_Area Then
                Me.IsiCboCustomerSalesArea()
            ElseIf Me.cboCustomerGroup.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Segmentation Then
                Me.IsiCboCustomerSegmentation()
            ElseIf Me.cboCustomerGroup.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Type Then
                Me.IsiCboCustomerTypeCode()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '--------------- ISI CBO CUSTOMER GROUP
    Private Sub IsiCboCustomerGroup()
        Try
            'Me.cboCustomerCode.DataSource = Nothing
            'Me.cboCustomerCode.Items.Clear()
            Me.cboCustomerCode.DataSource = Me.TBLCboCustomerGroup.DefaultView
            Me.TxtPrimaryCustGroup.Text = Me.TBLCboCustomerGroup.Rows(0).Item(0).ToString
            Me.TBLCboCustomerGroup.DefaultView.Sort = "[Code]"
            Me.cboCustomerCode.DisplayMember = "Code"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboCustomerSalesArea()
        Try
            'Me.cboCustomerCode.DataSource = Nothing
            'Me.cboCustomerCode.Items.Clear()

            Me.cboCustomerCode.DataSource = Me.TBLCBOCustomerSalesArea.DefaultView
            Me.TxtPrimaryCustGroup.Text = Me.TBLCBOCustomerSalesArea.Rows(0).Item(0).ToString
            Me.TBLCBOCustomerSalesArea.DefaultView.Sort = "Sales Area Code"
            Me.cboCustomerCode.DisplayMember = "Sales Area Code"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboCustomerSegmentation()
        Try
            'Me.cboCustomerCode.DataSource = Nothing
            'Me.cboCustomerCode.Items.Clear()

            Me.cboCustomerCode.DataSource = Me.TBLCBOCustomerSegmentation.DefaultView
            Me.TxtPrimaryCustGroup.Text = Me.TBLCBOCustomerSegmentation.Rows(0).Item(0).ToString
            Me.TBLCBOCustomerSegmentation.DefaultView.Sort = "Segmentation Code"
            Me.cboCustomerCode.DisplayMember = "Segmentation Code"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboCustomerTypeCode()
        Try
            'Me.cboCustomerCode.DataSource = Nothing
            'Me.cboCustomerCode.Items.Clear()


            Me.cboCustomerCode.DataSource = Me.TBLTCBOypeCode.DefaultView
            Me.TxtPrimaryCustGroup.Text = Me.TBLTCBOypeCode.Rows(0).Item(0).ToString
            Me.TBLTCBOypeCode.DefaultView.Sort = "Type Code"
            Me.cboCustomerCode.DisplayMember = "Type Code"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomer.Click
        'Fungsi Find Data
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Dim ObjCode As New clsMasterCustomer

        Try
            If Me.TxtCustomerCode.Text <> "" Then
                i.txtCari.Text = Me.TxtCustomerCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.txtprimaryCust.Text = i.PrimaryCustomer
                Me.TxtCustomerCode.Text = ObjCode.GetCustomerCode(Me.txtprimaryCust.Text)
                Me.TxtCustomerName.Text = ObjCode.GetCustName(Me.txtprimaryCust.Text)
                Me.DGVHarga.Focus()
            Else
                Me.TxtCustomerCode.Clear()
                Me.txtprimaryCust.Text = "-1"
                Me.TxtCustomerName.Text = " "
                Me.TxtCustomerCode.Focus()
            End If

            Me.DGVHarga.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            ObjCode = Nothing
        End Try
    End Sub

    Private Sub TxtCustomerCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCode.Enter
        Try
            Me.Buffer = Me.txtprimaryCust.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtCustomerCode.KeyDown
        Dim objCode As New clsMasterCustomer
        Dim i As New FrmCariMasterCustomer

        Try
            If e.KeyValue = 13 Then
                If Me.TxtCustomerCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtCustomerCode.Text) = True Then
                        Me.txtprimaryCust.Text = objCode.GetPrimary(Me.TxtCustomerCode.Text)
                        Me.TxtCustomerCode.Text = objCode.GetCustomerCode(Me.txtprimaryCust.Text)
                        Me.TxtCustomerName.Text = objCode.GetCustName(Me.txtprimaryCust.Text)
                        Me.DGVHarga.Focus()
                    Else
                        Me.btnFindCustomer.PerformClick()
                    End If
                Else
                    Me.btnFindCustomer.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
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
                    Me.txtprimaryCust.Text = objCode.GetPrimary(Me.TxtCustomerCode.Text)
                    Me.TxtCustomerCode.Text = objCode.GetCustomerCode(Me.txtprimaryCust.Text)
                    Me.TxtCustomerName.Text = objCode.GetCustName(Me.txtprimaryCust.Text)
                Else
                    If Me.TxtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtprimaryCust.Text = i.PrimaryCustomer
                        Me.TxtCustomerCode.Text = objCode.GetCustomerCode(Me.txtprimaryCust.Text)
                        Me.TxtCustomerName.Text = objCode.GetCustName(Me.txtprimaryCust.Text)
                    Else
                        Me.TxtCustomerCode.Clear()
                        Me.txtprimaryCust.Text = "-1"
                        Me.TxtCustomerName.Text = " "
                    End If
                End If
            Else
                Me.TxtCustomerCode.Clear()
                Me.txtprimaryCust.Text = "-1"
                Me.TxtCustomerName.Text = " "
            End If

            If Me.txtprimaryCust.Text <> Me.Buffer Then
                Me.DGVHarga.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnDelete_ByGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete_ByGroup.Click
        Dim Result As DialogResult
        Dim RowCustomerGroup As DseMasterProduct_EditPriceList.tblmaster_product_pricelist_groupRow
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Dim TBL As New DseMasterProduct_EditPriceList.tblmaster_product_pricelist_groupDataTable
        Try

            If Me.TxtPrimaryCustGroup.Text <> "-1" And Me.TxtPrimaryCustGroup.Text <> "" Then
                Result = MsgBox("Are you sure delete this data ? ", MsgBoxStyle.YesNo, "Information")
                If Result = Windows.Forms.DialogResult.Yes Then
                    Me.ObjMP.HapusGroupCustomer(Me.TxtPrimaryCustGroup.Text)
                    Me.DGVHarga.Rows.Clear()
                    MsgBox("Data Customer Group Deleted", MsgBoxStyle.Information, "Information")
                    DTA.Fill(TBL)
                    If TBL.Rows.Count <= 0 Then
                        Me.cboCustomerGroup.Enabled = True
                    Else
                        'kalo <>0, maka RowCustomergroup ambil ke index 0, Row.type = cboType.index, enable = false
                        RowCustomerGroup = TBL.Rows(0)
                        Me.cboCustomerGroup.Enabled = False
                    End If
                End If
            Else
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_ByCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete_ByCustomer.Click
        Dim Result As DialogResult
        Try
            If Me.txtprimaryCust.Text <> "-1" And Me.txtprimaryCust.Text <> "" Then
                Result = MsgBox("Are you sure delete this data ?  ", MsgBoxStyle.YesNo, "Information")
                If Result = Windows.Forms.DialogResult.Yes Then
                    Me.ObjMP.HapusCustomer(Me.txtprimaryCust.Text)
                    Me.DGVHarga.Rows.Clear()
                    MsgBox("Data Customer Deleted", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtKodeProduct_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtKodeProduk.GotFocus
        Try
            If Me.TxtKodeProduk.Enabled = True Then
                If Me.TxtKodeProduk.Enabled = True Then
                    If Me.TxtKodeProduk.ReadOnly = False Then
                        TxtKodeProduk.BackColor = Me.BackColorAll
                    Else
                        TxtKodeProduk.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtKodeProduk_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtKodeProduk.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtKodeProduk.Text <> "" Then
                    Me.BtnFilter.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtKodeProduk_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtKodeProduk.LostFocus
        Try
            If Me.TxtKodeProduk.Enabled = True Then
                If Me.TxtKodeProduk.Enabled = True Then
                    If Me.TxtKodeProduk.ReadOnly = False Then
                        TxtKodeProduk.BackColor = Me.txtPrimaryCurrency.BackColor
                    Else
                        TxtKodeProduk.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCode.GotFocus
        Try
            If Me.TxtCustomerCode.Enabled = True Then
                If Me.TxtCustomerCode.Enabled = True Then
                    If Me.TxtCustomerCode.ReadOnly = False Then
                        TxtCustomerCode.BackColor = Me.BackColorAll
                    Else
                        TxtCustomerCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CustomerCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCode.LostFocus
        Try
            If Me.TxtCustomerCode.Enabled = True Then
                If Me.TxtCustomerCode.Enabled = True Then
                    If Me.TxtCustomerCode.ReadOnly = False Then
                        TxtCustomerCode.BackColor = Me.txtprimaryCust.BackColor
                    Else
                        TxtCustomerCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCustomerCode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCustomerCode.SelectedIndexChanged
        Dim RowGroup As DseMasterProduct_EditPriceList.tblmaster_customer_groupRow
        Dim RowArea As DseMasterProduct_EditPriceList.tblmaster_customer_sales_areaRow
        Dim RowSegmentation As DseMasterProduct_EditPriceList.tblmaster_customer_segmentationRow
        Dim RowType As DseMasterProduct_EditPriceList.tblmaster_customer_typeRow
        Try
            If Me.cboCustomerCode.Enabled = True Then
                Me.DGVHarga.Rows.Clear()

                'Cbo Customer Group
                If Me.cboCustomerGroup.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Group Then
                    If Me.cboCustomerCode.SelectedIndex = -1 Then
                        Me.TxtPrimaryCustGroup.Text = "-1"
                    Else
                        RowGroup = Me.TBLCboCustomerGroup.DefaultView(Me.cboCustomerCode.SelectedIndex).Row
                        Me.TxtPrimaryCustGroup.Text = RowGroup.Primary
                    End If
                    'Cbo Customer Sales Area
                ElseIf Me.cboCustomerGroup.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Sales_Area Then
                    If Me.cboCustomerCode.SelectedIndex = -1 Then
                        Me.TxtPrimaryCustGroup.Text = "-1"
                    Else
                        RowArea = Me.TBLCBOCustomerSalesArea.DefaultView(Me.cboCustomerCode.SelectedIndex).Row
                        Me.TxtPrimaryCustGroup.Text = RowArea.Primary
                    End If
                    'Cbo Customer Segmentation
                ElseIf Me.cboCustomerGroup.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Segmentation Then
                    If Me.cboCustomerCode.SelectedIndex = -1 Then
                        Me.TxtPrimaryCustGroup.Text = "-1"
                    Else
                        RowSegmentation = Me.TBLCBOCustomerSegmentation.DefaultView(Me.cboCustomerCode.SelectedIndex).Row
                        Me.TxtPrimaryCustGroup.Text = RowSegmentation.Primary
                    End If
                    'Cbo Cstomer Type
                ElseIf Me.cboCustomerGroup.SelectedIndex = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Type Then
                    If Me.cboCustomerCode.SelectedIndex = -1 Then
                        Me.TxtPrimaryCustGroup.Text = "-1"
                    Else
                        RowType = Me.TBLTCBOypeCode.DefaultView(Me.cboCustomerCode.SelectedIndex).Row
                        Me.TxtPrimaryCustGroup.Text = RowType.Primary
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub



    Private Sub DGVHarga_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DGVHarga.UserDeletingRow
        Dim Row As DataGridViewRow = Me.DGVHarga.Rows(e.Row.Index)
        Dim rowDGVAtasnya As DataGridViewRow

        Try
            If Me.rboGeneral.Checked = True Then
                If Row.Cells(Me.ColMaxQty.Index).Value <> "-" Then
                    Me.ObjMP.HapusGeneral(Row.Cells(Me.ColPrimary.Index).Value)
                Else
                    MsgBox("This data can't delete", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                    Exit Sub
                End If
            ElseIf Me.rboCustomer.Checked = True Then
                If Row.Index = 0 Then
                    If Me.ObjMP.ApakahAdaCust(Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.txtPrimaryCurrency.Text, _
                    Me.txtprimaryCust.Text, Row.Cells(Me.ColPrimaryUnit.Index).Value) = True Then
                        Me.ObjMP.HapusItemCustomer(Row.Cells(Me.ColPrimary.Index).Value)
                    End If
                Else
                    rowDGVAtasnya = Me.DGVHarga.Rows(e.Row.Index - 1)
                    If Row.Cells(Me.ColMaxQty.Index).Value <> "-" Then
                        If Me.ObjMP.ApakahAdaCust(Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.txtPrimaryCurrency.Text, _
                        Me.txtprimaryCust.Text, Row.Cells(Me.ColPrimaryUnit.Index).Value) = True Then
                            Me.ObjMP.HapusItemCustomer(Row.Cells(Me.ColPrimary.Index).Value)
                        End If
                    Else
                        If rowDGVAtasnya.Cells(Me.ColPrimaryCodeNo.Index).Value = Row.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                        rowDGVAtasnya.Cells(Me.ColPrimaryUnit.Index).Value = Row.Cells(Me.ColPrimaryUnit.Index).Value Then
                            'jika qty max =0, maka itu harus baris terakhir baru baleh dihapuskan
                            e.Cancel = True
                        Else
                            If Me.ObjMP.ApakahAdaCust(Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.txtPrimaryCurrency.Text, _
                            Me.txtprimaryCust.Text, Row.Cells(Me.ColPrimaryUnit.Index).Value) = True Then
                                Me.ObjMP.HapusItemCustomer(Row.Cells(Me.ColPrimary.Index).Value)
                            End If
                        End If
                    End If
                End If
            ElseIf Me.rboCustomerGroup.Checked = True Then
                If Row.Index = 0 Then
                    If Me.ObjMP.ApakahAdaCustGroup(Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.txtPrimaryCurrency.Text, _
                    Me.TxtPrimaryCustGroup.Text, Row.Cells(Me.ColPrimaryUnit.Index).Value) = True Then
                        Me.ObjMP.HapusItemGroup(Row.Cells(Me.ColPrimary.Index).Value)
                    End If
                Else
                    rowDGVAtasnya = Me.DGVHarga.Rows(e.Row.Index - 1)
                    If Row.Cells(Me.ColMaxQty.Index).Value <> "-" Then
                        If Me.ObjMP.ApakahAdaCustGroup(Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.txtPrimaryCurrency.Text, _
                        Me.TxtPrimaryCustGroup.Text, Row.Cells(Me.ColPrimaryUnit.Index).Value) = True Then
                            Me.ObjMP.HapusItemGroup(Row.Cells(Me.ColPrimary.Index).Value)
                        End If
                    Else
                        If rowDGVAtasnya.Cells(Me.ColPrimaryCodeNo.Index).Value = Row.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                        rowDGVAtasnya.Cells(Me.ColPrimaryUnit.Index).Value = Row.Cells(Me.ColPrimaryUnit.Index).Value Then
                            'jika qty max =0, maka itu harus baris terakhir baru baleh dihapuskan
                            e.Cancel = True
                        Else
                            If Me.ObjMP.ApakahAdaCustGroup(Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.txtPrimaryCurrency.Text, _
                            Me.TxtPrimaryCustGroup.Text, Row.Cells(Me.ColPrimaryUnit.Index).Value) = True Then
                                Me.ObjMP.HapusItemGroup(Row.Cells(Me.ColPrimary.Index).Value)
                            End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim DTACusGroup As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Dim TBLCusGroup As New DseMasterProduct_EditPriceList.tblmaster_product_pricelist_groupDataTable
        Dim Result As DialogResult
        Try
            DTACusGroup.Fill(TBLCusGroup)

            If TBLCusGroup.Rows.Count > 0 Then
                Result = MsgBox("Are you sure delete all data customer group ? ", MsgBoxStyle.YesNo, "Information")
                If Result = Windows.Forms.DialogResult.Yes Then
                    DTACusGroup.DeleteQueryAll()
                    Me.cboCustomerGroup.Enabled = True
                    Me.DGVHarga.Rows.Clear()
                    MsgBox("All data customer group has been deleted", MsgBoxStyle.Information, "Information")
                Else
                End If
            Else

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLCusGroup.Dispose()
            DTACusGroup.Dispose()
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim DTACust As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_customerTableAdapter
        Dim TBLCust As New DseMasterProduct_EditPriceList.tblmaster_product_pricelist_customerDataTable
        Dim Result As DialogResult
        Try
            DTACust.Fill(TBLCust)

            If TBLCust.Rows.Count > 0 Then
                Result = MsgBox("Are you sure delete all data customer ? ", MsgBoxStyle.YesNo, "Information")
                If Result = Windows.Forms.DialogResult.Yes Then
                    DTACust.DeleteQueryAll()
                    Me.DGVHarga.Rows.Clear()
                    MsgBox("All data customer has been deleted", MsgBoxStyle.Information, "Information")
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub



    Private Sub BtnPrintCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrintCust.Click
        Dim Frm As New FrmCetakMasterProduct_PriceList
        Dim DTACust As New DseMasterProduct_EditPriceListTableAdapters.TblCetak_CustomerTableAdapter
        Dim TBLCust As New DseMasterProduct_EditPriceList.TblCetak_CustomerDataTable

        Try
            DTACust.FillBy_PrimaryCurrencyAndCustomer(TBLCust, Me.txtPrimaryCurrency.Text, Me.txtprimaryCust.Text)
            If TBLCust.Rows.Count > 0 Then
                If Me.txtprimaryCust.Text = "-1" Then
                    MsgBox("Please Input Customer Code", MsgBoxStyle.Information, "Information")
                Else
                    Frm.TypeRbo = FrmCetakMasterProduct_PriceList.EnumGeneral.CustomerSatuan
                    Frm.PrimaryCurrency = Me.txtPrimaryCurrency.Text
                    Frm.PrimaryCustomer = Me.txtprimaryCust.Text
                    Frm.ShowDialog()
                End If
            Else
                MsgBox("No Data in this customer", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim Frm As New FrmCetakMasterProduct_PriceList

        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.ChkGeneralAll.Checked = True Then
                Frm.TypeRbo = FrmCetakMasterProduct_PriceList.EnumGeneral.GeneralAll
                Frm.PrimaryCurrency = Me.txtPrimaryCurrency.Text
            ElseIf Me.rboGeneral.Checked = True Then
                Frm.TypeRbo = FrmCetakMasterProduct_PriceList.EnumGeneral.general
                Frm.PrimaryCurrency = Me.txtPrimaryCurrency.Text
                Frm.ProductCode = Me.TxtKodeProduk.Text & "%"
            ElseIf Me.rboCustomer.Checked = True Then
                Frm.TypeRbo = FrmCetakMasterProduct_PriceList.EnumGeneral.Customer
            ElseIf rboCustomerGroup.Checked = True Then
                Me.CetakCustomerGroup()
                Exit Sub

            End If
            Frm.PrimaryCurrency = Me.txtPrimaryCurrency.Text
            Frm.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Public Sub CetakCustomerGroup()
        Dim dsCetak As New DseMasterProduct_EditPriceList
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.DataTable_Cetak_GroupTableAdapter
        Dim Tbl As New DseMasterProduct_EditPriceList.DataTable_Cetak_GroupDataTable
        Dim row As DseMasterProduct_EditPriceList.DataTable_Cetak_GroupRow
        Dim tblCetak As DseMasterProduct_EditPriceList.TblCetakGroupDataTable = dsCetak.TblCetakGroup
        Dim rowCetak As DseMasterProduct_EditPriceList.TblCetakGroupRow
        Dim repottt As New CrpLapMasterProduct_PriceList_Group
        Dim ij As New FrmCetakMasterProduct_PriceList

        Try
            DTA.FillBy_Currency(Tbl, Me.txtPrimaryCurrency.Text)
            If DTA.FillBy_Currency(Tbl, Me.txtPrimaryCurrency.Text) > 0 Then
                For Each row In Tbl
                    With row
                        rowCetak = tblCetak.NewTblCetakGroupRow
                        rowCetak.Product_Code = .Product_Code
                        rowCetak.Description_Product = .Description
                        rowCetak.Unit_Code = .Unit_Code
                        rowCetak.Max_Qty = .max_qty
                        rowCetak.Unit_Price = .unit_price
                        rowCetak.Discount = .discount
                        rowCetak.Remarks = .Remarks
                        rowCetak.Currency = .Currency_Code
                        rowCetak.DefPrec_Currency = .Default_Precision
                        rowCetak.DefPrec_Qty = ._Decimal

                        '......
                        If Me.cboCustomerGroup.SelectedIndex = 0 Then
                            rowCetak.Tipe_Group = "Customer Group"
                            rowCetak.Group_Code = .Customer_Group_Code
                            rowCetak.Group_Desc = .Description_Group
                        ElseIf Me.cboCustomerGroup.SelectedIndex = 1 Then
                            rowCetak.Tipe_Group = "Customer Sales Area"
                            rowCetak.Group_Code = .Sales_Area_Code
                            rowCetak.Group_Desc = .Description_Sales_Area
                        ElseIf Me.cboCustomerGroup.SelectedIndex = 2 Then
                            rowCetak.Tipe_Group = "Customer Segmentation"
                            rowCetak.Group_Code = .Segmentation_Code
                            rowCetak.Group_Desc = .Description_Segmentation
                        ElseIf Me.cboCustomerGroup.SelectedIndex = 3 Then
                            rowCetak.Tipe_Group = "Customer Type"
                            rowCetak.Group_Code = .Type_Code
                            rowCetak.Group_Desc = .Description_Tipe
                        End If
                        tblCetak.AddTblCetakGroupRow(rowCetak)
                    End With
                Next
                ''CETAK BROOOOH !!! :::
                repottt.SetDataSource(dsCetak)
                ij.CrvLaporanMasterPriceList.ReportSource = repottt
                ij.CrvLaporanMasterPriceList.Show()
                ij.mShowCustomerGroup = True
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BtnPrintCusGrouop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrintCusGrouop.Click
        Dim dsCetak As New DseMasterProduct_EditPriceList
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.DataTable_Cetak_GroupTableAdapter
        Dim Tbl As New DseMasterProduct_EditPriceList.DataTable_Cetak_GroupDataTable
        Dim row As DseMasterProduct_EditPriceList.DataTable_Cetak_GroupRow
        Dim tblCetak As DseMasterProduct_EditPriceList.TblCetakGroupDataTable = dsCetak.TblCetakGroup
        Dim rowCetak As DseMasterProduct_EditPriceList.TblCetakGroupRow
        Dim repottt As New CrpLapMasterProduct_PriceList_Group
        'Dim Frm As New FrmCetakMasterProduct_PriceList
        Dim ij As New FrmCetakMasterProduct_PriceList

        Try
            DTA.FillBy_PrimaryCurrency(Tbl, Me.txtPrimaryCurrency.Text, Me.TxtPrimaryCustGroup.Text)
            If DTA.FillBy_PrimaryCurrency(Tbl, Me.txtPrimaryCurrency.Text, Me.TxtPrimaryCustGroup.Text) > 0 Then
                For Each row In Tbl
                    With row
                        rowCetak = tblCetak.NewTblCetakGroupRow
                        rowCetak.Product_Code = .Product_Code
                        rowCetak.Description_Product = .Description
                        rowCetak.Unit_Code = .Unit_Code
                        rowCetak.Max_Qty = .max_qty
                        rowCetak.Unit_Price = .unit_price
                        rowCetak.Discount = .discount
                        rowCetak.Remarks = .Remarks
                        rowCetak.Currency = .Currency_Code
                        rowCetak.DefPrec_Currency = .Default_Precision
                        rowCetak.DefPrec_Qty = ._Decimal

                        '......
                        If Me.cboCustomerGroup.SelectedIndex = 0 Then
                            rowCetak.Tipe_Group = "Customer Group"
                            rowCetak.Group_Code = .Customer_Group_Code
                            rowCetak.Group_Desc = .Description_Group
                        ElseIf Me.cboCustomerGroup.SelectedIndex = 1 Then
                            rowCetak.Tipe_Group = "Customer Sales Area"
                            rowCetak.Group_Code = .Sales_Area_Code
                            rowCetak.Group_Desc = .Description_Sales_Area
                        ElseIf Me.cboCustomerGroup.SelectedIndex = 2 Then
                            rowCetak.Tipe_Group = "Customer Segmentation"
                            rowCetak.Group_Code = .Segmentation_Code
                            rowCetak.Group_Desc = .Description_Segmentation
                        ElseIf Me.cboCustomerGroup.SelectedIndex = 3 Then
                            rowCetak.Tipe_Group = "Customer Type"
                            rowCetak.Group_Code = .Type_Code
                            rowCetak.Group_Desc = .Description_Tipe
                        End If
                        tblCetak.AddTblCetakGroupRow(rowCetak)
                    End With
                Next
                ''CETAK BROOOOH !!! :::
                repottt.SetDataSource(dsCetak)
                ij.CrvLaporanMasterPriceList.ReportSource = repottt
                ij.CrvLaporanMasterPriceList.Show()
                ij.mShowCustomerGroup = True
                ij.ShowDialog()
            Else
                MsgBox("No Data in this group code", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub
End Class