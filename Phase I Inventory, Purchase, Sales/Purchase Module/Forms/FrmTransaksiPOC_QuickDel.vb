Option Explicit On
Option Strict Off

Public Class FrmTransaksiPOC_QuickDel
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public UserCode As String = ""

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim ObjSettingProd As New Core_Module.ClsTransaksiSetting_ProdGroup
    Dim DTA As New DseTransaksiPOC_QuickTableAdapters.DataTable1TableAdapter
    Dim TBL As New DseTransaksiPOC_Quick.DataTable1DataTable
    Dim Row As DseTransaksiPOC_Quick.DataTable1Row
    Dim DTAUnit As New Core_Module.dseMasterProductTableAdapters.tblmaster_product_item_multiunitTableAdapter
    Dim TBLUnit As New Core_Module.dseMasterProduct.tblmaster_product_item_multiunitDataTable
    Dim rowUnit As Core_Module.dseMasterProduct.tblmaster_product_item_multiunitRow
    Dim DTAUnit1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_product1TableAdapter
    Dim TBLUnit1 As New Core_Module.dseMasterProduct.tblmaster_product1DataTable
    Dim rowUnit1 As Core_Module.dseMasterProduct.tblmaster_product1Row
    Dim DTAUnit2 As New DseTransaksiPOCTableAdapters.tblmaster_unitTableAdapter
    Dim TBLUnit2 As New DseTransaksiPOC.tblmaster_unitDataTable
    Dim rowUnit2 As DseTransaksiPOC.tblmaster_unitRow

    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjPO As New ClsTransaksiPO
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objItem As New Inventory_Module.ClsTransaksiMaterialPlanningGroupItem
    Dim Buffer As String = ""
    Dim DelIndex As Integer = -1
    Dim ObjUser As New Core_Module.ClsUser
    Dim DTAUser As New DseTransaksiPOCTableAdapters.tbltrans_poc_setting_userTableAdapter
    Dim TBLUser As New DseTransaksiPOC.tbltrans_poc_setting_userDataTable
    Dim rowUser As DseTransaksiPOC.tbltrans_poc_setting_userRow
    Dim EraseRowIndex As Integer = -1
    Dim EraseRowIndex_Supp As Integer = -1
    Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1 '
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim DummyIndex As Integer = -1

    Private Sub FrmTransaksiPOC_QuickDel_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim rowDGV As DataGridViewRow

        Try
            'simpan posisi terakhir format supplier
            Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.POC_Quick, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.POC_Quick_SupplierFormat, Me.cboSupplierFormat.SelectedIndex, Me.UserCode)
            'simpan data ref user
            Me.dgvCreatedBy.AllowUserToAddRows = False
            If Me.dgvCreatedBy.Rows.Count > 0 Then
                Me.DTAUser.Delete(Me.UserCode)
                For Each rowDGV In Me.dgvCreatedBy.Rows
                    If IsNothing(rowDGV.Cells(Me.ColUserCode.Index).Value) = False Then
                        Me.DTAUser.Insert(Me.UserCode, rowDGV.Cells(Me.ColUserCode.Index).Value)
                    End If
                Next
            End If
            Me.dgvCreatedBy.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPOC_QuickDel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim DummyVal As String = ""
        Dim rowDGV As DataGridViewRow

        Try
            Me.ShowIcon = False
            Me.dgvItem.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCreatedBy.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvSupplier.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.cboSupplierFormat.SelectedIndex = 0

            DummyVal = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.POC_Quick, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.POC_Quick_SupplierFormat, Me.UserCode)
            If DummyVal <> "" Then
                Me.cboSupplierFormat.SelectedIndex = DummyVal
            End If

            Core_Module.SetConstring_DTAMasterProduct(Me.DTAUnit1)
            Core_Module.SetConstring_DTAMasterProduct_ItemMU(Me.DTAUnit)
            Me.DTAUser.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTAUnit2.Connection.ConnectionString = DTA.Connection.ConnectionString

            'load data user
            Me.DTAUser.FillBy_UserRef(Me.TBLUser, Me.UserCode)
            If Me.TBLUser.Rows.Count > 0 Then
                Me.dgvCreatedBy.AllowUserToAddRows = False
                For Each rowUser In Me.TBLUser.Rows
                    Me.dgvCreatedBy.Rows.Add()
                    rowDGV = Me.dgvCreatedBy.Rows(Me.dgvCreatedBy.Rows.Count - 1)
                    rowDGV.Cells(Me.ColUserCode.Index).Value = rowUser.user_code_created_by
                    rowDGV.Cells(Me.ColUserName.Index).Value = Me.ObjUser.GetNama(rowUser.user_code_created_by)
                Next
                Me.dgvCreatedBy.AllowUserToAddRows = True
            End If
            Me.ObjCD.SetNumber(Me.dgvCreatedBy)
            Me.dgvCreatedBy.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dtpFrom.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function SetDecimalProduct(ByVal vTBL As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewDataTable) As Double
        Dim Row, Row1 As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim DummyInd As Double = 0

        Try

            For Each Row In vTBL.Rows
                If Row.IsDecimal_MultiNull = True Then
                    Row.Decimal_Multi = 0

                    'filter table multi unit
                    Me.TBLUnit.DefaultView.RowFilter = "[Primary Product] = " + Row.Primary_Product.ToString + " AND [Primary Unit] = " + Row.Primary_Unit_Multi.ToString
                    If Me.TBLUnit.DefaultView.Count > 0 Then
                        Me.rowUnit = Me.TBLUnit.DefaultView(0).Row
                        Row.Decimal_Multi = Me.rowUnit.Default_Precision
                    Else
                        'kalo ga ada filter tabel product
                        Me.TBLUnit1.DefaultView.RowFilter = "Primary = " + Row.Primary_Product.ToString
                        If Me.TBLUnit1.DefaultView.Count > 0 Then
                            Me.rowUnit1 = Me.TBLUnit1.DefaultView(0).Row
                            Row.Decimal_Multi = Me.rowUnit1._Decimal
                        End If
                    End If

                    'set semua decimal untuk row dibawahnya, jika ditemukan kombinasi primary product dan primary unit yang sama
                    If DummyInd + 1 <= vTBL.Rows.Count - 1 Then
                        For ij As Double = DummyInd + 1 To vTBL.Rows.Count - 1
                            Row1 = vTBL.Rows(ij)
                            If Row1.IsDecimal_MultiNull = True Then
                                If Row1.Primary_Product = Row.Primary_Product And Row1.Primary_Unit_Multi = Row.Primary_Unit_Multi Then
                                    Row1.Decimal_Multi = Row.Decimal_Multi
                                End If
                            End If
                        Next
                    End If
                End If

                'set ke baris selanjutnya
                DummyInd += 1
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetUnitCode(ByVal vPrimaryUnit As Double) As String

        Try
            Me.TBLUnit2.DefaultView.RowFilter = "Primary = " + vPrimaryUnit.ToString
            If Me.TBLUnit2.DefaultView.Count = 0 Then
                Return ""
            Else
                Me.rowUnit2 = Me.TBLUnit2.DefaultView(0).Row
                Return Me.rowUnit2.Unit_Code
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim UserCodeList As New List(Of String)
        Dim rowDGV As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.dgvCreatedBy.AllowUserToAddRows = False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'kosongkan dgv item
            Me.dgvItem.Rows.Clear()
            Me.TBL.Rows.Clear()
            Me.TBLUnit2.Rows.Clear()

            'load data-filter supplier
            If Me.TxtPrimaryPO.Text = "" Or Me.TxtPrimaryPO.Text = "-1" Then
                If Me.ChkIgnoreLowTolerance.Checked = True Then
                    DTA.FillBy_All_UpTolerance(TBL, Me.dtpFrom.Value, Me.dtpTo.Value)
                Else
                    DTA.FillBy_All(TBL, Me.dtpFrom.Value, Me.dtpTo.Value)
                End If
            Else
                If Me.ChkIgnoreLowTolerance.Checked = True Then
                    DTA.FillBy_PrimaryPO_UpTolerance(TBL, Me.dtpFrom.Value, Me.dtpTo.Value, Me.TxtPrimaryPO.Text)
                Else
                    DTA.FillBy_PrimaryPO(TBL, Me.dtpFrom.Value, Me.dtpTo.Value, Me.TxtPrimaryPO.Text)
                End If
            End If
            Me.TBLUnit.Rows.Clear()
            Me.TBLUnit1.Rows.Clear()
            Me.DTAUnit.Fill(TBLUnit)
            Me.DTAUnit1.Fill(TBLUnit1)
            Me.DTAUnit2.Fill(Me.TBLUnit2)

            'consider
            '1. supplier
            'If Me.TxtPrimarySupplier.Text = "-1" Or Me.TxtPrimarySupplier.Text = "" Then
            'Else
            '    For ij As Double = TBL.Rows.Count - 1 To 0 Step -1
            '        Row = TBL.Rows(ij)
            '        If Row.Primary_Supplier <> Me.TxtPrimarySupplier.Text Then
            '            TBL.Rows.Remove(Row)
            '        End If
            '    Next
            'End If
            ''2. code no
            'If Me.TxtPrimaryProduct.Text = "-1" Or Me.TxtPrimaryProduct.Text = "" Then
            'Else
            '    For ij As Double = TBL.Rows.Count - 1 To 0 Step -1
            '        Row = TBL.Rows(ij)
            '        If Row.Primary_Product <> Me.TxtPrimaryProduct.Text Then
            '            TBL.Rows.Remove(Row)
            '        End If
            '    Next
            'End If
            ''3. created by
            'If Me.dgvCreatedBy.Rows.Count > 0 Then
            '    UserCodeList.Clear()
            '    For Each rowDGV In Me.dgvCreatedBy.Rows
            '        UserCodeList.Add(rowDGV.Cells(Me.ColUserCode.Index).Value)
            '    Next
            '    For ij As Double = TBL.Rows.Count - 1 To 0 Step -1
            '        Row = TBL.Rows(ij)
            '        If UserCodeList.Contains(Row.Created_By) = False Then
            '            TBL.Rows.Remove(Row)
            '        End If
            '    Next
            'End If

            Me.GenerateFiltering()

            'set decimal multi
            'Me.SetDecimalProduct(TBL)

            'pindahkan ke dgv
            For Each Row In TBL.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Row.Amount, Row.Default_Precision)
                rowDGV.Cells(Me.ColCancel.Index).Value = Boolean.FalseString
                rowDGV.Cells(Me.ColCodeNo.Index).Value = Row.Product_Code
                rowDGV.Cells(Me.ColCurrency.Index).Value = Row.Currency_Code
                rowDGV.Cells(Me.ColDate.Index).Value = Row.Po_Date
                rowDGV.Cells(Me.ColDesc.Index).Value = Row.Product_Description
                rowDGV.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Row.Disc_Amount, Row.Default_Precision)
                rowDGV.Cells(Me.ColDiscPerct.Index).Value = Me.ObjCD.DigitGrouping(Row.Discount, 2)
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                rowDGV.Cells(Me.ColPONo.Index).Value = Row.PO_No
                rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(Row.POB_Multi, Row.Decimal_Multi)
                rowDGV.Cells(Me.ColDecimalMulti.Index).Value = Row.Decimal_Multi
                rowDGV.Cells(Me.ColQtyCancel.Index).Value = 0
                If Me.cboSupplierFormat.SelectedIndex = 0 Then
                    rowDGV.Cells(Me.ColSupplier.Index).Value = Row.Supplier_Code
                ElseIf Me.cboSupplierFormat.SelectedIndex = 1 Then
                    rowDGV.Cells(Me.ColSupplier.Index).Value = Row.Supplier_Name
                Else
                    rowDGV.Cells(Me.ColSupplier.Index).Value = Row.Supplier_Code + " - " + Row.Supplier_Name
                End If

                rowDGV.Cells(Me.ColUnit.Index).Value = Row.Unit_Code
                rowDGV.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(Row.Unit_Price, Row.Default_Precision)
                rowDGV.Cells(Me.ColPrimaryPoduct.Index).Value = Row.Primary_Product
                rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value = Row.Primary_Unit_Multi
                rowDGV.Cells(Me.ColPrimaryPO.Index).Value = Row.Primary_PO
                rowDGV.Cells(Me.ColRemarksItem.Index).Value = ""
                rowDGV.Cells(Me.ColPrimaruUnitDefault.Index).Value = Row.Primary_Unit_Default
                rowDGV.Cells(Me.ColQtyDefault.Index).Value = Me.ObjCD.DigitGrouping(Row.POB_Multi * Row.Ratio_Default / Row.Ratio_Reference, Row.Decimal_Product)
                rowDGV.Cells(Me.ColRatioDefault.Index).Value = Row.Ratio_Default
                rowDGV.Cells(Me.ColRatioReference.Index).Value = Row.Ratio_Reference
                rowDGV.Cells(Me.ColDefPrec_QtyDefault.Index).Value = Row.Decimal_Product
                rowDGV.Cells(Me.ColUnitDefault.Index).Value = Me.GetUnitCode(Row.Primary_Unit_Default)
            Next

            Me.dgvCreatedBy.AllowUserToAddRows = True
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            UserCodeList = Nothing
        End Try
    End Sub

    Private Sub btnFindPO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPO.Click
        Dim i As New FrmCariTransaksiPO
        i.BackColorAll = Me.BackColorAll

        Try
            'i.cboCriteria.SelectedIndex = 0
            If Me.TxtPONumber.Text <> "" Then
                i.txtCari.Text = Me.TxtPONumber.Text
            End If

            i.ShowDialog()
            If i.PrimaryPO > 0 Then
                Me.TxtPrimaryPO.Text = i.PrimaryPO
                Me.TxtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
            Else
                Me.TxtPrimaryPO.Text = "-1"
                Me.TxtPONumber.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtPONumber_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPONumber.KeyDown
        Try
            If e.KeyValue = 13 Then
                Dim Primary As Double = ObjPO.GetPrimary(Me.TxtPONumber.Text)
                If Primary <> -1 Then
                    Me.TxtPrimaryPO.Text = Primary
                    Me.TxtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                Else
                    Me.btnFindPO.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPONumber_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPONumber.Leave
        Try
            If Me.TxtPONumber.Text <> "" Then
                Dim Primary As Double = ObjPO.GetPrimary(Me.TxtPONumber.Text)
                If Primary <> -1 Then
                    Me.TxtPrimaryPO.Text = Primary
                    Me.TxtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                Else
                    Me.btnFindPO.PerformClick()
                End If
            Else
                Me.TxtPrimaryPO.Text = "-1"
                Me.TxtPONumber.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPONumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPONumber.GotFocus
        Try
            Me.TxtPONumber.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPONumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPONumber.LostFocus
        Try
            Me.TxtPONumber.BackColor = Me.TxtPrimaryPO.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCancellationReason_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCancellationReason.GotFocus
        Try
            Me.TxtCancellationReason.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCancellationReason_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCancellationReason.LostFocus
        Try
            Me.TxtCancellationReason.BackColor = Me.TxtPrimaryPO.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim AdaYgDicancel As Boolean = False
        Dim rowDGV As DataGridViewRow
        Dim ListPrimaryPO As New List(Of Double)
        Dim PrimaryPO As Double = -1
        Dim ij As New FrmTransaksiPOC
        Dim SupName, CPName, PhoneNo As String

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColCancel.Index).Value = Boolean.TrueString Then
                    AdaYgDicancel = True
                    Exit For
                End If
            Next
            If AdaYgDicancel = True Then
                'cancellation reason harus ada nilainya
                Me.TxtCancellationReason.Text = Me.TxtCancellationReason.Text.Trim
                If Me.TxtCancellationReason.Text = "" Then
                    MsgBox("Please specify cancellation reason.", MsgBoxStyle.Information, "Information")
                    Me.TxtCancellationReason.Focus()
                    Exit Sub
                End If

                If MsgBox("Start cancellation process?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    '1. isikan list po
                    For Each rowDGV In Me.dgvItem.Rows
                        If rowDGV.Cells(Me.ColCancel.Index).Value = Boolean.TrueString Then
                            If ListPrimaryPO.Contains(rowDGV.Cells(Me.ColPrimaryPO.Index).Value) = False Then
                                ListPrimaryPO.Add(rowDGV.Cells(Me.ColPrimaryPO.Index).Value)
                            End If
                        End If
                    Next

                    Me.prg01.Visible = True
                    Me.prg01.Value = 0
                    Me.prg01.Maximum = ListPrimaryPO.Count
                    ij.prgProgress.Value = 0
                    ij.prgProgress.Maximum = Me.prg01.Maximum
                    Me.LblProgress.Text = ""
                    Me.LblProgress.Visible = True

                    '2. untuk setiap list po, buat poc, lengkapi item nya lalu simpan
                    ij.vFromQuick = True
                    ij.vDirectlySave = False
                    ij.UserCode = Me.UserCode
                    ij.Show()
                    ij.Location = New Point((ij.Width * -1) - 5, (ij.Height * -1) - 5)
                    For Each PrimaryPO In ListPrimaryPO
                        ij.SetBaru()
                        ij.txtPrimaryPO.Text = PrimaryPO
                        ij.txtPONO.Text = Me.ObjPO.GetNoTransaksiPO(PrimaryPO)
                        ij.vShowDialog = False

                        Me.prg01.Value += 1
                        Me.LblProgress.Text = "Data " + Me.prg01.Value.ToString + " of " + Me.prg01.Maximum.ToString + " - PO Number: " + ij.txtPONO.Text
                        ij.prgProgress.Value = Me.prg01.Value
                        ij.lblProgress.Visible = True
                        ij.lblProgress.Text = Me.LblProgress.Text

                        ij.txtPrimarySupplier.Text = ObjPO.GetPrimarySupplier(PrimaryPO)
                        ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(ij.txtPrimarySupplier.Text, SupName, CPName, PhoneNo)
                        ij.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(ij.txtPrimarySupplier.Text)
                        ij.TxtSupplierName.Text = SupName
                        ij.txtContactPerson.Text = CPName
                        ij.txtPhoneNo.Text = PhoneNo

                        ij.txtCancellationReason.Text = Me.TxtCancellationReason.Text.Trim
                        ij.TxtRemark.Text = "Quick PO Cancellation"

                        For Each rowDGV1 As DataGridViewRow In Me.dgvItem.Rows
                            If rowDGV1.Cells(Me.ColPrimaryPO.Index).Value = PrimaryPO Then
                                ij.dgvItem.Rows.Add()
                                rowDGV = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)
                                With rowDGV
                                    .Cells(ij.ColNo.Index).Value = ij.dgvItem.Rows.Count
                                    .Cells(ij.ColCodeNo.Index).Value = rowDGV1.Cells(Me.ColCodeNo.Index).Value
                                    .Cells(ij.ColDescription.Index).Value = rowDGV1.Cells(Me.ColDesc.Index).Value
                                    .Cells(ij.ColQtyMaster.Index).Value = 0
                                    .Cells(ij.ColQtyPerMaster.Index).Value = 0
                                    .Cells(ij.ColMasterUnit.Index).Value = ""
                                    .Cells(ij.ColQty.Index).Value = 0

                                    .Cells(ij.ColTotalQty.Index).Value = rowDGV1.Cells(Me.ColQtyDefault.Index).Value
                                    .Cells(ij.ColUnit.Index).Value = rowDGV1.Cells(Me.ColUnitDefault.Index).Value
                                    .Cells(ij.ColRemark.Index).Value = rowDGV1.Cells(Me.ColRemarksItem.Index).Value
                                    .Cells(ij.ColPrimaryCodeNo.Index).Value = rowDGV1.Cells(Me.ColPrimaryPoduct.Index).Value
                                    .Cells(ij.ColPrimaryMasterUnit.Index).Value = -1
                                    .Cells(ij.ColPrimaryUnits.Index).Value = rowDGV1.Cells(Me.ColPrimaruUnitDefault.Index).Value
                                    .Cells(ij.ColQtyMulti.Index).Value = rowDGV1.Cells(Me.ColQtyCancel.Index).Value

                                    .Cells(ij.ColUnitCodeMulti.Index).Value = rowDGV1.Cells(Me.ColUnit.Index).Value
                                    If .Cells(ij.ColUnitCodeMulti.Index).Value <> .Cells(ij.ColUnit.Index).Value Then
                                        .Cells(ij.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(rowDGV1.Cells(Me.ColRatioReference.Index).Value, _
                                                Me.GetDefPrec(rowDGV1.Cells(Me.ColPrimaryPoduct.Index).Value, rowDGV1.Cells(Me.ColPrimaryUnitMulti.Index).Value)) _
                                                & " " & rowDGV1.Cells(Me.ColUnit.Index).Value & " = " & Me.ObjCD.DigitGrouping(rowDGV1.Cells(Me.ColRatioDefault.Index).Value, _
                                                rowDGV1.Cells(Me.ColDefPrec_QtyDefault.Index).Value) & " " & rowDGV1.Cells(Me.ColUnitDefault.Index).Value
                                    Else
                                        .Cells(ij.ColRatioDesc.Index).Value = "-"
                                    End If
                                    .Cells(ij.ColRatioReference.Index).Value = rowDGV1.Cells(Me.ColRatioReference.Index).Value
                                    .Cells(ij.ColRatioDefault.Index).Value = rowDGV1.Cells(Me.ColRatioDefault.Index).Value
                                    .Cells(ij.ColPrimaryUnitMulti.Index).Value = rowDGV1.Cells(Me.ColPrimaryUnitMulti.Index).Value
                                End With
                            End If
                        Next
                        ij.Hide()
                        ij.vDirectlySave = Me.chkDirectSave.Checked
                        ij.vShowDialog = True
                        ij.ShowDialog()
                        If ij.vStopCancellation = True Then
                            Exit For
                        End If
                    Next
                    '3. 
                End If

                Me.btnFilter.PerformClick()
            Else
                MsgBox("Please specify at least 1 item to cancel.", MsgBoxStyle.Information, "Information")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryPO = Nothing
            ij = Nothing
            Me.prg01.Value = 0
            Me.prg01.Visible = False
            Me.LblProgress.Visible = False
        End Try
    End Sub

    Private Sub dgvCreatedBy_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCreatedBy.CellBeginEdit
        If e.ColumnIndex = Me.ColUserCode.Index Then
            Me.Buffer = Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value
        Else
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub dgvCreatedBy_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCreatedBy.CellEndEdit
        Dim KodeUser As String = ""
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim RowIdx As Integer = e.RowIndex
        Try
            If IsNothing(Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColUserCode.Index Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    If e.RowIndex <> Me.dgvCreatedBy.RowCount - 1 Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                    Exit Sub
                End If
            Else
                ''data account number
                If e.ColumnIndex = Me.ColUserCode.Index Then
                    Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjUser.ApakahSudahAda(Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                        KodeUser = Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value
                        index = Me.GetIndexItem(KodeUser, e.RowIndex)
                        If index < 0 Then
                            Me.dgvCreatedBy.Item(Me.ColUserCode.Index, e.RowIndex).Value = Me.dgvCreatedBy.Item(Me.ColUserCode.Index, e.RowIndex).Value
                            Me.dgvCreatedBy.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.dgvCreatedBy.Item(Me.ColUserCode.Index, e.RowIndex).Value)
                        Else
                            If e.RowIndex <> index Then
                                MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                Me.dgvCreatedBy.Item(Me.ColUserCode.Index, e.RowIndex).Value = ""
                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                                Exit Sub
                            End If
                        End If
                    Else
                        '2. kode akun yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmCariMasterUser
                        ij.backColorAll = Me.BackColorAll
                        ij.TxtCari.Text = Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.Kodeuser <> "" Then
                            index = Me.GetIndexItem(ij.Kodeuser, e.RowIndex)
                            If index < 0 Then
                                Me.dgvCreatedBy.Item(Me.ColUserCode.Index, e.RowIndex).Value = ij.Kodeuser
                                Me.dgvCreatedBy.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Exit Sub
                                Else
                                    Me.dgvCreatedBy.Item(Me.ColUserCode.Index, e.RowIndex).Value = ij.Kodeuser
                                    Me.dgvCreatedBy.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                End If
                            End If
                            ''2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                            'Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = ij.Kodeuser.ToUpper
                            'Me.dgvCreatedBy.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value)
                        ElseIf ij.Kodeuser = "" Then
                            If e.ColumnIndex = Me.ColUserCode.Index Then
                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                    If Me.dgvCreatedBy.Rows.Count - 1 <> e.RowIndex Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                Else
                                    Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                            Else
                                Me.dgvCreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                        End If
                        ij = Nothing
                    End If
                End If
            End If
            Me.ObjCD.SetNumber(Me.dgvCreatedBy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vKodeUser As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvCreatedBy.Rows.Count - 2
                RowDGV = Me.dgvCreatedBy.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColUserCode.Index).Value = vKodeUser Then
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

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(200)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.EraseRowIndex <> -1 Then
                Me.dgvCreatedBy.Rows.RemoveAt(Me.EraseRowIndex)
                Me.EraseRowIndex = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            If MsgBox("Check all data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                For Each rowDGV In Me.dgvItem.Rows
                    rowDGV.Cells(Me.ColCancel.Index).Value = Boolean.TrueString
                    rowDGV.Cells(Me.ColQtyCancel.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQty.Index).Value, rowDGV.Cells(Me.ColDecimalMulti.Index).Value)
                Next
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            If MsgBox("Uncheck all data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                For Each rowDGV In Me.dgvItem.Rows
                    rowDGV.Cells(Me.ColCancel.Index).Value = Boolean.FalseString
                    rowDGV.Cells(Me.ColQtyCancel.Index).Value = 0
                Next
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetDefPrec(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double) As Double
        Try
            Me.TBL.DefaultView.RowFilter = "[Primary Product] = " + vPrimaryProduct.ToString + " AND [Primary Unit Multi] = " + vPrimaryUnit.ToString
            If Me.TBL.DefaultView.Count > 0 Then
                Me.Row = Me.TBL.DefaultView(0).Row
                Return Me.Row.Decimal_Multi
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yg bisa diedit:
            '1. chk cancel
            '2. qty cancel+remarks item (jika chk cancel di check)
            If e.ColumnIndex = Me.ColCancel.Index Then
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            ElseIf e.ColumnIndex = Me.ColQtyCancel.Index Or e.ColumnIndex = Me.ColRemarksItem.Index Then
                If Me.dgvItem.Item(Me.ColCancel.Index, e.RowIndex).Value = Boolean.TrueString Then
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColQtyCancel.Index Then
                    'kembalikan ke nilai buffer
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                ElseIf e.ColumnIndex = Me.ColRemarksItem.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Exit Sub
                End If
            Else
                If e.ColumnIndex = Me.ColCancel.Index Then
                    'kemungkinan
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString Then
                        'di cek
                        Me.dgvItem.Item(Me.ColQtyCancel.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(Me.ColQtyCancel.Index, e.RowIndex).Value) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString
                        End If
                    Else
                        'di un check
                        Me.dgvItem.Item(Me.ColQtyCancel.Index, e.RowIndex).Value = 0
                        Me.dgvItem.Item(Me.ColRemarksItem.Index, e.RowIndex).Value = ""
                    End If
                ElseIf e.ColumnIndex = Me.ColQtyCancel.Index Then
                    'pastikan non negatif
                    If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'pastikan nilainya lebih kecil dari qty
                    If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) Then
                        MsgBox("Maximum value your can enter is: " + Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value + " " + Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value, _
                        MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'digit grouping
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColDecimalMulti.Index, e.RowIndex).Value)
                    'hitung ulang qty default
                    Me.dgvItem.Item(Me.ColQtyDefault.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value * _
                    Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value / Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColDefPrec_QtyDefault.Index, e.RowIndex).Value)
                ElseIf e.ColumnIndex = Me.ColRemarksItem.Index Then
                    'trim
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    'truncate jika lebih panjang dr 150
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 150 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 150)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.ObjCD.setnumber(Me.dgvItem)
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim vPrimary As Double = -1

        Try
            If e.ColumnIndex = -1 Or e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColPONo.Index Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPO") = False Then
                    MsgBox("Insufficient acces.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                Else
                    vPrimary = Me.dgvItem.Item(Me.ColPrimaryPO.Index, e.RowIndex).Value
                    Dim ij As Purchase_Module.FrmTransaksiPO
                    Dim j As Windows.Forms.Form

                    For Each j In Me.ParentForm.MdiChildren
                        ij = TryCast(j, Purchase_Module.FrmTransaksiPO)
                        If IsNothing(ij) = False Then
                            ij.BringToFront()
                            ij.Focus()
                            ij.ShowTransaction(vPrimary)
                            Exit Sub
                        End If
                    Next

                    Dim i As New Purchase_Module.FrmTransaksiPO
                    ' Increment the caption counter.
                    i.UserCode = Me.UserCode.ToString
                    ' Set the caption to be unique.
                    i.BAckColorAll = Core_Module.GetSMESSBackColorAll
                    i.MdiParent = Me.ParentForm
                    i.Show()
                    i.ShowTransaction(vPrimary)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#Region "Datagridview Filtering"
    Private Sub DgvSupplier_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvSupplier.CellBeginEdit
        Try
            If Me.DgvSupplier.Item(Me.ColSupp_Primary.Index, e.RowIndex).Value <> -1 Then
                If e.ColumnIndex = Me.ColSupp_Code.Index Then
                    Me.Buffer = Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColSupp_Code.Index Then
                    Me.Buffer = Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSupplier_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvSupplier.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvSupplier.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim Supplier As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColSupp_Code.Index Then
                        Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjSupplier.ApakahSudahAda(Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Supplier = Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value
                            index = Me.GetIndexItem_Supp(Supplier, e.RowIndex)
                            If index < 0 Then
                                Me.DgvSupplier.Item(Me.ColSupp_Primary.Index, e.RowIndex).Value = Me.ObjSupplier.GetPrimary(Me.DgvSupplier.Item(Me.ColSupp_Code.Index, e.RowIndex).Value)
                                Me.DgvSupplier.Item(Me.ColSupp_Code.Index, e.RowIndex).Value = Me.DgvSupplier.Item(Me.ColSupp_Code.Index, e.RowIndex).Value
                                Me.DgvSupplier.Item(Me.ColSupp_Name.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(Me.DgvSupplier.Item(Me.ColSupp_Primary.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvSupplier.Item(Me.ColSupp_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex_Supp = e.RowIndex
                                        Me.BgwSupplier.RunWorkerAsync()
                                    Else
                                        Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterSupplier
                            ij.txtCari.Text = Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimarySupplier <> -1 Then
                                index = Me.GetIndexItem_Supp(ij.PrimarySupplier, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvSupplier.Item(Me.ColSupp_Code.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierCode(ij.PrimarySupplier)
                                    Me.DgvSupplier.Item(Me.ColSupp_Name.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(ij.PrimarySupplier)
                                    Me.DgvSupplier.Item(Me.ColSupp_Primary.Index, e.RowIndex).Value = ij.PrimarySupplier
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_Supp = e.RowIndex
                                            Me.BgwSupplier.RunWorkerAsync()
                                        Else
                                            Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Exit Sub
                                    Else
                                        Me.DgvSupplier.Item(Me.ColSupp_Code.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierCode(ij.PrimarySupplier)
                                        Me.DgvSupplier.Item(Me.ColSupp_Name.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(ij.PrimarySupplier)
                                        Me.DgvSupplier.Item(Me.ColSupp_Primary.Index, e.RowIndex).Value = ij.PrimarySupplier
                                    End If
                                End If
                            ElseIf ij.PrimarySupplier = -1 Then
                                If e.ColumnIndex = Me.ColSupp_Code.Index Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        If Me.DgvSupplier.Rows.Count - 1 <> e.RowIndex Then
                                            Me.EraseRowIndex_Supp = e.RowIndex
                                            Me.BgwSupplier.RunWorkerAsync()
                                        End If
                                    Else
                                        Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                Else
                                    Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvSupplier)
                    End If
                Else
                    Me.EraseRowIndex_Supp = e.RowIndex
                    Me.BgwSupplier.RunWorkerAsync()
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_Supp(ByVal vKodeSupplier As String, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvSupplier.Rows.Count - 2
                RowDGV = Me.DgvSupplier.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColSupp_Code.Index).Value = vKodeSupplier Then
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

    Private Sub BgwSupplier_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwSupplier.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwSupplier_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwSupplier.RunWorkerCompleted
        Try
            Me.DgvSupplier.Rows.RemoveAt(Me.EraseRowIndex_Supp)
            Me.ObjCD.SetNumber(Me.DgvSupplier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSupplier_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvSupplier.DataError

    End Sub

    Private Sub DgvSupplier_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvSupplier.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            rowDGV = Me.DgvSupplier.Rows(Me.DgvSupplier.RowCount - 1)
            With rowDGV
                .Cells(Me.ColSupp_No.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColSupp_Code.Index).Value = ""
                .Cells(Me.ColSupp_Name.Index).Value = ""
                .Cells(Me.ColSupp_Primary.Index).Value = "-1"
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSupplier_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvSupplier.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.DgvSupplier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSupplier_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvSupplier.Sorted
        Me.ObjCD.SetNumber(Me.DgvSupplier)
    End Sub

    Private Sub DgvSupplier_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvSupplier.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvSupplier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    ''
    Private Sub dgvItemProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemProductGroup.CellBeginEdit
        Dim rowDGV As DataGridViewRow

        Try
            'yang bisa diedit:
            '1. jika baris baru
            If e.RowIndex = Me.dgvItemProductGroup.Rows.Count - 1 Then
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'pastikan baris diatasnya sudah beres dahulu
                    If e.RowIndex > 0 Then
                        rowDGV = Me.dgvItemProductGroup.Rows(e.RowIndex - 1)
                        If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value.ToString.Trim = "" Then
                            MsgBox("Please specify " + rowDGV.Cells(Me.ColProductGroup_Level.Index).Value.ToString + _
                            " code on the previous row.", MsgBoxStyle.Information, "Information")
                            Me.dgvItemProductGroup.ClearSelection()
                            rowDGV.Selected = True
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = -1 Then
                    '2. jika bukan baris baru dan primary =-1
                    'yang bisa diedit hanya code product group dan level
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    '3. jika bukan baris baru dan primary <> -1
                    'yang bisa diedit: code, remarks
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or _
                    e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProductGroup.CellEndEdit
        Dim Level As Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand

        Try
            If IsNothing(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'Remarks
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    'kalo kode product group dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                    Exit Sub
                End If
            Else
                'level
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'do nothing

                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If

                    'cek ke master
                    Level = Me.objItem.GetProductLevel(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
                    If Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                        Me.EndEditProductGroup(e, Me.ProductGroup_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                        Me.EndEditProductGroup(e, Me.Brand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                        Me.EndEditProductGroup(e, Me.SubBrand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                        Me.EndEditCategory(e)
                    Else
                        Me.EndEditProduct(e)
                    End If
                ElseIf e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'remarks
                    'pastikan ditrim dan  max 250 karakter
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'tidak boleh ada karakter aneh2
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                        MsgBox("Please avoid using the following character: ', *, %, #.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex_ItemProductGroup(ByVal vCode As String, ByVal vRefIndex As Integer, ByVal vProductGroupLevelStr As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = vCode And _
                rowDGV.Index <> vRefIndex And _
                rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = vProductGroupLevelStr Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub EndEditProductGroup(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal vProductGroup_Level As Integer)
        Try
            If Me.ObjProductGroup.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level) = True Then
                '1. kode product group ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product group yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.FrmCariMasterProductGroup
                ij.ProductGroupLevel = vProductGroup_Level
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProductGroup <> -1 Then
                    '2.a. product group yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProductGroup.GetGroupCode(ij.PrimaryProductGroup)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(ij.PrimaryProductGroup)
                ElseIf ij.PrimaryProductGroup = -1 Then
                    '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Trim(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If
            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.Buffer, vProductGroup_Level)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditCategory(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjCategory.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode category ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode category yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterCategory
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryCategory <> -1 Then
                    '2.a. category yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryCategory
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                ElseIf ij.PrimaryCategory = -1 Then
                    '2.b. kode category yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.Buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditProduct(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjProduct.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode product ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterProduct
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProduct <> -1 Then
                    '2.a. product yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProduct
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                ElseIf ij.PrimaryProduct = -1 Then
                    '2.b. kode product yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.Buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemProductGroup.DataError

    End Sub

    Private Sub dgvItemProductGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemProductGroup.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemProductGroup.Rows(Me.dgvItemProductGroup.Rows.Count - 1)
            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = Me.objItem.GetProductLevelStr(Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup)
            rowDGV.Cells(Me.ColProductGroup_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value = -1
            rowDGV.Cells(Me.ColProductGroup_Remarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemProductGroup.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemProductGroup.Sorted
        Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
    End Sub

    Private Sub dgvItemProductGroup_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemProductGroup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
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
            If Me.DummyIndex <> -1 Then
                Me.dgvItemProductGroup.Rows.RemoveAt(Me.DummyIndex)
                DummyIndex = -1
                Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    Private Sub GenerateFiltering()
        Dim FilterProduct As String = ""
        Dim FilterSupplier As String = ""
        Dim FilterGroup As String = ""
        Dim FilterUser As String = ""
        Dim rowDGV As New DataGridViewRow
        Dim Filter As String = ""
        Dim row As DseTransaksiPOC_Quick.DataTable1Row
        Dim TBLCopyan As New DseTransaksiPOC_Quick.DataTable1DataTable
        Dim rowCopyan As DseTransaksiPOC_Quick.DataTable1Row
     
        Try
            If Me.dgvCreatedBy.Rows.Count > 0 Then
                For Each rowDGV In Me.dgvCreatedBy.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColUserCode.Index).Value) = False Then
                        If FilterUser = "" Then
                            FilterUser = "[Created By] = '" & rowDGV.Cells(Me.ColUserCode.Index).Value & "'"
                        Else
                            FilterUser &= " Or [Created By] = '" & rowDGV.Cells(Me.ColUserCode.Index).Value & "'"
                        End If
                    End If
                Next
            End If

            If Me.DgvSupplier.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvSupplier.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColSupp_Code.Index).Value) = False Then
                        If FilterSupplier = "" Then
                            FilterSupplier = "[Primary Supplier] = " & rowDGV.Cells(Me.ColSupp_Primary.Index).Value
                        Else
                            FilterSupplier &= " Or [Primary Supplier] = " & rowDGV.Cells(Me.ColSupp_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
                    TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)
                    If FilterGroup = "" Then
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            FilterGroup &= " [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            FilterGroup &= " [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            FilterGroup &= " [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        Else
                            FilterGroup &= " [Primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        End If
                    Else
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " OR [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            FilterGroup &= " OR [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            FilterGroup &= " OR [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            FilterGroup &= " OR [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        Else
                            FilterGroup &= " OR [Primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        End If
                    End If
                End If
            Next

            If FilterUser <> "" Then
                Filter = FilterUser
            End If
            
            If FilterSupplier <> "" Then
                If Filter = "" Then
                    Filter = FilterSupplier
                Else
                    Filter = "(" & Filter & ") AND (" & FilterSupplier & ") "
                End If
            End If

            If FilterGroup <> "" Then
                If Filter = "" Then
                    Filter = FilterGroup
                Else
                    Filter = "(" & Filter & ") AND (" & FilterGroup & ") "
                End If
            End If

            If Filter <> "" Then
                Me.TBL.DefaultView.RowFilter = Filter
                For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                    rowCopyan = Me.TBL.DefaultView(i).Row
                    TBLCopyan.ImportRow(rowCopyan)
                Next

                Me.TBL.Rows.Clear()
                For Each row In TBLCopyan.Rows
                    TBL.ImportRow(row)
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCreatedBy_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvCreatedBy.DataError

    End Sub

    Private Sub dgvCreatedBy_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvCreatedBy.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            rowDGV = Me.dgvCreatedBy.Rows(Me.dgvCreatedBy.RowCount - 1)
            With rowDGV
                .Cells(Me.ColUserNo.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColUserCode.Index).Value = ""
                .Cells(Me.ColUserName.Index).Value = ""
                .Cells(Me.ColUser_Primary.Index).Value = "-1"
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCreatedBy_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvCreatedBy.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvCreatedBy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCreatedBy_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCreatedBy.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvCreatedBy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbClearFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbClearFilter.Click
        Try
            Me.dgvCreatedBy.Rows.Clear()
            Me.DgvSupplier.Rows.Clear()
            Me.dgvItemProductGroup.Rows.Clear()
            Me.TxtPrimaryPO.Text = "-1"
            Me.TxtPONumber.Clear()
            Me.cboSupplierFormat.SelectedIndex = 0
            Me.dtpFrom.Value = Now
            Me.dtpTo.Value = Now
            Me.chkDirectSave.Checked = False
            Me.ChkIgnoreLowTolerance.Checked = False
            Me.btnFilter.Focus()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class